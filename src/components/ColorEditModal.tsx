import React, { useState, useEffect, useRef, useCallback } from 'react';
import { X, Plus } from 'lucide-react';

interface ColorEditModalProps {
  isOpen: boolean;
  initialColor: string;
  onSelectColor: (hex: string) => void;
  onClose: () => void;
}

// 윈도우 그림판 기본 색상 40개 (4행 x 10열)
const BASIC_COLORS: string[] = [
  // 1행 (파스텔/기본 레드~블루)
  '#f87171', '#ef4444', '#b91c1c', '#7f1d1d', '#450a0a', '#67e8f9', '#06b6d4', '#0284c7', '#1d4ed8', '#1e3a8a',
  // 2행 (옐로우, 오렌지, 브라운, 퍼플)
  '#fef08a', '#eab308', '#fb923c', '#ea580c', '#854d0e', '#65a30d', '#a855f7', '#7c3aed', '#4338ca', '#312e81',
  // 3행 (그린, 핑크, 마젠타)
  '#86efac', '#22c55e', '#16a34a', '#4ade80', '#15803d', '#a3e635', '#f472b6', '#ec4899', '#db2777', '#831843',
  // 4행 (다크그린, 모노크롬, 무채색)
  '#14532d', '#166534', '#0f766e', '#115e59', '#064e3b', '#3b0764', '#000000', '#525252', '#a3a3a3', '#ffffff',
];

const CUSTOM_COLORS_STORAGE_KEY = 'mymemo_paint_custom_colors';

// 색상 변환 유틸리티
function hexToRgb(hex: string): { r: number; g: number; b: number } {
  let clean = hex.replace('#', '').trim();
  if (clean.length === 3) {
    clean = clean.split('').map((c) => c + c).join('');
  }
  const num = parseInt(clean, 16);
  if (isNaN(num) || clean.length !== 6) {
    return { r: 0, g: 0, b: 0 };
  }
  return {
    r: (num >> 16) & 255,
    g: (num >> 8) & 255,
    b: num & 255,
  };
}

function rgbToHex(r: number, g: number, b: number): string {
  const clamp = (v: number) => Math.max(0, Math.min(255, Math.round(v)));
  const toHex = (v: number) => clamp(v).toString(16).padStart(2, '0');
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

// RGB -> HSL (h: 0~360, s: 0~1, l: 0~1)
function rgbToHsl(r: number, g: number, b: number): { h: number; s: number; l: number } {
  r /= 255;
  g /= 255;
  b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let h = 0;
  let s = 0;
  const l = (max + min) / 2;

  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    switch (max) {
      case r:
        h = (g - b) / d + (g < b ? 6 : 0);
        break;
      case g:
        h = (b - r) / d + 2;
        break;
      case b:
        h = (r - g) / d + 4;
        break;
    }
    h *= 60;
  }
  return { h, s, l };
}

// HSL -> RGB (h: 0~360, s: 0~1, l: 0~1)
function hslToRgb(h: number, s: number, l: number): { r: number; g: number; b: number } {
  h = ((h % 360) + 360) % 360;
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
  const m = l - c / 2;
  let r1 = 0, g1 = 0, b1 = 0;

  if (h >= 0 && h < 60) {
    r1 = c; g1 = x; b1 = 0;
  } else if (h >= 60 && h < 120) {
    r1 = x; g1 = c; b1 = 0;
  } else if (h >= 120 && h < 180) {
    r1 = 0; g1 = c; b1 = x;
  } else if (h >= 180 && h < 240) {
    r1 = 0; g1 = x; b1 = c;
  } else if (h >= 240 && h < 300) {
    r1 = x; g1 = 0; b1 = c;
  } else {
    r1 = c; g1 = 0; b1 = x;
  }

  return {
    r: Math.round((r1 + m) * 255),
    g: Math.round((g1 + m) * 255),
    b: Math.round((b1 + m) * 255),
  };
}

export const ColorEditModal: React.FC<ColorEditModalProps> = ({
  isOpen,
  initialColor,
  onSelectColor,
  onClose,
}) => {
  const [hexInput, setHexInput] = useState<string>('#000000');
  const [rgb, setRgb] = useState<{ r: number; g: number; b: number }>({ r: 0, g: 0, b: 0 });
  const [hsl, setHsl] = useState<{ h: number; s: number; l: number }>({ h: 0, s: 1, l: 0 });
  const [colorModel, setColorModel] = useState<'RGB' | 'HSL'>('RGB');

  // 사용자 지정 색 (24개 슬롯)
  const [customColors, setCustomColors] = useState<(string | null)[]>(() => {
    try {
      const saved = localStorage.getItem(CUSTOM_COLORS_STORAGE_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) return parsed.slice(0, 24);
      }
    } catch {}
    return Array(24).fill(null);
  });
  const [selectedCustomIndex, setSelectedCustomIndex] = useState<number | null>(null);

  const spectrumCanvasRef = useRef<HTMLCanvasElement>(null);
  const lumCanvasRef = useRef<HTMLCanvasElement>(null);
  const isDraggingSpectrum = useRef<boolean>(false);
  const isDraggingLum = useRef<boolean>(false);

  // 모달 열릴 때 초기화
  useEffect(() => {
    if (!isOpen) return;
    const initialHex = initialColor && initialColor.startsWith('#') ? initialColor : '#000000';
    const initialRgb = hexToRgb(initialHex);
    const initialHsl = rgbToHsl(initialRgb.r, initialRgb.g, initialRgb.b);
    setHexInput(initialHex);
    setRgb(initialRgb);
    setHsl(initialHsl);
  }, [isOpen, initialColor]);

  // 2D 스펙트럼 캔버스 렌더링 (X=Hue 0~360, Y=Saturation 100%~0%)
  const drawSpectrumCanvas = useCallback(() => {
    const canvas = spectrumCanvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const w = canvas.width;
    const h = canvas.height;

    // 1. 가로 무지개 (Hue) 그라디언트
    const hueGrad = ctx.createLinearGradient(0, 0, w, 0);
    hueGrad.addColorStop(0 / 6, '#ff0000');
    hueGrad.addColorStop(1 / 6, '#ffff00');
    hueGrad.addColorStop(2 / 6, '#00ff00');
    hueGrad.addColorStop(3 / 6, '#00ffff');
    hueGrad.addColorStop(4 / 6, '#0000ff');
    hueGrad.addColorStop(5 / 6, '#ff00ff');
    hueGrad.addColorStop(6 / 6, '#ff0000');

    ctx.fillStyle = hueGrad;
    ctx.fillRect(0, 0, w, h);

    // 2. 세로 흰색 (Saturation) 그라디언트 (상단: 투명, 하단: 흰색)
    const satGrad = ctx.createLinearGradient(0, 0, 0, h);
    satGrad.addColorStop(0, 'rgba(255, 255, 255, 0)');
    satGrad.addColorStop(1, 'rgba(255, 255, 255, 1)');

    ctx.fillStyle = satGrad;
    ctx.fillRect(0, 0, w, h);
  }, []);

  // 수직 명도(Luminosity) 캔버스 렌더링 (상단: 흰색, 중간: 순수 Hue/Sat 색상, 하단: 검정)
  const drawLumCanvas = useCallback(() => {
    const canvas = lumCanvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const w = canvas.width;
    const h = canvas.height;

    // 현재 HSL에서 L=0.5(순수 색상) 추출
    const pureRgb = hslToRgb(hsl.h, hsl.s, 0.5);
    const pureColor = rgbToHex(pureRgb.r, pureRgb.g, pureRgb.b);

    const lumGrad = ctx.createLinearGradient(0, 0, 0, h);
    lumGrad.addColorStop(0, '#ffffff'); // 최상단: 흰색
    lumGrad.addColorStop(0.5, pureColor); // 중간: 현재 순수 색상
    lumGrad.addColorStop(1, '#000000'); // 최하단: 검정

    ctx.fillStyle = lumGrad;
    ctx.fillRect(0, 0, w, h);
  }, [hsl.h, hsl.s]);

  useEffect(() => {
    if (isOpen) {
      drawSpectrumCanvas();
      drawLumCanvas();
    }
  }, [isOpen, drawSpectrumCanvas, drawLumCanvas]);

  // H, S, L 값이 바뀔 때 캔버스 및 RGB/Hex 동기화 (사용자가 입력하거나 드래그할 때)
  const updateFromHsl = (newH: number, newS: number, newL: number) => {
    const clampedH = ((newH % 360) + 360) % 360;
    const clampedS = Math.max(0, Math.min(1, newS));
    const clampedL = Math.max(0, Math.min(1, newL));

    const newHslObj = { h: clampedH, s: clampedS, l: clampedL };
    const newRgbObj = hslToRgb(clampedH, clampedS, clampedL);
    const newHex = rgbToHex(newRgbObj.r, newRgbObj.g, newRgbObj.b);

    setHsl(newHslObj);
    setRgb(newRgbObj);
    setHexInput(newHex);
  };

  const updateFromRgb = (r: number, g: number, b: number) => {
    const clampedR = Math.max(0, Math.min(255, r));
    const clampedG = Math.max(0, Math.min(255, g));
    const clampedB = Math.max(0, Math.min(255, b));

    const newRgb = { r: clampedR, g: clampedG, b: clampedB };
    const newHex = rgbToHex(clampedR, clampedG, clampedB);
    const newHsl = rgbToHsl(clampedR, clampedG, clampedB);

    setRgb(newRgb);
    setHexInput(newHex);
    setHsl(newHsl);
  };

  const updateFromHex = (hex: string) => {
    setHexInput(hex);
    if (/^#[0-9A-Fa-f]{6}$/.test(hex)) {
      const newRgb = hexToRgb(hex);
      const newHsl = rgbToHsl(newRgb.r, newRgb.g, newRgb.b);
      setRgb(newRgb);
      setHsl(newHsl);
    }
  };

  // 2D 스펙트럼 마우스 이벤트
  const handleSpectrumPointer = (clientX: number, clientY: number) => {
    const canvas = spectrumCanvasRef.current;
    if (!canvas) return;
    const rect = canvas.getBoundingClientRect();
    const x = Math.max(0, Math.min(rect.width, clientX - rect.left));
    const y = Math.max(0, Math.min(rect.height, clientY - rect.top));

    const newH = (x / rect.width) * 360;
    const newS = 1 - y / rect.height; // 위쪽이 100%, 아래쪽이 0%(흰색)

    updateFromHsl(newH, newS, hsl.l);
  };

  // 수직 슬라이더 마우스 이벤트 (Luminance)
  const handleLumPointer = (clientY: number) => {
    const canvas = lumCanvasRef.current;
    if (!canvas) return;
    const rect = canvas.getBoundingClientRect();
    const y = Math.max(0, Math.min(rect.height, clientY - rect.top));
    const newL = 1 - y / rect.height; // 위쪽이 1(흰색), 아래쪽이 0(검정)
    updateFromHsl(hsl.h, hsl.s, newL);
  };

  // 사용자 지정 색 추가 ([+] 버튼)
  const handleAddCustomColor = () => {
    const nextColors = [...customColors];
    const targetIdx =
      selectedCustomIndex !== null && selectedCustomIndex >= 0
        ? selectedCustomIndex
        : nextColors.findIndex((c) => c === null);

    const insertIdx = targetIdx >= 0 ? targetIdx : 0;
    nextColors[insertIdx] = hexInput;
    setCustomColors(nextColors);
    setSelectedCustomIndex(insertIdx);
    try {
      localStorage.setItem(CUSTOM_COLORS_STORAGE_KEY, JSON.stringify(nextColors));
    } catch {}
  };

  if (!isOpen) return null;

  // 2D 스펙트럼 마커 위치 (%)
  const markerX = (hsl.h / 360) * 100;
  const markerY = (1 - hsl.s) * 100;
  // 수직 슬라이더 마커 위치 (%)
  const lumMarkerY = (1 - hsl.l) * 100;

  return (
    <div
      className="fixed inset-0 z-70 flex items-center justify-center bg-black/75 backdrop-blur-xs p-3 animate-in fade-in select-none"
      onClick={onClose}
    >
      <div
        className="bg-slate-900 border border-slate-700/80 rounded-2xl p-5 w-full max-w-xl shadow-2xl space-y-4 text-slate-200"
        onClick={(e) => e.stopPropagation()}
      >
        {/* 헤더 */}
        <div className="flex items-center justify-between pb-2 border-b border-slate-800">
          <span className="font-bold text-base text-slate-100 flex items-center gap-2">
            <span>🎨 색 편집</span>
            <span className="text-xs text-slate-400 font-normal">(그림판 스타일 색상 선택)</span>
          </span>
          <button
            type="button"
            onClick={onClose}
            className="p-1 hover:bg-slate-800 rounded-lg text-slate-400 hover:text-white transition-colors"
          >
            <X size={18} />
          </button>
        </div>

        {/* 상단 2D 스펙트럼 + 슬라이더 + 입력 제어 영역 */}
        <div className="grid grid-cols-1 md:grid-cols-12 gap-4 items-start">
          {/* 1. 2D 색상 스펙트럼 (7 cols) */}
          <div className="md:col-span-6 flex gap-3 h-[200px]">
            <div
              className="relative flex-1 h-full rounded-xl overflow-hidden border border-slate-700 cursor-crosshair shadow-inner"
              onPointerDown={(e) => {
                isDraggingSpectrum.current = true;
                e.currentTarget.setPointerCapture(e.pointerId);
                handleSpectrumPointer(e.clientX, e.clientY);
              }}
              onPointerMove={(e) => {
                if (isDraggingSpectrum.current) {
                  handleSpectrumPointer(e.clientX, e.clientY);
                }
              }}
              onPointerUp={(e) => {
                isDraggingSpectrum.current = false;
                try {
                  e.currentTarget.releasePointerCapture(e.pointerId);
                } catch {}
              }}
            >
              <canvas
                ref={spectrumCanvasRef}
                width={240}
                height={200}
                className="w-full h-full block"
              />
              {/* 원형 마커 */}
              <div
                className="absolute w-4 h-4 rounded-full border-2 border-white shadow-md pointer-events-none -translate-x-1/2 -translate-y-1/2"
                style={{
                  left: `${markerX}%`,
                  top: `${markerY}%`,
                  backgroundColor: 'transparent',
                  boxShadow: '0 0 0 1px rgba(0,0,0,0.8)',
                }}
              />
            </div>

            {/* 수직 명도(Luminance) 슬라이더 (18px 너비) */}
            <div
              className="relative w-6 h-full rounded-xl overflow-hidden border border-slate-700 cursor-ns-resize shadow-inner"
              onPointerDown={(e) => {
                isDraggingLum.current = true;
                e.currentTarget.setPointerCapture(e.pointerId);
                handleLumPointer(e.clientY);
              }}
              onPointerMove={(e) => {
                if (isDraggingLum.current) {
                  handleLumPointer(e.clientY);
                }
              }}
              onPointerUp={(e) => {
                isDraggingLum.current = false;
                try {
                  e.currentTarget.releasePointerCapture(e.pointerId);
                } catch {}
              }}
            >
              <canvas
                ref={lumCanvasRef}
                width={24}
                height={200}
                className="w-full h-full block"
              />
              {/* 슬라이더 인디케이터 화살표/바 */}
              <div
                className="absolute left-0 right-0 h-2 -translate-y-1/2 border-y-2 border-white pointer-events-none"
                style={{
                  top: `${lumMarkerY}%`,
                  boxShadow: '0 0 0 1px rgba(0,0,0,0.9)',
                }}
              />
            </div>
          </div>

          {/* 2. 미리보기 박스 & RGB / HEX 입력 영역 (6 cols) */}
          <div className="md:col-span-6 flex flex-col justify-between h-[200px] bg-slate-950/70 p-3.5 rounded-xl border border-slate-800">
            {/* 현재 색상 미리보기 & Hex */}
            <div className="flex items-center gap-3">
              <div
                className="w-14 h-14 rounded-xl border-2 border-slate-600 shadow-md shrink-0 flex items-center justify-center"
                style={{ backgroundColor: hexInput }}
              />
              <div className="flex-1 space-y-1">
                <span className="text-[11px] text-slate-400 font-semibold block">현재 선택 색상</span>
                <input
                  type="text"
                  value={hexInput}
                  onChange={(e) => updateFromHex(e.target.value)}
                  placeholder="#000000"
                  maxLength={7}
                  className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1 text-sm font-mono font-bold text-white uppercase focus:outline-none focus:border-indigo-500"
                />
              </div>
            </div>

            {/* 모델 드롭다운 및 RGB/HSL 수치 입력 */}
            <div className="space-y-1.5 pt-2 border-t border-slate-800/80">
              <div className="flex items-center justify-between text-xs">
                <span className="text-slate-400 font-medium">색상 모델</span>
                <select
                  value={colorModel}
                  onChange={(e) => setColorModel(e.target.value as 'RGB' | 'HSL')}
                  className="bg-slate-900 border border-slate-700 rounded px-2 py-0.5 text-xs font-semibold text-slate-200"
                >
                  <option value="RGB">RGB</option>
                  <option value="HSL">HSL</option>
                </select>
              </div>

              {colorModel === 'RGB' ? (
                <div className="grid grid-cols-3 gap-2">
                  <div className="space-y-0.5">
                    <span className="text-[10px] text-red-400 font-bold block">빨강 (R)</span>
                    <input
                      type="number"
                      min={0}
                      max={255}
                      value={rgb.r}
                      onChange={(e) => updateFromRgb(Number(e.target.value), rgb.g, rgb.b)}
                      className="w-full bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs font-bold text-center text-white focus:outline-none focus:border-red-500"
                    />
                  </div>
                  <div className="space-y-0.5">
                    <span className="text-[10px] text-green-400 font-bold block">녹색 (G)</span>
                    <input
                      type="number"
                      min={0}
                      max={255}
                      value={rgb.g}
                      onChange={(e) => updateFromRgb(rgb.r, Number(e.target.value), rgb.b)}
                      className="w-full bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs font-bold text-center text-white focus:outline-none focus:border-green-500"
                    />
                  </div>
                  <div className="space-y-0.5">
                    <span className="text-[10px] text-blue-400 font-bold block">파랑 (B)</span>
                    <input
                      type="number"
                      min={0}
                      max={255}
                      value={rgb.b}
                      onChange={(e) => updateFromRgb(rgb.r, rgb.g, Number(e.target.value))}
                      className="w-full bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs font-bold text-center text-white focus:outline-none focus:border-blue-500"
                    />
                  </div>
                </div>
              ) : (
                <div className="grid grid-cols-3 gap-2">
                  <div className="space-y-0.5">
                    <span className="text-[10px] text-amber-400 font-bold block">색상 (H)</span>
                    <input
                      type="number"
                      min={0}
                      max={360}
                      value={Math.round(hsl.h)}
                      onChange={(e) => updateFromHsl(Number(e.target.value), hsl.s, hsl.l)}
                      className="w-full bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs font-bold text-center text-white focus:outline-none focus:border-amber-500"
                    />
                  </div>
                  <div className="space-y-0.5">
                    <span className="text-[10px] text-teal-400 font-bold block">채도 (S)</span>
                    <input
                      type="number"
                      min={0}
                      max={100}
                      value={Math.round(hsl.s * 100)}
                      onChange={(e) => updateFromHsl(hsl.h, Number(e.target.value) / 100, hsl.l)}
                      className="w-full bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs font-bold text-center text-white focus:outline-none focus:border-teal-500"
                    />
                  </div>
                  <div className="space-y-0.5">
                    <span className="text-[10px] text-purple-400 font-bold block">명도 (L)</span>
                    <input
                      type="number"
                      min={0}
                      max={100}
                      value={Math.round(hsl.l * 100)}
                      onChange={(e) => updateFromHsl(hsl.h, hsl.s, Number(e.target.value) / 100)}
                      className="w-full bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs font-bold text-center text-white focus:outline-none focus:border-purple-500"
                    />
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* 하단 1: 기본 색 (40개 그리드) */}
        <div className="space-y-1.5 pt-2 border-t border-slate-800">
          <span className="text-xs font-semibold text-slate-300 block">기본 색 (40종)</span>
          <div className="grid grid-cols-10 gap-1.5 bg-slate-950/60 p-2.5 rounded-xl border border-slate-800/80">
            {BASIC_COLORS.map((c, idx) => (
              <button
                key={`${c}-${idx}`}
                type="button"
                onClick={() => updateFromHex(c)}
                style={{ backgroundColor: c }}
                title={c}
                className={`w-6 h-6 rounded-full border transition-all ${
                  hexInput.toLowerCase() === c.toLowerCase()
                    ? 'border-white scale-125 shadow-md ring-2 ring-indigo-500'
                    : 'border-slate-700/80 hover:scale-110 hover:border-slate-400'
                }`}
              />
            ))}
          </div>
        </div>

        {/* 하단 2: 사용자 지정 색 (24개 슬롯 + [+] 버튼) */}
        <div className="space-y-1.5 pt-1">
          <div className="flex items-center justify-between">
            <span className="text-xs font-semibold text-slate-300">사용자 지정 색</span>
            <button
              type="button"
              onClick={handleAddCustomColor}
              className="flex items-center gap-1 text-[11px] font-bold text-indigo-400 hover:text-indigo-300 bg-indigo-950/50 hover:bg-indigo-900/60 px-2 py-0.5 rounded-md border border-indigo-500/30 transition-colors"
              title="현재 선택한 색상을 사용자 지정 색 슬롯에 등록"
            >
              <Plus size={13} />
              <span>현재 색상 추가</span>
            </button>
          </div>

          <div className="grid grid-cols-12 gap-1.5 bg-slate-950/60 p-2.5 rounded-xl border border-slate-800/80">
            {customColors.map((c, idx) => (
              <button
                key={`custom-${idx}`}
                type="button"
                onClick={() => {
                  setSelectedCustomIndex(idx);
                  if (c) updateFromHex(c);
                }}
                style={{ backgroundColor: c || 'transparent' }}
                title={c ? `사용자 색상: ${c}` : `빈 슬롯 ${idx + 1}`}
                className={`w-6 h-6 rounded-full border transition-all flex items-center justify-center ${
                  c
                    ? hexInput.toLowerCase() === c.toLowerCase()
                      ? 'border-white scale-125 shadow-md ring-2 ring-indigo-500'
                      : 'border-slate-700 hover:scale-110'
                    : selectedCustomIndex === idx
                    ? 'border-dashed border-indigo-400 bg-indigo-950/30'
                    : 'border-dashed border-slate-700 hover:border-slate-500 bg-slate-900/50'
                }`}
              >
                {!c && <span className="text-[9px] text-slate-600 font-mono">+</span>}
              </button>
            ))}
          </div>
        </div>

        {/* 버튼 영역: 확인 / 취소 */}
        <div className="flex items-center justify-end gap-2 pt-3 border-t border-slate-800">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-xs font-semibold transition-colors"
          >
            취소
          </button>
          <button
            type="button"
            onClick={() => {
              onSelectColor(hexInput);
              onClose();
            }}
            className="px-6 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-xs font-bold transition-all shadow-md shadow-blue-600/30"
          >
            확인
          </button>
        </div>
      </div>
    </div>
  );
};
