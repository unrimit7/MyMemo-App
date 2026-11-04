import React, { useState, useEffect, useRef, useCallback } from 'react';
import {
  X,
  Save,
  Download,
  PenTool,
  Highlighter,
  Type,
  Stamp,
  RotateCcw,
  Trash2,
  ZoomIn,
  ZoomOut,
  Loader2,
  CheckCircle,
  Move,
  Crop,
  Pipette,
  Palette,
  Copy,
  RotateCw,
  FlipHorizontal,
  FlipVertical,
  Minus,
  Circle,
  Square as RectIcon,
  ArrowUpRight,
  EyeOff
} from 'lucide-react';
import { ColorEditModal } from './ColorEditModal';

interface ImageEditorModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: string; // image data URL
  title: string;
  attId?: string;
  onSave?: (updatedDataUrl: string) => Promise<void> | void;
}

interface Stroke {
  id: string;
  type: 'pen' | 'highlighter' | 'eraser';
  color: string;
  width: number;
  opacity: number;
  points: { x: number; y: number }[];
}

interface ShapeItem {
  id: string;
  type: 'rect' | 'circle' | 'arrow' | 'line';
  color: string;
  width: number;
  startX: number;
  startY: number;
  endX: number;
  endY: number;
  fillColor?: string;
}

interface TextAnnotation {
  id: string;
  x: number; // 0~1 normalized
  y: number; // 0~1 normalized
  text: string;
  fontSize: number;
  color: string;
  hasBox: boolean;
  boxBorderColor: string;
  boxBgColor: string;
}

interface StampAnnotation {
  id: string;
  x: number; // 0~1 normalized
  y: number; // 0~1 normalized
  width: number;
  height: number;
  dataUrl: string;
}

interface HistoryAction {
  type: 'stroke' | 'shape' | 'text' | 'stamp' | 'base_snapshot';
  id?: string;
  snapshotDataUrl?: string; // 베이스 이미지(회전/반전/자르기) 스냅샷
}

const PEN_COLORS = ['#dc2626', '#2563eb', '#16a34a', '#000000', '#ffffff', '#9333ea', '#f59e0b'];
const HIGHLIGHTER_COLORS = ['#fde047', '#86efac', '#f472b6', '#67e8f9'];
const TEXT_COLORS = ['#000000', '#ffffff', '#dc2626', '#2563eb', '#16a34a', '#9333ea'];

export const ImageEditorModal: React.FC<ImageEditorModalProps> = ({
  isOpen,
  onClose,
  data,
  title,
  attId: _attId,
  onSave,
}) => {
  // 베이스 이미지 상태 (회전, 반전, 크롭 등에 의해 변형 가능)
  const [baseImageSrc, setBaseImageSrc] = useState<string>(data);
  const [imgNaturalSize, setImgNaturalSize] = useState<{ w: number; h: number }>({ w: 800, h: 600 });
  const [zoom, setZoom] = useState<number>(1.0);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [saveSuccess, setSaveSuccess] = useState<boolean>(false);

  // 툴 선택
  type Tool = 'pan' | 'select' | 'pen' | 'highlighter' | 'shape' | 'mosaic' | 'text' | 'stamp' | 'eyedropper';
  const [activeTool, setActiveTool] = useState<Tool>('pen');
  const previousToolRef = useRef<Tool>('pen');

  // 드로잉 속성
  const [penColor, setPenColor] = useState<string>('#dc2626');
  const [penWidth, setPenWidth] = useState<number>(3);
  const [highlighterColor, setHighlighterColor] = useState<string>('#fde047');
  const [highlighterWidth, setHighlighterWidth] = useState<number>(20);

  // 도형 도구 설정
  const [selectedShapeType, setSelectedShapeType] = useState<'rect' | 'circle' | 'arrow' | 'line'>('rect');
  const [shapeWidth, setShapeWidth] = useState<number>(3);
  const [shapeColor, setShapeColor] = useState<string>('#dc2626');

  // 텍스트 속성
  const [textFontSize, setTextFontSize] = useState<number>(20);
  const [textColor, setTextColor] = useState<string>('#000000');
  const [textHasBox, setTextHasBox] = useState<boolean>(false);
  const [textBoxBorderColor, setTextBoxBorderColor] = useState<string>('#000000');
  const [textBoxBgColor, setTextBoxBgColor] = useState<string>('transparent');

  // 모달 팝업 상태들
  const [showTextModal, setShowTextModal] = useState<boolean>(false);
  const [pendingTextPos, setPendingTextPos] = useState<{ x: number; y: number } | null>(null);
  const [inputModalText, setInputModalText] = useState<string>('');

  const [showSignPad, setShowSignPad] = useState<boolean>(false);
  const [pendingStampPos, setPendingStampPos] = useState<{ x: number; y: number } | null>(null);

  const [showColorEditModal, setShowColorEditModal] = useState<boolean>(false);
  const [colorEditTarget, setColorEditTarget] = useState<'pen' | 'highlighter' | 'shape' | 'text' | 'boxBorder'>('pen');
  const [eyedropperToast, setEyedropperToast] = useState<string | null>(null);

  // 사각형 영역 선택 상태
  const [selectionRect, setSelectionRect] = useState<{ x: number; y: number; w: number; h: number } | null>(null);
  const selectionStartPosRef = useRef<{ x: number; y: number } | null>(null);

  // 주석 컬렉션 & 히스토리
  const [strokes, setStrokes] = useState<Stroke[]>([]);
  const [shapes, setShapes] = useState<ShapeItem[]>([]);
  const [texts, setTexts] = useState<TextAnnotation[]>([]);
  const [stamps, setStamps] = useState<StampAnnotation[]>([]);
  const [history, setHistory] = useState<HistoryAction[]>([]);

  // 캔버스 & 컨테이너 Refs
  const baseCanvasRef = useRef<HTMLCanvasElement>(null);
  const overlayCanvasRef = useRef<HTMLCanvasElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const signCanvasRef = useRef<HTMLCanvasElement>(null);

  // 드로잉 중 임시 객체들
  const currentStrokeRef = useRef<Stroke | null>(null);
  const currentShapeRef = useRef<ShapeItem | null>(null);
  const mosaicStartPosRef = useRef<{ x: number; y: number } | null>(null);
  const currentMosaicRectRef = useRef<{ x: number; y: number; w: number; h: number } | null>(null);

  // 드래그 중인 아이템 추적
  const draggingItemRef = useRef<{
    type: 'stamp' | 'text';
    id: string;
    offsetX: number;
    offsetY: number;
  } | null>(null);

  // [이동] 손 도구 팬 드래그
  const [isPanDragging, setIsPanDragging] = useState<boolean>(false);
  const panDragStartRef = useRef<{ x: number; y: number; scrollLeft: number; scrollTop: number } | null>(null);

  // 원본 이미지 로드
  useEffect(() => {
    if (!isOpen || !data) return;
    setBaseImageSrc(data);
    setStrokes([]);
    setShapes([]);
    setTexts([]);
    setStamps([]);
    setHistory([]);
    setSelectionRect(null);
    setZoom(1.0);
  }, [isOpen, data]);

  // 베이스 이미지 로드 및 기본 캔버스 렌더링
  useEffect(() => {
    if (!isOpen || !baseImageSrc) return;
    setIsLoading(true);

    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.onload = () => {
      setImgNaturalSize({ w: img.naturalWidth, h: img.naturalHeight });
      const baseCanvas = baseCanvasRef.current;
      const overlayCanvas = overlayCanvasRef.current;
      if (baseCanvas) {
        baseCanvas.width = img.naturalWidth;
        baseCanvas.height = img.naturalHeight;
        const ctx = baseCanvas.getContext('2d');
        if (ctx) {
          ctx.drawImage(img, 0, 0);
        }
      }
      if (overlayCanvas) {
        overlayCanvas.width = img.naturalWidth;
        overlayCanvas.height = img.naturalHeight;
      }
      setIsLoading(false);
      redrawOverlay();
    };
    img.onerror = () => {
      setIsLoading(false);
    };
    img.src = baseImageSrc;
  }, [isOpen, baseImageSrc]);

  // 오버레이 다시 그리기 (스트로크, 도형, 텍스트, 스탬프, 선택 마키)
  const redrawOverlay = useCallback(() => {
    const overlay = overlayCanvasRef.current;
    if (!overlay) return;
    const ctx = overlay.getContext('2d');
    if (!ctx) return;

    const w = overlay.width;
    const h = overlay.height;
    ctx.clearRect(0, 0, w, h);

    // 1. 스트로크 (펜, 형광펜)
    strokes.forEach((stroke) => {
      if (stroke.points.length < 2) return;
      ctx.save();
      ctx.beginPath();
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.strokeStyle = stroke.color;
      ctx.lineWidth = stroke.width * (w / 800);
      ctx.globalAlpha = stroke.opacity;

      ctx.moveTo(stroke.points[0].x * w, stroke.points[0].y * h);
      for (let i = 1; i < stroke.points.length; i++) {
        ctx.lineTo(stroke.points[i].x * w, stroke.points[i].y * h);
      }
      ctx.stroke();
      ctx.restore();
    });

    // 2. 도형 (사각형, 원, 화살표, 직선)
    shapes.forEach((s) => {
      ctx.save();
      ctx.strokeStyle = s.color;
      ctx.lineWidth = s.width * (w / 800);
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';

      const x1 = s.startX * w;
      const y1 = s.startY * h;
      const x2 = s.endX * w;
      const y2 = s.endY * h;

      if (s.type === 'rect') {
        ctx.strokeRect(Math.min(x1, x2), Math.min(y1, y2), Math.abs(x2 - x1), Math.abs(y2 - y1));
      } else if (s.type === 'circle') {
        const cx = (x1 + x2) / 2;
        const cy = (y1 + y2) / 2;
        const rx = Math.abs(x2 - x1) / 2;
        const ry = Math.abs(y2 - y1) / 2;
        ctx.beginPath();
        ctx.ellipse(cx, cy, Math.max(1, rx), Math.max(1, ry), 0, 0, Math.PI * 2);
        ctx.stroke();
      } else if (s.type === 'line') {
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);
        ctx.stroke();
      } else if (s.type === 'arrow') {
        // 화살표 본체 선
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);
        ctx.stroke();

        // 화살표 촉
        const angle = Math.atan2(y2 - y1, x2 - x1);
        const headLen = Math.max(12, s.width * (w / 800) * 3);
        ctx.fillStyle = s.color;
        ctx.beginPath();
        ctx.moveTo(x2, y2);
        ctx.lineTo(x2 - headLen * Math.cos(angle - Math.PI / 6), y2 - headLen * Math.sin(angle - Math.PI / 6));
        ctx.lineTo(x2 - headLen * Math.cos(angle + Math.PI / 6), y2 - headLen * Math.sin(angle + Math.PI / 6));
        ctx.closePath();
        ctx.fill();
      }
      ctx.restore();
    });

    // 3. 스탬프 및 오려낸 패치
    stamps.forEach((stamp) => {
      const img = new Image();
      img.onload = () => {
        const sx = stamp.x * w;
        const sy = stamp.y * h;
        const sw = stamp.width * w;
        const sh = stamp.height * h;
        ctx.drawImage(img, sx, sy, sw, sh);

        // 호버 가이드
        ctx.save();
        ctx.strokeStyle = 'rgba(99, 102, 241, 0.4)';
        ctx.setLineDash([4, 4]);
        ctx.lineWidth = 1;
        ctx.strokeRect(sx - 2, sy - 2, sw + 4, sh + 4);
        ctx.restore();
      };
      img.src = stamp.dataUrl;
    });

    // 4. 텍스트 박스
    texts.forEach((textItem) => {
      ctx.save();
      const scaledSize = Math.max(10, textItem.fontSize * (w / 800));
      ctx.font = `bold ${scaledSize}px Pretendard, -apple-system, BlinkMacSystemFont, sans-serif`;
      const x = textItem.x * w;
      const y = textItem.y * h;

      const lines = textItem.text.split('\n');
      const lineHeight = scaledSize * 1.3;

      let maxLineWidth = 0;
      lines.forEach((l) => {
        maxLineWidth = Math.max(maxLineWidth, ctx.measureText(l).width);
      });
      if (maxLineWidth === 0) {
        maxLineWidth = scaledSize * 3;
      }

      const padX = Math.max(6, scaledSize * 0.4);
      const padY = Math.max(4, scaledSize * 0.3);
      const boxX = x - padX;
      const boxY = y - scaledSize - padY;
      const boxW = maxLineWidth + padX * 2;
      const boxH = lines.length * lineHeight + padY * 2;

      // 배경
      if (textItem.boxBgColor && textItem.boxBgColor !== 'transparent') {
        ctx.fillStyle = textItem.boxBgColor;
        ctx.fillRect(boxX, boxY, boxW, boxH);
      }

      // 테두리
      if (textItem.hasBox) {
        ctx.strokeStyle = textItem.boxBorderColor || textItem.color || '#000000';
        ctx.lineWidth = Math.max(1.5, 2 * (w / 800));
        ctx.strokeRect(boxX, boxY, boxW, boxH);
      }

      // 글자
      ctx.fillStyle = textItem.color;
      lines.forEach((line, idx) => {
        ctx.fillText(line, x, y + idx * lineHeight);
      });
      ctx.restore();
    });

    // 5. 사각형 영역 선택 (Selection Marquee) 마키 박스
    if (selectionRect && selectionRect.w > 0 && selectionRect.h > 0) {
      ctx.save();
      ctx.setLineDash([6, 4]);
      ctx.strokeStyle = '#2563eb';
      ctx.lineWidth = 2;
      ctx.fillStyle = 'rgba(37, 99, 235, 0.15)';
      const sx = selectionRect.x * w;
      const sy = selectionRect.y * h;
      const sw = selectionRect.w * w;
      const sh = selectionRect.h * h;
      ctx.fillRect(sx, sy, sw, sh);
      ctx.strokeRect(sx, sy, sw, sh);

      // 네 모서리 핸들
      ctx.setLineDash([]);
      ctx.fillStyle = '#ffffff';
      ctx.strokeStyle = '#2563eb';
      ctx.lineWidth = 1.5;
      const hSize = 6;
      [
        [sx, sy],
        [sx + sw, sy],
        [sx, sy + sh],
        [sx + sw, sy + sh],
      ].forEach(([hx, hy]) => {
        ctx.fillRect(hx - hSize / 2, hy - hSize / 2, hSize, hSize);
        ctx.strokeRect(hx - hSize / 2, hy - hSize / 2, hSize, hSize);
      });
      ctx.restore();
    }
  }, [strokes, shapes, texts, stamps, selectionRect]);

  useEffect(() => {
    redrawOverlay();
  }, [redrawOverlay]);

  // 단색 패치 이미지 생성
  const createSolidColorDataUrl = (width: number, height: number, color: string = '#ffffff'): string => {
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(10, Math.round(width));
    canvas.height = Math.max(10, Math.round(height));
    const ctx = canvas.getContext('2d');
    if (ctx) {
      ctx.fillStyle = color;
      ctx.fillRect(0, 0, canvas.width, canvas.height);
    }
    return canvas.toDataURL('image/png');
  };

  // 선택 영역 이미지 캡처
  const captureSelectionPatch = (rect: { x: number; y: number; w: number; h: number }): string | null => {
    const baseCanvas = baseCanvasRef.current;
    const overlayCanvas = overlayCanvasRef.current;
    if (!baseCanvas) return null;

    const pixelX = Math.round(rect.x * baseCanvas.width);
    const pixelY = Math.round(rect.y * baseCanvas.height);
    const pixelW = Math.round(rect.w * baseCanvas.width);
    const pixelH = Math.round(rect.h * baseCanvas.height);

    if (pixelW <= 2 || pixelH <= 2) return null;

    const patchCanvas = document.createElement('canvas');
    patchCanvas.width = pixelW;
    patchCanvas.height = pixelH;
    const patchCtx = patchCanvas.getContext('2d');
    if (!patchCtx) return null;

    patchCtx.drawImage(baseCanvas, pixelX, pixelY, pixelW, pixelH, 0, 0, pixelW, pixelH);
    if (overlayCanvas) {
      patchCtx.drawImage(overlayCanvas, pixelX, pixelY, pixelW, pixelH, 0, 0, pixelW, pixelH);
    }

    return patchCanvas.toDataURL('image/png');
  };

  // 화면 스포이트 샘플링
  const sampleColorFromCanvas = (clientX: number, clientY: number): string | null => {
    const baseCanvas = baseCanvasRef.current;
    const overlayCanvas = overlayCanvasRef.current;
    if (!baseCanvas) return null;

    const rect = baseCanvas.getBoundingClientRect();
    const normX = Math.max(0, Math.min(1, (clientX - rect.left) / rect.width));
    const normY = Math.max(0, Math.min(1, (clientY - rect.top) / rect.height));

    const pixelX = Math.floor(normX * baseCanvas.width);
    const pixelY = Math.floor(normY * baseCanvas.height);

    const tempCanvas = document.createElement('canvas');
    tempCanvas.width = 1;
    tempCanvas.height = 1;
    const tempCtx = tempCanvas.getContext('2d');
    if (!tempCtx) return null;

    tempCtx.drawImage(baseCanvas, pixelX, pixelY, 1, 1, 0, 0, 1, 1);
    if (overlayCanvas) {
      tempCtx.drawImage(overlayCanvas, pixelX, pixelY, 1, 1, 0, 0, 1, 1);
    }

    const p = tempCtx.getImageData(0, 0, 1, 1).data;
    const toHex = (v: number) => v.toString(16).padStart(2, '0');
    return `#${toHex(p[0])}${toHex(p[1])}${toHex(p[2])}`;
  };

  // 1. 선택 영역 이동
  const handleMoveSelection = () => {
    if (!selectionRect) return;
    const rect = selectionRect;
    const patchDataUrl = captureSelectionPatch(rect);
    if (!patchDataUrl) return;

    const whiteDataUrl = createSolidColorDataUrl(rect.w * 800, rect.h * 600, '#ffffff');
    const whiteId = `erase_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
    const liftedId = `lifted_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;

    const whiteStamp: StampAnnotation = {
      id: whiteId,
      x: rect.x,
      y: rect.y,
      width: rect.w,
      height: rect.h,
      dataUrl: whiteDataUrl,
    };

    const liftedStamp: StampAnnotation = {
      id: liftedId,
      x: rect.x,
      y: rect.y,
      width: rect.w,
      height: rect.h,
      dataUrl: patchDataUrl,
    };

    setStamps((prev) => [...prev, whiteStamp, liftedStamp]);
    setHistory((prev) => [
      ...prev,
      { type: 'stamp', id: whiteId },
      { type: 'stamp', id: liftedId },
    ]);
    setSelectionRect(null);
  };

  // 2. 선택 영역 지우기
  const handleEraseSelection = () => {
    if (!selectionRect) return;
    const rect = selectionRect;
    const whiteDataUrl = createSolidColorDataUrl(rect.w * 800, rect.h * 600, '#ffffff');
    const eraseId = `erase_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;

    const eraseStamp: StampAnnotation = {
      id: eraseId,
      x: rect.x,
      y: rect.y,
      width: rect.w,
      height: rect.h,
      dataUrl: whiteDataUrl,
    };

    setStamps((prev) => [...prev, eraseStamp]);
    setHistory((prev) => [...prev, { type: 'stamp', id: eraseId }]);
    setSelectionRect(null);
  };

  // 3. 선택 영역 복제
  const handleDuplicateSelection = () => {
    if (!selectionRect) return;
    const rect = selectionRect;
    const patchDataUrl = captureSelectionPatch(rect);
    if (!patchDataUrl) return;

    const dupId = `dup_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
    const dupStamp: StampAnnotation = {
      id: dupId,
      x: Math.min(0.8, rect.x + 0.03),
      y: Math.min(0.8, rect.y + 0.03),
      width: rect.w,
      height: rect.h,
      dataUrl: patchDataUrl,
    };

    setStamps((prev) => [...prev, dupStamp]);
    setHistory((prev) => [...prev, { type: 'stamp', id: dupId }]);
    setSelectionRect(null);
  };

  // 4. 선택 영역으로 이미지 자르기 (Crop)
  const handleCropToSelection = () => {
    if (!selectionRect) return;
    const baseCanvas = baseCanvasRef.current;
    if (!baseCanvas) return;

    const rect = selectionRect;
    const pixelX = Math.round(rect.x * baseCanvas.width);
    const pixelY = Math.round(rect.y * baseCanvas.height);
    const pixelW = Math.round(rect.w * baseCanvas.width);
    const pixelH = Math.round(rect.h * baseCanvas.height);

    if (pixelW < 10 || pixelH < 10) return;

    // 현재 캔버스 전체(베이스+주석) 합성
    const combinedCanvas = document.createElement('canvas');
    combinedCanvas.width = baseCanvas.width;
    combinedCanvas.height = baseCanvas.height;
    const cCtx = combinedCanvas.getContext('2d');
    if (!cCtx) return;

    cCtx.drawImage(baseCanvas, 0, 0);
    if (overlayCanvasRef.current) {
      cCtx.drawImage(overlayCanvasRef.current, 0, 0);
    }

    // 자를 영역 오프스크린 캔버스에 추출
    const cropCanvas = document.createElement('canvas');
    cropCanvas.width = pixelW;
    cropCanvas.height = pixelH;
    const cropCtx = cropCanvas.getContext('2d');
    if (!cropCtx) return;

    cropCtx.drawImage(combinedCanvas, pixelX, pixelY, pixelW, pixelH, 0, 0, pixelW, pixelH);

    const croppedDataUrl = cropCanvas.toDataURL('image/png');

    // 이전 상태 히스토리 보관
    setHistory((prev) => [...prev, { type: 'base_snapshot', snapshotDataUrl: baseImageSrc }]);

    // 새 베이스 이미지 적용 및 주석 초기화
    setBaseImageSrc(croppedDataUrl);
    setStrokes([]);
    setShapes([]);
    setTexts([]);
    setStamps([]);
    setSelectionRect(null);
  };

  // 5. 모자이크 처리 (드래그 영역 픽셀화)
  const applyMosaicToRect = (rect: { x: number; y: number; w: number; h: number }) => {
    const baseCanvas = baseCanvasRef.current;
    if (!baseCanvas) return;

    const pixelX = Math.round(rect.x * baseCanvas.width);
    const pixelY = Math.round(rect.y * baseCanvas.height);
    const pixelW = Math.round(rect.w * baseCanvas.width);
    const pixelH = Math.round(rect.h * baseCanvas.height);

    if (pixelW < 4 || pixelH < 4) return;

    // 픽셀레이션(모자이크) 패치 생성
    const offCanvas = document.createElement('canvas');
    offCanvas.width = pixelW;
    offCanvas.height = pixelH;
    const offCtx = offCanvas.getContext('2d');
    if (!offCtx) return;

    // 1) 1/10 크기로 줄여서 그리기 (저해상도 화)
    const sampleScale = Math.max(0.04, Math.min(0.12, 10 / Math.max(pixelW, pixelH)));
    const smallW = Math.max(3, Math.round(pixelW * sampleScale));
    const smallH = Math.max(3, Math.round(pixelH * sampleScale));

    const smallCanvas = document.createElement('canvas');
    smallCanvas.width = smallW;
    smallCanvas.height = smallH;
    const smallCtx = smallCanvas.getContext('2d');
    if (!smallCtx) return;

    smallCtx.drawImage(baseCanvas, pixelX, pixelY, pixelW, pixelH, 0, 0, smallW, smallH);

    // 2) 다시 확대하여 모자이크 블록 효과 생성
    offCtx.imageSmoothingEnabled = false;
    offCtx.drawImage(smallCanvas, 0, 0, smallW, smallH, 0, 0, pixelW, pixelH);

    const mosaicDataUrl = offCanvas.toDataURL('image/png');
    const mosaicId = `mosaic_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;

    const mosaicStamp: StampAnnotation = {
      id: mosaicId,
      x: rect.x,
      y: rect.y,
      width: rect.w,
      height: rect.h,
      dataUrl: mosaicDataUrl,
    };

    setStamps((prev) => [...prev, mosaicStamp]);
    setHistory((prev) => [...prev, { type: 'stamp', id: mosaicId }]);
  };

  // 이미지 회전 및 반전 변형
  const handleTransformImage = (type: 'rotate_cw' | 'rotate_ccw' | 'flip_h' | 'flip_v') => {
    const baseCanvas = baseCanvasRef.current;
    if (!baseCanvas) return;

    // 전체 합성 캔버스 준비
    const combinedCanvas = document.createElement('canvas');
    combinedCanvas.width = baseCanvas.width;
    combinedCanvas.height = baseCanvas.height;
    const cCtx = combinedCanvas.getContext('2d');
    if (!cCtx) return;

    cCtx.drawImage(baseCanvas, 0, 0);
    if (overlayCanvasRef.current) {
      cCtx.drawImage(overlayCanvasRef.current, 0, 0);
    }

    const tCanvas = document.createElement('canvas');
    const isRotate = type === 'rotate_cw' || type === 'rotate_ccw';
    tCanvas.width = isRotate ? combinedCanvas.height : combinedCanvas.width;
    tCanvas.height = isRotate ? combinedCanvas.width : combinedCanvas.height;

    const tCtx = tCanvas.getContext('2d');
    if (!tCtx) return;

    if (type === 'rotate_cw') {
      tCtx.translate(tCanvas.width, 0);
      tCtx.rotate((90 * Math.PI) / 180);
    } else if (type === 'rotate_ccw') {
      tCtx.translate(0, tCanvas.height);
      tCtx.rotate((-90 * Math.PI) / 180);
    } else if (type === 'flip_h') {
      tCtx.translate(tCanvas.width, 0);
      tCtx.scale(-1, 1);
    } else if (type === 'flip_v') {
      tCtx.translate(0, tCanvas.height);
      tCtx.scale(1, -1);
    }

    tCtx.drawImage(combinedCanvas, 0, 0);

    const transformedDataUrl = tCanvas.toDataURL('image/png');
    setHistory((prev) => [...prev, { type: 'base_snapshot', snapshotDataUrl: baseImageSrc }]);

    setBaseImageSrc(transformedDataUrl);
    setStrokes([]);
    setShapes([]);
    setTexts([]);
    setStamps([]);
    setSelectionRect(null);
  };

  // 마우스 팬(Pan) 이동
  const handleContainerMouseDown = (e: React.MouseEvent<HTMLDivElement>) => {
    if (e.button !== 0 || activeTool !== 'pan') return;
    const container = containerRef.current;
    if (!container) return;

    e.preventDefault();
    panDragStartRef.current = {
      x: e.clientX,
      y: e.clientY,
      scrollLeft: container.scrollLeft,
      scrollTop: container.scrollTop,
    };
    setIsPanDragging(true);
  };

  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (!panDragStartRef.current || !containerRef.current) return;
      const dx = e.clientX - panDragStartRef.current.x;
      const dy = e.clientY - panDragStartRef.current.y;
      containerRef.current.scrollLeft = panDragStartRef.current.scrollLeft - dx;
      containerRef.current.scrollTop = panDragStartRef.current.scrollTop - dy;
    };

    const handleMouseUp = () => {
      panDragStartRef.current = null;
      setIsPanDragging(false);
    };

    if (isPanDragging) {
      window.addEventListener('mousemove', handleMouseMove);
      window.addEventListener('mouseup', handleMouseUp);
    }

    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [isPanDragging]);

  // 마우스/터치 포인터 이벤트
  const handlePointerDown = (e: React.PointerEvent<HTMLCanvasElement>) => {
    if (activeTool === 'pan') return;

    const overlay = overlayCanvasRef.current;
    if (!overlay) return;
    const rect = overlay.getBoundingClientRect();

    const normX = Math.max(0, Math.min(1, (e.clientX - rect.left) / rect.width));
    const normY = Math.max(0, Math.min(1, (e.clientY - rect.top) / rect.height));

    // 0. 스포이트
    if (activeTool === 'eyedropper') {
      const picked = sampleColorFromCanvas(e.clientX, e.clientY);
      if (picked) {
        if (previousToolRef.current === 'highlighter') setHighlighterColor(picked);
        else if (previousToolRef.current === 'shape') setShapeColor(picked);
        else if (previousToolRef.current === 'text') setTextColor(picked);
        else setPenColor(picked);

        setEyedropperToast(`추출한 색상: ${picked}`);
        setTimeout(() => setEyedropperToast(null), 2500);
      }
      setActiveTool(previousToolRef.current || 'pen');
      return;
    }

    // 1. 영역 선택 도구
    if (activeTool === 'select') {
      if (selectionRect && selectionRect.w > 0.01 && selectionRect.h > 0.01) {
        const inX = normX >= selectionRect.x && normX <= selectionRect.x + selectionRect.w;
        const inY = normY >= selectionRect.y && normY <= selectionRect.y + selectionRect.h;
        if (inX && inY) {
          handleMoveSelection();
          return;
        }
      }
      overlay.setPointerCapture(e.pointerId);
      selectionStartPosRef.current = { x: normX, y: normY };
      setSelectionRect({ x: normX, y: normY, w: 0, h: 0 });
      return;
    }

    // 2. 모자이크 도구
    if (activeTool === 'mosaic') {
      overlay.setPointerCapture(e.pointerId);
      mosaicStartPosRef.current = { x: normX, y: normY };
      currentMosaicRectRef.current = { x: normX, y: normY, w: 0, h: 0 };
      return;
    }

    // 3. 기존 도장/패치 위 클릭 -> 위치 이동 시작
    if (stamps.length > 0) {
      const clickedStamp = [...stamps].reverse().find((s) => {
        return normX >= s.x && normX <= s.x + s.width && normY >= s.y && normY <= s.y + s.height;
      });
      if (clickedStamp) {
        overlay.setPointerCapture(e.pointerId);
        draggingItemRef.current = {
          type: 'stamp',
          id: clickedStamp.id,
          offsetX: normX - clickedStamp.x,
          offsetY: normY - clickedStamp.y,
        };
        return;
      }
    }

    // 4. 기존 텍스트 박스 위 클릭 -> 위치 이동 시작
    if (texts.length > 0) {
      const w = overlay.width;
      const clickedText = [...texts].reverse().find((t) => {
        const scaledSize = Math.max(10, t.fontSize * (w / 800));
        const lineHeight = scaledSize * 1.3;
        const lines = t.text.split('\n');
        const padX = Math.max(5, scaledSize * 0.4) / w;
        const padY = Math.max(4, scaledSize * 0.3) / overlay.height;
        const boxX = t.x - padX;
        const boxY = t.y - scaledSize / overlay.height - padY;
        const boxW = (t.fontSize * 0.7 * t.text.length + 30) / w;
        const boxH = (lines.length * lineHeight) / overlay.height + padY * 2;
        return normX >= boxX && normX <= boxX + boxW && normY >= boxY && normY <= boxY + boxH;
      });
      if (clickedText) {
        overlay.setPointerCapture(e.pointerId);
        draggingItemRef.current = {
          type: 'text',
          id: clickedText.id,
          offsetX: normX - clickedText.x,
          offsetY: normY - clickedText.y,
        };
        return;
      }
    }

    // 5. 서명/도장 모달
    if (activeTool === 'stamp') {
      setPendingStampPos({ x: normX, y: normY });
      setShowSignPad(true);
      return;
    }

    // 6. 텍스트 모달
    if (activeTool === 'text') {
      setPendingTextPos({ x: normX, y: normY });
      setShowTextModal(true);
      return;
    }

    // 7. 도형 그리기 시작
    if (activeTool === 'shape') {
      overlay.setPointerCapture(e.pointerId);
      currentShapeRef.current = {
        id: `shape_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
        type: selectedShapeType,
        color: shapeColor,
        width: shapeWidth,
        startX: normX,
        startY: normY,
        endX: normX,
        endY: normY,
      };
      return;
    }

    // 8. 펜 / 형광펜 드로잉 시작
    if (activeTool === 'pen' || activeTool === 'highlighter') {
      overlay.setPointerCapture(e.pointerId);
      const isHighlighter = activeTool === 'highlighter';
      currentStrokeRef.current = {
        id: `stroke_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
        type: isHighlighter ? 'highlighter' : 'pen',
        color: isHighlighter ? highlighterColor : penColor,
        width: isHighlighter ? highlighterWidth : penWidth,
        opacity: isHighlighter ? 0.4 : 1.0,
        points: [{ x: normX, y: normY }],
      };
    }
  };

  const handlePointerMove = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const overlay = overlayCanvasRef.current;
    if (!overlay) return;
    const rect = overlay.getBoundingClientRect();

    const normX = Math.max(0, Math.min(1, (e.clientX - rect.left) / rect.width));
    const normY = Math.max(0, Math.min(1, (e.clientY - rect.top) / rect.height));

    // A. 영역 선택 드래그
    if (activeTool === 'select' && selectionStartPosRef.current) {
      const start = selectionStartPosRef.current;
      setSelectionRect({
        x: Math.min(start.x, normX),
        y: Math.min(start.y, normY),
        w: Math.abs(normX - start.x),
        h: Math.abs(normY - start.y),
      });
      return;
    }

    // B. 모자이크 드래그
    if (activeTool === 'mosaic' && mosaicStartPosRef.current) {
      const start = mosaicStartPosRef.current;
      currentMosaicRectRef.current = {
        x: Math.min(start.x, normX),
        y: Math.min(start.y, normY),
        w: Math.abs(normX - start.x),
        h: Math.abs(normY - start.y),
      };
      const ctx = overlay.getContext('2d');
      if (ctx) {
        redrawOverlay();
        const m = currentMosaicRectRef.current;
        ctx.save();
        ctx.strokeStyle = '#f59e0b';
        ctx.setLineDash([4, 4]);
        ctx.strokeRect(m.x * overlay.width, m.y * overlay.height, m.w * overlay.width, m.h * overlay.height);
        ctx.restore();
      }
      return;
    }

    // C. 도장 또는 텍스트 아이템 드래그 이동
    if (draggingItemRef.current) {
      const drag = draggingItemRef.current;
      if (drag.type === 'stamp') {
        setStamps((prev) =>
          prev.map((s) =>
            s.id === drag.id
              ? {
                  ...s,
                  x: Math.max(0, Math.min(1 - s.width, normX - drag.offsetX)),
                  y: Math.max(0, Math.min(1 - s.height, normY - drag.offsetY)),
                }
              : s
          )
        );
      } else if (drag.type === 'text') {
        setTexts((prev) =>
          prev.map((t) =>
            t.id === drag.id
              ? {
                  ...t,
                  x: Math.max(0, Math.min(0.9, normX - drag.offsetX)),
                  y: Math.max(0.04, Math.min(0.96, normY - drag.offsetY)),
                }
              : t
          )
        );
      }
      return;
    }

    // D. 도형 드래그 실시간 미리보기
    if (activeTool === 'shape' && currentShapeRef.current) {
      currentShapeRef.current.endX = normX;
      currentShapeRef.current.endY = normY;
      redrawOverlay();
      const ctx = overlay.getContext('2d');
      if (ctx) {
        const s = currentShapeRef.current;
        ctx.save();
        ctx.strokeStyle = s.color;
        ctx.lineWidth = s.width * (overlay.width / 800);
        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';
        const x1 = s.startX * overlay.width;
        const y1 = s.startY * overlay.height;
        const x2 = s.endX * overlay.width;
        const y2 = s.endY * overlay.height;

        if (s.type === 'rect') {
          ctx.strokeRect(Math.min(x1, x2), Math.min(y1, y2), Math.abs(x2 - x1), Math.abs(y2 - y1));
        } else if (s.type === 'circle') {
          const cx = (x1 + x2) / 2;
          const cy = (y1 + y2) / 2;
          ctx.beginPath();
          ctx.ellipse(cx, cy, Math.max(1, Math.abs(x2 - x1) / 2), Math.max(1, Math.abs(y2 - y1) / 2), 0, 0, Math.PI * 2);
          ctx.stroke();
        } else if (s.type === 'line' || s.type === 'arrow') {
          ctx.beginPath();
          ctx.moveTo(x1, y1);
          ctx.lineTo(x2, y2);
          ctx.stroke();
        }
        ctx.restore();
      }
      return;
    }

    // E. 펜 / 형광펜 드로잉
    if (!currentStrokeRef.current || (activeTool !== 'pen' && activeTool !== 'highlighter')) return;

    currentStrokeRef.current.points.push({ x: normX, y: normY });
    const ctx = overlay.getContext('2d');
    if (!ctx) return;
    const stroke = currentStrokeRef.current;
    const len = stroke.points.length;
    if (len < 2) return;

    const p1 = stroke.points[len - 2];
    const p2 = stroke.points[len - 1];

    ctx.save();
    ctx.beginPath();
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.strokeStyle = stroke.color;
    ctx.lineWidth = stroke.width * (overlay.width / 800);
    ctx.globalAlpha = stroke.opacity;
    ctx.moveTo(p1.x * overlay.width, p1.y * overlay.height);
    ctx.lineTo(p2.x * overlay.width, p2.y * overlay.height);
    ctx.stroke();
    ctx.restore();
  };

  const handlePointerUp = (e: React.PointerEvent<HTMLCanvasElement>) => {
    // 1. 영역 선택 완료
    if (activeTool === 'select' && selectionStartPosRef.current) {
      selectionStartPosRef.current = null;
      try {
        e.currentTarget.releasePointerCapture(e.pointerId);
      } catch {}
      if (selectionRect && (selectionRect.w < 0.008 || selectionRect.h < 0.008)) {
        setSelectionRect(null);
      }
      return;
    }

    // 2. 모자이크 완료
    if (activeTool === 'mosaic' && mosaicStartPosRef.current) {
      mosaicStartPosRef.current = null;
      try {
        e.currentTarget.releasePointerCapture(e.pointerId);
      } catch {}
      if (currentMosaicRectRef.current && currentMosaicRectRef.current.w > 0.01 && currentMosaicRectRef.current.h > 0.01) {
        applyMosaicToRect(currentMosaicRectRef.current);
      }
      currentMosaicRectRef.current = null;
      redrawOverlay();
      return;
    }

    // 3. 아이템 드래그 완료
    if (draggingItemRef.current) {
      draggingItemRef.current = null;
      try {
        e.currentTarget.releasePointerCapture(e.pointerId);
      } catch {}
      return;
    }

    // 4. 도형 그리기 완료
    if (activeTool === 'shape' && currentShapeRef.current) {
      const s = currentShapeRef.current;
      currentShapeRef.current = null;
      try {
        e.currentTarget.releasePointerCapture(e.pointerId);
      } catch {}

      if (Math.hypot(s.endX - s.startX, s.endY - s.startY) > 0.01) {
        setShapes((prev) => [...prev, s]);
        setHistory((prev) => [...prev, { type: 'shape', id: s.id }]);
      }
      return;
    }

    // 5. 펜 드로잉 완료
    if (!currentStrokeRef.current) return;
    try {
      e.currentTarget.releasePointerCapture(e.pointerId);
    } catch {}

    const finished = currentStrokeRef.current;
    currentStrokeRef.current = null;

    if (finished.points.length > 1) {
      setStrokes((prev) => [...prev, finished]);
      setHistory((prev) => [...prev, { type: 'stroke', id: finished.id }]);
    }
  };

  // 실행 취소 (Undo) - 정확한 LIFO 역순
  const handleUndo = () => {
    if (history.length === 0) return;
    const last = history[history.length - 1];
    const newHist = history.slice(0, -1);

    if (last.type === 'stroke') {
      setStrokes((prev) => prev.filter((s) => s.id !== last.id));
    } else if (last.type === 'shape') {
      setShapes((prev) => prev.filter((s) => s.id !== last.id));
    } else if (last.type === 'text') {
      setTexts((prev) => prev.filter((t) => t.id !== last.id));
    } else if (last.type === 'stamp') {
      setStamps((prev) => prev.filter((s) => s.id !== last.id));
    } else if (last.type === 'base_snapshot' && last.snapshotDataUrl) {
      setBaseImageSrc(last.snapshotDataUrl);
    }

    setHistory(newHist);
  };

  // 전체 지우기 / 원본 복원
  const handleClearAll = () => {
    if (!window.confirm('모든 편집 내용과 주석을 초기화하시겠습니까?')) return;
    setBaseImageSrc(data);
    setStrokes([]);
    setShapes([]);
    setTexts([]);
    setStamps([]);
    setHistory([]);
    setSelectionRect(null);
  };

  // 텍스트 모달 확인
  const handleConfirmInsertText = () => {
    const trimmed = inputModalText.trim();
    if (!trimmed && textBoxBgColor === 'transparent') return;

    const finalText = trimmed || '   ';
    const pos = pendingTextPos || { x: 0.5, y: 0.3 };
    const newId = `text_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;

    const newText: TextAnnotation = {
      id: newId,
      x: pos.x,
      y: pos.y,
      text: finalText,
      fontSize: textFontSize,
      color: textColor,
      hasBox: textHasBox,
      boxBorderColor: textBoxBorderColor,
      boxBgColor: textBoxBgColor,
    };

    setTexts((prev) => [...prev, newText]);
    setHistory((prev) => [...prev, { type: 'text', id: newId }]);

    setInputModalText('');
    setShowTextModal(false);
    setPendingTextPos(null);
  };

  // 손글씨 서명 삽입
  const handleInsertSignature = () => {
    const canvas = signCanvasRef.current;
    if (!canvas) return;
    const signDataUrl = canvas.toDataURL('image/png');

    const pos = pendingStampPos || { x: 0.5, y: 0.5 };
    const stampId = `stamp_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
    const newStamp: StampAnnotation = {
      id: stampId,
      x: Math.max(0, Math.min(0.8, pos.x - 0.1)),
      y: Math.max(0, Math.min(0.88, pos.y - 0.05)),
      width: 0.25,
      height: 0.15,
      dataUrl: signDataUrl,
    };

    setStamps((prev) => [...prev, newStamp]);
    setHistory((prev) => [...prev, { type: 'stamp', id: stampId }]);

    setShowSignPad(false);
    setPendingStampPos(null);
  };

  // 프리셋 도장 날인
  const handleInsertPresetStamp = (presetText: string, color: string = '#dc2626') => {
    const offCanvas = document.createElement('canvas');
    offCanvas.width = 160;
    offCanvas.height = 160;
    const ctx = offCanvas.getContext('2d');
    if (!ctx) return;

    ctx.strokeStyle = color;
    ctx.lineWidth = 6;
    ctx.beginPath();
    ctx.arc(80, 80, 72, 0, Math.PI * 2);
    ctx.stroke();

    ctx.fillStyle = color;
    ctx.font = 'bold 36px sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(presetText, 80, 80);

    const stampDataUrl = offCanvas.toDataURL('image/png');
    const pos = pendingStampPos || { x: 0.5, y: 0.5 };
    const stampId = `stamp_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;

    const newStamp: StampAnnotation = {
      id: stampId,
      x: Math.max(0, Math.min(0.85, pos.x - 0.075)),
      y: Math.max(0, Math.min(0.88, pos.y - 0.05)),
      width: 0.18,
      height: 0.18,
      dataUrl: stampDataUrl,
    };

    setStamps((prev) => [...prev, newStamp]);
    setHistory((prev) => [...prev, { type: 'stamp', id: stampId }]);

    setShowSignPad(false);
    setPendingStampPos(null);
  };

  // 최종 편집 이미지 바이너리 생성 (베이스 + 오버레이 합성)
  const generateEditedImageDataUrl = async (): Promise<string> => {
    const baseCanvas = baseCanvasRef.current;
    if (!baseCanvas) return baseImageSrc;

    const finalCanvas = document.createElement('canvas');
    finalCanvas.width = baseCanvas.width;
    finalCanvas.height = baseCanvas.height;
    const fCtx = finalCanvas.getContext('2d');
    if (!fCtx) return baseImageSrc;

    // 1. 베이스 이미지 렌더링
    fCtx.drawImage(baseCanvas, 0, 0);

    // 2. 스트로크 렌더링
    strokes.forEach((stroke) => {
      if (stroke.points.length < 2) return;
      fCtx.save();
      fCtx.beginPath();
      fCtx.lineCap = 'round';
      fCtx.lineJoin = 'round';
      fCtx.strokeStyle = stroke.color;
      fCtx.lineWidth = stroke.width * (finalCanvas.width / 800);
      fCtx.globalAlpha = stroke.opacity;

      fCtx.moveTo(stroke.points[0].x * finalCanvas.width, stroke.points[0].y * finalCanvas.height);
      for (let i = 1; i < stroke.points.length; i++) {
        fCtx.lineTo(stroke.points[i].x * finalCanvas.width, stroke.points[i].y * finalCanvas.height);
      }
      fCtx.stroke();
      fCtx.restore();
    });

    // 3. 도형 렌더링
    shapes.forEach((s) => {
      fCtx.save();
      fCtx.strokeStyle = s.color;
      fCtx.lineWidth = s.width * (finalCanvas.width / 800);
      fCtx.lineCap = 'round';
      fCtx.lineJoin = 'round';

      const x1 = s.startX * finalCanvas.width;
      const y1 = s.startY * finalCanvas.height;
      const x2 = s.endX * finalCanvas.width;
      const y2 = s.endY * finalCanvas.height;

      if (s.type === 'rect') {
        fCtx.strokeRect(Math.min(x1, x2), Math.min(y1, y2), Math.abs(x2 - x1), Math.abs(y2 - y1));
      } else if (s.type === 'circle') {
        const cx = (x1 + x2) / 2;
        const cy = (y1 + y2) / 2;
        fCtx.beginPath();
        fCtx.ellipse(cx, cy, Math.max(1, Math.abs(x2 - x1) / 2), Math.max(1, Math.abs(y2 - y1) / 2), 0, 0, Math.PI * 2);
        fCtx.stroke();
      } else if (s.type === 'line') {
        fCtx.beginPath();
        fCtx.moveTo(x1, y1);
        fCtx.lineTo(x2, y2);
        fCtx.stroke();
      } else if (s.type === 'arrow') {
        fCtx.beginPath();
        fCtx.moveTo(x1, y1);
        fCtx.lineTo(x2, y2);
        fCtx.stroke();

        const angle = Math.atan2(y2 - y1, x2 - x1);
        const headLen = Math.max(12, s.width * (finalCanvas.width / 800) * 3);
        fCtx.fillStyle = s.color;
        fCtx.beginPath();
        fCtx.moveTo(x2, y2);
        fCtx.lineTo(x2 - headLen * Math.cos(angle - Math.PI / 6), y2 - headLen * Math.sin(angle - Math.PI / 6));
        fCtx.lineTo(x2 - headLen * Math.cos(angle + Math.PI / 6), y2 - headLen * Math.sin(angle + Math.PI / 6));
        fCtx.closePath();
        fCtx.fill();
      }
      fCtx.restore();
    });

    // 4. 스탬프 렌더링
    for (const stamp of stamps) {
      await new Promise<void>((resolve) => {
        const img = new Image();
        img.onload = () => {
          fCtx.drawImage(
            img,
            stamp.x * finalCanvas.width,
            stamp.y * finalCanvas.height,
            stamp.width * finalCanvas.width,
            stamp.height * finalCanvas.height
          );
          resolve();
        };
        img.onerror = () => resolve();
        img.src = stamp.dataUrl;
      });
    }

    // 5. 텍스트 렌더링
    texts.forEach((textItem) => {
      fCtx.save();
      const scaledSize = Math.max(10, textItem.fontSize * (finalCanvas.width / 800));
      fCtx.font = `bold ${scaledSize}px Pretendard, -apple-system, BlinkMacSystemFont, sans-serif`;
      const x = textItem.x * finalCanvas.width;
      const y = textItem.y * finalCanvas.height;

      const lines = textItem.text.split('\n');
      const lineHeight = scaledSize * 1.3;

      let maxLineWidth = 0;
      lines.forEach((l) => {
        maxLineWidth = Math.max(maxLineWidth, fCtx.measureText(l).width);
      });
      if (maxLineWidth === 0) maxLineWidth = scaledSize * 3;

      const padX = Math.max(6, scaledSize * 0.4);
      const padY = Math.max(4, scaledSize * 0.3);
      const boxX = x - padX;
      const boxY = y - scaledSize - padY;
      const boxW = maxLineWidth + padX * 2;
      const boxH = lines.length * lineHeight + padY * 2;

      if (textItem.boxBgColor && textItem.boxBgColor !== 'transparent') {
        fCtx.fillStyle = textItem.boxBgColor;
        fCtx.fillRect(boxX, boxY, boxW, boxH);
      }

      if (textItem.hasBox) {
        fCtx.strokeStyle = textItem.boxBorderColor || textItem.color || '#000000';
        fCtx.lineWidth = 2;
        fCtx.strokeRect(boxX, boxY, boxW, boxH);
      }

      fCtx.fillStyle = textItem.color;
      lines.forEach((line, idx) => {
        fCtx.fillText(line, x, y + idx * lineHeight);
      });
      fCtx.restore();
    });

    return finalCanvas.toDataURL('image/png');
  };

  // 저장 및 메모 반영
  const handleSave = async () => {
    if (isSaving) return;
    setIsSaving(true);
    try {
      const updatedUrl = await generateEditedImageDataUrl();
      if (onSave) {
        await onSave(updatedUrl);
      }
      setBaseImageSrc(updatedUrl);
      setStrokes([]);
      setShapes([]);
      setTexts([]);
      setStamps([]);
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 2500);
    } catch (err) {
      console.error('이미지 저장 오류:', err);
      alert('이미지를 저장하지 못했습니다.');
    } finally {
      setIsSaving(false);
    }
  };

  // 다운로드
  const handleDownload = async () => {
    try {
      const updatedUrl = await generateEditedImageDataUrl();
      const a = document.createElement('a');
      a.href = updatedUrl;
      const safeTitle = title.replace(/\.[^/.]+$/, '') || 'edited_image';
      a.download = `${safeTitle}.png`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    } catch (err) {
      alert('다운로드에 실패했습니다.');
    }
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex flex-col bg-slate-950/95 backdrop-blur-md text-slate-100 animate-in fade-in select-none">
      {/* 1. 상단 툴바: 제목, 도구 선택, 저장 및 닫기 */}
      <div className="flex items-center justify-between px-3 py-2 bg-slate-900 border-b border-slate-800 gap-2 shrink-0 overflow-x-auto">
        <div className="flex items-center gap-2 shrink-0">
          <span className="text-indigo-400 font-bold text-xs sm:text-sm">🖼️ 이미지 편집기</span>
          <span className="text-xs sm:text-sm font-semibold truncate text-slate-200 max-w-[120px] sm:max-w-[180px]" title={title}>
            {title}
          </span>
          <span className="text-[10px] text-slate-400 font-mono hidden md:inline">
            ({imgNaturalSize.w} × {imgNaturalSize.h}px)
          </span>
          {saveSuccess && (
            <span className="flex items-center gap-1 text-[11px] text-emerald-400 bg-emerald-500/10 px-2 py-0.5 rounded-full border border-emerald-500/20 font-bold">
              <CheckCircle size={12} /> 저장됨
            </span>
          )}
        </div>

        {/* 중앙: 도구 팔레트 */}
        <div className="flex items-center gap-1 bg-slate-950 p-1 rounded-xl border border-slate-800 shrink-0">
          {/* 손 / 이동 */}
          <button
            type="button"
            onClick={() => setActiveTool('pan')}
            className={`p-1.5 rounded-lg text-xs font-semibold flex items-center gap-1 transition-all ${
              activeTool === 'pan' ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
            title="화면 이동 (마우스로 드래그하여 이동)"
          >
            <Move size={14} />
            <span className="hidden lg:inline">이동</span>
          </button>

          {/* 사각형 영역 선택 */}
          <button
            type="button"
            onClick={() => setActiveTool('select')}
            className={`p-1.5 rounded-lg text-xs font-semibold flex items-center gap-1 transition-all ${
              activeTool === 'select' ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
            title="영역 선택 (오려내어 이동, 지우기, 복제, 자르기/크롭)"
          >
            <Crop size={14} />
            <span className="hidden lg:inline">영역선택</span>
          </button>

          {/* 펜 */}
          <button
            type="button"
            onClick={() => setActiveTool('pen')}
            className={`p-1.5 rounded-lg text-xs font-semibold flex items-center gap-1 transition-all ${
              activeTool === 'pen' ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
            title="자유 펜 드로잉"
          >
            <PenTool size={14} />
            <span className="hidden lg:inline">펜</span>
          </button>

          {/* 형광펜 */}
          <button
            type="button"
            onClick={() => setActiveTool('highlighter')}
            className={`p-1.5 rounded-lg text-xs font-semibold flex items-center gap-1 transition-all ${
              activeTool === 'highlighter' ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
            title="형광펜 (반투명 강조)"
          >
            <Highlighter size={14} />
            <span className="hidden lg:inline">형광펜</span>
          </button>

          {/* 도형 */}
          <button
            type="button"
            onClick={() => setActiveTool('shape')}
            className={`p-1.5 rounded-lg text-xs font-semibold flex items-center gap-1 transition-all ${
              activeTool === 'shape' ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
            title="도형 (사각형, 원, 화살표, 직선)"
          >
            <RectIcon size={14} />
            <span className="hidden lg:inline">도형</span>
          </button>

          {/* 모자이크 */}
          <button
            type="button"
            onClick={() => setActiveTool('mosaic')}
            className={`p-1.5 rounded-lg text-xs font-semibold flex items-center gap-1 transition-all ${
              activeTool === 'mosaic' ? 'bg-amber-600 text-white shadow-sm' : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
            title="모자이크 (드래그하여 민감한 개인정보/얼굴 픽셀화 블러 처리)"
          >
            <EyeOff size={14} />
            <span className="hidden lg:inline">모자이크</span>
          </button>

          {/* 글자 */}
          <button
            type="button"
            onClick={() => setActiveTool('text')}
            className={`p-1.5 rounded-lg text-xs font-semibold flex items-center gap-1 transition-all ${
              activeTool === 'text' ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
            title="글자 / 말풍선 / 박스 입력"
          >
            <Type size={14} />
            <span className="hidden lg:inline">글자</span>
          </button>

          {/* 서명/도장 */}
          <button
            type="button"
            onClick={() => setActiveTool('stamp')}
            className={`p-1.5 rounded-lg text-xs font-semibold flex items-center gap-1 transition-all ${
              activeTool === 'stamp' ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
            title="서명 및 인장 날인"
          >
            <Stamp size={14} className={activeTool === 'stamp' ? 'text-white' : 'text-rose-400'} />
            <span className="hidden lg:inline">도장</span>
          </button>

          {/* 스포이트 */}
          <button
            type="button"
            onClick={() => {
              previousToolRef.current = activeTool === 'eyedropper' ? 'pen' : activeTool;
              setActiveTool('eyedropper');
            }}
            className={`p-1.5 rounded-lg text-xs font-semibold flex items-center gap-1 transition-all ${
              activeTool === 'eyedropper' ? 'bg-amber-600 text-white shadow-sm' : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
            title="스포이트 (이미지에서 색상 추출)"
          >
            <Pipette size={14} />
          </button>

          <div className="h-4 w-px bg-slate-800 mx-1" />

          {/* 실행 취소 (Undo) */}
          <button
            type="button"
            onClick={handleUndo}
            disabled={history.length === 0}
            className="p-1.5 rounded-lg text-xs text-slate-400 hover:text-white hover:bg-slate-800 transition-all disabled:opacity-30"
            title="되돌리기 (Undo)"
          >
            <RotateCcw size={14} />
          </button>

          {/* 전체 초기화 */}
          <button
            type="button"
            onClick={handleClearAll}
            className="p-1.5 rounded-lg text-xs text-slate-400 hover:text-rose-400 hover:bg-slate-800 transition-all"
            title="전체 편집 초기화"
          >
            <Trash2 size={14} />
          </button>
        </div>

        {/* 우측: 저장, 다운로드, 닫기 */}
        <div className="flex items-center gap-1.5 shrink-0">
          <button
            type="button"
            onClick={handleSave}
            disabled={isSaving}
            className="flex items-center gap-1 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-bold transition-all shadow-md active:scale-95 disabled:opacity-50"
          >
            {isSaving ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />}
            <span>저장 및 반영</span>
          </button>

          <button
            type="button"
            onClick={handleDownload}
            className="flex items-center gap-1 px-2.5 py-1.5 bg-slate-800 hover:bg-slate-750 text-slate-200 rounded-lg text-xs font-semibold border border-slate-700 transition-all active:scale-95"
            title="편집 이미지 다운로드"
          >
            <Download size={14} />
            <span className="hidden sm:inline">다운로드</span>
          </button>

          <button
            type="button"
            onClick={onClose}
            className="p-1.5 hover:bg-rose-500/20 text-slate-400 hover:text-rose-400 rounded-lg transition-colors"
            title="닫기"
          >
            <X size={18} />
          </button>
        </div>
      </div>

      {/* 2. 보조 옵션 툴바 (도구별 세부 옵션 + 변형 컨트롤) */}
      <div className="flex items-center justify-between px-3 py-1.5 bg-slate-900/70 border-b border-slate-800/80 gap-2 text-xs overflow-x-auto shrink-0">
        <div className="flex items-center gap-2 shrink-0">
          {/* A. 영역 선택 도구 옵션 */}
          {activeTool === 'select' && (
            <div className="flex items-center gap-1.5">
              {selectionRect && selectionRect.w > 0.01 && selectionRect.h > 0.01 ? (
                <>
                  <span className="text-blue-400 font-bold text-[11px] bg-blue-950/60 px-2 py-0.5 rounded border border-blue-500/30">
                    📐 선택됨
                  </span>
                  <button
                    type="button"
                    onClick={handleMoveSelection}
                    className="px-2 py-0.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded text-[11px] font-bold flex items-center gap-1"
                    title="선택 영역 오려내어 이동"
                  >
                    <Move size={12} />
                    <span>영역 이동</span>
                  </button>
                  <button
                    type="button"
                    onClick={handleCropToSelection}
                    className="px-2 py-0.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded text-[11px] font-bold flex items-center gap-1"
                    title="선택 영역으로 이미지 전체 자르기"
                  >
                    <Crop size={12} />
                    <span>자르기(크롭)</span>
                  </button>
                  <button
                    type="button"
                    onClick={handleEraseSelection}
                    className="px-2 py-0.5 bg-rose-600 hover:bg-rose-500 text-white rounded text-[11px] font-bold flex items-center gap-1"
                    title="선택 영역 화이트로 지우기"
                  >
                    <Trash2 size={12} />
                    <span>지우기</span>
                  </button>
                  <button
                    type="button"
                    onClick={handleDuplicateSelection}
                    className="px-2 py-0.5 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded text-[11px] font-bold flex items-center gap-1 border border-slate-700"
                    title="선택 영역 복제"
                  >
                    <Copy size={12} />
                    <span>복제</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectionRect(null)}
                    className="p-1 text-slate-400 hover:text-white"
                  >
                    <X size={14} />
                  </button>
                </>
              ) : (
                <span className="text-indigo-300 text-[11px] font-medium">
                  🖱️ 마우스로 사각형 영역을 드래그하세요 (이동, 크롭, 지우기, 복제 가능)
                </span>
              )}
            </div>
          )}

          {/* B. 모자이크 도구 옵션 */}
          {activeTool === 'mosaic' && (
            <span className="text-amber-400 text-[11px] font-medium flex items-center gap-1.5">
              <EyeOff size={13} />
              <span>가리고 싶은 영역을 마우스로 드래그하면 즉시 모자이크(블러) 처리됩니다.</span>
            </span>
          )}

          {/* C. 펜 도구 옵션 */}
          {activeTool === 'pen' && (
            <div className="flex items-center gap-1.5">
              <span className="text-slate-400 text-[11px]">색상:</span>
              <div className="flex items-center gap-1">
                {PEN_COLORS.map((c) => (
                  <button
                    key={c}
                    type="button"
                    onClick={() => setPenColor(c)}
                    style={{ backgroundColor: c }}
                    className={`w-4 h-4 rounded-full border ${penColor === c ? 'border-white scale-125 shadow-sm' : 'border-slate-700 opacity-70'}`}
                  />
                ))}
                <button
                  type="button"
                  onClick={() => {
                    setColorEditTarget('pen');
                    setShowColorEditModal(true);
                  }}
                  className="px-1.5 py-0.5 bg-slate-800 hover:bg-slate-750 text-slate-200 rounded text-[10px] font-bold flex items-center gap-1 border border-slate-700 ml-0.5"
                >
                  <Palette size={11} className="text-indigo-400" />
                  <span>색 편집...</span>
                </button>
              </div>

              <span className="text-slate-400 text-[11px] ml-1">두께:</span>
              {[2, 4, 7, 12].map((w) => (
                <button
                  key={w}
                  type="button"
                  onClick={() => setPenWidth(w)}
                  className={`px-1.5 py-0.5 rounded text-[11px] font-bold ${penWidth === w ? 'bg-indigo-600 text-white' : 'bg-slate-800 text-slate-400'}`}
                >
                  {w}px
                </button>
              ))}
              <div className="flex items-center gap-0.5 bg-slate-950 px-1 py-0.5 rounded border border-slate-800 ml-0.5">
                <input
                  type="number"
                  min={1}
                  max={60}
                  value={penWidth}
                  onChange={(e) => setPenWidth(Math.max(1, Math.min(60, Number(e.target.value) || 1)))}
                  className="w-9 bg-slate-900 border border-slate-700 text-center text-xs font-bold text-white rounded px-0.5 py-0.5"
                />
                <span className="text-[10px] text-slate-400 pr-0.5">px</span>
              </div>
            </div>
          )}

          {/* D. 형광펜 도구 옵션 */}
          {activeTool === 'highlighter' && (
            <div className="flex items-center gap-1.5">
              <span className="text-slate-400 text-[11px]">형광:</span>
              <div className="flex items-center gap-1">
                {HIGHLIGHTER_COLORS.map((c) => (
                  <button
                    key={c}
                    type="button"
                    onClick={() => setHighlighterColor(c)}
                    style={{ backgroundColor: c }}
                    className={`w-4 h-4 rounded-full border ${highlighterColor === c ? 'border-white scale-125 shadow-sm' : 'border-transparent opacity-70'}`}
                  />
                ))}
                <button
                  type="button"
                  onClick={() => {
                    setColorEditTarget('highlighter');
                    setShowColorEditModal(true);
                  }}
                  className="px-1.5 py-0.5 bg-slate-800 text-slate-200 rounded text-[10px] font-bold flex items-center gap-1 border border-slate-700 ml-0.5"
                >
                  <Palette size={11} className="text-indigo-400" />
                  <span>색 편집...</span>
                </button>
              </div>

              <span className="text-slate-400 text-[11px] ml-1">굵기:</span>
              {[15, 24, 32].map((w) => (
                <button
                  key={w}
                  type="button"
                  onClick={() => setHighlighterWidth(w)}
                  className={`px-1.5 py-0.5 rounded text-[11px] font-bold ${highlighterWidth === w ? 'bg-indigo-600 text-white' : 'bg-slate-800 text-slate-400'}`}
                >
                  {w}px
                </button>
              ))}
            </div>
          )}

          {/* E. 도형 도구 옵션 */}
          {activeTool === 'shape' && (
            <div className="flex items-center gap-2">
              <div className="flex items-center gap-1 bg-slate-950 p-0.5 rounded-lg border border-slate-800">
                {[
                  { id: 'rect', label: '사각형', icon: RectIcon },
                  { id: 'circle', label: '원', icon: Circle },
                  { id: 'arrow', label: '화살표', icon: ArrowUpRight },
                  { id: 'line', label: '직선', icon: Minus },
                ].map((s) => {
                  const Icon = s.icon;
                  return (
                    <button
                      key={s.id}
                      type="button"
                      onClick={() => setSelectedShapeType(s.id as any)}
                      className={`px-2 py-1 rounded text-xs font-semibold flex items-center gap-1 ${
                        selectedShapeType === s.id ? 'bg-indigo-600 text-white' : 'text-slate-400 hover:text-white'
                      }`}
                    >
                      <Icon size={13} />
                      <span>{s.label}</span>
                    </button>
                  );
                })}
              </div>

              <div className="flex items-center gap-1">
                {PEN_COLORS.map((c) => (
                  <button
                    key={c}
                    type="button"
                    onClick={() => setShapeColor(c)}
                    style={{ backgroundColor: c }}
                    className={`w-4 h-4 rounded-full border ${shapeColor === c ? 'border-white scale-125' : 'border-slate-700 opacity-70'}`}
                  />
                ))}
              </div>

              <div className="flex items-center gap-1">
                <span className="text-slate-400 text-[11px]">선 굵기:</span>
                {[2, 4, 7].map((w) => (
                  <button
                    key={w}
                    type="button"
                    onClick={() => setShapeWidth(w)}
                    className={`px-1.5 py-0.5 rounded text-[11px] font-bold ${shapeWidth === w ? 'bg-indigo-600 text-white' : 'bg-slate-800 text-slate-400'}`}
                  >
                    {w}px
                  </button>
                ))}
              </div>
            </div>
          )}

          {/* F. 텍스트 도구 옵션 */}
          {activeTool === 'text' && (
            <div className="flex items-center gap-2">
              <span className="text-slate-400 text-[11px]">크기:</span>
              {[14, 20, 28, 36].map((s) => (
                <button
                  key={s}
                  type="button"
                  onClick={() => setTextFontSize(s)}
                  className={`px-1.5 py-0.5 rounded text-[11px] font-bold ${textFontSize === s ? 'bg-indigo-600 text-white' : 'bg-slate-800 text-slate-400'}`}
                >
                  {s}pt
                </button>
              ))}

              <button
                type="button"
                onClick={() => {
                  setTextHasBox(false);
                  setTextBoxBgColor('#ffffff');
                  setTextColor('#000000');
                  setPendingTextPos({ x: 0.5, y: 0.3 });
                  setShowTextModal(true);
                }}
                className="px-2 py-0.5 bg-amber-600/30 hover:bg-amber-600 text-amber-200 hover:text-white rounded text-[11px] font-bold border border-amber-500/30 flex items-center gap-1"
                title="흰색 불투명 배경으로 기존 이미지 글자 가리고 새 글자 덮어쓰기"
              >
                <span>🧼 화이트 덮어쓰기</span>
              </button>

              <button
                type="button"
                onClick={() => {
                  setPendingTextPos({ x: 0.5, y: 0.3 });
                  setShowTextModal(true);
                }}
                className="px-2 py-0.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded text-[11px] font-bold shadow"
              >
                + 글자 작성
              </button>
            </div>
          )}

          {/* G. 도장 옵션 */}
          {activeTool === 'stamp' && (
            <div className="flex items-center gap-2">
              <span className="text-emerald-400 text-[11px] font-semibold">📍 서명 또는 도장을 찍을 위치를 클릭하세요.</span>
              <button
                type="button"
                onClick={() => {
                  setPendingStampPos({ x: 0.5, y: 0.5 });
                  setShowSignPad(true);
                }}
                className="px-2 py-0.5 bg-rose-600/30 hover:bg-rose-600 text-rose-200 hover:text-white rounded text-[11px] font-bold border border-rose-500/30"
              >
                + 서명/도장 만들기
              </button>
            </div>
          )}

          {/* H. 스포이트 알림 */}
          {activeTool === 'eyedropper' && (
            <span className="text-amber-400 font-bold text-[11px] flex items-center gap-1 animate-pulse">
              <Pipette size={12} />
              <span>이미지에서 추출할 색상의 위치를 클릭하세요.</span>
            </span>
          )}
        </div>

        {/* 우측: 이미지 변형 (회전/반전) & 줌 */}
        <div className="flex items-center gap-2 shrink-0">
          <div className="flex items-center gap-1 bg-slate-950 p-1 rounded-lg border border-slate-800">
            <button
              type="button"
              onClick={() => handleTransformImage('rotate_cw')}
              className="p-1 hover:text-white text-slate-400"
              title="시계 방향 90° 회전"
            >
              <RotateCw size={13} />
            </button>
            <button
              type="button"
              onClick={() => handleTransformImage('flip_h')}
              className="p-1 hover:text-white text-slate-400"
              title="좌우 반전"
            >
              <FlipHorizontal size={13} />
            </button>
            <button
              type="button"
              onClick={() => handleTransformImage('flip_v')}
              className="p-1 hover:text-white text-slate-400"
              title="상하 반전"
            >
              <FlipVertical size={13} />
            </button>
          </div>

          <div className="flex items-center gap-1 bg-slate-950 px-1.5 py-0.5 rounded-lg border border-slate-800">
            <button
              type="button"
              onClick={() => setZoom((z) => Math.max(0.3, Math.round((z - 0.2) * 10) / 10))}
              className="p-1 hover:text-white"
              title="축소"
            >
              <ZoomOut size={13} />
            </button>
            <span className="text-[11px] font-bold w-10 text-center">{Math.round(zoom * 100)}%</span>
            <button
              type="button"
              onClick={() => setZoom((z) => Math.min(3.0, Math.round((z + 0.2) * 10) / 10))}
              className="p-1 hover:text-white"
              title="확대"
            >
              <ZoomIn size={13} />
            </button>
          </div>
        </div>
      </div>

      {/* 3. 메인 이미지 작업 캔버스 뷰포트 */}
      <div
        ref={containerRef}
        onMouseDown={handleContainerMouseDown}
        className={`flex-1 overflow-auto bg-slate-950 p-4 relative select-none flex items-center justify-center ${
          isPanDragging
            ? 'cursor-grabbing'
            : activeTool === 'pan'
            ? 'cursor-grab'
            : ''
        }`}
      >
        {isLoading && (
          <div className="flex flex-col items-center justify-center gap-2 text-slate-400">
            <Loader2 className="w-8 h-8 animate-spin text-indigo-500" />
            <span className="text-xs">이미지를 불러오는 중입니다...</span>
          </div>
        )}

        <div
          className="relative shrink-0 shadow-2xl rounded-xl overflow-visible transition-transform duration-75"
          style={{
            transform: `scale(${zoom})`,
            transformOrigin: 'center center',
          }}
        >
          {/* 베이스 이미지 렌더링 캔버스 */}
          <canvas ref={baseCanvasRef} className="rounded-xl block max-w-none shadow-md bg-white" />

          {/* 주석 상호작용 오버레이 캔버스 */}
          <canvas
            ref={overlayCanvasRef}
            onPointerDown={handlePointerDown}
            onPointerMove={handlePointerMove}
            onPointerUp={handlePointerUp}
            onPointerCancel={handlePointerUp}
            className={`absolute inset-0 rounded-xl ${
              activeTool === 'pan'
                ? 'pointer-events-none'
                : 'touch-none cursor-crosshair'
            }`}
          />

          {/* 사각형 영역 선택 플로팅 액션 바 */}
          {selectionRect && selectionRect.w > 0.01 && selectionRect.h > 0.01 && (
            <div
              className="absolute z-30 flex items-center gap-1 bg-slate-900/95 border border-indigo-500/80 px-2.5 py-1.5 rounded-xl shadow-2xl backdrop-blur-xs text-xs font-bold text-white pointer-events-auto -translate-x-1/2"
              style={{
                left: `${(selectionRect.x + selectionRect.w / 2) * 100}%`,
                top:
                  selectionRect.y > 0.1
                    ? `${selectionRect.y * 100 - 5.5}%`
                    : `${(selectionRect.y + selectionRect.h) * 100 + 1.5}%`,
              }}
            >
              <button
                type="button"
                onClick={handleMoveSelection}
                className="px-2.5 py-1 bg-indigo-600 hover:bg-indigo-500 rounded-lg flex items-center gap-1 shadow-xs transition-colors"
                title="선택 영역 오려내어 이동"
              >
                <Move size={12} />
                <span>이동</span>
              </button>
              <button
                type="button"
                onClick={handleCropToSelection}
                className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-500 rounded-lg flex items-center gap-1 shadow-xs transition-colors"
                title="이 영역만 남기고 자르기"
              >
                <Crop size={12} />
                <span>자르기</span>
              </button>
              <button
                type="button"
                onClick={handleEraseSelection}
                className="px-2.5 py-1 bg-rose-600 hover:bg-rose-500 rounded-lg flex items-center gap-1 shadow-xs transition-colors"
                title="선택 영역 화이트로 지우기"
              >
                <Trash2 size={12} />
                <span>지우기</span>
              </button>
              <button
                type="button"
                onClick={handleDuplicateSelection}
                className="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 rounded-lg flex items-center gap-1 border border-slate-700 shadow-xs transition-colors"
                title="선택 영역 복제"
              >
                <Copy size={12} />
                <span>복제</span>
              </button>
              <button
                type="button"
                onClick={() => setSelectionRect(null)}
                className="p-1 hover:bg-slate-800 rounded-lg text-slate-400 hover:text-white transition-colors"
              >
                <X size={14} />
              </button>
            </div>
          )}
        </div>
      </div>

      {/* 4. 스포이트 알림 토스트 */}
      {eyedropperToast && (
        <div className="fixed top-16 left-1/2 -translate-x-1/2 z-60 bg-amber-600 text-white px-4 py-2 rounded-xl shadow-2xl text-xs font-bold flex items-center gap-2 animate-in fade-in slide-in-from-top-2 border border-amber-400">
          <Pipette size={14} />
          <span>{eyedropperToast}</span>
        </div>
      )}

      {/* 5. 텍스트 삽입 모달 */}
      {showTextModal && (
        <div
          className="fixed inset-0 z-60 flex items-center justify-center bg-black/70 backdrop-blur-xs p-4 animate-in fade-in"
          onClick={() => setShowTextModal(false)}
        >
          <div
            className="bg-slate-900 border border-slate-800 rounded-2xl p-5 w-full max-w-md shadow-2xl space-y-4"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between pb-2 border-b border-slate-800">
              <span className="font-bold text-sm text-slate-200">📝 이미지 텍스트 삽입</span>
              <button
                type="button"
                onClick={() => setShowTextModal(false)}
                className="p-1 hover:bg-slate-800 rounded text-slate-400"
              >
                <X size={16} />
              </button>
            </div>

            {/* 빠른 프리셋 */}
            <div className="grid grid-cols-2 gap-2 text-xs">
              <button
                type="button"
                onClick={() => {
                  setTextHasBox(false);
                  setTextBoxBgColor('transparent');
                }}
                className={`p-2 rounded-xl border text-left flex flex-col gap-0.5 ${
                  !textHasBox && textBoxBgColor === 'transparent'
                    ? 'bg-indigo-600/20 border-indigo-500 text-indigo-200'
                    : 'bg-slate-950 border-slate-800 text-slate-400'
                }`}
              >
                <span className="font-bold text-slate-200">📝 일반 글자</span>
                <span className="text-[10px]">배경 투명 일반 텍스트</span>
              </button>

              <button
                type="button"
                onClick={() => {
                  setTextHasBox(false);
                  setTextBoxBgColor('#ffffff');
                  setTextColor('#000000');
                }}
                className={`p-2 rounded-xl border text-left flex flex-col gap-0.5 ${
                  !textHasBox && textBoxBgColor === '#ffffff'
                    ? 'bg-amber-600/20 border-amber-500 text-amber-200'
                    : 'bg-slate-950 border-slate-800 text-slate-400'
                }`}
              >
                <span className="font-bold text-amber-300">🧼 화이트 덮어쓰기</span>
                <span className="text-[10px]">기존 글자 가리고 새 글자</span>
              </button>
            </div>

            {/* 글자 내용 입력 */}
            <div className="space-y-1.5">
              <label className="text-xs font-semibold text-slate-300">글자 내용:</label>
              <textarea
                value={inputModalText}
                onChange={(e) => setInputModalText(e.target.value)}
                placeholder="삽입할 글자를 입력하세요"
                rows={2}
                autoFocus
                className="w-full bg-slate-950 border border-slate-700 rounded-xl p-3 text-sm text-slate-100 focus:outline-none focus:border-indigo-500 resize-none font-medium"
              />
            </div>

            {/* 글자 크기 & 색상 */}
            <div className="grid grid-cols-2 gap-3 bg-slate-950/60 p-3 rounded-xl border border-slate-800/80">
              <div className="space-y-1">
                <span className="text-[11px] text-slate-400 font-semibold block">크기</span>
                <div className="flex items-center gap-1">
                  {[14, 20, 28, 36].map((s) => (
                    <button
                      key={s}
                      type="button"
                      onClick={() => setTextFontSize(s)}
                      className={`px-2 py-0.5 rounded text-xs font-bold ${
                        textFontSize === s ? 'bg-indigo-600 text-white' : 'bg-slate-800 text-slate-400'
                      }`}
                    >
                      {s}pt
                    </button>
                  ))}
                </div>
              </div>

              <div className="space-y-1">
                <span className="text-[11px] text-slate-400 font-semibold block">색상</span>
                <div className="flex items-center gap-1.5 pt-0.5">
                  {TEXT_COLORS.map((c) => (
                    <button
                      key={c}
                      type="button"
                      onClick={() => setTextColor(c)}
                      style={{ backgroundColor: c }}
                      className={`w-5 h-5 rounded-full border ${textColor === c ? 'border-white scale-125' : 'border-slate-700'}`}
                    />
                  ))}
                </div>
              </div>
            </div>

            {/* 버튼 */}
            <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-800">
              <button
                type="button"
                onClick={() => setShowTextModal(false)}
                className="px-3 py-1.5 bg-slate-800 text-slate-300 rounded-lg text-xs font-semibold"
              >
                취소
              </button>
              <button
                type="button"
                onClick={handleConfirmInsertText}
                disabled={!inputModalText.trim() && textBoxBgColor === 'transparent'}
                className="px-4 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-bold transition-all shadow disabled:opacity-40"
              >
                이미지에 삽입
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 6. 서명 / 도장 모달 */}
      {showSignPad && (
        <div
          className="fixed inset-0 z-60 flex items-center justify-center bg-black/70 backdrop-blur-xs p-4 animate-in fade-in"
          onClick={() => setShowSignPad(false)}
        >
          <div
            className="bg-slate-900 border border-slate-800 rounded-2xl p-5 w-full max-w-sm shadow-2xl space-y-4"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between pb-2 border-b border-slate-800">
              <span className="font-bold text-sm text-slate-200">✍️ 서명 또는 도장 만들기</span>
              <button
                type="button"
                onClick={() => setShowSignPad(false)}
                className="p-1 hover:bg-slate-800 rounded text-slate-400"
              >
                <X size={16} />
              </button>
            </div>

            {/* 서명 드로잉 캔버스 */}
            <div className="space-y-1.5">
              <span className="text-xs text-slate-400 block">아래 박스에 마우스나 손가락으로 서명하세요:</span>
              <div className="bg-white rounded-xl border border-slate-700 p-1 flex justify-center">
                <canvas
                  ref={signCanvasRef}
                  width={300}
                  height={140}
                  className="touch-none bg-white rounded-lg cursor-crosshair"
                  onPointerDown={(e) => {
                    const canvas = signCanvasRef.current;
                    if (!canvas) return;
                    canvas.setPointerCapture(e.pointerId);
                    const ctx = canvas.getContext('2d');
                    if (!ctx) return;
                    const r = canvas.getBoundingClientRect();
                    ctx.beginPath();
                    ctx.moveTo(e.clientX - r.left, e.clientY - r.top);
                  }}
                  onPointerMove={(e) => {
                    if (e.buttons !== 1) return;
                    const canvas = signCanvasRef.current;
                    if (!canvas) return;
                    const ctx = canvas.getContext('2d');
                    if (!ctx) return;
                    const r = canvas.getBoundingClientRect();
                    ctx.lineWidth = 3;
                    ctx.lineCap = 'round';
                    ctx.strokeStyle = '#dc2626';
                    ctx.lineTo(e.clientX - r.left, e.clientY - r.top);
                    ctx.stroke();
                  }}
                />
              </div>
              <div className="flex items-center justify-between pt-1">
                <button
                  type="button"
                  onClick={() => {
                    const canvas = signCanvasRef.current;
                    if (canvas) {
                      const ctx = canvas.getContext('2d');
                      if (ctx) ctx.clearRect(0, 0, canvas.width, canvas.height);
                    }
                  }}
                  className="text-xs text-slate-400 hover:text-rose-400 flex items-center gap-1"
                >
                  <Trash2 size={12} /> 다시 쓰기
                </button>
                <button
                  type="button"
                  onClick={handleInsertSignature}
                  className="px-3 py-1 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-bold"
                >
                  서명 삽입
                </button>
              </div>
            </div>

            {/* 프리셋 인장 */}
            <div className="border-t border-slate-800 pt-3 space-y-2">
              <span className="text-xs text-slate-400 block">원클릭 인장 날인:</span>
              <div className="grid grid-cols-2 gap-2">
                <button
                  type="button"
                  onClick={() => handleInsertPresetStamp('서명')}
                  className="py-2 bg-rose-600/20 hover:bg-rose-600/30 text-rose-300 border border-rose-500/40 rounded-xl text-xs font-bold"
                >
                  🔴 [서명] 인장
                </button>
                <button
                  type="button"
                  onClick={() => handleInsertPresetStamp('확인')}
                  className="py-2 bg-blue-600/20 hover:bg-blue-600/30 text-blue-300 border border-blue-500/40 rounded-xl text-xs font-bold"
                >
                  🔵 [확인] 인장
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* 7. 그림판 스타일 색 편집 모달 */}
      {showColorEditModal && (
        <ColorEditModal
          isOpen={showColorEditModal}
          initialColor={
            colorEditTarget === 'pen'
              ? penColor
              : colorEditTarget === 'highlighter'
              ? highlighterColor
              : colorEditTarget === 'shape'
              ? shapeColor
              : colorEditTarget === 'text'
              ? textColor
              : textBoxBorderColor
          }
          onSelectColor={(hex) => {
            if (colorEditTarget === 'pen') setPenColor(hex);
            else if (colorEditTarget === 'highlighter') setHighlighterColor(hex);
            else if (colorEditTarget === 'shape') setShapeColor(hex);
            else if (colorEditTarget === 'text') setTextColor(hex);
            else if (colorEditTarget === 'boxBorder') setTextBoxBorderColor(hex);
          }}
          onClose={() => setShowColorEditModal(false)}
        />
      )}
    </div>
  );
};
