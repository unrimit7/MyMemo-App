import React, { useState, useEffect, useRef } from 'react';
import {
  X,
  Edit3,
  Download,
  FileDown,
  ChevronLeft,
  ChevronRight,
  Maximize2,
  Minimize2,
  Presentation,
  Loader2,
  Sidebar,
  ImageIcon,
  Layout,
  List
} from 'lucide-react';
import {
  parsePptxPresentation,
  type PptxPresentation,
  type PptxSlide,
  type PptxSlideElement,
} from '../utils/pptxViewer';
import { useFloatingWindow } from './useFloatingWindow';

interface PptxViewerModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: string;
  title: string;
  attId?: string;
  onOpenEditor?: () => void;
  onImportMarkdown?: (markdown: string) => void;
}

// 색상의 밝기 판별 함수 (글자 대비색 계산용)
function isDarkColor(colorStr: string): boolean {
  if (!colorStr) return false;
  let hex = colorStr.replace('#', '');
  if (hex.length === 3) {
    hex = hex.split('').map((c) => c + c).join('');
  }
  if (hex.length !== 6) return false;
  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);
  const brightness = (r * 299 + g * 587 + b * 114) / 1000;
  return brightness < 140;
}

export const PptxViewerModal: React.FC<PptxViewerModalProps> = ({
  isOpen,
  onClose,
  data,
  title,
  onOpenEditor,
  onImportMarkdown,
}) => {
  const [presentation, setPresentation] = useState<PptxPresentation | null>(null);
  const [currentSlideIndex, setCurrentSlideIndex] = useState<number>(0);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  // 뷰어 모드 및 스타일 상태
  const [showSidebar, setShowSidebar] = useState<boolean>(true);
  const [isSlideShow, setIsSlideShow] = useState<boolean>(false);
  const [isDarkMode, setIsDarkMode] = useState<boolean>(false);
  const [renderMode, setRenderMode] = useState<'canvas' | 'outline'>('canvas'); // 'canvas': 2D 정밀 좌표 레이아웃, 'outline': 텍스트 개요 뷰
  const [lightboxImage, setLightboxImage] = useState<string | null>(null);

  const containerRef = useRef<HTMLDivElement>(null);
  const floating = useFloatingWindow({ id: 'pptx-viewer', defaultWidth: 940, defaultHeight: 820 });

  // PPTX 데이터 로드
  useEffect(() => {
    if (!isOpen || !data) return;

    let isMounted = true;
    setLoading(true);
    setError(null);

    parsePptxPresentation(data)
      .then((pres) => {
        if (!isMounted) return;
        setPresentation(pres);
        setCurrentSlideIndex(0);
        setLoading(false);
      })
      .catch((err) => {
        if (!isMounted) return;
        console.error('파워포인트 파싱 실패:', err);
        setError('파워포인트(.pptx) 파일을 불러오거나 파싱하지 못했습니다.');
        setLoading(false);
      });

    return () => {
      isMounted = false;
    };
  }, [isOpen, data]);

  // 키보드 네비게이션
  useEffect(() => {
    if (!isOpen) return;

    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') {
        e.preventDefault();
        setCurrentSlideIndex((prev) => {
          if (!presentation) return prev;
          return Math.min(presentation.slides.length - 1, prev + 1);
        });
      } else if (e.key === 'ArrowLeft' || e.key === 'PageUp' || e.key === 'Backspace') {
        e.preventDefault();
        setCurrentSlideIndex((prev) => Math.max(0, prev - 1));
      } else if (e.key === 'Home') {
        e.preventDefault();
        setCurrentSlideIndex(0);
      } else if (e.key === 'End') {
        e.preventDefault();
        if (presentation) {
          setCurrentSlideIndex(presentation.slides.length - 1);
        }
      } else if (e.key === 'F5') {
        e.preventDefault();
        setIsSlideShow((prev) => !prev);
      } else if (e.key === 'Escape') {
        if (lightboxImage) {
          e.preventDefault();
          setLightboxImage(null);
        } else if (isSlideShow) {
          e.preventDefault();
          setIsSlideShow(false);
        } else {
          onClose();
        }
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, isSlideShow, presentation, lightboxImage, onClose]);

  // 메모로 가져오기 핸들러
  const handleImport = () => {
    if (!presentation || !onImportMarkdown) return;

    const mdLines: string[] = [];
    mdLines.push(`### 📊 ${presentation.title || title}\n`);

    for (const s of presentation.slides) {
      mdLines.push(`#### 슬라이드 ${s.slideNumber}: ${s.title}`);
      for (const p of s.paragraphs) {
        if (p.isTitle && p.text === s.title) continue;
        const indent = '  '.repeat(Math.max(0, p.level));
        mdLines.push(`${indent}- ${p.text}`);
      }
      for (const t of s.tables) {
        if (t.headers.length > 0) {
          mdLines.push('\n| ' + t.headers.join(' | ') + ' |');
          mdLines.push('| ' + t.headers.map(() => '---').join(' | ') + ' |');
          for (const row of t.rows) {
            mdLines.push('| ' + row.join(' | ') + ' |');
          }
          mdLines.push('');
        }
      }
      mdLines.push('');
    }

    onImportMarkdown(mdLines.join('\n'));
  };

  // 다운로드
  const handleDownload = () => {
    const a = document.createElement('a');
    a.href = data;
    a.download = title.endsWith('.pptx') ? title : `${title}.pptx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  if (!isOpen) return null;

  const currentSlide: PptxSlide | undefined = presentation?.slides[currentSlideIndex];
  const totalSlides = presentation?.slides.length || 0;
  const slideAspect = presentation?.aspectRatio || 16 / 9;

  // 2D 캔버스 요소 렌더러 (한쇼/파워포인트 원본 레이아웃 복원)
  const renderCanvasElement = (el: PptxSlideElement, isDark: boolean) => {
    if (el.type === 'shape') {
      // 1. 연결선 / 구분선 렌더링
      if (el.shapeType === 'line') {
        const isHoriz = el.height <= 0.8;
        return (
          <div
            key={el.id}
            style={{
              position: 'absolute',
              left: `${el.left}%`,
              top: `${el.top}%`,
              width: `${Math.max(0.1, el.width)}%`,
              height: isHoriz ? `${Math.max(1, el.borderWidth || 1)}px` : `${Math.max(0.1, el.height)}%`,
              backgroundColor: el.borderColor || el.fillColor || (isDark ? '#64748B' : '#334155'),
              zIndex: el.zIndex,
            }}
          />
        );
      }

      const hasBg = Boolean(el.fillColor && el.fillColor !== 'transparent');
      const isDarkBg = hasBg && isDarkColor(el.fillColor!);

      return (
        <div
          key={el.id}
          style={{
            position: 'absolute',
            left: `${el.left}%`,
            top: `${el.top}%`,
            width: `${el.width}%`,
            height: `${el.height}%`,
            backgroundColor: el.fillColor || 'transparent',
            border: el.borderColor ? `${el.borderWidth || 1}px solid ${el.borderColor}` : undefined,
            borderRadius: el.isRounded ? '16px' : el.shapeType === 'ellipse' ? '9999px' : '3px',
            zIndex: el.zIndex,
            overflow: 'hidden',
            display: 'flex',
            flexDirection: 'column',
            justifyContent:
              el.verticalAlign === 'middle' ? 'center' : el.verticalAlign === 'bottom' ? 'flex-end' : 'flex-start',
            padding: hasBg ? '8px 12px' : '2px 4px',
            boxSizing: 'border-box',
            transform: el.rotation ? `rotate(${el.rotation}deg)` : undefined,
            transformOrigin: 'center center',
            writingMode: el.isVerticalText ? 'vertical-rl' : undefined,
            columnCount: el.columnCount && el.columnCount > 1 ? el.columnCount : undefined,
            columnGap: el.columnCount && el.columnCount > 1 ? '16px' : undefined,
          }}
          className="pointer-events-auto leading-normal select-text shadow-xs"
        >
          {el.paragraphs.map((p, pIdx) => {
            const textAlign: any = p.align || 'left';
            const isBlank = !p.text || !p.text.trim();
            const indentPx =
              p.level > 0
                ? p.level * 16
                : p.marL && p.marL > 0 && (!p.indent || p.indent >= 0)
                ? Math.min(48, Math.round(p.marL / 12700))
                : undefined;

            if (isBlank) {
              return (
                <div
                  key={pIdx}
                  style={{ minHeight: '1.2em', whiteSpace: 'pre-wrap' }}
                  className="leading-snug select-none"
                >
                  &nbsp;
                </div>
              );
            }

            return (
              <div
                key={pIdx}
                style={{
                  textAlign,
                  whiteSpace: 'pre-wrap',
                  paddingLeft: indentPx ? `${indentPx}px` : undefined,
                }}
                className={`leading-snug ${el.isTitle ? 'font-extrabold' : ''}`}
              >
                {p.bulletChar && (
                  <span className="mr-1 text-slate-500 font-sans select-none">{p.bulletChar}</span>
                )}
                {p.runs.map((r, rIdx) => {
                  const textColor =
                    r.color ||
                    (hasBg ? (isDarkBg ? '#FFFFFF' : '#0F172A') : isDark ? '#F1F5F9' : '#0F172A');
                  const fontSize = r.fontSize
                    ? `${Math.max(10, Math.min(32, r.fontSize * 0.8))}px`
                    : el.isTitle
                    ? '24px'
                    : '12px';

                  return (
                    <span
                      key={rIdx}
                      style={{
                        color: textColor,
                        backgroundColor: r.highlightColor || undefined,
                        fontSize,
                        fontWeight: r.isBold || el.isTitle ? 'bold' : 'normal',
                        fontStyle: r.isItalic ? 'italic' : 'normal',
                        textDecoration: r.isStrike
                          ? 'line-through'
                          : r.isUnderline
                          ? 'underline'
                          : 'none',
                        verticalAlign:
                          r.baseline && r.baseline > 0
                            ? 'super'
                            : r.baseline && r.baseline < 0
                            ? 'sub'
                            : undefined,
                      }}
                      className={r.hyperlink ? 'underline text-blue-600 hover:text-blue-800 cursor-pointer' : undefined}
                      title={r.hyperlink ? `링크: ${r.hyperlink}` : undefined}
                    >
                      {r.text}
                    </span>
                  );
                })}
              </div>
            );
          })}
        </div>
      );
    }

    if (el.type === 'table') {
      return (
        <div
          key={el.id}
          style={{
            position: 'absolute',
            left: `${el.left}%`,
            top: `${el.top}%`,
            width: `${el.width}%`,
            height: `${el.height}%`,
            zIndex: el.zIndex,
            overflow: 'hidden',
          }}
          className="pointer-events-auto rounded-lg shadow-xs"
        >
          <table className="w-full h-full border-collapse text-xs table-fixed">
            <colgroup>
              {el.colWidths.map((w, wIdx) => (
                <col key={wIdx} style={{ width: `${w}%` }} />
              ))}
            </colgroup>
            <tbody>
              {el.rows.map((row, rIdx) => (
                <tr key={rIdx} style={row.height ? { height: `${row.height}%` } : undefined}>
                  {row.cells.map((cell, cIdx) => (
                    <td
                      key={cIdx}
                      colSpan={cell.colSpan}
                      rowSpan={cell.rowSpan}
                      style={{
                        backgroundColor:
                          cell.fillColor ||
                          (rIdx === 0
                            ? isDark
                              ? '#1E293B'
                              : '#F8FAFC'
                            : isDark
                            ? '#0F172A'
                            : '#FFFFFF'),
                        border: `1px solid ${isDark ? '#334155' : '#CBD5E1'}`,
                        verticalAlign: cell.verticalAlign || 'middle',
                        padding: '4px 6px',
                      }}
                      className="leading-snug overflow-hidden select-text"
                    >
                      {cell.paragraphs.map((p, pIdx) => {
                        const isBlank = !p.text || !p.text.trim();
                        const indentPx =
                          p.level > 0
                            ? p.level * 16
                            : p.marL && p.marL > 0 && (!p.indent || p.indent >= 0)
                            ? Math.min(48, Math.round(p.marL / 12700))
                            : undefined;

                        if (isBlank) {
                          return (
                            <div
                              key={pIdx}
                              style={{ minHeight: '1.1em', whiteSpace: 'pre-wrap' }}
                              className="mb-0.5 last:mb-0 select-none"
                            >
                              &nbsp;
                            </div>
                          );
                        }

                        return (
                          <div
                            key={pIdx}
                            style={{
                              textAlign: p.align || 'left',
                              whiteSpace: 'pre-wrap',
                              paddingLeft: indentPx ? `${indentPx}px` : undefined,
                            }}
                            className="mb-0.5 last:mb-0"
                          >
                            {p.runs.map((r, rIdx) => {
                              const textColor = r.color || (isDark ? '#E2E8F0' : '#0F172A');
                              const fontSize = r.fontSize
                                ? `${Math.max(9, Math.min(18, r.fontSize * 0.7))}px`
                                : '11px';

                              return (
                                <span
                                  key={rIdx}
                                  style={{
                                    color: textColor,
                                    fontSize,
                                    fontWeight: r.isBold ? 'bold' : 'normal',
                                    fontStyle: r.isItalic ? 'italic' : 'normal',
                                  }}
                                >
                                  {r.text}
                                </span>
                              );
                            })}
                          </div>
                        );
                      })}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      );
    }

    if (el.type === 'image') {
      const isFullBackground = el.width >= 90 && el.height >= 90;

      return (
        <div
          key={el.id}
          style={{
            position: 'absolute',
            left: `${el.left}%`,
            top: `${el.top}%`,
            width: `${el.width}%`,
            height: `${el.height}%`,
            zIndex: el.zIndex,
          }}
          className={`flex items-center justify-center ${
            isFullBackground ? 'pointer-events-none' : 'pointer-events-auto'
          }`}
        >
          <img
            src={el.dataUrl}
            alt="Slide graphic"
            onClick={() => !isFullBackground && setLightboxImage(el.dataUrl)}
            style={{
              width: '100%',
              height: '100%',
              objectFit: isFullBackground ? 'cover' : 'contain',
            }}
            className={isFullBackground ? '' : 'cursor-zoom-in hover:opacity-95 transition-opacity'}
            title={isFullBackground ? undefined : '클릭하여 원본 고해상도로 확대 보기'}
          />
        </div>
      );
    }

    if (el.type === 'chart') {
      const palette = ['#3B82F6', '#F59E0B', '#10B981', '#EC4899', '#8B5CF6', '#06B6D4'];
      const { chartType, title: chartTitle, categories, series } = el;

      let maxVal = 1;
      for (const s of series) {
        for (const v of s.values) {
          if (v > maxVal) maxVal = v;
        }
      }

      return (
        <div
          key={el.id}
          style={{
            position: 'absolute',
            left: `${el.left}%`,
            top: `${el.top}%`,
            width: `${el.width}%`,
            height: `${el.height}%`,
            zIndex: el.zIndex,
            overflow: 'hidden',
          }}
          className="bg-white/95 dark:bg-slate-800/95 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 p-2 flex flex-col pointer-events-auto"
        >
          {chartTitle && (
            <div className="text-xs font-bold text-slate-800 dark:text-slate-100 text-center mb-1 truncate">
              {chartTitle}
            </div>
          )}

          <div className="flex-1 w-full h-full min-h-0 relative">
            {chartType === 'bar' && (
              <svg viewBox="0 0 300 180" className="w-full h-full" preserveAspectRatio="none">
                <line x1="30" y1="20" x2="290" y2="20" stroke="#CBD5E1" strokeDasharray="2,2" strokeWidth="0.5" />
                <line x1="30" y1="80" x2="290" y2="80" stroke="#CBD5E1" strokeDasharray="2,2" strokeWidth="0.5" />
                <line x1="30" y1="140" x2="290" y2="140" stroke="#64748B" strokeWidth="1" />
                {categories.map((cat, cIdx) => {
                  const groupWidth = 260 / Math.max(1, categories.length);
                  const groupX = 30 + cIdx * groupWidth;
                  const barWidth = Math.max(4, (groupWidth * 0.7) / Math.max(1, series.length));

                  return (
                    <g key={cIdx}>
                      {series.map((s, sIdx) => {
                        const val = s.values[cIdx] || 0;
                        const barH = maxVal > 0 ? (val / maxVal) * 115 : 0;
                        const bx = groupX + 5 + sIdx * barWidth;
                        const by = 140 - barH;
                        const color = palette[sIdx % palette.length];

                        return (
                          <rect
                            key={sIdx}
                            x={bx}
                            y={by}
                            width={barWidth - 1}
                            height={Math.max(1, barH)}
                            fill={color}
                            rx="1"
                          >
                            <title>{`${s.name || `시리즈 ${sIdx + 1}`}: ${val}`}</title>
                          </rect>
                        );
                      })}
                      <text
                        x={groupX + groupWidth / 2}
                        y="155"
                        textAnchor="middle"
                        fontSize="8"
                        fill="#64748B"
                      >
                        {cat}
                      </text>
                    </g>
                  );
                })}
              </svg>
            )}

            {chartType === 'line' && (
              <svg viewBox="0 0 300 180" className="w-full h-full" preserveAspectRatio="none">
                <line x1="30" y1="20" x2="290" y2="20" stroke="#CBD5E1" strokeDasharray="2,2" strokeWidth="0.5" />
                <line x1="30" y1="80" x2="290" y2="80" stroke="#CBD5E1" strokeDasharray="2,2" strokeWidth="0.5" />
                <line x1="30" y1="140" x2="290" y2="140" stroke="#64748B" strokeWidth="1" />
                {series.map((s, sIdx) => {
                  const color = palette[sIdx % palette.length];
                  const points = s.values
                    .map((v, vIdx) => {
                      const step = categories.length > 1 ? 260 / (categories.length - 1) : 130;
                      const x = 30 + vIdx * step;
                      const y = 140 - (maxVal > 0 ? (v / maxVal) * 115 : 0);
                      return `${x},${y}`;
                    })
                    .join(' ');

                  return (
                    <g key={sIdx}>
                      <polyline
                        fill="none"
                        stroke={color}
                        strokeWidth="2.5"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        points={points}
                      />
                      {s.values.map((v, vIdx) => {
                        const step = categories.length > 1 ? 260 / (categories.length - 1) : 130;
                        const x = 30 + vIdx * step;
                        const y = 140 - (maxVal > 0 ? (v / maxVal) * 115 : 0);
                        return (
                          <circle key={vIdx} cx={x} cy={y} r="3" fill={color}>
                            <title>{`${categories[vIdx] || ''}: ${v}`}</title>
                          </circle>
                        );
                      })}
                    </g>
                  );
                })}
                {categories.map((cat, cIdx) => {
                  const step = categories.length > 1 ? 260 / (categories.length - 1) : 130;
                  const x = 30 + cIdx * step;
                  return (
                    <text key={cIdx} x={x} y="155" textAnchor="middle" fontSize="8" fill="#64748B">
                      {cat}
                    </text>
                  );
                })}
              </svg>
            )}

            {chartType === 'pie' && (
              <svg viewBox="0 0 200 180" className="w-full h-full">
                {(() => {
                  const s = series[0];
                  if (!s || s.values.length === 0) return null;
                  const total = s.values.reduce((a, b) => a + b, 0) || 1;
                  let cumulativeAngle = 0;
                  const cx = 100;
                  const cy = 90;
                  const r = 65;

                  return s.values.map((val, idx) => {
                    const angle = (val / total) * 360;
                    const startAngle = cumulativeAngle;
                    const endAngle = cumulativeAngle + angle;
                    cumulativeAngle = endAngle;

                    const startRad = (startAngle - 90) * (Math.PI / 180);
                    const endRad = (endAngle - 90) * (Math.PI / 180);

                    const x1 = cx + r * Math.cos(startRad);
                    const y1 = cy + r * Math.sin(startRad);
                    const x2 = cx + r * Math.cos(endRad);
                    const y2 = cy + r * Math.sin(endRad);

                    const largeArc = angle > 180 ? 1 : 0;
                    const d = `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${largeArc} 1 ${x2} ${y2} Z`;
                    const color = palette[idx % palette.length];

                    return (
                      <path key={idx} d={d} fill={color} stroke="#FFFFFF" strokeWidth="1">
                        <title>{`${categories[idx] || `항목 ${idx + 1}`}: ${val} (${Math.round((val / total) * 100)}%)`}</title>
                      </path>
                    );
                  });
                })()}
              </svg>
            )}
          </div>

          <div className="flex flex-wrap items-center justify-center gap-2 mt-1 shrink-0 text-[9px] text-slate-600 dark:text-slate-300">
            {chartType === 'pie'
              ? categories.map((cat, idx) => (
                  <div key={idx} className="flex items-center gap-1">
                    <span className="w-2 h-2 rounded-xs inline-block" style={{ backgroundColor: palette[idx % palette.length] }} />
                    <span>{cat}</span>
                  </div>
                ))
              : series.map((s, idx) => (
                  <div key={idx} className="flex items-center gap-1">
                    <span className="w-2 h-2 rounded-xs inline-block" style={{ backgroundColor: palette[idx % palette.length] }} />
                    <span>{s.name || `시리즈 ${idx + 1}`}</span>
                  </div>
                ))}
          </div>
        </div>
      );
    }

    return null;
  };

  return (
    <div
      ref={containerRef}
      className={isSlideShow ? 'fixed inset-0 z-50 flex flex-col bg-black/95 select-none p-1 sm:p-2' : floating.modalWrapperClass}
    >
      <div
        {...(isSlideShow ? {} : floating.windowContainerProps)}
        className={isSlideShow ? 'w-full h-full flex flex-col bg-slate-900 overflow-hidden' : floating.windowContainerProps.className}
      >
        {/* 상단 툴바 (슬라이드 쇼 모드가 아닐 때) */}
        {!isSlideShow ? (
          <div 
            {...floating.headerProps}
            className="px-4 py-2.5 bg-slate-900 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3 shrink-0"
          >
            {/* 좌측: 타이틀 및 슬라이드 카운터 */}
            <div className="flex items-center gap-3 min-w-0">
              <span className="p-1.5 rounded-lg bg-orange-500/15 text-orange-400 border border-orange-500/30">
                <Presentation size={18} />
              </span>
              <div className="truncate">
                <div className="flex items-center gap-2">
                  <h2 className="text-sm md:text-base font-bold text-slate-100 truncate" title={title}>
                    {title}
                  </h2>
                  <span className="text-[10px] px-1.5 py-0.5 rounded bg-orange-500/20 text-orange-300 border border-orange-500/40 font-bold uppercase shrink-0">
                    PPTX
                  </span>
                  {totalSlides > 0 && (
                    <span className="text-xs text-slate-400 font-mono hidden sm:inline">
                      ({currentSlideIndex + 1} / {totalSlides} 슬라이드)
                    </span>
                  )}
                </div>
              </div>
            </div>

            {/* 우측: 액션 버튼들 */}
            <div className="flex items-center gap-2 shrink-0">
              {/* 직접 편집 버튼 */}
              {onOpenEditor && (
                <button
                  type="button"
                  onClick={onOpenEditor}
                  className="px-3 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-semibold shadow-md flex items-center gap-1.5 transition-all active:scale-95 cursor-pointer"
                  title="웹에서 슬라이드 내용 직접 수정하기"
                >
                  <Edit3 size={13} />
                  <span>직접 편집</span>
                </button>
              )}

              {/* 슬라이드 쇼 (전체화면) 버튼 */}
              <button
                type="button"
                onClick={() => setIsSlideShow(true)}
                className="px-3 py-1.5 bg-orange-600/30 hover:bg-orange-600 text-orange-200 hover:text-white rounded-lg text-xs font-semibold border border-orange-500/40 flex items-center gap-1.5 transition-all active:scale-95 cursor-pointer"
                title="발표용 슬라이드 쇼 모드 (F5)"
              >
                <Maximize2 size={13} />
                <span>슬라이드 쇼</span>
              </button>

              {/* 메모로 가져오기 버튼 */}
              {onImportMarkdown && (
                <button
                  type="button"
                  onClick={handleImport}
                  className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-750 text-indigo-300 hover:text-indigo-200 rounded-lg text-xs font-medium border border-slate-700 flex items-center gap-1.5 transition-colors cursor-pointer"
                  title="슬라이드 본문 텍스트를 메모에 마크다운으로 삽입"
                >
                  <FileDown size={13} />
                  <span className="hidden sm:inline">메모로 가져오기</span>
                </button>
              )}

              {/* 다운로드 버튼 */}
              <button
                type="button"
                onClick={handleDownload}
                className="p-1.5 bg-slate-800 hover:bg-slate-750 text-slate-300 hover:text-white rounded-lg border border-slate-700 transition-colors cursor-pointer"
                title="파워포인트 파일 다운로드"
              >
                <Download size={15} />
              </button>

              {floating.toggleFloatingButton}

              {/* 닫기 버튼 */}
              <button
                type="button"
                onClick={onClose}
                className="p-1.5 bg-slate-800 hover:bg-rose-900/40 text-slate-400 hover:text-rose-300 rounded-lg border border-slate-700 transition-colors cursor-pointer"
                title="닫기 (Esc)"
              >
                <X size={16} />
              </button>
            </div>
          </div>
        ) : (
          /* 슬라이드 쇼 상단 플로팅 미니 바 */
          <div className="absolute top-3 right-4 z-50 flex items-center gap-2 bg-slate-900/90 backdrop-blur-md px-3.5 py-1.5 rounded-full border border-slate-700/60 shadow-xl text-xs text-slate-300">
            <button
              type="button"
              onClick={() => setCurrentSlideIndex((prev) => Math.max(0, prev - 1))}
              disabled={currentSlideIndex <= 0}
              className="hover:text-white disabled:opacity-20 cursor-pointer"
              title="이전 슬라이드 (←)"
            >
              <ChevronLeft size={16} />
            </button>

            <span className="font-mono text-orange-400 font-bold px-1">
              {currentSlideIndex + 1} / {totalSlides}
            </span>

            <button
              type="button"
              onClick={() => setCurrentSlideIndex((prev) => Math.min(totalSlides - 1, prev + 1))}
              disabled={currentSlideIndex >= totalSlides - 1}
              className="hover:text-white disabled:opacity-20 cursor-pointer"
              title="다음 슬라이드 (→)"
            >
              <ChevronRight size={16} />
            </button>

            <span className="text-slate-600">|</span>

            {/* 레이아웃 전환: 정밀 캔버스 vs 텍스트 개요 */}
            <button
              type="button"
              onClick={() => setRenderMode((prev) => (prev === 'canvas' ? 'outline' : 'canvas'))}
              className="px-2 py-0.5 rounded bg-slate-800 hover:bg-slate-700 text-slate-300 text-[11px] cursor-pointer"
              title="원본 캔버스 레이아웃 / 텍스트 개요 뷰 전환"
            >
              {renderMode === 'canvas' ? '📐 원본 레이아웃' : '📋 텍스트 개요'}
            </button>

            {/* 배경 테마 토글 */}
            <button
              type="button"
              onClick={() => setIsDarkMode((prev) => !prev)}
              className="px-2 py-0.5 rounded bg-slate-800 hover:bg-slate-700 text-slate-300 text-[11px] cursor-pointer"
              title="배경 테마 전환"
            >
              {isDarkMode ? '다크' : '화이트'}
            </button>

            <span className="text-slate-600">|</span>

            <button
              type="button"
              onClick={() => setIsSlideShow(false)}
              className="text-slate-300 hover:text-white flex items-center gap-1 font-medium cursor-pointer"
              title="슬라이드 쇼 종료 (Esc)"
            >
              <Minimize2 size={13} />
              <span>종료</span>
            </button>
          </div>
        )}

        {/* 메인 뷰: 좌측 썸네일 탐색기 + 중앙 슬라이드 스테이지 */}
        <div className="flex-1 flex overflow-hidden min-h-0 relative">
          {/* 로딩 / 에러 상태 */}
          {loading && (
            <div className="absolute inset-0 z-30 flex flex-col items-center justify-center bg-slate-950/80 gap-3">
              <Loader2 size={32} className="animate-spin text-orange-400" />
              <span className="text-xs text-slate-400 font-medium">파워포인트 슬라이드 파싱 중...</span>
            </div>
          )}

          {error && (
            <div className="absolute inset-0 z-30 flex flex-col items-center justify-center bg-slate-950 p-6 text-center">
              <span className="text-3xl mb-2">⚠️</span>
              <p className="text-rose-400 font-semibold text-sm mb-4">{error}</p>
              <button
                type="button"
                onClick={handleDownload}
                className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-white rounded-lg text-xs font-semibold border border-slate-700 flex items-center gap-2"
              >
                <Download size={14} />
                <span>기기로 다운로드하여 열기</span>
              </button>
            </div>
          )}

          {/* 좌측 슬라이드 썸네일 바 */}
          {!isSlideShow && showSidebar && presentation && (
            <div className="w-48 sm:w-56 bg-slate-950 border-r border-slate-800 flex flex-col shrink-0 overflow-y-auto p-2.5 gap-2.5 select-none">
              <div className="flex items-center justify-between px-1 text-[11px] font-bold text-slate-400 uppercase tracking-wider">
                <span>슬라이드 목록 ({totalSlides})</span>
              </div>

              {presentation.slides.map((s, idx) => {
                const isActive = idx === currentSlideIndex;
                return (
                  <div
                    key={s.slideNumber}
                    onClick={() => setCurrentSlideIndex(idx)}
                    className={`p-2 rounded-xl border transition-all cursor-pointer flex flex-col gap-1 ${
                      isActive
                        ? 'bg-orange-950/30 border-orange-500/60 ring-2 ring-orange-500/40'
                        : 'bg-slate-900/60 border-slate-800 hover:bg-slate-850 hover:border-slate-700'
                    }`}
                  >
                    <div className="flex items-center justify-between text-[11px] font-mono">
                      <div className="flex items-center gap-1">
                        <span className={`px-1.5 py-0.2 rounded font-bold ${isActive ? 'bg-orange-500 text-white' : 'bg-slate-800 text-slate-400'}`}>
                          #{s.slideNumber}
                        </span>
                        {s.isHidden && (
                          <span className="px-1 py-0.2 rounded text-[9px] bg-rose-900/60 text-rose-300 border border-rose-700/50 font-sans">
                            숨김
                          </span>
                        )}
                      </div>
                      {s.images.length > 0 && (
                        <span className="text-slate-500 flex items-center gap-0.5 text-[10px]">
                          <ImageIcon size={10} /> {s.images.length}
                        </span>
                      )}
                    </div>
                    {/* 미니 16:9 썸네일 카드 */}
                    <div className="w-full aspect-video bg-slate-950 rounded border border-slate-800/80 p-1.5 flex flex-col justify-center overflow-hidden">
                      <span className="text-[10px] font-semibold text-slate-200 line-clamp-2 leading-tight">
                        {s.title}
                      </span>
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {/* 중앙 슬라이드 스테이지 */}
          <div className="flex-1 flex flex-col bg-slate-950 overflow-hidden relative">
            {/* 상단 뷰어 옵션 바 (사이드바 토글, 레이아웃 모드 토글, 테마 토글) */}
            {!isSlideShow && (
              <div className="px-3 py-1.5 bg-slate-900/60 border-b border-slate-800/80 flex items-center justify-between text-xs text-slate-400 shrink-0">
                <button
                  type="button"
                  onClick={() => setShowSidebar((prev) => !prev)}
                  className="flex items-center gap-1.5 hover:text-slate-200 cursor-pointer"
                  title="슬라이드 썸네일 목록 보이기/숨기기"
                >
                  <Sidebar size={13} />
                  <span>{showSidebar ? '목록 숨기기' : '슬라이드 목록'}</span>
                </button>

                <div className="flex items-center gap-2">
                  {/* 원본 캔버스 레이아웃 vs 텍스트 개요 뷰 전환 */}
                  <div className="flex items-center bg-slate-950 p-0.5 rounded-lg border border-slate-800">
                    <button
                      type="button"
                      onClick={() => setRenderMode('canvas')}
                      className={`px-2 py-0.5 rounded text-[11px] flex items-center gap-1 cursor-pointer transition-colors ${
                        renderMode === 'canvas'
                          ? 'bg-orange-600 text-white font-semibold'
                          : 'text-slate-400 hover:text-slate-200'
                      }`}
                      title="한쇼/파워포인트 원본 2차원 그래픽 캔버스 레이아웃"
                    >
                      <Layout size={12} />
                      <span>원본 레이아웃</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setRenderMode('outline')}
                      className={`px-2 py-0.5 rounded text-[11px] flex items-center gap-1 cursor-pointer transition-colors ${
                        renderMode === 'outline'
                          ? 'bg-orange-600 text-white font-semibold'
                          : 'text-slate-400 hover:text-slate-200'
                      }`}
                      title="단락 및 글머리 기호 텍스트 개요 뷰"
                    >
                      <List size={12} />
                      <span>텍스트 개요</span>
                    </button>
                  </div>

                  <button
                    type="button"
                    onClick={() => setIsDarkMode((prev) => !prev)}
                    className="px-2 py-0.5 rounded bg-slate-800 hover:bg-slate-750 text-slate-300 text-[11px] border border-slate-700 cursor-pointer"
                    title="슬라이드 배경 테마 전환"
                  >
                    {isDarkMode ? '다크 테마' : '화이트 테마'}
                  </button>
                  <span className="text-slate-500 hidden sm:inline">단축키: ← / → / Space</span>
                </div>
              </div>
            )}

            {/* 슬라이드 렌더링 영역 */}
            <div className="flex-1 flex items-center justify-center p-2 sm:p-4 overflow-auto">
              {currentSlide ? (
                renderMode === 'canvas' && currentSlide.elements && currentSlide.elements.length > 0 ? (
                  /* 1. 한쇼/파워포인트 원본 2차원 절대 좌표 캔버스 렌더링 */
                  <div className="flex flex-col items-center justify-center w-full">
                  <div
                    style={{
                      aspectRatio: `${slideAspect}`,
                      maxHeight: isSlideShow ? '96vh' : '82vh',
                      maxWidth: isSlideShow
                        ? `min(98vw, calc(96vh * ${slideAspect}))`
                        : `min(96vw, calc(82vh * ${slideAspect}))`,
                      width: isSlideShow
                        ? `min(98vw, calc(96vh * ${slideAspect}))`
                        : `min(96vw, calc(82vh * ${slideAspect}))`,
                      height: 'auto',
                      backgroundColor: currentSlide.backgroundColor || (isDarkMode ? '#0F172A' : '#FFFFFF'),
                    }}
                    className={`rounded-2xl shadow-2xl relative overflow-hidden select-text border transition-all ${
                      isDarkMode ? 'border-slate-800 text-slate-100' : 'border-slate-300 text-slate-900'
                    }`}
                  >
                    {/* 상단 얇은 슬라이드 액센트 바 */}
                    <div className="absolute top-0 left-0 right-0 h-1.5 bg-gradient-to-r from-orange-500 via-amber-500 to-indigo-600 z-30" />

                    {/* 숨김 슬라이드 배지 */}
                    {currentSlide.isHidden && (
                      <div className="absolute top-3 right-4 z-30 px-2.5 py-1 rounded-full bg-rose-600/90 text-white text-[11px] font-bold shadow-md flex items-center gap-1.5 backdrop-blur-xs">
                        <span>숨김 슬라이드</span>
                      </div>
                    )}

                    {/* 모든 2D 엘리먼트 (도형, 표, 이미지, 차트) 좌표 배치 */}
                    {currentSlide.elements.map((el) => renderCanvasElement(el, isDarkMode))}

                    {/* 하단 우측 슬라이드 번호 */}
                    <div className="absolute bottom-2 right-4 text-[10px] text-slate-400 font-mono z-20 pointer-events-none">
                      {currentSlide.slideNumber}
                    </div>
                  </div>
                  {currentSlide.notes && !isSlideShow && (
                    <div className="mt-3 p-3 bg-slate-900/90 border border-slate-700/80 rounded-xl text-xs text-slate-200 max-w-2xl w-full shadow-lg">
                      <div className="font-bold text-orange-400 mb-1 flex items-center gap-1.5">
                        <span>📝 발표자 메모</span>
                      </div>
                      <div className="whitespace-pre-wrap leading-relaxed text-slate-300">{currentSlide.notes}</div>
                    </div>
                  )}
                  </div>
                ) : (
                  /* 2. 텍스트 개요 뷰 (Fallback / 개요 읽기 모드) */
                  <div
                    style={{
                      aspectRatio: `${slideAspect}`,
                      maxHeight: isSlideShow ? '94vh' : '78vh',
                      maxWidth: isSlideShow
                        ? `min(98vw, calc(94vh * ${slideAspect}))`
                        : `min(95vw, calc(78vh * ${slideAspect}))`,
                      width: isSlideShow
                        ? `min(98vw, calc(94vh * ${slideAspect}))`
                        : `min(95vw, calc(78vh * ${slideAspect}))`,
                      height: 'auto',
                    }}
                    className={`rounded-2xl shadow-2xl flex flex-col border transition-all overflow-hidden relative ${
                      isDarkMode
                        ? 'bg-slate-900 border-slate-750 text-slate-100'
                        : 'bg-white border-slate-200 text-slate-900 shadow-slate-950/40'
                    }`}
                  >
                    <div className="h-1.5 bg-gradient-to-r from-orange-500 via-amber-500 to-indigo-600 shrink-0" />
                    <div className="flex-1 p-5 sm:p-7 flex flex-col justify-between overflow-y-auto">
                      <div>
                        <div className="border-b border-orange-500/30 pb-2 mb-3 flex items-center justify-between">
                          <h1 className="text-lg sm:text-xl font-extrabold tracking-tight truncate">
                            {currentSlide.title}
                          </h1>
                          <span className="text-xs font-mono font-bold px-2 py-0.5 rounded bg-orange-100 text-orange-700">
                            {currentSlide.slideNumber}
                          </span>
                        </div>

                        <div className="space-y-2">
                          {currentSlide.paragraphs.map((p, pIdx) => (
                            <div key={pIdx} className="flex items-start gap-2 text-xs sm:text-sm leading-relaxed">
                              <span className="text-orange-500 font-bold select-none">•</span>
                              <p>{p.text}</p>
                            </div>
                          ))}
                        </div>

                        {currentSlide.images.length > 0 && (
                          <div className="mt-3 flex flex-wrap gap-2.5 items-center justify-center">
                            {currentSlide.images.map((img, imgIdx) => (
                              <img
                                key={img.id || imgIdx}
                                src={img.dataUrl}
                                alt="Slide graphic"
                                onClick={() => setLightboxImage(img.dataUrl)}
                                className="max-h-28 rounded-lg border border-slate-300 shadow-xs object-contain cursor-zoom-in"
                              />
                            ))}
                          </div>
                        )}
                      </div>

                      <div className="pt-2 mt-auto flex items-center justify-between text-[10px] text-slate-400 select-none">
                        <span>{presentation?.title}</span>
                        <span>Slide {currentSlide.slideNumber}</span>
                      </div>
                    </div>
                  </div>
                )
              ) : null}
            </div>

            {/* 하단 슬라이드 네비게이션 컨트롤러 */}
            <div className="px-4 py-2.5 bg-slate-900 border-t border-slate-800 flex items-center justify-between shrink-0">
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setCurrentSlideIndex((prev) => Math.max(0, prev - 1))}
                  disabled={currentSlideIndex <= 0}
                  className="px-3 py-1.5 bg-slate-800 hover:bg-slate-750 disabled:opacity-30 text-white rounded-lg text-xs font-semibold flex items-center gap-1 transition-all cursor-pointer"
                  title="이전 슬라이드 (←, PageUp)"
                >
                  <ChevronLeft size={14} />
                  <span>이전</span>
                </button>

                <button
                  type="button"
                  onClick={() =>
                    setCurrentSlideIndex((prev) =>
                      Math.min(totalSlides - 1, prev + 1)
                    )
                  }
                  disabled={currentSlideIndex >= totalSlides - 1}
                  className="px-3 py-1.5 bg-slate-800 hover:bg-slate-750 disabled:opacity-30 text-white rounded-lg text-xs font-semibold flex items-center gap-1 transition-all cursor-pointer"
                  title="다음 슬라이드 (→, Space, PageDown)"
                >
                  <span>다음</span>
                  <ChevronRight size={14} />
                </button>
              </div>

              {/* 중앙 슬라이더 & 인디케이터 */}
              <div className="flex items-center gap-3">
                <input
                  type="range"
                  min={0}
                  max={Math.max(0, totalSlides - 1)}
                  value={currentSlideIndex}
                  onChange={(e) => setCurrentSlideIndex(parseInt(e.target.value, 10))}
                  className="w-28 sm:w-48 accent-orange-500 cursor-pointer"
                />
                <span className="text-xs font-mono font-bold text-orange-400">
                  {currentSlideIndex + 1} / {totalSlides}
                </span>
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setIsSlideShow((prev) => !prev)}
                  className="p-1.5 bg-slate-800 hover:bg-slate-750 text-slate-300 hover:text-white rounded-lg border border-slate-700 cursor-pointer"
                  title="슬라이드 쇼 모드 토글 (F5)"
                >
                  {isSlideShow ? <Minimize2 size={14} /> : <Maximize2 size={14} />}
                </button>
              </div>
            </div>
          </div>
        </div>

        {floating.resizeHandle}
      </div>

      {/* 고해상도 이미지 라이트박스 */}
      {lightboxImage && (
        <div
          className="fixed inset-0 z-[100] bg-black/95 backdrop-blur-md flex flex-col items-center justify-center p-4 select-none animate-fadeIn"
          onClick={() => setLightboxImage(null)}
        >
          <div
            className="relative max-w-5xl w-full max-h-[92vh] flex flex-col items-center"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="w-full flex items-center justify-between pb-2 text-xs text-slate-300">
              <span className="font-semibold text-orange-400 flex items-center gap-1.5">
                <ImageIcon size={14} />
                <span>슬라이드 포함 이미지 원본 고해상도 확대</span>
              </span>
              <div className="flex items-center gap-2">
                <a
                  href={lightboxImage}
                  download={`slide_image_${Date.now()}.png`}
                  className="px-2.5 py-1 bg-slate-800 hover:bg-slate-750 text-white rounded-lg border border-slate-700 flex items-center gap-1 text-xs no-underline cursor-pointer"
                  title="이미지 파일로 다운로드"
                >
                  <Download size={12} />
                  <span>다운로드</span>
                </a>
                <button
                  type="button"
                  onClick={() => setLightboxImage(null)}
                  className="p-1 bg-slate-800 hover:bg-rose-900/40 text-slate-400 hover:text-rose-300 rounded-lg border border-slate-700 cursor-pointer"
                  title="닫기 (Esc)"
                >
                  <X size={16} />
                </button>
              </div>
            </div>
            <img
              src={lightboxImage}
              alt="Expanded slide element"
              className="max-w-full max-h-[84vh] rounded-xl shadow-2xl object-contain border border-slate-700 bg-slate-900"
            />
          </div>
        </div>
      )}
    </div>
  );
};
