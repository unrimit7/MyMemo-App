// PPT 슬라이드 한 장을 원본 좌표대로 그리는 공통 렌더러 (큰 화면 뷰어와 메모 카드 첫 슬라이드 미리보기가 같이 사용)
import React from 'react';
import type { PptxSlide, PptxSlideElement, PptxParagraph, PptxTextRun } from '../utils/pptxViewer';

// 색상의 밝기 판별 함수 (글자 대비색 계산용)
export function isDarkColor(colorStr: string): boolean {
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

export function renderPptxSlideElement(el: PptxSlideElement, isDark: boolean, onImageClick?: (dataUrl: string) => void): React.ReactNode {
  if (el.type === 'shape') {
    // 1. 연결선 / 구분선 렌더링
    if (el.shapeType === 'line') {
      const isHoriz = el.height <= 0.8;
      const lineColor = el.borderColor || el.fillColor || (isDark ? '#64748B' : '#334155');
      const lineWidth = Math.max(1, el.borderWidth || 1);
      const arrowSize = Math.max(5, lineWidth * 3);

      return (
        <div
          key={el.id}
          style={{
            position: 'absolute',
            left: `${el.left}%`,
            top: `${el.top}%`,
            width: `${Math.max(0.1, el.width)}%`,
            height: isHoriz ? `${lineWidth}px` : `${Math.max(0.1, el.height)}%`,
            backgroundColor: lineColor,
            zIndex: el.zIndex,
            overflow: 'visible',
          }}
        >
          {isHoriz && el.tailEnd && (
            <div
              style={{
                position: 'absolute',
                right: -arrowSize,
                top: '50%',
                transform: 'translateY(-50%)',
                width: 0,
                height: 0,
                borderTop: `${arrowSize}px solid transparent`,
                borderBottom: `${arrowSize}px solid transparent`,
                borderLeft: `${Math.round(arrowSize * 1.5)}px solid ${lineColor}`,
              }}
            />
          )}
          {isHoriz && el.headEnd && (
            <div
              style={{
                position: 'absolute',
                left: -arrowSize,
                top: '50%',
                transform: 'translateY(-50%)',
                width: 0,
                height: 0,
                borderTop: `${arrowSize}px solid transparent`,
                borderBottom: `${arrowSize}px solid transparent`,
                borderRight: `${Math.round(arrowSize * 1.5)}px solid ${lineColor}`,
              }}
            />
          )}
        </div>
      );
    }

    const hasBg = Boolean(el.fillColor && el.fillColor !== 'transparent');
    const isDarkBg = hasBg && isDarkColor(el.fillColor!);
    const isRightArrow = el.shapeType === 'rightArrow';

    const renderParagraph = (p: PptxParagraph, pIdx: number) => {
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
          {p.runs.map((r: PptxTextRun, rIdx: number) => {
            const textColor =
              r.color ||
              (r.highlightColor ? '#000000' : hasBg ? (isDarkBg ? '#FFFFFF' : '#0F172A') : isDark ? '#F1F5F9' : '#0F172A');
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
                className={
                  r.hyperlink
                    ? 'underline text-blue-600 hover:text-blue-800 cursor-pointer'
                    : undefined
                }
                title={r.hyperlink ? `링크: ${r.hyperlink}` : undefined}
              >
                {r.text}
              </span>
            );
          })}
        </div>
      );
    };

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
          border:
            !isRightArrow && el.borderColor
              ? `${el.borderWidth || 1}px solid ${el.borderColor}`
              : undefined,
          borderRadius: isRightArrow
            ? undefined
            : el.shapeType === 'ellipse'
            ? '50%'
            : el.isRounded
            ? '16px'
            : '3px',
          clipPath: isRightArrow
            ? 'polygon(0% 25%, 65% 25%, 65% 0%, 100% 50%, 65% 100%, 65% 75%, 0% 75%)'
            : el.shapeType === 'ellipse'
            ? 'ellipse(50% 50% at 50% 50%)'
            : undefined,
          zIndex: el.zIndex,
          overflow: 'hidden',
          display: 'flex',
          flexDirection: 'column',
          justifyContent:
            el.verticalAlign === 'middle'
              ? 'center'
              : el.verticalAlign === 'bottom'
              ? 'flex-end'
              : 'flex-start',
          padding: hasBg ? (isRightArrow ? '8px 24px 8px 12px' : '8px 12px') : '2px 4px',
          boxSizing: 'border-box',
          transform: el.rotation ? `rotate(${el.rotation}deg)` : undefined,
          transformOrigin: 'center center',
          writingMode: el.isVerticalText ? 'vertical-rl' : undefined,
        }}
        className="pointer-events-auto leading-normal select-text shadow-xs"
      >
        {el.columnCount && el.columnCount > 1 ? (
          <div
            style={{
              columnCount: el.columnCount,
              columnGap: '16px',
              width: '100%',
              display: 'block',
            }}
          >
            {el.paragraphs.map(renderParagraph)}
          </div>
        ) : (
          el.paragraphs.map(renderParagraph)
        )}
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
                          {p.runs.map((r: PptxTextRun, rIdx: number) => {
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
          onClick={() => !isFullBackground && onImageClick?.(el.dataUrl)}
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
}

/** 슬라이드 설계 너비(px). 글자 크기가 px 고정이라, 이 너비로 그린 뒤 transform으로 줄여 원본 비율을 유지한다 */
export const PPTX_DESIGN_WIDTH = 960;

/** 슬라이드 한 장을 주어진 너비에 맞춰 축소해 그린다 (메모 카드 미리보기용) */
export const PptxSlideThumbnail: React.FC<{ slide: PptxSlide; aspect: number; width: number; isDark?: boolean }> = ({ slide, aspect, width, isDark = false }) => {
  const designH = PPTX_DESIGN_WIDTH / aspect;
  const scale = width / PPTX_DESIGN_WIDTH;
  return (
    <div style={{ width, height: designH * scale, overflow: 'hidden', position: 'relative' }}>
      <div
        data-testid="pptx-card-slide"
        style={{
          width: PPTX_DESIGN_WIDTH,
          height: designH,
          transform: `scale(${scale})`,
          transformOrigin: 'top left',
          position: 'relative',
          overflow: 'hidden',
          backgroundColor: slide.backgroundColor || (isDark ? '#0F172A' : '#FFFFFF'),
          color: isDark ? '#F1F5F9' : '#0F172A',
        }}
      >
        {slide.elements.map((el) => renderPptxSlideElement(el, isDark))}
      </div>
    </div>
  );
};
