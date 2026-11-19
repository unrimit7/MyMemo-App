// 오피스 미리보기·편집 화면 왼쪽에 쪽(시트) 썸네일 목록을 보여주는 공용 사이드바
import React, { useEffect, useRef, useState } from 'react';
import { PanelLeftClose, PanelLeftOpen } from 'lucide-react';

/**
 * 문서 미리보기/편집 창 왼쪽에 쪽(슬라이드·시트) 썸네일을 순서대로 보여주는 공용 사이드바.
 * - 화면에 보이는 썸네일만 그린다(IntersectionObserver). 수백 쪽짜리 문서도 처음 여는 속도가 느려지지 않는다.
 * - 현재 쪽 썸네일은 강조하고 목록에서 자동으로 보이게 스크롤한다.
 * - 접기/펴기 상태는 문서 종류별로 기억한다(좁은 화면은 기본 접힘).
 */
export interface PageThumbnailSidebarProps {
  /** 전체 쪽 수 */
  count: number;
  /** 현재 쪽 (0부터) */
  current: number;
  /** 썸네일 클릭 시 (0부터) */
  onSelect: (index: number) => void;
  /** 썸네일 내용. 화면에 보이게 된 썸네일만 호출된다 */
  renderThumb: (index: number) => React.ReactNode;
  /** 썸네일 아래 이름 (기본: 쪽 번호) */
  getLabel?: (index: number) => string;
  /** 머리글 (예: '쪽', '슬라이드', '시트') */
  unit?: string;
  /** 그리기 전 자리표시 박스의 가로/세로 비 (기본 A4 세로) */
  placeholderAspect?: number;
  /** 접힘 상태 기억 키 */
  storageKey: string;
  /** 같은 쪽이라도 내용이 바뀌면 다시 그리도록 하는 값(편집 화면용) */
  version?: number | string;
  /** 사이드바 아래쪽에 붙일 안내 문구 */
  footnote?: string;
  className?: string;
}

const THUMB_WIDTH = 112;

function readCollapsed(key: string): boolean {
  try {
    const v = localStorage.getItem(`mymemo.thumbs.${key}`);
    if (v === '1') return true;
    if (v === '0') return false;
  } catch {
    // 저장소를 못 쓰는 환경에서는 기본값
  }
  return typeof window !== 'undefined' && window.innerWidth < 640;
}

function writeCollapsed(key: string, collapsed: boolean) {
  try {
    localStorage.setItem(`mymemo.thumbs.${key}`, collapsed ? '1' : '0');
  } catch {
    // ignore
  }
}

const ThumbItem: React.FC<{
  index: number;
  active: boolean;
  label: string;
  placeholderAspect: number;
  rootRef: React.RefObject<HTMLDivElement | null>;
  onSelect: (index: number) => void;
  renderThumb: (index: number) => React.ReactNode;
  version?: number | string;
}> = ({ index, active, label, placeholderAspect, rootRef, onSelect, renderThumb, version }) => {
  const itemRef = useRef<HTMLButtonElement>(null);
  const [visible, setVisible] = useState(false);

  useEffect(() => {
    const el = itemRef.current;
    if (!el || visible) return;
    if (typeof IntersectionObserver === 'undefined') {
      setVisible(true);
      return;
    }
    const io = new IntersectionObserver(
      (entries) => {
        if (entries.some((e) => e.isIntersecting)) {
          setVisible(true);
          io.disconnect();
        }
      },
      { root: rootRef.current, rootMargin: '300px 0px' }
    );
    io.observe(el);
    return () => io.disconnect();
  }, [visible, rootRef]);

  useEffect(() => {
    if (active && itemRef.current) {
      itemRef.current.scrollIntoView({ block: 'nearest' });
    }
  }, [active]);

  return (
    <button
      ref={itemRef}
      type="button"
      data-testid={`page-thumb-${index}`}
      aria-current={active ? 'page' : undefined}
      onClick={() => onSelect(index)}
      className="group flex flex-col items-center gap-1 w-full py-1.5 cursor-pointer focus:outline-none"
      title={label}
    >
      <div
        className={`relative bg-white overflow-hidden rounded-sm transition-shadow ${
          active
            ? 'ring-[3px] ring-sky-400 shadow-lg shadow-sky-900/40'
            : 'ring-1 ring-slate-700 group-hover:ring-2 group-hover:ring-slate-400'
        }`}
        style={{ width: THUMB_WIDTH, minHeight: visible ? undefined : THUMB_WIDTH / placeholderAspect }}
      >
        {visible ? (
          <React.Fragment key={version}>{renderThumb(index)}</React.Fragment>
        ) : null}
      </div>
      <span
        className={`max-w-full truncate px-1 text-[11px] leading-tight ${
          active ? 'text-sky-300 font-bold' : 'text-slate-400 group-hover:text-slate-200'
        }`}
      >
        {label}
      </span>
    </button>
  );
};

export const PageThumbnailSidebar: React.FC<PageThumbnailSidebarProps> = ({
  count,
  current,
  onSelect,
  renderThumb,
  getLabel,
  unit = '쪽',
  placeholderAspect = 1 / Math.SQRT2,
  storageKey,
  version,
  footnote,
  className = '',
}) => {
  const [collapsed, setCollapsed] = useState<boolean>(() => readCollapsed(storageKey));
  const listRef = useRef<HTMLDivElement>(null);

  const toggle = () => {
    setCollapsed((prev) => {
      writeCollapsed(storageKey, !prev);
      return !prev;
    });
  };

  if (collapsed) {
    return (
      <div className={`shrink-0 w-8 bg-slate-900 border-r border-slate-800 flex flex-col items-center py-2 gap-2 ${className}`}>
        <button
          type="button"
          onClick={toggle}
          className="p-1 text-slate-400 hover:text-white hover:bg-slate-800 rounded"
          title={`${unit} 썸네일 보이기`}
          aria-label={`${unit} 썸네일 보이기`}
        >
          <PanelLeftOpen size={16} />
        </button>
        {count > 0 && (
          <span className="text-[10px] text-slate-500 [writing-mode:vertical-rl]">
            {current + 1} / {count}
          </span>
        )}
      </div>
    );
  }

  return (
    <div
      className={`shrink-0 w-[140px] bg-slate-900 border-r border-slate-800 flex flex-col min-h-0 select-none ${className}`}
      data-testid="page-thumbnail-sidebar"
    >
      <div className="flex items-center justify-between px-2 py-1.5 border-b border-slate-800 text-[11px] text-slate-400">
        <span className="font-semibold">
          {unit} {count > 0 ? `${current + 1}/${count}` : ''}
        </span>
        <button
          type="button"
          onClick={toggle}
          className="p-0.5 text-slate-400 hover:text-white hover:bg-slate-800 rounded"
          title={`${unit} 썸네일 숨기기`}
          aria-label={`${unit} 썸네일 숨기기`}
        >
          <PanelLeftClose size={15} />
        </button>
      </div>
      <div ref={listRef} className="flex-1 min-h-0 overflow-y-auto overflow-x-hidden px-1 py-1.5">
        {Array.from({ length: count }, (_, i) => (
          <ThumbItem
            key={i}
            index={i}
            active={i === current}
            label={getLabel ? getLabel(i) : String(i + 1)}
            placeholderAspect={placeholderAspect}
            rootRef={listRef}
            onSelect={onSelect}
            renderThumb={renderThumb}
            version={version}
          />
        ))}
      </div>
      {footnote && (
        <div className="px-2 py-1.5 border-t border-slate-800 text-[10px] leading-snug text-slate-500">{footnote}</div>
      )}
    </div>
  );
};

/** 썸네일 너비(px). 각 형식별 썸네일 렌더러가 이 너비에 맞춰 그린다 */
export const PAGE_THUMB_WIDTH = THUMB_WIDTH;

/* ───────────────────────── 형식별 썸네일 ───────────────────────── */

/** PDF 쪽 썸네일 (pdf.js로 작은 캔버스에 그림) */
export const PdfPageThumb: React.FC<{
  pdfDoc: import('pdfjs-dist').PDFDocumentProxy;
  pageNumber: number;
}> = ({ pdfDoc, pageNumber }) => {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  useEffect(() => {
    let alive = true;
    const canvas = canvasRef.current;
    if (!canvas) return;
    import('../utils/pdfViewer')
      .then(({ renderPdfPageToCanvas }) => {
        if (!alive) return;
        // renderPdfPageToCanvas는 (targetWidth - 10)에 맞춰 그린다
        return renderPdfPageToCanvas(pdfDoc, pageNumber, canvas, THUMB_WIDTH + 10, 1, 0);
      })
      .then(() => {
        if (alive && canvas) {
          canvas.style.width = '100%';
          canvas.style.height = 'auto';
        }
      })
      .catch((err) => {
        if (alive && err?.name !== 'RenderingCancelledException') console.warn('PDF 썸네일 렌더링 실패:', err);
      });
    return () => {
      alive = false;
    };
  }, [pdfDoc, pageNumber]);
  return <canvas ref={canvasRef} className="block w-full h-auto" />;
};

/**
 * SVG 문자열 썸네일. 여러 SVG를 한 문서에 직접 넣으면 내부 id(clipPath 등)가 서로 엉키므로
 * Blob URL 이미지로 격리해서 보여준다.
 */
export const SvgStringThumb: React.FC<{ getSvg: () => string | Promise<string> }> = ({ getSvg }) => {
  const [url, setUrl] = useState<string | null>(null);
  const [failed, setFailed] = useState(false);
  useEffect(() => {
    let alive = true;
    let created: string | null = null;
    Promise.resolve()
      .then(() => getSvg())
      .then((svg) => {
        if (!alive || !svg) return;
        created = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml' }));
        setUrl(created);
      })
      .catch((err) => {
        console.warn('SVG 썸네일 생성 실패:', err);
        if (alive) setFailed(true);
      });
    return () => {
      alive = false;
      if (created) URL.revokeObjectURL(created);
    };
    // getSvg는 매 렌더 새로 만들어지므로 의존성에서 뺀다 (다시 그리기는 sidebar의 version으로 제어)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  if (failed) return <div className="w-full aspect-[1/1.414] flex items-center justify-center text-[10px] text-slate-400">미리보기 없음</div>;
  if (!url) return <div className="w-full aspect-[1/1.414] bg-slate-100 animate-pulse" />;
  return <img src={url} alt="" draggable={false} className="block w-full h-auto" />;
};

/**
 * 이미 화면에 그려진 쪽 DOM(워드 미리보기, 워드 편집기)을 복제해 줄여 보여준다.
 * 원본 쪽의 부모 클래스를 감싸개에 붙여 라이브러리 CSS 선택자가 그대로 적용되게 한다.
 */
export const DomCloneThumb: React.FC<{
  source: HTMLElement | null | undefined;
  /** 긴 구역을 여러 쪽으로 나눠 보일 때: 이 쪽이 시작하는 세로 위치와 쪽 높이(원본 px) */
  sliceTop?: number;
  sliceHeight?: number;
  /** 화면에서 떼어 둔(분리된) 복제본일 때 원래 크기 */
  sourceWidth?: number;
  sourceHeight?: number;
  /** 분리된 복제본일 때 원래 부모 요소의 class (라이브러리 CSS 선택자 유지용) */
  parentClassName?: string;
}> = ({ source, sliceTop = 0, sliceHeight, sourceWidth, sourceHeight, parentClassName }) => {
  const hostRef = useRef<HTMLDivElement>(null);
  const [height, setHeight] = useState<number>(THUMB_WIDTH * Math.SQRT2);
  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    host.innerHTML = '';
    if (!source) return;
    const w = sourceWidth || source.offsetWidth || 794;
    const h = sourceHeight || source.offsetHeight || w * Math.SQRT2;
    const scale = THUMB_WIDTH / w;
    // 쪽 높이를 모르면 A4 비율보다 길게 그려진 부분은 잘라서 보여준다
    const visibleH = sliceHeight ? Math.min(sliceHeight, Math.max(20, h - sliceTop)) : Math.min(h, w * 1.6);
    setHeight(Math.max(20, visibleH * scale));

    const parentWrap = document.createElement('div');
    const pc = parentClassName ?? source.parentElement?.className;
    if (pc && typeof pc === 'string') parentWrap.className = pc;
    parentWrap.style.cssText = `all:unset;display:block;color:#0f172a;background:#fff;position:absolute;left:0;top:${-sliceTop * scale}px;width:${w}px;transform:scale(${scale});transform-origin:top left;pointer-events:none;`;
    const clone = source.cloneNode(true) as HTMLElement;
    clone.removeAttribute('id');
    clone.querySelectorAll('[id]').forEach((n) => n.removeAttribute('id'));
    clone.querySelectorAll('[contenteditable]').forEach((n) => n.setAttribute('contenteditable', 'false'));
    clone.style.margin = '0';
    clone.style.boxShadow = 'none';
    clone.style.transform = 'none';
    clone.style.position = 'relative';
    clone.style.left = '0';
    clone.style.top = '0';
    parentWrap.appendChild(clone);
    host.appendChild(parentWrap);
  }, [source, sliceTop, sliceHeight, sourceWidth, sourceHeight, parentClassName]);
  return (
    <div
      ref={hostRef}
      aria-hidden="true"
      className="relative w-full overflow-hidden bg-white"
      style={{ height }}
    />
  );
};

/** 엑셀 시트 썸네일: 왼쪽 위 일부 셀을 작은 격자로 그린다 */
export const SheetGridThumb: React.FC<{ rows: (string | number | null | undefined)[][]; empty?: boolean }> = ({ rows, empty }) => {
  const maxCols = Math.max(1, Math.min(6, rows.reduce((m, r) => Math.max(m, r.length), 0)));
  const shown = rows.slice(0, 14);
  return (
    <div className="w-full aspect-[4/3] bg-white overflow-hidden">
      {empty || shown.length === 0 ? (
        <div className="w-full h-full flex items-center justify-center text-[10px] text-slate-400">빈 시트</div>
      ) : (
        <table className="w-full table-fixed border-collapse text-[5px] leading-[7px] text-slate-800">
          <tbody>
            {shown.map((r, ri) => (
              <tr key={ri} className={ri === 0 ? 'bg-slate-100 font-bold' : ''}>
                {Array.from({ length: maxCols }, (_, ci) => (
                  <td key={ci} className="border border-slate-200 px-[1px] truncate whitespace-nowrap overflow-hidden">
                    {r[ci] == null ? '' : String(r[ci])}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
};

/** 화면에 그려진 쪽 요소 하나가 실제로는 여러 쪽 분량일 때(워드 미리보기) 나눠 보일 조각 */
export interface DomPageSlice {
  el: HTMLElement;
  top: number;
  height: number;
}

/**
 * 쪽 요소들을 썸네일 단위로 나눈다. 쪽 높이(min-height 또는 A4 비율)를 넘는 요소는
 * 그 높이마다 한 쪽으로 계산한다. (docx-preview는 쪽 나눔·구역 나눔에서만 쪽을 나누기 때문)
 */
export function splitDomPages(els: HTMLElement[]): DomPageSlice[] {
  const out: DomPageSlice[] = [];
  for (const el of els) {
    const w = el.offsetWidth || 794;
    const h = el.offsetHeight || w * Math.SQRT2;
    const minH = parseFloat(el.style.minHeight || '') || 0;
    const pageH = minH > 100 ? minH : w * Math.SQRT2;
    const n = Math.max(1, Math.min(200, Math.ceil((h - 4) / pageH)));
    for (let k = 0; k < n; k++) out.push({ el, top: k * pageH, height: pageH });
  }
  return out;
}

/** 스크롤되는 가장 가까운 조상 요소 */
export function findScrollParent(el: HTMLElement | null, stopAt?: HTMLElement | null): HTMLElement | null {
  let cur = el?.parentElement ?? null;
  while (cur && cur !== document.body) {
    const oy = getComputedStyle(cur).overflowY;
    // overflow:hidden 이라도 내용이 넘치면 스크롤 위치는 코드로 옮길 수 있다
    if (oy !== 'visible' && oy !== 'clip' && cur.scrollHeight > cur.clientHeight + 1) return cur;
    if (cur === stopAt) break;
    cur = cur.parentElement;
  }
  return null;
}

/** 스크롤 영역 안에서 el(의 top + offset)이 맨 위에 오도록 스크롤 */
export function scrollElementIntoScroller(scroller: HTMLElement, el: HTMLElement, offset = 0, smooth = true) {
  const delta = el.getBoundingClientRect().top + offset - scroller.getBoundingClientRect().top;
  scroller.scrollTo({ top: scroller.scrollTop + delta - 8, behavior: smooth ? 'smooth' : 'auto' });
}
