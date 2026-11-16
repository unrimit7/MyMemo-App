// 메모 본문 PPT 카드의 고정 안내 그림 자리에 실제 첫 슬라이드를 그려 넣는 함수 (메인 편집기와 창 분리가 같이 사용, K8)
import { createRoot, type Root } from 'react-dom/client';
import { getCachedPptxPresentation } from '../utils/pptxViewer';
import { PptxSlideThumbnail } from './PptxSlideRenderer';

// 카드 칸마다 React 루트와 그린 데이터 표식 (같은 데이터면 다시 그리지 않음)
const roots = new WeakMap<HTMLElement, { root: Root; key: string }>();

function dataKey(attId: string, data: string): string {
  return `${attId}:${data.length}:${data.slice(-64)}`;
}

/**
 * preview 안의 모든 .pptx-inline-viewer 카드에 첫 슬라이드를 그린다.
 * 읽기에 실패하면 원래 안내 그림을 그대로 두고 작은 안내만 붙인다.
 */
export function hydratePptxInlineCards(preview: HTMLElement, getData: (attId: string) => string | undefined): void {
  preview.querySelectorAll<HTMLElement>('.pptx-inline-viewer').forEach((card) => {
    const attId = card.getAttribute('data-att-id') || '';
    const data = attId ? getData(attId) : undefined;
    const box = card.querySelector<HTMLElement>('.pptx-container');
    if (!data || !box) return;
    const key = dataKey(attId, data);
    const prev = roots.get(box);
    if (prev && prev.key === key) return;

    getCachedPptxPresentation(attId, data)
      .then((pres) => {
        const slide = pres.slides[0];
        if (!slide || !box.isConnected) return;
        // 카드 안쪽 너비 (안내 그림과 같은 최대 672px)
        const width = Math.max(200, Math.min(672, box.clientWidth - 32));
        let host = box.querySelector<HTMLElement>(':scope > .pptx-card-slide-host');
        if (!host) {
          host = document.createElement('div');
          host.className = 'pptx-card-slide-host rounded-xl overflow-hidden shadow-xl border border-slate-300 bg-white';
          box.querySelectorAll(':scope > *').forEach((c) => ((c as HTMLElement).style.display = 'none'));
          box.appendChild(host);
        }
        const entry = roots.get(box);
        const root = entry?.root || createRoot(host);
        roots.set(box, { root, key });
        root.render(<PptxSlideThumbnail slide={slide} aspect={pres.aspectRatio || 16 / 9} width={width} />);
        box.setAttribute('data-slide-count', String(pres.slides.length));
        box.title = `첫 슬라이드 미리보기 (전체 ${pres.slides.length}장). 클릭하면 전면 뷰어로 크게 볼 수 있습니다.`;
      })
      .catch((err) => {
        console.warn('PPT 카드 첫 슬라이드를 그리지 못했습니다 (안내 그림 유지):', err);
        box.title = '첫 슬라이드를 읽지 못해 안내 그림을 표시합니다. 클릭하면 전면 뷰어로 열립니다.';
      });
  });
}
