import { useEffect, useRef } from 'react';

export interface UseWheelPageNavigationOptions {
  containerRef: React.RefObject<HTMLDivElement | null>;
  currentPage: number;
  totalPages: number;
  onPageChange: (newPage: number) => void;
  enabled?: boolean;
  threshold?: number; // 휠 누적 임계치 (기본값: 180, 일반 마우스 휠 2회 회전 수준)
}

export function useWheelPageNavigation({
  containerRef,
  currentPage,
  totalPages,
  onPageChange,
  enabled = true,
  threshold = 180,
}: UseWheelPageNavigationOptions) {
  const accRef = useRef<{
    delta: number;
    dir: 'up' | 'down' | null;
    lastTime: number;
    cooldownUntil: number;
  }>({
    delta: 0,
    dir: null,
    lastTime: 0,
    cooldownUntil: 0,
  });

  const pendingScrollRef = useRef<'top' | 'bottom' | null>(null);

  useEffect(() => {
    const container = containerRef.current;
    if (!container || !enabled) return;

    const handleWheel = (e: WheelEvent) => {
      // Ctrl 또는 Meta(Cmd) 키가 눌려있으면 핀치 줌 / 배율 확대·축소 동작이므로 제외
      if (e.ctrlKey || e.metaKey) return;

      const now = Date.now();
      const acc = accRef.current;

      // 페이지 전환 직후 쿨다운 중이면 추가 휠 무시 (급격한 다중 페이지 건너뜀 방지)
      if (now < acc.cooldownUntil) {
        return;
      }

      // 이전 휠 입력 후 500ms 이상 지났으면 누적치 리셋
      if (now - acc.lastTime > 500) {
        acc.delta = 0;
        acc.dir = null;
      }

      const scrollTop = container.scrollTop;
      const scrollHeight = container.scrollHeight;
      const clientHeight = container.clientHeight;

      // 서브픽셀 렌더링 및 줌 배율을 감안한 4px 허용 마진
      const isAtTop = scrollTop <= 4;
      const isAtBottom = scrollTop + clientHeight >= scrollHeight - 4;

      if (e.deltaY < 0) {
        // 휠을 위로 올림 (이전 페이지 방향)
        if (isAtTop) {
          // 첫 페이지인 경우 더 이상 앞 페이지로 이동하지 않음 (변화 없음)
          if (currentPage <= 1) {
            acc.delta = 0;
            acc.dir = null;
            return;
          }

          if (acc.dir !== 'up') {
            acc.delta = 0;
            acc.dir = 'up';
          }
          acc.delta += Math.abs(e.deltaY);
          acc.lastTime = now;

          // 상단에서 휠을 몇 번 더 올렸을 때(임계치 충족) 이전 페이지로 이동
          if (acc.delta >= threshold) {
            acc.delta = 0;
            acc.dir = null;
            acc.cooldownUntil = now + 400; // 400ms 쿨다운

            const targetPage = currentPage - 1;
            pendingScrollRef.current = 'bottom';
            onPageChange(targetPage);
          }
        } else {
          // 본문 내부 스크롤 중이므로 누적치 리셋
          acc.delta = 0;
          acc.dir = null;
        }
      } else if (e.deltaY > 0) {
        // 휠을 아래로 내림 (다음 페이지 방향)
        if (isAtBottom) {
          // 마지막 페이지인 경우 더 이상 다음 페이지로 이동하지 않음 (변화 없음)
          if (currentPage >= totalPages) {
            acc.delta = 0;
            acc.dir = null;
            return;
          }

          if (acc.dir !== 'down') {
            acc.delta = 0;
            acc.dir = 'down';
          }
          acc.delta += Math.abs(e.deltaY);
          acc.lastTime = now;

          // 하단에서 휠을 몇 번 더 내렸을 때(임계치 충족) 다음 페이지로 이동
          if (acc.delta >= threshold) {
            acc.delta = 0;
            acc.dir = null;
            acc.cooldownUntil = now + 400; // 400ms 쿨다운

            const targetPage = currentPage + 1;
            pendingScrollRef.current = 'top';
            onPageChange(targetPage);
          }
        } else {
          // 본문 내부 스크롤 중이므로 누적치 리셋
          acc.delta = 0;
          acc.dir = null;
        }
      }
    };

    container.addEventListener('wheel', handleWheel, { passive: true });
    return () => {
      container.removeEventListener('wheel', handleWheel);
    };
  }, [containerRef, currentPage, totalPages, enabled, threshold, onPageChange]);

  return { pendingScrollRef };
}
