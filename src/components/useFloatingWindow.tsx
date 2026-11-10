import React, { useState, useEffect, useRef, useCallback } from 'react';
import { Maximize2, Minimize2 } from 'lucide-react';

export interface UseFloatingWindowOptions {
  id?: string;
  defaultWidth?: number;
  defaultHeight?: number;
  minWidth?: number;
  minHeight?: number;
}

export function useFloatingWindow(options: UseFloatingWindowOptions = {}) {
  const {
    id = 'doc-modal',
    defaultWidth = 860,
    defaultHeight = 780,
    minWidth = 380,
    minHeight = 320,
  } = options;

  const storageKey = `mymemo_floating_${id}`;

  // 플로팅 창 모드 활성화 여부
  const [isFloating, setIsFloating] = useState<boolean>(() => {
    try {
      return localStorage.getItem(storageKey) === 'true';
    } catch {
      return false;
    }
  });

  // 초기 위치 계산: 대화면인 경우 우측에 배치하여 좌측 메모가 보이도록 설정
  const getDefaultPos = useCallback(() => {
    const screenW = typeof window !== 'undefined' ? window.innerWidth : 1200;
    const initialW = Math.min(defaultWidth, screenW - 40);
    const x = screenW >= 1024 ? Math.max(20, screenW - initialW - 30) : 20;
    const y = 50;
    return { x, y };
  }, [defaultWidth]);

  const [position, setPosition] = useState<{ x: number; y: number }>(getDefaultPos);
  const [size, setSize] = useState<{ width: number; height: number }>(() => {
    const screenW = typeof window !== 'undefined' ? window.innerWidth : 1200;
    const screenH = typeof window !== 'undefined' ? window.innerHeight : 900;
    return {
      width: Math.min(defaultWidth, screenW - 40),
      height: Math.min(defaultHeight, screenH - 80),
    };
  });

  const [isDragging, setIsDragging] = useState(false);
  const [isResizing, setIsResizing] = useState(false);

  const dragRef = useRef({
    startX: 0,
    startY: 0,
    initialPosX: 0,
    initialPosY: 0,
  });

  const resizeRef = useRef({
    startX: 0,
    startY: 0,
    initialW: 0,
    initialH: 0,
  });

  // 플로팅 모드 토글 및 브라우저 저장
  const toggleFloating = useCallback(() => {
    setIsFloating((prev) => {
      const next = !prev;
      try {
        localStorage.setItem(storageKey, String(next));
      } catch {}
      return next;
    });
  }, [storageKey]);

  // 상단 헤더 드래그 시작 (버튼 및 입력란 클릭 시에는 드래그 방지)
  const handleHeaderPointerDown = useCallback((e: React.PointerEvent) => {
    if (!isFloating) return;
    const target = e.target as HTMLElement;
    if (target.closest('button, input, select, textarea, a, .no-drag')) {
      return;
    }
    e.preventDefault();
    setIsDragging(true);
    dragRef.current = {
      startX: e.clientX,
      startY: e.clientY,
      initialPosX: position.x,
      initialPosY: position.y,
    };
  }, [isFloating, position]);

  // 우측 하단 크기 조절 시작
  const handleResizePointerDown = useCallback((e: React.PointerEvent) => {
    if (!isFloating) return;
    e.preventDefault();
    e.stopPropagation();
    setIsResizing(true);
    resizeRef.current = {
      startX: e.clientX,
      startY: e.clientY,
      initialW: size.width,
      initialH: size.height,
    };
  }, [isFloating, size]);

  // 전역 마우스/포인터 이동 및 해제 이벤트
  useEffect(() => {
    if (!isDragging && !isResizing) return;

    const handlePointerMove = (e: PointerEvent) => {
      if (isDragging) {
        const deltaX = e.clientX - dragRef.current.startX;
        const deltaY = e.clientY - dragRef.current.startY;
        const newX = dragRef.current.initialPosX + deltaX;
        const newY = dragRef.current.initialPosY + deltaY;

        const maxW = typeof window !== 'undefined' ? window.innerWidth : 1200;
        const maxH = typeof window !== 'undefined' ? window.innerHeight : 900;
        const clampedX = Math.max(-size.width + 120, Math.min(maxW - 100, newX));
        const clampedY = Math.max(0, Math.min(maxH - 60, newY));

        setPosition({ x: clampedX, y: clampedY });
      }

      if (isResizing) {
        const deltaX = e.clientX - resizeRef.current.startX;
        const deltaY = e.clientY - resizeRef.current.startY;
        const newW = resizeRef.current.initialW + deltaX;
        const newH = resizeRef.current.initialH + deltaY;

        const maxW = typeof window !== 'undefined' ? window.innerWidth - position.x - 10 : 1200;
        const maxH = typeof window !== 'undefined' ? window.innerHeight - position.y - 10 : 900;

        setSize({
          width: Math.max(minWidth, Math.min(maxW, newW)),
          height: Math.max(minHeight, Math.min(maxH, newH)),
        });
      }
    };

    const handlePointerUp = () => {
      setIsDragging(false);
      setIsResizing(false);
    };

    window.addEventListener('pointermove', handlePointerMove);
    window.addEventListener('pointerup', handlePointerUp);

    return () => {
      window.removeEventListener('pointermove', handlePointerMove);
      window.removeEventListener('pointerup', handlePointerUp);
    };
  }, [isDragging, isResizing, size.width, position.x, position.y, minWidth, minHeight]);

  return {
    isFloating,
    setIsFloating,
    toggleFloating,
    position,
    size,
    isDragging,
    isResizing,
    headerProps: {
      onPointerDown: handleHeaderPointerDown,
      style: {
        cursor: isFloating ? (isDragging ? 'grabbing' : 'grab') : 'default',
        userSelect: isDragging ? ('none' as const) : undefined,
      },
    },
    modalWrapperClass: isFloating
      ? 'fixed inset-0 z-50 pointer-events-none select-none overflow-hidden'
      : 'fixed inset-0 z-50 bg-black/85 backdrop-blur-sm flex flex-col items-center justify-center p-2 md:p-4 select-none animate-fadeIn',
    windowContainerProps: {
      style: isFloating
        ? {
            position: 'fixed' as const,
            left: `${position.x}px`,
            top: `${position.y}px`,
            width: `${size.width}px`,
            height: `${size.height}px`,
            maxWidth: '100vw',
            maxHeight: '100vh',
            zIndex: 55,
          }
        : undefined,
      className: isFloating
        ? 'pointer-events-auto bg-slate-900 border border-slate-700/80 rounded-2xl shadow-[0_25px_60px_-15px_rgba(0,0,0,0.8)] flex flex-col overflow-hidden ring-1 ring-white/10 relative'
        : 'w-full max-w-6xl h-[95vh] bg-slate-900 border border-slate-800 rounded-2xl shadow-2xl flex flex-col overflow-hidden relative',
    },
    toggleFloatingButton: (
      <button
        type="button"
        onClick={toggleFloating}
        className="p-1.5 bg-slate-800 hover:bg-slate-750 text-slate-300 hover:text-white rounded-lg border border-slate-700 transition-colors"
        title={isFloating ? '전체화면 모달로 전환' : '창 분리 (화면 이동 및 메모와 동시 비교)'}
      >
        {isFloating ? <Maximize2 size={15} /> : <Minimize2 size={15} />}
      </button>
    ),
    resizeHandle: isFloating ? (
      <div
        onPointerDown={handleResizePointerDown}
        className="absolute bottom-0 right-0 w-5 h-5 cursor-nwse-resize z-30 flex items-center justify-center text-slate-500 hover:text-indigo-400 pointer-events-auto select-none"
        title="드래그하여 창 크기 조절"
      >
        <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor">
          <circle cx="10" cy="10" r="1.2" />
          <circle cx="6" cy="10" r="1.2" />
          <circle cx="10" cy="6" r="1.2" />
          <circle cx="2" cy="10" r="1.2" />
          <circle cx="6" cy="6" r="1.2" />
          <circle cx="10" cy="2" r="1.2" />
        </svg>
      </div>
    ) : null,
  };
}
