import React, { useEffect, useState, useRef } from 'react';
import { createPortal } from 'react-dom';

export interface PopoutWindowProps {
  title: string;
  width?: number;
  height?: number;
  onClose: () => void;
  children: React.ReactNode;
}

export const PopoutWindow: React.FC<PopoutWindowProps> = ({
  title,
  width = 960,
  height = 760,
  onClose,
  children,
}) => {
  const [container, setContainer] = useState<HTMLDivElement | null>(null);
  const popoutWindowRef = useRef<Window | null>(null);
  const onCloseRef = useRef(onClose);
  onCloseRef.current = onClose;

  useEffect(() => {
    // 듀얼 모니터 및 현재 모니터 위치를 고려한 중앙 좌표 계산
    const screenX = window.screenX !== undefined ? window.screenX : window.screenLeft;
    const screenY = window.screenY !== undefined ? window.screenY : window.screenTop;
    const outerW = window.outerWidth || window.innerWidth;
    const outerH = window.outerHeight || window.innerHeight;
    const left = Math.max(0, screenX + Math.round((outerW - width) / 2));
    const top = Math.max(0, screenY + Math.round((outerH - height) / 2));

    const win = window.open(
      '',
      '_blank',
      `popup=yes,width=${width},height=${height},left=${left},top=${top},menubar=no,toolbar=no,location=no,status=no,resizable=yes,scrollbars=yes`
    );

    if (!win) {
      alert('⚠️ 브라우저에 의해 새 창(팝업)이 차단되었습니다.\n주소창 오른쪽 끝에서 "팝업 항상 허용"을 클릭해 주세요.');
      onCloseRef.current();
      return;
    }

    popoutWindowRef.current = win;
    win.document.title = title ? `${title} - MyMemo` : 'MyMemo';

    // Base 태그 설정 (상대 경로 에셋 및 폰트 깨짐 방지)
    const base = win.document.createElement('base');
    base.href = window.location.href;
    win.document.head.appendChild(base);

    // 메타 뷰포트 설정
    const meta = win.document.createElement('meta');
    meta.name = 'viewport';
    meta.content = 'width=device-width, initial-scale=1.0';
    win.document.head.appendChild(meta);

    // 부모 창의 모든 스타일시트 및 인라인 style 태그 복사 (절대 경로 보존)
    document.querySelectorAll('link[rel="stylesheet"], style').forEach((node) => {
      const clone = node.cloneNode(true) as HTMLElement;
      if (clone.tagName === 'LINK' && (clone as HTMLLinkElement).href) {
        (clone as HTMLLinkElement).href = (node as HTMLLinkElement).href;
      }
      win.document.head.appendChild(clone);
    });

    // body 기본 스타일
    win.document.body.className = 'bg-slate-950 text-slate-100 font-sans m-0 p-0 overflow-hidden select-text';
    win.document.body.style.margin = '0';
    win.document.body.style.backgroundColor = '#020617';
    win.document.body.style.width = '100vw';
    win.document.body.style.height = '100vh';
    win.document.body.style.overflow = 'hidden';

    // React 포털 마운트용 루트 컨테이너 생성
    const cont = win.document.createElement('div');
    cont.id = 'popout-root';
    cont.className = 'w-full h-full flex flex-col bg-slate-950 text-slate-100 overflow-hidden';
    cont.style.width = '100%';
    cont.style.height = '100%';
    win.document.body.appendChild(cont);
    setContainer(cont);

    // 사용자가 OS 팝업 창의 우측 상단 'X'를 눌렀을 때
    const handlePopoutUnload = () => {
      onCloseRef.current();
    };
    win.addEventListener('beforeunload', handlePopoutUnload);

    // 메인 창이 닫히거나 새로고침될 때 팝업 창도 함께 정리
    const handleParentUnload = () => {
      if (win && !win.closed) {
        win.close();
      }
    };
    window.addEventListener('beforeunload', handleParentUnload);

    return () => {
      win.removeEventListener('beforeunload', handlePopoutUnload);
      window.removeEventListener('beforeunload', handleParentUnload);
      if (win && !win.closed) {
        win.close();
      }
    };
  }, [width, height, title]);

  // 제목 동적 갱신
  useEffect(() => {
    if (popoutWindowRef.current && !popoutWindowRef.current.closed) {
      popoutWindowRef.current.document.title = title ? `${title} - MyMemo` : 'MyMemo';
    }
  }, [title]);

  if (!container) return null;
  return createPortal(children, container);
};
