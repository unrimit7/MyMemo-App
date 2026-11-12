import { useEffect, useRef } from 'react';

export interface UseOfficeShortcutsOptions {
  enabled?: boolean;
  onSave?: () => void | Promise<void>;
  onClose?: () => void;
  hasChanges?: boolean;
  onConfirmCloseNeeded?: () => void;
  isEditingCellOrDialog?: () => boolean;
  toggleFullscreen?: () => void;
}

/**
 * 오피스 편집기 4종(Excel, Word, PPT, HWP) 공통 단축키 훅
 *
 * - Ctrl+S: 저장 실행 (이중 실행 방지 가드 포함)
 * - Ctrl+Z, Ctrl+Y, Ctrl+F: 가로채지 않고 편집기 고유 동작 100% 보존
 * - Alt+Enter: 전체 화면 토글
 * - Esc: 셀 편집/대화상자 상태일 때는 취소만 수행하고 모달은 닫지 않음
 */
export function useOfficeShortcuts({
  enabled = true,
  onSave,
  onClose,
  hasChanges = false,
  onConfirmCloseNeeded,
  isEditingCellOrDialog,
  toggleFullscreen,
}: UseOfficeShortcutsOptions) {
  const isSavingRef = useRef(false);

  useEffect(() => {
    if (!enabled) return;

    const handleKeyDown = async (e: KeyboardEvent) => {
      const isCtrlOrMeta = e.ctrlKey || e.metaKey;
      const key = e.key.toLowerCase();

      // 1. Ctrl+S: 저장 (이중 실행 방지)
      if (isCtrlOrMeta && key === 's') {
        e.preventDefault();
        e.stopPropagation();

        if (isSavingRef.current) return;
        isSavingRef.current = true;
        try {
          if (onSave) {
            await onSave();
          }
        } finally {
          setTimeout(() => {
            isSavingRef.current = false;
          }, 500);
        }
        return;
      }

      // 2. Alt+Enter: 전체 화면 토글 (F11 배제)
      if (e.altKey && e.key === 'Enter') {
        e.preventDefault();
        e.stopPropagation();
        if (toggleFullscreen) {
          toggleFullscreen();
        }
        return;
      }

      // 3. Esc: 셀 편집 / 대화상자 상태 판별
      if (e.key === 'Escape') {
        // 셀 편집 중이거나 에디터 내부 대화상자가 열려 있는 경우 모달을 닫지 않음
        if (isEditingCellOrDialog && isEditingCellOrDialog()) {
          // 셀 편집이나 대화상자가 자체적으로 Esc를 소화하도록 모달 닫기 중단
          return;
        }

        // 변경 사항이 있는 경우 닫기 확인 창 노출
        if (hasChanges && onConfirmCloseNeeded) {
          e.preventDefault();
          e.stopPropagation();
          onConfirmCloseNeeded();
          return;
        }

        if (onClose) {
          e.preventDefault();
          e.stopPropagation();
          onClose();
        }
      }
    };

    window.addEventListener('keydown', handleKeyDown, true);
    return () => {
      window.removeEventListener('keydown', handleKeyDown, true);
    };
  }, [enabled, onSave, onClose, hasChanges, onConfirmCloseNeeded, isEditingCellOrDialog, toggleFullscreen]);
}
