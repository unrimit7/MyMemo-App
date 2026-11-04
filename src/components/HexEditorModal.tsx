import React, { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import {
  X,
  Save,
  Download,
  ArrowUp,
  ArrowDown,
  Binary,
  CheckCircle,
  AlertCircle,
  Loader2,
  Copy,
  Check,
  Undo2,
  Redo2,
  Navigation
} from 'lucide-react';
import {
  dataUrlToUint8Array,
  uint8ArrayToDataUrl,
  formatHexOffset,
  formatByteToHex,
  byteToAsciiChar,
  searchHexBytes,
  searchStringBytes,
  inspectByteValues,
  type InspectorData,
} from '../utils/hexViewer';

interface HexEditorModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: string;
  title: string;
  attId?: string;
  onSave?: (updatedDataUrl: string) => Promise<void> | void;
}

const ROW_HEIGHT = 26; // 한 행 높이(px)
const BYTES_PER_ROW = 16;
const OVERSCAN_ROWS = 25; // 상하 가상 스크롤 여유 행

export const HexEditorModal: React.FC<HexEditorModalProps> = ({
  isOpen,
  onClose,
  data,
  title,
  onSave,
}) => {
  // 바이너리 데이터 버퍼
  const [buffer, setBuffer] = useState<Uint8Array>(new Uint8Array(0));
  const [originalBuffer, setOriginalBuffer] = useState<Uint8Array>(new Uint8Array(0));
  const [modifiedSet, setModifiedSet] = useState<Set<number>>(new Set());

  // 커서 및 편집 상태
  const [selectedOffset, setSelectedOffset] = useState<number>(0);
  const [activeColumn, setActiveColumn] = useState<'hex' | 'ascii'>('hex');
  const [hexNibble, setHexNibble] = useState<0 | 1>(0); // 0: 상위 4비트, 1: 하위 4비트

  // Undo / Redo 스택
  const historyRef = useRef<{ offset: number; oldVal: number; newVal: number }[]>([]);
  const historyIndexRef = useRef<number>(-1);

  // 검색 상태
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [searchMode, setSearchMode] = useState<'text' | 'hex'>('text');
  const [matchOffset, setMatchOffset] = useState<number | null>(null);

  // 주소 이동(Go to Offset) 상태
  const [showGotoDialog, setShowGotoDialog] = useState<boolean>(false);
  const [gotoInputVal, setGotoInputVal] = useState<string>('');

  // 상태 플래그
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [hasChanges, setHasChanges] = useState<boolean>(false);
  const [saveToast, setSaveToast] = useState<string | null>(null);
  const [errorToast, setErrorToast] = useState<string | null>(null);
  const [copiedToast, setCopiedToast] = useState<string | null>(null);

  // 스크롤 및 뷰포트
  const scrollContainerRef = useRef<HTMLDivElement>(null);
  const [scrollTop, setScrollTop] = useState<number>(0);
  const [viewportHeight, setViewportHeight] = useState<number>(600);

  // 데이터 로드
  useEffect(() => {
    if (!isOpen || !data) return;

    try {
      const bytes = dataUrlToUint8Array(data);
      setBuffer(new Uint8Array(bytes));
      setOriginalBuffer(new Uint8Array(bytes));
      setModifiedSet(new Set());
      setSelectedOffset(0);
      setHexNibble(0);
      setActiveColumn('hex');
      setHasChanges(false);
      setMatchOffset(null);
      setSearchQuery('');
      historyRef.current = [];
      historyIndexRef.current = -1;
    } catch (err) {
      console.error('바이너리 로드 실패:', err);
      setErrorToast('바이너리 데이터를 불러오지 못했습니다.');
    }
  }, [isOpen, data]);

  // 컨테이너 크기 감지
  useEffect(() => {
    if (!isOpen) return;
    const container = scrollContainerRef.current;
    if (!container) return;

    const updateHeight = () => {
      setViewportHeight(container.clientHeight || 600);
    };
    updateHeight();
    window.addEventListener('resize', updateHeight);
    return () => window.removeEventListener('resize', updateHeight);
  }, [isOpen]);

  const totalBytes = buffer.length;
  const totalRows = Math.ceil(totalBytes / BYTES_PER_ROW);

  // 가상 스크롤 계산
  const { startIndex, endIndex, topPadding, bottomPadding } = useMemo(() => {
    const startRow = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - OVERSCAN_ROWS);
    const visibleRowCount = Math.ceil(viewportHeight / ROW_HEIGHT);
    const endRow = Math.min(totalRows, startRow + visibleRowCount + OVERSCAN_ROWS * 2);

    const topPad = startRow * ROW_HEIGHT;
    const bottomPad = Math.max(0, (totalRows - endRow) * ROW_HEIGHT);

    return {
      startIndex: startRow,
      endIndex: endRow,
      topPadding: topPad,
      bottomPadding: bottomPad,
    };
  }, [scrollTop, viewportHeight, totalRows]);

  // 커서가 화면 밖으로 벗어났을 때 자동 스크롤
  const scrollToOffset = useCallback((offset: number) => {
    const container = scrollContainerRef.current;
    if (!container) return;

    const targetRow = Math.floor(offset / BYTES_PER_ROW);
    const targetTop = targetRow * ROW_HEIGHT;
    const currentScroll = container.scrollTop;
    const containerHeight = container.clientHeight;

    if (targetTop < currentScroll) {
      container.scrollTop = targetTop;
    } else if (targetTop + ROW_HEIGHT > currentScroll + containerHeight) {
      container.scrollTop = targetTop + ROW_HEIGHT - containerHeight;
    }
  }, []);

  // 바이트 값 수정 적용
  const updateByteAt = useCallback((offset: number, newByte: number) => {
    if (offset < 0 || offset >= buffer.length) return;
    const oldByte = buffer[offset];
    if (oldByte === newByte) return;

    const newBuffer = new Uint8Array(buffer);
    newBuffer[offset] = newByte;
    setBuffer(newBuffer);

    setModifiedSet((prev) => {
      const next = new Set(prev);
      if (newByte === originalBuffer[offset]) {
        next.delete(offset);
      } else {
        next.add(offset);
      }
      return next;
    });

    setHasChanges(true);

    // 히스토리 추가
    const nextIdx = historyIndexRef.current + 1;
    historyRef.current = [
      ...historyRef.current.slice(0, nextIdx),
      { offset, oldVal: oldByte, newVal: newByte },
    ];
    historyIndexRef.current = nextIdx;
  }, [buffer, originalBuffer]);

  // 되돌리기(Undo)
  const handleUndo = () => {
    if (historyIndexRef.current >= 0) {
      const item = historyRef.current[historyIndexRef.current];
      historyIndexRef.current -= 1;

      const newBuffer = new Uint8Array(buffer);
      newBuffer[item.offset] = item.oldVal;
      setBuffer(newBuffer);

      setModifiedSet((prev) => {
        const next = new Set(prev);
        if (item.oldVal === originalBuffer[item.offset]) {
          next.delete(item.offset);
        } else {
          next.add(item.offset);
        }
        return next;
      });

      setSelectedOffset(item.offset);
      scrollToOffset(item.offset);
    }
  };

  // 다시 실행(Redo)
  const handleRedo = () => {
    if (historyIndexRef.current < historyRef.current.length - 1) {
      historyIndexRef.current += 1;
      const item = historyRef.current[historyIndexRef.current];

      const newBuffer = new Uint8Array(buffer);
      newBuffer[item.offset] = item.newVal;
      setBuffer(newBuffer);

      setModifiedSet((prev) => {
        const next = new Set(prev);
        if (item.newVal === originalBuffer[item.offset]) {
          next.delete(item.offset);
        } else {
          next.add(item.offset);
        }
        return next;
      });

      setSelectedOffset(item.offset);
      scrollToOffset(item.offset);
    }
  };

  // 키보드 네비게이션 및 직접 입력 처리
  const handleKeyDown = (e: React.KeyboardEvent) => {
    // Ctrl + Z / Ctrl + Y
    if ((e.ctrlKey || e.metaKey) && e.key === 'z') {
      e.preventDefault();
      if (e.shiftKey) handleRedo();
      else handleUndo();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key === 'y') {
      e.preventDefault();
      handleRedo();
      return;
    }

    // Ctrl + S (저장)
    if ((e.ctrlKey || e.metaKey) && e.key === 's') {
      e.preventDefault();
      handleSave();
      return;
    }

    // Ctrl + F (검색)
    if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
      e.preventDefault();
      const input = document.getElementById('hex-search-input');
      input?.focus();
      return;
    }

    // Ctrl + G (주소 이동)
    if ((e.ctrlKey || e.metaKey) && e.key === 'g') {
      e.preventDefault();
      setShowGotoDialog(true);
      return;
    }

    // Tab (HEX <-> ASCII 열 전환)
    if (e.key === 'Tab') {
      e.preventDefault();
      setActiveColumn((prev) => (prev === 'hex' ? 'ascii' : 'hex'));
      setHexNibble(0);
      return;
    }

    // 방향키 이동
    if (e.key === 'ArrowLeft') {
      e.preventDefault();
      if (activeColumn === 'hex' && hexNibble === 1) {
        setHexNibble(0);
      } else {
        const next = Math.max(0, selectedOffset - 1);
        setSelectedOffset(next);
        setHexNibble(activeColumn === 'hex' ? 1 : 0);
        scrollToOffset(next);
      }
      return;
    }
    if (e.key === 'ArrowRight') {
      e.preventDefault();
      if (activeColumn === 'hex' && hexNibble === 0) {
        setHexNibble(1);
      } else {
        const next = Math.min(totalBytes - 1, selectedOffset + 1);
        setSelectedOffset(next);
        setHexNibble(0);
        scrollToOffset(next);
      }
      return;
    }
    if (e.key === 'ArrowUp') {
      e.preventDefault();
      const next = Math.max(0, selectedOffset - BYTES_PER_ROW);
      setSelectedOffset(next);
      scrollToOffset(next);
      return;
    }
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      const next = Math.min(totalBytes - 1, selectedOffset + BYTES_PER_ROW);
      setSelectedOffset(next);
      scrollToOffset(next);
      return;
    }
    if (e.key === 'PageUp') {
      e.preventDefault();
      const next = Math.max(0, selectedOffset - BYTES_PER_ROW * 16);
      setSelectedOffset(next);
      scrollToOffset(next);
      return;
    }
    if (e.key === 'PageDown') {
      e.preventDefault();
      const next = Math.min(totalBytes - 1, selectedOffset + BYTES_PER_ROW * 16);
      setSelectedOffset(next);
      scrollToOffset(next);
      return;
    }
    if (e.key === 'Home') {
      e.preventDefault();
      const rowStart = Math.floor(selectedOffset / BYTES_PER_ROW) * BYTES_PER_ROW;
      setSelectedOffset(rowStart);
      setHexNibble(0);
      return;
    }
    if (e.key === 'End') {
      e.preventDefault();
      const rowEnd = Math.min(totalBytes - 1, Math.floor(selectedOffset / BYTES_PER_ROW) * BYTES_PER_ROW + 15);
      setSelectedOffset(rowEnd);
      setHexNibble(0);
      return;
    }

    // HEX 컬럼에서 16진수 키보드 입력 (0-9, a-f, A-F)
    if (activeColumn === 'hex' && /^[0-9a-fA-F]$/.test(e.key)) {
      e.preventDefault();
      const digit = parseInt(e.key, 16);
      const currentByte = buffer[selectedOffset] || 0;

      if (hexNibble === 0) {
        // 상위 4비트 입력 -> 하위 비트로 전진
        const newByte = (digit << 4) | (currentByte & 0x0f);
        updateByteAt(selectedOffset, newByte);
        setHexNibble(1);
      } else {
        // 하위 4비트 입력 -> 다음 바이트로 전진
        const newByte = (currentByte & 0xf0) | digit;
        updateByteAt(selectedOffset, newByte);
        setHexNibble(0);
        if (selectedOffset < totalBytes - 1) {
          const next = selectedOffset + 1;
          setSelectedOffset(next);
          scrollToOffset(next);
        }
      }
      return;
    }

    // ASCII 컬럼에서 직접 문자 입력 (인쇄 가능한 문자)
    if (activeColumn === 'ascii' && e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      e.preventDefault();
      const code = e.key.charCodeAt(0) & 0xff;
      updateByteAt(selectedOffset, code);
      if (selectedOffset < totalBytes - 1) {
        const next = selectedOffset + 1;
        setSelectedOffset(next);
        scrollToOffset(next);
      }
      return;
    }
  };

  // 검색 실행 (다음/이전)
  const handleSearch = (forward: boolean = true) => {
    if (!searchQuery.trim()) return;

    let found = -1;
    const startPos = matchOffset !== null ? (forward ? matchOffset + 1 : matchOffset - 1) : selectedOffset;

    if (searchMode === 'hex') {
      // "4D 5A" 또는 "4D5A" 문자열을 Uint8Array로 변환
      const hexClean = searchQuery.replace(/\s+/g, '');
      if (hexClean.length % 2 !== 0) {
        alert('16진수 검색어는 2자리 단위(예: 4D 5A)여야 합니다.');
        return;
      }
      const pattern = new Uint8Array(hexClean.length / 2);
      for (let i = 0; i < pattern.length; i++) {
        pattern[i] = parseInt(hexClean.substr(i * 2, 2), 16);
      }
      found = searchHexBytes(buffer, pattern, startPos, forward);
    } else {
      found = searchStringBytes(buffer, searchQuery, startPos, true, forward);
    }

    if (found !== -1) {
      setMatchOffset(found);
      setSelectedOffset(found);
      setHexNibble(0);
      scrollToOffset(found);
    } else {
      alert(`'${searchQuery}' 일치하는 바이너리 패턴을 더 이상 찾을 수 없습니다.`);
    }
  };

  // 주소 이동(Go to Offset) 실행
  const handleGotoOffset = () => {
    let offset = 0;
    const trimmed = gotoInputVal.trim();
    if (trimmed.startsWith('0x') || trimmed.startsWith('0X')) {
      offset = parseInt(trimmed.slice(2), 16);
    } else {
      offset = parseInt(trimmed, 10);
    }

    if (isNaN(offset) || offset < 0 || offset >= totalBytes) {
      alert(`0 ~ ${totalBytes - 1} (0x${formatHexOffset(totalBytes - 1)}) 범위의 주소를 입력하세요.`);
      return;
    }

    setSelectedOffset(offset);
    setHexNibble(0);
    scrollToOffset(offset);
    setShowGotoDialog(false);
    setGotoInputVal('');
  };

  // 16진수 또는 텍스트 클립보드 복사
  const handleCopy = (type: 'hex' | 'ascii') => {
    if (type === 'hex') {
      const hexArr = Array.from(buffer.slice(0, Math.min(buffer.length, 2048))).map(formatByteToHex);
      navigator.clipboard.writeText(hexArr.join(' '));
      setCopiedToast('HEX 바이트가 복사되었습니다 (최대 2KB)');
    } else {
      const asciiArr = Array.from(buffer.slice(0, Math.min(buffer.length, 2048))).map(byteToAsciiChar);
      navigator.clipboard.writeText(asciiArr.join(''));
      setCopiedToast('ASCII 텍스트가 복사되었습니다 (최대 2KB)');
    }
    setTimeout(() => setCopiedToast(null), 2500);
  };

  // 저장 및 메모 반영
  const handleSave = async () => {
    setIsSaving(true);
    setSaveToast(null);
    setErrorToast(null);

    try {
      const mime = title.toLowerCase().endsWith('.bin') ? 'application/octet-stream' : 'application/octet-stream';
      const updatedDataUrl = await uint8ArrayToDataUrl(buffer, mime);
      if (onSave) {
        await onSave(updatedDataUrl);
      }
      setOriginalBuffer(new Uint8Array(buffer));
      setModifiedSet(new Set());
      setHasChanges(false);
      setSaveToast('바이너리 파일이 성공적으로 저장 및 메모에 반영되었습니다!');
      setTimeout(() => setSaveToast(null), 3000);
    } catch (err) {
      console.error('바이너리 저장 실패:', err);
      setErrorToast('바이너리 파일을 저장하는 중 오류가 발생했습니다.');
    } finally {
      setIsSaving(false);
    }
  };

  // 다운로드
  const handleDownload = async () => {
    try {
      const dataUrl = await uint8ArrayToDataUrl(buffer);
      const link = document.createElement('a');
      link.href = dataUrl;
      link.download = title;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (err) {
      console.error('다운로드 생성 실패:', err);
      alert('파일을 다운로드하지 못했습니다.');
    }
  };

  // 닫기
  const handleClose = () => {
    if (hasChanges) {
      if (!confirm('저장되지 않은 바이너리 수정 사항이 있습니다. 닫으시겠습니까?')) {
        return;
      }
    }
    onClose();
  };

  // 현재 커서 위치의 Data Inspector 값
  const inspector: InspectorData | null = useMemo(() => {
    return inspectByteValues(buffer, selectedOffset);
  }, [buffer, selectedOffset]);

  if (!isOpen) return null;

  return (
    <div
      className="fixed inset-0 z-50 bg-black/90 backdrop-blur-sm flex flex-col items-center justify-center p-2 md:p-4 select-none animate-fadeIn"
      onKeyDown={handleKeyDown}
      tabIndex={0}
    >
      {/* 팝업 컨테이너 (울트라에디트 다크 테마) */}
      <div className="w-full max-w-7xl h-[95vh] bg-slate-900 border border-slate-800 rounded-2xl shadow-2xl flex flex-col overflow-hidden text-slate-200">
        {/* 상단 1단: 제목 & 저장 / 다운로드 / 닫기 헤더 */}
        <div className="px-4 py-2.5 bg-slate-900 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3 shrink-0">
          <div className="flex items-center gap-2.5 min-w-0">
            <span className="p-1.5 rounded-lg bg-indigo-500/10 text-indigo-400 border border-indigo-500/20">
              <Binary size={18} />
            </span>
            <div className="truncate">
              <div className="flex items-center gap-2">
                <h2 className="text-sm md:text-base font-bold text-slate-100 truncate font-mono" title={title}>
                  {title}
                </h2>
                <span className="text-[10px] px-1.5 py-0.5 rounded bg-indigo-500/15 text-indigo-300 border border-indigo-500/30 font-bold uppercase shrink-0">
                  UltraEdit HEX
                </span>
                <span className="text-xs text-slate-400 font-mono hidden sm:inline">
                  ({totalBytes.toLocaleString()} bytes / {(totalBytes / 1024).toFixed(1)} KB)
                </span>
                {hasChanges && (
                  <span className="text-[11px] text-amber-400 flex items-center gap-1 font-medium">
                    <span className="w-1.5 h-1.5 rounded-full bg-amber-400 animate-pulse"></span>
                    {modifiedSet.size}바이트 수정됨
                  </span>
                )}
              </div>
            </div>
          </div>

          {/* 우측 상단 액션 */}
          <div className="flex items-center gap-2 shrink-0">
            {saveToast && (
              <div className="hidden sm:flex items-center gap-1.5 text-xs text-emerald-300 bg-emerald-950/40 border border-emerald-500/30 px-2.5 py-1 rounded-lg">
                <CheckCircle size={13} />
                <span>{saveToast}</span>
              </div>
            )}
            {errorToast && (
              <div className="hidden sm:flex items-center gap-1.5 text-xs text-rose-300 bg-rose-950/40 border border-rose-500/30 px-2.5 py-1 rounded-lg">
                <AlertCircle size={13} />
                <span>{errorToast}</span>
              </div>
            )}
            {copiedToast && (
              <div className="hidden sm:flex items-center gap-1.5 text-xs text-indigo-300 bg-indigo-950/40 border border-indigo-500/30 px-2.5 py-1 rounded-lg">
                <Check size={13} />
                <span>{copiedToast}</span>
              </div>
            )}

            {/* 저장 및 반영 버튼 */}
            <button
              type="button"
              onClick={handleSave}
              disabled={isSaving}
              className="px-3.5 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-semibold shadow-md flex items-center gap-1.5 transition-all active:scale-95 disabled:opacity-50"
              title="수정한 바이너리를 메모 첨부파일에 저장 (Ctrl + S)"
            >
              {isSaving ? <Loader2 size={13} className="animate-spin" /> : <Save size={13} />}
              <span>{isSaving ? '저장 중...' : '💾 저장 및 반영'}</span>
            </button>

            {/* 다운로드 버튼 */}
            <button
              type="button"
              onClick={handleDownload}
              className="p-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white rounded-lg border border-slate-700 transition-colors"
              title="바이너리 파일 다운로드"
            >
              <Download size={15} />
            </button>

            {/* 닫기 버튼 */}
            <button
              type="button"
              onClick={handleClose}
              className="p-1.5 bg-slate-800 hover:bg-rose-900/40 text-slate-400 hover:text-rose-300 rounded-lg border border-slate-700 transition-colors"
              title="닫기 (Esc)"
            >
              <X size={16} />
            </button>
          </div>
        </div>

        {/* 상단 2단: 검색 & 오프셋 이동 툴바 */}
        <div className="px-3 py-1.5 bg-slate-950 border-b border-slate-800 flex flex-wrap items-center justify-between gap-2 shrink-0 text-xs font-sans">
          {/* 좌측: 검색 컨트롤 */}
          <div className="flex items-center gap-2 flex-1 max-w-xl">
            <div className="flex items-center bg-slate-900 border border-slate-800 rounded-lg overflow-hidden flex-1 max-w-sm">
              <button
                type="button"
                onClick={() => setSearchMode((prev) => (prev === 'text' ? 'hex' : 'text'))}
                className="px-2 py-1 bg-slate-800 text-[11px] font-mono font-bold text-indigo-300 border-r border-slate-700 hover:bg-slate-750"
                title="클릭하여 텍스트 검색 / HEX 16진수 검색 모드 전환"
              >
                {searchMode.toUpperCase()}
              </button>
              <input
                id="hex-search-input"
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    e.preventDefault();
                    handleSearch(!e.shiftKey);
                  }
                }}
                placeholder={searchMode === 'hex' ? '16진수 검색 (예: 4D 5A 90)...' : '텍스트 검색 (예: PNG, header)...'}
                className="w-full px-2.5 py-1 bg-transparent text-xs text-slate-200 placeholder-slate-500 focus:outline-none font-mono"
              />
              {searchQuery && (
                <button
                  type="button"
                  onClick={() => setSearchQuery('')}
                  className="px-1.5 text-slate-500 hover:text-slate-300"
                >
                  <X size={12} />
                </button>
              )}
            </div>

            <button
              type="button"
              onClick={() => handleSearch(true)}
              className="px-2 py-1 bg-slate-800 hover:bg-slate-750 text-slate-200 border border-slate-700 rounded-lg flex items-center gap-1 font-medium"
              title="다음 찾기 (Enter)"
            >
              <ArrowDown size={12} />
              <span>다음</span>
            </button>
            <button
              type="button"
              onClick={() => handleSearch(false)}
              className="px-2 py-1 bg-slate-800 hover:bg-slate-750 text-slate-200 border border-slate-700 rounded-lg flex items-center gap-1 font-medium"
              title="이전 찾기 (Shift + Enter)"
            >
              <ArrowUp size={12} />
              <span>이전</span>
            </button>

            {/* 주소 이동 버튼 */}
            <button
              type="button"
              onClick={() => setShowGotoDialog(true)}
              className="px-2 py-1 bg-indigo-950/60 hover:bg-indigo-900/60 text-indigo-300 border border-indigo-500/30 rounded-lg flex items-center gap-1 font-medium"
              title="특정 오프셋 주소로 바로 점프 (Ctrl + G)"
            >
              <Navigation size={12} />
              <span>주소 이동</span>
            </button>
          </div>

          {/* 우측 되돌리기 / 복사 도구 */}
          <div className="flex items-center gap-1.5 shrink-0">
            <button
              type="button"
              onClick={handleUndo}
              disabled={historyIndexRef.current < 0}
              className="p-1.5 bg-slate-800 hover:bg-slate-750 text-slate-300 hover:text-white rounded-lg border border-slate-700 disabled:opacity-30"
              title="실행 취소 (Ctrl + Z)"
            >
              <Undo2 size={13} />
            </button>
            <button
              type="button"
              onClick={handleRedo}
              disabled={historyIndexRef.current >= historyRef.current.length - 1}
              className="p-1.5 bg-slate-800 hover:bg-slate-750 text-slate-300 hover:text-white rounded-lg border border-slate-700 disabled:opacity-30"
              title="다시 실행 (Ctrl + Y)"
            >
              <Redo2 size={13} />
            </button>

            <span className="w-px h-4 bg-slate-800 mx-1"></span>

            <button
              type="button"
              onClick={() => handleCopy('hex')}
              className="px-2 py-1 bg-slate-800 hover:bg-slate-750 text-slate-300 hover:text-white rounded-lg border border-slate-700 flex items-center gap-1 text-[11px]"
              title="HEX 바이트 문자열 복사"
            >
              <Copy size={11} />
              <span>HEX 복사</span>
            </button>
            <button
              type="button"
              onClick={() => handleCopy('ascii')}
              className="px-2 py-1 bg-slate-800 hover:bg-slate-750 text-slate-300 hover:text-white rounded-lg border border-slate-700 flex items-center gap-1 text-[11px]"
              title="ASCII 문자열 복사"
            >
              <Copy size={11} />
              <span>ASCII 복사</span>
            </button>
          </div>
        </div>

        {/* 주소 이동(Go to Offset) 팝업 인풋 */}
        {showGotoDialog && (
          <div className="px-4 py-2 bg-indigo-950/80 border-b border-indigo-500/30 flex items-center gap-3 animate-fadeIn">
            <span className="text-xs text-indigo-200 font-semibold">이동할 오프셋 주소 입력:</span>
            <input
              type="text"
              value={gotoInputVal}
              onChange={(e) => setGotoInputVal(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') handleGotoOffset();
                if (e.key === 'Escape') setShowGotoDialog(false);
              }}
              placeholder="예: 0x1A0 또는 4096..."
              className="px-2.5 py-1 bg-slate-900 border border-indigo-500/50 rounded-lg text-xs font-mono text-white placeholder-slate-500 focus:outline-none focus:border-indigo-400 w-64"
              autoFocus
            />
            <button
              type="button"
              onClick={handleGotoOffset}
              className="px-3 py-1 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-semibold"
            >
              이동
            </button>
            <button
              type="button"
              onClick={() => setShowGotoDialog(false)}
              className="text-slate-400 hover:text-slate-200"
            >
              <X size={14} />
            </button>
          </div>
        )}

        {/* 헥스 그리드 컬럼 헤더 */}
        <div className="px-4 py-1.5 bg-slate-950 border-b border-slate-800 flex items-center font-mono text-[11px] font-bold text-slate-500 shrink-0 select-none">
          {/* 주소 헤더 */}
          <div className="w-24 shrink-0 text-slate-400">Offset (h)</div>
          {/* 16진수 헤더 (8바이트 + 8바이트) */}
          <div className="flex-1 max-w-[460px] flex items-center shrink-0">
            <div className="grid grid-cols-8 gap-1.5 w-[216px] text-center text-slate-400">
              <span>00</span><span>01</span><span>02</span><span>03</span>
              <span>04</span><span>05</span><span>06</span><span>07</span>
            </div>
            <span className="w-4 text-center text-slate-700">|</span>
            <div className="grid grid-cols-8 gap-1.5 w-[216px] text-center text-slate-400">
              <span>08</span><span>09</span><span>0A</span><span>0B</span>
              <span>0C</span><span>0D</span><span>0E</span><span>0F</span>
            </div>
          </div>
          {/* ASCII 텍스트 헤더 */}
          <div className="w-44 shrink-0 pl-4 text-slate-400">
            Decoded text
          </div>
        </div>

        {/* 메인 헥스 에디터 그리드 (고성능 가상 스크롤) */}
        <div
          ref={scrollContainerRef}
          onScroll={(e) => setScrollTop((e.target as HTMLElement).scrollTop)}
          className="flex-1 overflow-y-auto overflow-x-auto bg-slate-950 px-4 font-mono text-[13px] leading-[26px] select-none"
        >
          {totalBytes === 0 ? (
            <div className="h-full flex items-center justify-center text-slate-500 text-xs">
              바이너리 데이터가 비어 있습니다.
            </div>
          ) : (
            <div style={{ paddingTop: `${topPadding}px`, paddingBottom: `${bottomPadding}px` }}>
              {Array.from({ length: endIndex - startIndex }).map((_, idx) => {
                const rowIdx = startIndex + idx;
                const rowOffset = rowIdx * BYTES_PER_ROW;

                return (
                  <div
                    key={rowOffset}
                    className="flex items-center hover:bg-slate-900/40 transition-colors h-[26px]"
                  >
                    {/* 오프셋 주소 (8자리 16진수) */}
                    <div className="w-24 shrink-0 text-indigo-400/80 font-bold select-text">
                      {formatHexOffset(rowOffset, 8)}
                    </div>

                    {/* 16바이트 헥스 열 */}
                    <div className="flex-1 max-w-[460px] flex items-center shrink-0">
                      {/* 앞 8바이트 */}
                      <div className="grid grid-cols-8 gap-1.5 w-[216px] text-center">
                        {Array.from({ length: 8 }).map((_, cIdx) => {
                          const byteOffset = rowOffset + cIdx;
                          if (byteOffset >= totalBytes) {
                            return <span key={cIdx} className="text-slate-800">..</span>;
                          }
                          const val = buffer[byteOffset];
                          const isSelected = selectedOffset === byteOffset;
                          const isModified = modifiedSet.has(byteOffset);
                          const isMatch = matchOffset === byteOffset;

                          return (
                            <span
                              key={cIdx}
                              onClick={() => {
                                setSelectedOffset(byteOffset);
                                setActiveColumn('hex');
                                setHexNibble(0);
                              }}
                              className={`cursor-pointer rounded px-0.5 transition-all ${
                                isSelected && activeColumn === 'hex'
                                  ? 'bg-indigo-600 text-white font-bold ring-2 ring-indigo-400'
                                  : isSelected
                                  ? 'bg-indigo-950 text-indigo-300 font-bold'
                                  : isMatch
                                  ? 'bg-amber-500/30 text-amber-300 font-bold'
                                  : isModified
                                  ? 'text-amber-400 font-bold'
                                  : 'text-slate-300 hover:bg-slate-800 hover:text-white'
                              }`}
                            >
                              {formatByteToHex(val)}
                            </span>
                          );
                        })}
                      </div>

                      {/* 8바이트 구분자 */}
                      <span className="w-4 text-center text-slate-800">|</span>

                      {/* 뒤 8바이트 */}
                      <div className="grid grid-cols-8 gap-1.5 w-[216px] text-center">
                        {Array.from({ length: 8 }).map((_, cIdx) => {
                          const byteOffset = rowOffset + 8 + cIdx;
                          if (byteOffset >= totalBytes) {
                            return <span key={cIdx} className="text-slate-800">..</span>;
                          }
                          const val = buffer[byteOffset];
                          const isSelected = selectedOffset === byteOffset;
                          const isModified = modifiedSet.has(byteOffset);
                          const isMatch = matchOffset === byteOffset;

                          return (
                            <span
                              key={cIdx}
                              onClick={() => {
                                setSelectedOffset(byteOffset);
                                setActiveColumn('hex');
                                setHexNibble(0);
                              }}
                              className={`cursor-pointer rounded px-0.5 transition-all ${
                                isSelected && activeColumn === 'hex'
                                  ? 'bg-indigo-600 text-white font-bold ring-2 ring-indigo-400'
                                  : isSelected
                                  ? 'bg-indigo-950 text-indigo-300 font-bold'
                                  : isMatch
                                  ? 'bg-amber-500/30 text-amber-300 font-bold'
                                  : isModified
                                  ? 'text-amber-400 font-bold'
                                  : 'text-slate-300 hover:bg-slate-800 hover:text-white'
                              }`}
                            >
                              {formatByteToHex(val)}
                            </span>
                          );
                        })}
                      </div>
                    </div>

                    {/* ASCII 텍스트 열 */}
                    <div className="w-44 shrink-0 pl-4 flex items-center font-mono tracking-wider">
                      {Array.from({ length: 16 }).map((_, cIdx) => {
                        const byteOffset = rowOffset + cIdx;
                        if (byteOffset >= totalBytes) {
                          return <span key={cIdx} className="text-slate-800">.</span>;
                        }
                        const val = buffer[byteOffset];
                        const char = byteToAsciiChar(val);
                        const isSelected = selectedOffset === byteOffset;
                        const isModified = modifiedSet.has(byteOffset);

                        return (
                          <span
                            key={cIdx}
                            onClick={() => {
                              setSelectedOffset(byteOffset);
                              setActiveColumn('ascii');
                            }}
                            className={`cursor-pointer px-[2px] rounded ${
                              isSelected && activeColumn === 'ascii'
                                ? 'bg-indigo-600 text-white font-bold ring-1 ring-indigo-400'
                                : isSelected
                                ? 'bg-indigo-950 text-indigo-300 font-bold'
                                : isModified
                                ? 'text-amber-400 font-bold'
                                : 'text-slate-400 hover:bg-slate-800 hover:text-white'
                            }`}
                          >
                            {char}
                          </span>
                        );
                      })}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        {/* 하단 패널: UltraEdit Data Inspector & 상태 바 */}
        <div className="px-4 py-2.5 bg-slate-900 border-t border-slate-800 flex flex-wrap items-center justify-between gap-3 shrink-0 text-xs font-mono">
          {/* 좌측: 현재 커서 위치 및 기본 정보 */}
          <div className="flex items-center gap-3">
            <span className="bg-slate-950 px-2 py-1 rounded border border-slate-800 text-indigo-300 font-bold">
              Offset: 0x{formatHexOffset(selectedOffset, 8)} ({selectedOffset.toLocaleString()})
            </span>
            <span className="text-slate-400 hidden sm:inline">
              열: {activeColumn.toUpperCase()} {activeColumn === 'hex' ? `(Nibble: ${hexNibble === 0 ? '상위' : '하위'})` : ''}
            </span>
          </div>

          {/* 중앙: Data Inspector (자료형별 해석) */}
          {inspector && (
            <div className="flex items-center gap-2 overflow-x-auto text-[11px] py-0.5">
              <span className="bg-slate-950 px-2 py-0.5 rounded border border-slate-800 text-slate-300">
                Uint8: <strong className="text-emerald-400">{inspector.uint8}</strong> (Int8: {inspector.int8})
              </span>
              {inspector.uint16LE !== null && (
                <span className="bg-slate-950 px-2 py-0.5 rounded border border-slate-800 text-slate-300">
                  Uint16 LE: <strong className="text-cyan-400">{inspector.uint16LE}</strong> (BE: {inspector.uint16BE})
                </span>
              )}
              {inspector.uint32LE !== null && (
                <span className="bg-slate-950 px-2 py-0.5 rounded border border-slate-800 text-slate-300 hidden md:inline">
                  Uint32 LE: <strong className="text-amber-400">{inspector.uint32LE}</strong>
                </span>
              )}
              <span className="bg-slate-950 px-2 py-0.5 rounded border border-slate-800 text-slate-300 hidden lg:inline">
                2진수: <strong className="text-purple-400">{inspector.binary}</strong>
              </span>
            </div>
          )}

          {/* 우측 단축키 안내 */}
          <div className="text-[11px] text-slate-500 hidden xl:flex items-center gap-2 shrink-0">
            <span>Tab: 열 전환</span>
            <span>•</span>
            <span>방향키: 이동</span>
            <span>•</span>
            <span>0-F/문자: 직접 입력 수정</span>
          </div>
        </div>
      </div>
    </div>
  );
};
