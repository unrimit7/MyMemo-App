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
  Navigation,
  WrapText,
  FileText
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
  decodeBytesToString,
  encodeStringToBytes,
  type InspectorData,
} from '../utils/hexViewer';
import { useFloatingWindow } from './useFloatingWindow';

// 울트라에디트(UltraEdit)의 상징적인 10/010 16진수 아이콘 (Image 2 & 3)
export const UltraEditHexIcon: React.FC<{ size?: number; className?: string }> = ({ size = 18, className = '' }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" className={className}>
    {/* 접힌 모서리가 있는 문서 윤곽 */}
    <path
      d="M4 3C4 2.44772 4.44772 2 5 2H14L20 8V21C20 21.5523 19.5523 22 19 22H5C4.44772 22 4 21.5523 4 21V3Z"
      fill="currentColor"
      fillOpacity="0.12"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinejoin="round"
    />
    <path d="M14 2V8H20" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" />
    {/* 내부 10 / 010 폰트 */}
    <text x="6" y="13" fontFamily="monospace" fontSize="6.5" fontWeight="900" fill="currentColor" letterSpacing="-0.5">
      10
    </text>
    <text x="5.5" y="19.5" fontFamily="monospace" fontSize="6" fontWeight="900" fill="currentColor" letterSpacing="-0.5">
      010
    </text>
  </svg>
);

// 울트라에디트 일반 텍스트 문서 아이콘
export const UltraEditTextIcon: React.FC<{ size?: number; className?: string }> = ({ size = 18, className = '' }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" className={className}>
    <path
      d="M4 3C4 2.44772 4.44772 2 5 2H14L20 8V21C20 21.5523 19.5523 22 19 22H5C4.44772 22 4 21.5523 4 21V3Z"
      fill="currentColor"
      fillOpacity="0.12"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinejoin="round"
    />
    <path d="M14 2V8H20" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" />
    <line x1="7.5" y1="12" x2="16.5" y2="12" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
    <line x1="7.5" y1="15.5" x2="16.5" y2="15.5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
    <line x1="7.5" y1="19" x2="13.5" y2="19" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
  </svg>
);

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
  // 1. 보기 모드: 'hex' (울트라에디트 16진수 모드) | 'text' (일반 텍스트/코드 편집 모드)
  const floating = useFloatingWindow({ id: 'hex-editor', defaultWidth: 960, defaultHeight: 820 });
  const [viewMode, setViewMode] = useState<'hex' | 'text'>('hex');
  const [textContent, setTextContent] = useState<string>('');
  const [encoding, setEncoding] = useState<'utf-8' | 'euc-kr' | 'latin1'>('utf-8');
  const [wordWrap, setWordWrap] = useState<boolean>(true);
  const [cursorLine, setCursorLine] = useState<number>(1);
  const [cursorCol, setCursorCol] = useState<number>(1);
  const [textMatchIndex, setTextMatchIndex] = useState<number | null>(null);

  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const gutterRef = useRef<HTMLDivElement>(null);

  // 2. 바이너리 데이터 버퍼
  const [buffer, setBuffer] = useState<Uint8Array>(new Uint8Array(0));
  const [originalBuffer, setOriginalBuffer] = useState<Uint8Array>(new Uint8Array(0));
  const [modifiedSet, setModifiedSet] = useState<Set<number>>(new Set());

  // 3. 커서 및 편집 상태 (헥스 모드)
  const [selectedOffset, setSelectedOffset] = useState<number>(0);
  const [activeColumn, setActiveColumn] = useState<'hex' | 'ascii'>('hex');
  const [hexNibble, setHexNibble] = useState<0 | 1>(0); // 0: 상위 4비트, 1: 하위 4비트

  // 4. Undo / Redo 스택 (헥스 모드)
  const historyRef = useRef<{ offset: number; oldVal: number; newVal: number }[]>([]);
  const historyIndexRef = useRef<number>(-1);

  // 5. 검색 상태
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [searchMode, setSearchMode] = useState<'text' | 'hex'>('text');
  const [matchOffset, setMatchOffset] = useState<number | null>(null);

  // 6. 주소 이동(Go to Offset) 상태
  const [showGotoDialog, setShowGotoDialog] = useState<boolean>(false);
  const [gotoInputVal, setGotoInputVal] = useState<string>('');

  // 7. 상태 플래그
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [hasChanges, setHasChanges] = useState<boolean>(false);
  const [saveToast, setSaveToast] = useState<string | null>(null);
  const [errorToast, setErrorToast] = useState<string | null>(null);
  const [copiedToast, setCopiedToast] = useState<string | null>(null);

  // 8. 스크롤 및 뷰포트 (가상 스크롤용)
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
      setTextContent(decodeBytesToString(bytes, encoding));
      setModifiedSet(new Set());
      setSelectedOffset(0);
      setHexNibble(0);
      setActiveColumn('hex');
      setHasChanges(false);
      setMatchOffset(null);
      setTextMatchIndex(null);
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
  }, [isOpen, viewMode]);

  const totalBytes = buffer.length;
  const totalRows = Math.ceil(totalBytes / BYTES_PER_ROW);

  // 텍스트 모드 총 라인 수
  const totalTextLines = useMemo(() => textContent.split('\n').length, [textContent]);
  const lineNumbersText = useMemo(() => {
    let res = '';
    for (let i = 1; i <= totalTextLines; i++) {
      res += (i === 1 ? '' : '\n') + i;
    }
    return res;
  }, [totalTextLines]);

  // 가상 스크롤 계산 (헥스 모드)
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

  // 커서가 화면 밖으로 벗어났을 때 자동 스크롤 (헥스 모드)
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

  // 텍스트 모드로 전환
  const switchToTextMode = (targetEncoding = encoding) => {
    const decoded = decodeBytesToString(buffer, targetEncoding);
    setTextContent(decoded);
    setViewMode('text');
    setTextMatchIndex(null);
  };

  // 울트라에디트 헥스 모드로 전환
  const switchToHexMode = () => {
    const encodedBytes = encodeStringToBytes(textContent, encoding);
    setBuffer(encodedBytes);

    // 원본 버퍼 대비 수정된 오프셋 계산
    const newModified = new Set<number>();
    const minLen = Math.min(encodedBytes.length, originalBuffer.length);
    for (let i = 0; i < minLen; i++) {
      if (encodedBytes[i] !== originalBuffer[i]) {
        newModified.add(i);
      }
    }
    for (let i = minLen; i < encodedBytes.length; i++) {
      newModified.add(i);
    }
    setModifiedSet(newModified);

    if (newModified.size > 0 || encodedBytes.length !== originalBuffer.length) {
      setHasChanges(true);
    }

    if (selectedOffset >= encodedBytes.length) {
      setSelectedOffset(Math.max(0, encodedBytes.length - 1));
    }
    setViewMode('hex');
  };

  // 모드 토글
  const toggleViewMode = () => {
    if (viewMode === 'hex') {
      switchToTextMode();
    } else {
      switchToHexMode();
    }
  };

  // 인코딩 변경 (텍스트 모드)
  const handleEncodingChange = (newEnc: 'utf-8' | 'euc-kr' | 'latin1') => {
    setEncoding(newEnc);
    const decoded = decodeBytesToString(buffer, newEnc);
    setTextContent(decoded);
    setSaveToast(`인코딩이 ${newEnc.toUpperCase()}(으)로 변경되었습니다.`);
    setTimeout(() => setSaveToast(null), 2500);
  };

  // 바이트 값 수정 적용 (헥스 모드)
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
    if (viewMode === 'hex') {
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
    } else {
      document.execCommand('undo');
    }
  };

  // 다시 실행(Redo)
  const handleRedo = () => {
    if (viewMode === 'hex') {
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
    } else {
      document.execCommand('redo');
    }
  };

  // 텍스트 커서 위치 업데이트
  const updateCursorPos = () => {
    const textarea = textareaRef.current;
    if (!textarea) return;
    const pos = textarea.selectionStart;
    const textBefore = textarea.value.substring(0, pos);
    const lines = textBefore.split('\n');
    setCursorLine(lines.length);
    setCursorCol((lines[lines.length - 1] || '').length + 1);
  };

  // 텍스트 모드 스크롤 동기화
  const handleTextareaScroll = (e: React.UIEvent<HTMLTextAreaElement>) => {
    if (gutterRef.current) {
      gutterRef.current.scrollTop = e.currentTarget.scrollTop;
    }
  };

  // 텍스트 모드 키 입력 처리
  const handleTextareaKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    // Ctrl + H: 울트라에디트 모드 전환
    if ((e.ctrlKey || e.metaKey) && (e.key === 'h' || e.key === 'H')) {
      e.preventDefault();
      switchToHexMode();
      return;
    }
    // Ctrl + S: 저장
    if ((e.ctrlKey || e.metaKey) && (e.key === 's' || e.key === 'S')) {
      e.preventDefault();
      handleSave();
      return;
    }
    // Tab: 2칸 들여쓰기 삽입
    if (e.key === 'Tab') {
      e.preventDefault();
      const textarea = e.currentTarget;
      const start = textarea.selectionStart;
      const end = textarea.selectionEnd;
      const val = textarea.value;
      const insert = '  ';
      const newVal = val.substring(0, start) + insert + val.substring(end);
      setTextContent(newVal);
      setHasChanges(true);
      requestAnimationFrame(() => {
        textarea.selectionStart = textarea.selectionEnd = start + insert.length;
        updateCursorPos();
      });
    }
  };

  // 키보드 네비게이션 및 단축키 (헥스 모드 및 모달 전체)
  const handleKeyDown = (e: React.KeyboardEvent) => {
    // Ctrl + H: 울트라에디트 표준 단축키 (텍스트 <-> 10 010 헥스 모드 전환)
    if ((e.ctrlKey || e.metaKey) && (e.key === 'h' || e.key === 'H')) {
      e.preventDefault();
      toggleViewMode();
      return;
    }

    // Ctrl + S (저장)
    if ((e.ctrlKey || e.metaKey) && (e.key === 's' || e.key === 'S')) {
      e.preventDefault();
      handleSave();
      return;
    }

    // 텍스트 모드에서는 텍스트 영역 내부 입력이 우선
    if (viewMode === 'text') return;

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
      const step = Math.floor(viewportHeight / ROW_HEIGHT) * BYTES_PER_ROW;
      const next = Math.max(0, selectedOffset - step);
      setSelectedOffset(next);
      scrollToOffset(next);
      return;
    }
    if (e.key === 'PageDown') {
      e.preventDefault();
      const step = Math.floor(viewportHeight / ROW_HEIGHT) * BYTES_PER_ROW;
      const next = Math.min(totalBytes - 1, selectedOffset + step);
      setSelectedOffset(next);
      scrollToOffset(next);
      return;
    }
    if (e.key === 'Home') {
      e.preventDefault();
      setSelectedOffset(0);
      scrollToOffset(0);
      return;
    }
    if (e.key === 'End') {
      e.preventDefault();
      setSelectedOffset(totalBytes - 1);
      scrollToOffset(totalBytes - 1);
      return;
    }

    // HEX 열에서 0-9, a-f 문자 직접 수정
    if (activeColumn === 'hex') {
      const key = e.key.toLowerCase();
      if (/^[0-9a-f]$/.test(key)) {
        e.preventDefault();
        const nibbleVal = parseInt(key, 16);
        const currentByte = buffer[selectedOffset] || 0;
        let newByte = currentByte;

        if (hexNibble === 0) {
          // 상위 4비트 수정
          newByte = (nibbleVal << 4) | (currentByte & 0x0f);
          updateByteAt(selectedOffset, newByte);
          setHexNibble(1);
        } else {
          // 하위 4비트 수정 후 다음 바이트로 이동
          newByte = (currentByte & 0xf0) | nibbleVal;
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
    }

    // ASCII 열에서 인쇄 가능한 단일 문자 직접 수정
    if (activeColumn === 'ascii') {
      if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        e.preventDefault();
        const charCode = e.key.charCodeAt(0);
        if (charCode >= 0 && charCode <= 255) {
          updateByteAt(selectedOffset, charCode);
          if (selectedOffset < totalBytes - 1) {
            const next = selectedOffset + 1;
            setSelectedOffset(next);
            scrollToOffset(next);
          }
        }
        return;
      }
    }
  };

  // 검색 실행 (모드별 대응)
  const handleSearch = (forward: boolean) => {
    if (!searchQuery.trim()) return;

    if (viewMode === 'text') {
      // 텍스트 모드 검색
      const lowerText = textContent.toLowerCase();
      const lowerQuery = searchQuery.toLowerCase();
      const textarea = textareaRef.current;
      const curPos = textarea?.selectionStart || 0;

      if (forward) {
        const startFrom = textMatchIndex !== null && textMatchIndex === curPos ? curPos + 1 : curPos;
        let nextIdx = lowerText.indexOf(lowerQuery, startFrom);
        if (nextIdx === -1) {
          nextIdx = lowerText.indexOf(lowerQuery, 0); // 처음부터 랩어라운드
        }
        if (nextIdx !== -1) {
          setTextMatchIndex(nextIdx);
          highlightTextSelection(nextIdx, nextIdx + lowerQuery.length);
        } else {
          setErrorToast('일치하는 텍스트가 없습니다.');
          setTimeout(() => setErrorToast(null), 2000);
        }
      } else {
        const startFrom = textMatchIndex !== null && textMatchIndex === curPos ? curPos - 1 : curPos - 1;
        let prevIdx = lowerText.lastIndexOf(lowerQuery, Math.max(0, startFrom));
        if (prevIdx === -1) {
          prevIdx = lowerText.lastIndexOf(lowerQuery); // 끝에서부터 랩어라운드
        }
        if (prevIdx !== -1) {
          setTextMatchIndex(prevIdx);
          highlightTextSelection(prevIdx, prevIdx + lowerQuery.length);
        } else {
          setErrorToast('일치하는 텍스트가 없습니다.');
          setTimeout(() => setErrorToast(null), 2000);
        }
      }
    } else {
      // 헥스 모드 검색
      let start = selectedOffset;
      if (matchOffset !== null) {
        start = forward ? matchOffset + 1 : matchOffset - 1;
      }

      let found = -1;
      if (searchMode === 'hex') {
        const hexTokens = searchQuery
          .trim()
          .replace(/0x/gi, '')
          .split(/[\s,]+/)
          .filter(Boolean)
          .map((h) => parseInt(h, 16))
          .filter((n) => !isNaN(n) && n >= 0 && n <= 255);

        if (hexTokens.length === 0) {
          setErrorToast('올바른 16진수 바이트를 입력하세요 (예: 4D 5A)');
          setTimeout(() => setErrorToast(null), 2500);
          return;
        }

        const pattern = new Uint8Array(hexTokens);
        found = searchHexBytes(buffer, pattern, start, forward);
      } else {
        found = searchStringBytes(buffer, searchQuery, start, forward, false);
      }

      if (found !== -1) {
        setMatchOffset(found);
        setSelectedOffset(found);
        setHexNibble(0);
        scrollToOffset(found);
      } else {
        setErrorToast('일치하는 항목을 찾지 못했습니다.');
        setTimeout(() => setErrorToast(null), 2500);
      }
    }
  };

  const highlightTextSelection = (start: number, end: number) => {
    const textarea = textareaRef.current;
    if (!textarea) return;
    textarea.focus();
    textarea.setSelectionRange(start, end);
    const textBefore = textContent.substring(0, start);
    const lineNum = textBefore.split('\n').length;
    const approxLineHeight = 24;
    const targetScroll = Math.max(0, (lineNum - 6) * approxLineHeight);
    textarea.scrollTop = targetScroll;
    if (gutterRef.current) gutterRef.current.scrollTop = targetScroll;
    updateCursorPos();
  };

  // 특정 오프셋 주소로 이동 (Go to Offset)
  const handleGotoOffset = () => {
    const trimmed = gotoInputVal.trim();
    if (!trimmed) return;

    let offset = -1;
    if (trimmed.startsWith('0x') || trimmed.startsWith('0X') || /^[0-9a-fA-F]+h$/i.test(trimmed)) {
      offset = parseInt(trimmed.replace(/h$/i, '').replace(/^0x/i, ''), 16);
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

  // 클립보드 복사
  const handleCopy = (type: 'hex' | 'ascii' | 'text') => {
    if (type === 'hex') {
      const hexArr = Array.from(buffer.slice(0, Math.min(buffer.length, 2048))).map(formatByteToHex);
      navigator.clipboard.writeText(hexArr.join(' '));
      setCopiedToast('HEX 바이트가 복사되었습니다 (최대 2KB)');
    } else if (type === 'ascii') {
      const asciiArr = Array.from(buffer.slice(0, Math.min(buffer.length, 2048))).map(byteToAsciiChar);
      navigator.clipboard.writeText(asciiArr.join(''));
      setCopiedToast('ASCII 텍스트가 복사되었습니다 (최대 2KB)');
    } else {
      navigator.clipboard.writeText(textContent);
      setCopiedToast('전체 텍스트가 복사되었습니다.');
    }
    setTimeout(() => setCopiedToast(null), 2500);
  };

  // 저장 및 메모 반영
  const handleSave = async () => {
    setIsSaving(true);
    setSaveToast(null);
    setErrorToast(null);

    try {
      let currentBytes = buffer;
      if (viewMode === 'text') {
        currentBytes = encodeStringToBytes(textContent, encoding);
        setBuffer(currentBytes);
      }
      const updatedDataUrl = await uint8ArrayToDataUrl(currentBytes);
      if (onSave) {
        await onSave(updatedDataUrl);
      }
      setOriginalBuffer(new Uint8Array(currentBytes));
      setModifiedSet(new Set());
      setHasChanges(false);
      setSaveToast('성공적으로 저장 및 메모에 반영되었습니다!');
      setTimeout(() => setSaveToast(null), 3000);
    } catch (err) {
      console.error('저장 실패:', err);
      setErrorToast('저장하는 중 오류가 발생했습니다.');
    } finally {
      setIsSaving(false);
    }
  };

  // 기기로 다운로드
  const handleDownload = async () => {
    try {
      let currentBytes = buffer;
      if (viewMode === 'text') {
        currentBytes = encodeStringToBytes(textContent, encoding);
      }
      const dataUrl = await uint8ArrayToDataUrl(currentBytes);
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
      if (!confirm('저장되지 않은 수정 사항이 있습니다. 닫으시겠습니까?')) {
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
      className={floating.modalWrapperClass}
      onKeyDown={handleKeyDown}
      tabIndex={0}
    >
      {/* 팝업 컨테이너 (울트라에디트 다크 테마) */}
      <div {...floating.windowContainerProps} className={floating.windowContainerProps.className + ' text-slate-200'}>
        {/* 상단 1단: 제목 & 울트라에디트 보기 전환(10 010 <-> 텍스트) & 액션 */}
        <div 
          {...floating.headerProps}
          className="px-4 py-2.5 bg-slate-900 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3 shrink-0"
        >
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
                  {viewMode === 'hex' ? 'UltraEdit HEX' : 'UltraEdit TEXT'}
                </span>
                <span className="text-xs text-slate-400 font-mono hidden sm:inline">
                  ({(viewMode === 'text' ? new TextEncoder().encode(textContent).length : totalBytes).toLocaleString()} bytes / {((viewMode === 'text' ? new TextEncoder().encode(textContent).length : totalBytes) / 1024).toFixed(1)} KB)
                </span>
                {hasChanges && (
                  <span className="text-[11px] text-amber-400 flex items-center gap-1 font-medium">
                    <span className="w-1.5 h-1.5 rounded-full bg-amber-400 animate-pulse"></span>
                    수정됨
                  </span>
                )}
              </div>
            </div>
          </div>

          {/* 중앙: 울트라에디트 상징적인 보기 모드 전환 컨트롤 (Image 2 & 3 대응) */}
          <div className="flex items-center bg-slate-950/90 p-1 rounded-xl border border-slate-750 shadow-inner gap-1">
            <button
              type="button"
              onClick={() => {
                if (viewMode !== 'text') switchToTextMode();
              }}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-all cursor-pointer ${
                viewMode === 'text'
                  ? 'bg-indigo-600 text-white shadow-sm ring-1 ring-indigo-400 font-bold'
                  : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800'
              }`}
              title="일반 텍스트/코드 보기 및 편집 모드로 전환 (단축키: Ctrl + H)"
            >
              <UltraEditTextIcon size={16} />
              <span>텍스트 모드</span>
            </button>

            <button
              type="button"
              onClick={() => {
                if (viewMode !== 'hex') switchToHexMode();
              }}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-all cursor-pointer ${
                viewMode === 'hex'
                  ? 'bg-indigo-600 text-white shadow-sm ring-1 ring-indigo-400 font-bold'
                  : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800'
              }`}
              title="울트라에디트 16진수(HEX) 뷰어 및 바이트 직접 편집 모드로 전환 (단축키: Ctrl + H)"
            >
              <UltraEditHexIcon size={16} />
              <span>10 010 헥스 모드</span>
            </button>
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
              className="px-3.5 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-semibold shadow-md flex items-center gap-1.5 transition-all active:scale-95 disabled:opacity-50 cursor-pointer"
              title="수정한 내용을 메모 첨부파일에 저장 (Ctrl + S)"
            >
              {isSaving ? <Loader2 size={13} className="animate-spin" /> : <Save size={13} />}
              <span>{isSaving ? '저장 중...' : '💾 저장 및 반영'}</span>
            </button>

            {/* 다운로드 버튼 */}
            <button
              type="button"
              onClick={handleDownload}
              className="p-1.5 bg-slate-800 hover:bg-slate-750 text-slate-300 hover:text-white rounded-lg border border-slate-700 transition-colors cursor-pointer"
              title="기기로 파일 다운로드"
            >
              <Download size={15} />
            </button>

            {floating.toggleFloatingButton}

            {/* 닫기 버튼 */}
            <button
              type="button"
              onClick={handleClose}
              className="p-1.5 bg-slate-800 hover:bg-rose-900/40 text-slate-400 hover:text-rose-300 rounded-lg border border-slate-700 transition-colors cursor-pointer"
              title="닫기 (Esc)"
            >
              <X size={16} />
            </button>
          </div>
        </div>

        {/* 상단 2단: 모드별 검색 & 도구 툴바 */}
        <div className="px-3 py-1.5 bg-slate-950 border-b border-slate-800 flex flex-wrap items-center justify-between gap-2 shrink-0 text-xs font-sans">
          {/* 좌측: 검색 컨트롤 */}
          <div className="flex items-center gap-2 flex-1 max-w-xl">
            <div className="flex items-center bg-slate-900 border border-slate-800 rounded-lg overflow-hidden flex-1 max-w-sm">
              {viewMode === 'hex' ? (
                <button
                  type="button"
                  onClick={() => setSearchMode((prev) => (prev === 'text' ? 'hex' : 'text'))}
                  className="px-2 py-1 bg-slate-800 text-[11px] font-mono font-bold text-indigo-300 border-r border-slate-700 hover:bg-slate-750 cursor-pointer"
                  title="클릭하여 텍스트 검색 / HEX 16진수 검색 모드 전환"
                >
                  {searchMode.toUpperCase()}
                </button>
              ) : (
                <span className="px-2 py-1 bg-slate-800 text-[11px] font-mono font-bold text-indigo-300 border-r border-slate-700 flex items-center gap-1">
                  <FileText size={11} />
                  <span>TEXT</span>
                </span>
              )}
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
                placeholder={
                  viewMode === 'hex'
                    ? searchMode === 'hex'
                      ? '16진수 검색 (예: 4D 5A 90)...'
                      : '텍스트 검색 (예: PNG, header)...'
                    : '텍스트 검색 (Enter / Shift+Enter)...'
                }
                className="w-full px-2.5 py-1 bg-transparent text-xs text-slate-200 placeholder-slate-500 focus:outline-none font-mono"
              />
              {searchQuery && (
                <button
                  type="button"
                  onClick={() => setSearchQuery('')}
                  className="px-1.5 text-slate-500 hover:text-slate-300 cursor-pointer"
                >
                  <X size={12} />
                </button>
              )}
            </div>

            <button
              type="button"
              onClick={() => handleSearch(true)}
              className="px-2.5 py-1 bg-slate-800 hover:bg-slate-750 text-slate-200 border border-slate-700 rounded-lg flex items-center gap-1 font-medium cursor-pointer"
              title="다음 찾기 (Enter)"
            >
              <ArrowDown size={12} />
              <span>다음</span>
            </button>
            <button
              type="button"
              onClick={() => handleSearch(false)}
              className="px-2.5 py-1 bg-slate-800 hover:bg-slate-750 text-slate-200 border border-slate-700 rounded-lg flex items-center gap-1 font-medium cursor-pointer"
              title="이전 찾기 (Shift + Enter)"
            >
              <ArrowUp size={12} />
              <span>이전</span>
            </button>

            {/* 헥스 모드 전용: 주소 이동 버튼 */}
            {viewMode === 'hex' && (
              <button
                type="button"
                onClick={() => setShowGotoDialog(true)}
                className="px-2.5 py-1 bg-indigo-950/60 hover:bg-indigo-900/60 text-indigo-300 border border-indigo-500/30 rounded-lg flex items-center gap-1 font-medium cursor-pointer"
                title="특정 오프셋 주소로 바로 점프 (Ctrl + G)"
              >
                <Navigation size={12} />
                <span>주소 이동</span>
              </button>
            )}
          </div>

          {/* 우측 도구 모음 */}
          <div className="flex items-center gap-1.5 shrink-0">
            {/* 텍스트 모드 전용 도구: 인코딩 선택 & 자동 줄바꿈 */}
            {viewMode === 'text' && (
              <>
                <div className="flex items-center gap-1 text-[11px] text-slate-400 bg-slate-900 border border-slate-800 rounded-lg px-2 py-0.5">
                  <span>인코딩:</span>
                  <select
                    value={encoding}
                    onChange={(e) => handleEncodingChange(e.target.value as any)}
                    className="bg-transparent text-indigo-300 font-semibold focus:outline-none cursor-pointer"
                    title="문서 문자 인코딩 선택 (한글 깨짐 시 EUC-KR 선택)"
                  >
                    <option value="utf-8" className="bg-slate-900 text-slate-200">UTF-8</option>
                    <option value="euc-kr" className="bg-slate-900 text-slate-200">EUC-KR (한국어)</option>
                    <option value="latin1" className="bg-slate-900 text-slate-200">Latin-1 (Binary 1:1)</option>
                  </select>
                </div>

                <button
                  type="button"
                  onClick={() => setWordWrap((prev) => !prev)}
                  className={`px-2 py-1 rounded-lg border text-[11px] flex items-center gap-1 cursor-pointer transition-colors ${
                    wordWrap
                      ? 'bg-indigo-950/70 text-indigo-300 border-indigo-500/40'
                      : 'bg-slate-800 text-slate-400 border-slate-700 hover:text-slate-200'
                  }`}
                  title="긴 줄 자동 줄 바꿈 켜기/끄기"
                >
                  <WrapText size={12} />
                  <span>줄 바꿈 {wordWrap ? 'ON' : 'OFF'}</span>
                </button>

                <span className="w-px h-4 bg-slate-800 mx-1"></span>
              </>
            )}

            {/* 실행 취소 / 다시 실행 */}
            <button
              type="button"
              onClick={handleUndo}
              disabled={viewMode === 'hex' ? historyIndexRef.current < 0 : false}
              className="p-1.5 bg-slate-800 hover:bg-slate-750 text-slate-300 hover:text-white rounded-lg border border-slate-700 disabled:opacity-30 cursor-pointer"
              title="실행 취소 (Ctrl + Z)"
            >
              <Undo2 size={13} />
            </button>
            <button
              type="button"
              onClick={handleRedo}
              disabled={viewMode === 'hex' ? historyIndexRef.current >= historyRef.current.length - 1 : false}
              className="p-1.5 bg-slate-800 hover:bg-slate-750 text-slate-300 hover:text-white rounded-lg border border-slate-700 disabled:opacity-30 cursor-pointer"
              title="다시 실행 (Ctrl + Y)"
            >
              <Redo2 size={13} />
            </button>

            <span className="w-px h-4 bg-slate-800 mx-1"></span>

            {/* 복사 버튼 */}
            {viewMode === 'hex' ? (
              <>
                <button
                  type="button"
                  onClick={() => handleCopy('hex')}
                  className="px-2 py-1 bg-slate-800 hover:bg-slate-750 text-slate-300 hover:text-white rounded-lg border border-slate-700 flex items-center gap-1 text-[11px] cursor-pointer"
                  title="HEX 바이트 문자열 복사"
                >
                  <Copy size={11} />
                  <span>HEX 복사</span>
                </button>
                <button
                  type="button"
                  onClick={() => handleCopy('ascii')}
                  className="px-2 py-1 bg-slate-800 hover:bg-slate-750 text-slate-300 hover:text-white rounded-lg border border-slate-700 flex items-center gap-1 text-[11px] cursor-pointer"
                  title="ASCII 문자열 복사"
                >
                  <Copy size={11} />
                  <span>ASCII 복사</span>
                </button>
              </>
            ) : (
              <button
                type="button"
                onClick={() => handleCopy('text')}
                className="px-2 py-1 bg-slate-800 hover:bg-slate-750 text-slate-300 hover:text-white rounded-lg border border-slate-700 flex items-center gap-1 text-[11px] cursor-pointer"
                title="전체 텍스트 복사"
              >
                <Copy size={11} />
                <span>텍스트 복사</span>
              </button>
            )}
          </div>
        </div>

        {/* 주소 이동(Go to Offset) 팝업 인풋 */}
        {showGotoDialog && viewMode === 'hex' && (
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
              className="px-3 py-1 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-semibold cursor-pointer"
            >
              이동
            </button>
            <button
              type="button"
              onClick={() => setShowGotoDialog(false)}
              className="px-2.5 py-1 text-slate-400 hover:text-slate-200 text-xs cursor-pointer"
            >
              취소
            </button>
          </div>
        )}

        {/* 메인 뷰 본문: 16진수 헥스 모드 vs 텍스트 모드 분기 */}
        {viewMode === 'hex' ? (
          <>
            {/* 가상 스크롤 헥스 덤프 컨테이너 */}
            <div
              ref={scrollContainerRef}
              onScroll={(e) => setScrollTop(e.currentTarget.scrollTop)}
              className="flex-1 overflow-y-auto bg-slate-950 font-mono text-xs sm:text-sm select-text focus:outline-none relative"
              tabIndex={0}
            >
              {/* 열 헤더 (고정) */}
              <div className="sticky top-0 z-20 bg-slate-900/95 backdrop-blur-xs border-b border-slate-800 text-slate-400 text-xs font-bold py-1.5 px-4 flex items-center">
                <span className="w-24 shrink-0 text-indigo-300">Offset (h)</span>
                <div className="w-[380px] shrink-0 flex items-center">
                  <span className="w-[180px] tracking-widest text-slate-300">
                    00 01 02 03 04 05 06 07
                  </span>
                  <span className="text-slate-600 px-2 font-light">|</span>
                  <span className="w-[180px] tracking-widest text-slate-300">
                    08 09 0A 0B 0C 0D 0E 0F
                  </span>
                </div>
                <span className="pl-6 text-indigo-300">Decoded text</span>
              </div>

              {/* 가상 스크롤 상단 여백 */}
              <div style={{ height: topPadding }} />

              {/* 렌더링된 행 목록 */}
              <div className="px-4 py-1">
                {Array.from({ length: Math.max(0, endIndex - startIndex) }).map((_, rIdx) => {
                  const rowIndex = startIndex + rIdx;
                  const rowOffset = rowIndex * BYTES_PER_ROW;
                  if (rowOffset >= totalBytes && totalBytes > 0) return null;

                  return (
                    <div
                      key={rowOffset}
                      className="flex items-center leading-[26px] h-[26px] hover:bg-slate-900/60 transition-colors group"
                    >
                      {/* 1. 오프셋 주소 열 */}
                      <span className="w-24 shrink-0 text-indigo-400/90 font-bold select-none group-hover:text-indigo-300">
                        {formatHexOffset(rowOffset, 8)}
                      </span>

                      {/* 2. HEX 바이트 열 (8 + 8 분할) */}
                      <div className="w-[380px] shrink-0 flex items-center font-mono">
                        {/* 첫 8바이트 */}
                        <div className="w-[180px] flex items-center justify-between">
                          {Array.from({ length: 8 }).map((_, col) => {
                            const offset = rowOffset + col;
                            if (offset >= totalBytes) {
                              return (
                                <span key={col} className="w-5 text-center text-slate-700 select-none">
                                  ..
                                </span>
                              );
                            }
                            const byte = buffer[offset];
                            const isSelected = offset === selectedOffset;
                            const isMatch = offset === matchOffset;
                            const isModified = modifiedSet.has(offset);

                            let cellStyle = 'text-slate-300 hover:text-white';
                            if (isModified) cellStyle = 'text-amber-400 font-bold';
                            if (isMatch) cellStyle = 'bg-yellow-500/30 text-yellow-200 ring-1 ring-yellow-400 font-bold';
                            if (isSelected) {
                              cellStyle =
                                activeColumn === 'hex'
                                  ? 'bg-indigo-600 text-white font-bold rounded ring-2 ring-indigo-400 shadow-md'
                                  : 'bg-indigo-900/50 text-indigo-200 ring-1 ring-indigo-500 rounded';
                            }

                            return (
                              <span
                                key={col}
                                onClick={() => {
                                  setSelectedOffset(offset);
                                  setActiveColumn('hex');
                                  setHexNibble(0);
                                }}
                                className={`w-5 text-center cursor-pointer select-none rounded transition-all text-xs ${cellStyle}`}
                                title={`오프셋: 0x${formatHexOffset(offset)} (${offset})\n값: 0x${formatByteToHex(byte)} (${byte})`}
                              >
                                {formatByteToHex(byte)}
                              </span>
                            );
                          })}
                        </div>

                        <span className="text-slate-700 px-2 font-light select-none">|</span>

                        {/* 뒤 8바이트 */}
                        <div className="w-[180px] flex items-center justify-between">
                          {Array.from({ length: 8 }).map((_, col) => {
                            const offset = rowOffset + 8 + col;
                            if (offset >= totalBytes) {
                              return (
                                <span key={col} className="w-5 text-center text-slate-700 select-none">
                                  ..
                                </span>
                              );
                            }
                            const byte = buffer[offset];
                            const isSelected = offset === selectedOffset;
                            const isMatch = offset === matchOffset;
                            const isModified = modifiedSet.has(offset);

                            let cellStyle = 'text-slate-300 hover:text-white';
                            if (isModified) cellStyle = 'text-amber-400 font-bold';
                            if (isMatch) cellStyle = 'bg-yellow-500/30 text-yellow-200 ring-1 ring-yellow-400 font-bold';
                            if (isSelected) {
                              cellStyle =
                                activeColumn === 'hex'
                                  ? 'bg-indigo-600 text-white font-bold rounded ring-2 ring-indigo-400 shadow-md'
                                  : 'bg-indigo-900/50 text-indigo-200 ring-1 ring-indigo-500 rounded';
                            }

                            return (
                              <span
                                key={col}
                                onClick={() => {
                                  setSelectedOffset(offset);
                                  setActiveColumn('hex');
                                  setHexNibble(0);
                                }}
                                className={`w-5 text-center cursor-pointer select-none rounded transition-all text-xs ${cellStyle}`}
                                title={`오프셋: 0x${formatHexOffset(offset)} (${offset})\n값: 0x${formatByteToHex(byte)} (${byte})`}
                              >
                                {formatByteToHex(byte)}
                              </span>
                            );
                          })}
                        </div>
                      </div>

                      {/* 3. Decoded text (ASCII) 열 */}
                      <div className="pl-6 flex items-center font-mono">
                        {Array.from({ length: 16 }).map((_, col) => {
                          const offset = rowOffset + col;
                          if (offset >= totalBytes) {
                            return (
                              <span key={col} className="w-3 text-center text-slate-700 select-none">
                                .
                              </span>
                            );
                          }
                          const byte = buffer[offset];
                          const char = byteToAsciiChar(byte);
                          const isSelected = offset === selectedOffset;
                          const isMatch = offset === matchOffset;
                          const isModified = modifiedSet.has(offset);

                          let cellStyle = byte >= 32 && byte <= 126 ? 'text-slate-200' : 'text-slate-600';
                          if (isModified) cellStyle = 'text-amber-400 font-bold';
                          if (isMatch) cellStyle = 'bg-yellow-500/40 text-yellow-100 font-bold';
                          if (isSelected) {
                            cellStyle =
                              activeColumn === 'ascii'
                                ? 'bg-indigo-600 text-white font-bold rounded ring-2 ring-indigo-400 shadow-md'
                                : 'bg-indigo-900/50 text-indigo-200 ring-1 ring-indigo-500 rounded';
                          }

                          return (
                            <span
                              key={col}
                              onClick={() => {
                                setSelectedOffset(offset);
                                setActiveColumn('ascii');
                              }}
                              className={`w-3.5 text-center cursor-pointer select-none transition-all text-xs ${cellStyle}`}
                              title={`문자: '${char}' (코드: ${byte})`}
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

              {/* 가상 스크롤 하단 여백 */}
              <div style={{ height: bottomPadding }} />
            </div>

            {/* 하단: 울트라에디트 Data Inspector 패널 */}
            <div className="px-4 py-2 bg-slate-900 border-t border-slate-800 flex flex-wrap items-center justify-between gap-3 text-xs font-mono shrink-0">
              {/* 왼쪽: 현재 선택 오프셋 정보 */}
              <div className="flex items-center gap-3">
                <span className="text-slate-400">
                  오프셋: <strong className="text-indigo-300">0x{formatHexOffset(selectedOffset)}</strong> ({selectedOffset} / {totalBytes})
                </span>
                <span className="text-slate-600">|</span>
                <span className="text-slate-400">
                  열: <strong className="text-slate-200 uppercase">{activeColumn}</strong>
                </span>
                {activeColumn === 'hex' && (
                  <span className="text-[11px] text-slate-500">({hexNibble === 0 ? '상위' : '하위'} 니블)</span>
                )}
              </div>

              {/* 중앙: 실시간 Data Inspector 파서 */}
              {inspector && (
                <div className="flex flex-wrap items-center gap-x-4 gap-y-1 text-[11px]">
                  <div className="flex items-center gap-1">
                    <span className="text-slate-500">Uint8:</span>
                    <span className="text-emerald-400 font-bold">{inspector.uint8}</span>
                  </div>
                  <div className="flex items-center gap-1">
                    <span className="text-slate-500">Int8:</span>
                    <span className="text-emerald-400 font-bold">{inspector.int8}</span>
                  </div>
                  {inspector.uint16LE !== null && (
                    <div className="flex items-center gap-1" title="16-bit Unsigned Little/Big Endian">
                      <span className="text-slate-500">Uint16:</span>
                      <span className="text-sky-300 font-bold">{inspector.uint16LE}</span>
                      <span className="text-slate-600 text-[10px]">(BE: {inspector.uint16BE})</span>
                    </div>
                  )}
                  {inspector.uint32LE !== null && (
                    <div className="flex items-center gap-1" title="32-bit Unsigned Little/Big Endian">
                      <span className="text-slate-500">Uint32:</span>
                      <span className="text-violet-300 font-bold">{inspector.uint32LE}</span>
                      <span className="text-slate-600 text-[10px]">(BE: {inspector.uint32BE})</span>
                    </div>
                  )}
                  <div className="flex items-center gap-1" title="8-bit Binary representation">
                    <span className="text-slate-500">Bit:</span>
                    <span className="text-amber-300 font-bold">{inspector.binary}</span>
                  </div>
                </div>
              )}

              {/* 오른쪽: 단축키 안내 */}
              <div className="text-[11px] text-slate-500 hidden xl:flex items-center gap-2">
                <span><kbd className="px-1 py-0.5 rounded bg-slate-800 text-slate-300 border border-slate-700">Tab</kbd> 열 전환</span>
                <span><kbd className="px-1 py-0.5 rounded bg-slate-800 text-slate-300 border border-slate-700">Ctrl+H</kbd> 모드 전환</span>
                <span><kbd className="px-1 py-0.5 rounded bg-slate-800 text-slate-300 border border-slate-700">Ctrl+G</kbd> 이동</span>
                <span><kbd className="px-1 py-0.5 rounded bg-slate-800 text-slate-300 border border-slate-700">Ctrl+S</kbd> 저장</span>
              </div>
            </div>
          </>
        ) : (
          /* 텍스트 편집기 뷰 (UltraEdit Text Mode) */
          <div className="flex-1 flex flex-col min-h-0 bg-slate-950 text-slate-100 relative">
            <div className="flex-1 flex overflow-hidden">
              {/* 좌측 행 번호 (Gutter) */}
              <div
                ref={gutterRef}
                className="w-14 sm:w-16 bg-slate-925/90 text-slate-500 font-mono text-xs text-right pr-3 pt-3 select-none overflow-hidden border-r border-slate-800 leading-6 shrink-0"
              >
                <pre className="m-0 p-0 font-mono text-xs leading-6 select-none text-slate-500 pointer-events-none">
                  {lineNumbersText}
                </pre>
              </div>

              {/* 우측 텍스트 본문 편집기 (Textarea) */}
              <textarea
                ref={textareaRef}
                value={textContent}
                onChange={(e) => {
                  setTextContent(e.target.value);
                  setHasChanges(true);
                  updateCursorPos();
                }}
                onKeyDown={handleTextareaKeyDown}
                onKeyUp={updateCursorPos}
                onClick={updateCursorPos}
                onSelect={updateCursorPos}
                onScroll={handleTextareaScroll}
                spellCheck={false}
                autoCapitalize="off"
                autoComplete="off"
                placeholder="텍스트 내용이 여기에 표시됩니다..."
                className={`flex-1 p-3 bg-transparent text-slate-200 font-mono text-xs sm:text-sm leading-6 resize-none focus:outline-none selection:bg-indigo-600/50 caret-indigo-400 overflow-auto ${
                  wordWrap ? 'whitespace-pre-wrap break-words' : 'whitespace-pre overflow-x-auto'
                }`}
              />
            </div>

            {/* 텍스트 모드 하단 상태 바 */}
            <div className="px-4 py-2 bg-slate-900 border-t border-slate-800 flex flex-wrap items-center justify-between text-xs text-slate-400 font-mono shrink-0 gap-2">
              <div className="flex items-center gap-4">
                <span className="text-slate-300">
                  줄 <strong className="text-indigo-400">{cursorLine}</strong>, 열 <strong className="text-indigo-400">{cursorCol}</strong>
                </span>
                <span>
                  총 <strong className="text-slate-200">{totalTextLines.toLocaleString()}</strong>줄 / <strong className="text-slate-200">{textContent.length.toLocaleString()}</strong>자
                </span>
                <span>
                  크기: <strong className="text-slate-200">{(new TextEncoder().encode(textContent).length).toLocaleString()}</strong> bytes
                </span>
              </div>

              <div className="flex items-center gap-3">
                <span className="text-slate-400">
                  인코딩: <strong className="text-indigo-300 uppercase">{encoding}</strong>
                </span>
                <span className="text-slate-600">|</span>
                <span className="text-slate-400">
                  모드 전환: <kbd className="px-1.5 py-0.5 rounded bg-slate-800 text-slate-200 text-[10px] border border-slate-700">Ctrl + H</kbd>
                </span>
              </div>
            </div>
          </div>
        )}

        {floating.resizeHandle}
      </div>
    </div>
  );
};
