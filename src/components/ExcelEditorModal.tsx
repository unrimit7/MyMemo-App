import React, { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import {
  X,
  Save,
  Download,
  Plus,
  Trash2,
  Undo2,
  Redo2,
  ArrowUp,
  ArrowDown,
  FileSpreadsheet,
  CheckCircle,
  AlertCircle,
  Loader2
} from 'lucide-react';
import { parseExcelWorkbook, generateExcelDataUrl } from '../utils/excelViewer';

interface ExcelEditorModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: string;
  title: string;
  attId?: string;
  onSave?: (updatedDataUrl: string) => Promise<void> | void;
}

// 열 인덱스 -> 알파벳 변환 (0 -> 'A', 25 -> 'Z', 26 -> 'AA')
function getColumnLabel(index: number): string {
  let label = '';
  let num = index;
  while (num >= 0) {
    label = String.fromCharCode((num % 26) + 65) + label;
    num = Math.floor(num / 26) - 1;
  }
  return label;
}

export const ExcelEditorModal: React.FC<ExcelEditorModalProps> = ({
  isOpen,
  onClose,
  data,
  title,
  attId,
  onSave,
}) => {
  const isCsv = title.toLowerCase().endsWith('.csv');
  const format: 'xlsx' | 'csv' = isCsv ? 'csv' : 'xlsx';

  // 시트별 2차원 데이터 상태
  const [sheets, setSheets] = useState<Record<string, (string | number)[][]>>({});
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [activeSheetName, setActiveSheetName] = useState<string>('');

  // 선택된 셀 및 편집 중인 셀
  const [selectedCell, setSelectedCell] = useState<{ r: number; c: number } | null>(null);
  const [editingCell, setEditingCell] = useState<{ r: number; c: number } | null>(null);
  const [editInputVal, setEditInputVal] = useState<string>('');

  // 상태 관리: 변경 여부, 저장 진행 상태, 알림 메시지
  const [hasChanges, setHasChanges] = useState<boolean>(false);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [saveToast, setSaveToast] = useState<string | null>(null);
  const [errorToast, setErrorToast] = useState<string | null>(null);

  // Undo / Redo 스택
  const historyRef = useRef<Record<string, (string | number)[][]>[]>([]);
  const historyIndexRef = useRef<number>(-1);
  const isInternalUpdateRef = useRef<boolean>(false);

  const editInputRef = useRef<HTMLInputElement>(null);
  const gridContainerRef = useRef<HTMLDivElement>(null);

  // 데이터 초기화
  useEffect(() => {
    if (!isOpen || !data) return;

    try {
      const parsed = parseExcelWorkbook(data, attId);
      const initialSheets = JSON.parse(JSON.stringify(parsed.sheets));
      const initialSheetNames = [...parsed.sheetNames];
      const initialActiveSheet = initialSheetNames[0] || 'Sheet1';

      // 빈 데이터 보정 (최소 5행 5열)
      initialSheetNames.forEach((name) => {
        const rows = initialSheets[name] || [];
        if (rows.length === 0) {
          initialSheets[name] = [
            ['', '', '', '', ''],
            ['', '', '', '', ''],
            ['', '', '', '', ''],
          ];
        }
      });

      setSheets(initialSheets);
      setSheetNames(initialSheetNames);
      setActiveSheetName(initialActiveSheet);
      setSelectedCell({ r: 0, c: 0 });
      setEditingCell(null);
      setHasChanges(false);

      // 히스토리 초기화
      historyRef.current = [JSON.parse(JSON.stringify(initialSheets))];
      historyIndexRef.current = 0;
    } catch (err) {
      console.error('엑셀 데이터 로드 실패:', err);
      setErrorToast('스프레드시트 데이터를 로드하지 못했습니다.');
    }
  }, [isOpen, data, attId]);

  // 히스토리에 현재 스냅샷 푸시
  const pushHistory = useCallback((newSheets: Record<string, (string | number)[][]>) => {
    if (isInternalUpdateRef.current) return;
    const nextIndex = historyIndexRef.current + 1;
    const cloned = JSON.parse(JSON.stringify(newSheets));
    historyRef.current = [...historyRef.current.slice(0, nextIndex), cloned];
    historyIndexRef.current = nextIndex;
    setHasChanges(true);
  }, []);

  // 현재 활성 시트 행 데이터
  const currentRows = useMemo(() => {
    return sheets[activeSheetName] || [];
  }, [sheets, activeSheetName]);

  // 현재 활성 시트 최대 열 수
  const colCount = useMemo(() => {
    let max = 0;
    currentRows.forEach((r) => {
      max = Math.max(max, r.length);
    });
    return Math.max(max, 5); // 최소 5개 열 보장
  }, [currentRows]);

  // 현재 활성 시트 표시 행 수 (최소 8개 행 보장)
  const rowCount = Math.max(currentRows.length, 8);

  // 특정 셀 값 가져오기
  const getCellValue = (r: number, c: number): string => {
    const row = currentRows[r];
    if (!row) return '';
    const val = row[c];
    return val !== undefined && val !== null ? String(val) : '';
  };

  // 셀 편집 시작
  const startEditing = (r: number, c: number) => {
    setSelectedCell({ r, c });
    setEditingCell({ r, c });
    setEditInputVal(getCellValue(r, c));
  };

  // 셀 편집 완료 및 적용
  const commitEditing = () => {
    if (!editingCell) return;
    const { r, c } = editingCell;
    const oldVal = getCellValue(r, c);

    if (oldVal !== editInputVal) {
      const newRows: (string | number)[][] = currentRows.map((row) => [...row]);

      // 필요한 행과 열까지 확장
      while (newRows.length <= r) {
        newRows.push([]);
      }
      const targetRow = newRows[r];
      while (targetRow.length <= c) {
        targetRow.push('');
      }

      // 숫자 변환 시도
      const trimmed = editInputVal.trim();
      const num = Number(trimmed);
      targetRow[c] = trimmed !== '' && !isNaN(num) && !trimmed.startsWith('0') ? num : editInputVal;

      const newSheets = {
        ...sheets,
        [activeSheetName]: newRows,
      };

      setSheets(newSheets);
      pushHistory(newSheets);
    }

    setEditingCell(null);
  };

  // 편집 인풋 포커스
  useEffect(() => {
    if (editingCell && editInputRef.current) {
      editInputRef.current.focus();
      editInputRef.current.select();
    }
  }, [editingCell]);

  // 되돌리기(Undo)
  const handleUndo = () => {
    if (historyIndexRef.current > 0) {
      historyIndexRef.current -= 1;
      const prevSheets = JSON.parse(JSON.stringify(historyRef.current[historyIndexRef.current]));
      isInternalUpdateRef.current = true;
      setSheets(prevSheets);
      setEditingCell(null);
      setTimeout(() => {
        isInternalUpdateRef.current = false;
      }, 50);
    }
  };

  // 다시 실행(Redo)
  const handleRedo = () => {
    if (historyIndexRef.current < historyRef.current.length - 1) {
      historyIndexRef.current += 1;
      const nextSheets = JSON.parse(JSON.stringify(historyRef.current[historyIndexRef.current]));
      isInternalUpdateRef.current = true;
      setSheets(nextSheets);
      setEditingCell(null);
      setTimeout(() => {
        isInternalUpdateRef.current = false;
      }, 50);
    }
  };

  // 행 추가 (+ 행 추가)
  const handleAddRow = (insertBelow: boolean = true) => {
    commitEditing();
    const newRows: (string | number)[][] = currentRows.map((r) => [...r]);
    const insertIndex = selectedCell && insertBelow ? selectedCell.r + 1 : newRows.length;
    const emptyRow = Array.from({ length: colCount }).map(() => '');

    newRows.splice(insertIndex, 0, emptyRow);

    const newSheets = { ...sheets, [activeSheetName]: newRows };
    setSheets(newSheets);
    pushHistory(newSheets);
    setSelectedCell({ r: insertIndex, c: selectedCell?.c || 0 });
  };

  // 행 삭제 (- 행 삭제)
  const handleDeleteRow = () => {
    if (currentRows.length <= 1) {
      alert('최소 1개의 행은 유지되어야 합니다.');
      return;
    }
    commitEditing();
    const rIdx = selectedCell?.r ?? currentRows.length - 1;
    const newRows = currentRows.filter((_, idx) => idx !== rIdx);

    const newSheets = { ...sheets, [activeSheetName]: newRows };
    setSheets(newSheets);
    pushHistory(newSheets);
    setSelectedCell({ r: Math.max(0, Math.min(newRows.length - 1, rIdx)), c: selectedCell?.c || 0 });
  };

  // 열 추가 (+ 열 추가)
  const handleAddCol = (insertRight: boolean = true) => {
    commitEditing();
    const insertIndex = selectedCell && insertRight ? selectedCell.c + 1 : colCount;
    const newRows = currentRows.map((row) => {
      const cloned = [...row];
      while (cloned.length < insertIndex) {
        cloned.push('');
      }
      cloned.splice(insertIndex, 0, '');
      return cloned;
    });

    const newSheets = { ...sheets, [activeSheetName]: newRows };
    setSheets(newSheets);
    pushHistory(newSheets);
    setSelectedCell({ r: selectedCell?.r || 0, c: insertIndex });
  };

  // 열 삭제 (- 열 삭제)
  const handleDeleteCol = () => {
    if (colCount <= 1) {
      alert('최소 1개의 열은 유지되어야 합니다.');
      return;
    }
    commitEditing();
    const cIdx = selectedCell?.c ?? colCount - 1;
    const newRows = currentRows.map((row) => row.filter((_, idx) => idx !== cIdx));

    const newSheets = { ...sheets, [activeSheetName]: newRows };
    setSheets(newSheets);
    pushHistory(newSheets);
    setSelectedCell({ r: selectedCell?.r || 0, c: Math.max(0, cIdx - 1) });
  };

  // 정렬 (선택된 열 기준 정렬, 1행 헤더 보존)
  const handleSort = (ascending: boolean) => {
    commitEditing();
    const cIdx = selectedCell?.c || 0;
    if (currentRows.length <= 2) return;

    const header = currentRows[0];
    const dataRows = currentRows.slice(1);

    dataRows.sort((a, b) => {
      const valA = String(a[cIdx] ?? '').trim();
      const valB = String(b[cIdx] ?? '').trim();

      const numA = Number(valA);
      const numB = Number(valB);

      if (!isNaN(numA) && !isNaN(numB) && valA !== '' && valB !== '') {
        return ascending ? numA - numB : numB - numA;
      }
      return ascending ? valA.localeCompare(valB, 'ko') : valB.localeCompare(valA, 'ko');
    });

    const newRows = [header, ...dataRows];
    const newSheets = { ...sheets, [activeSheetName]: newRows };
    setSheets(newSheets);
    pushHistory(newSheets);
  };

  // 시트 추가
  const handleAddSheet = () => {
    let num = sheetNames.length + 1;
    let newName = `Sheet${num}`;
    while (sheetNames.includes(newName)) {
      num++;
      newName = `Sheet${num}`;
    }

    const newSheetNames = [...sheetNames, newName];
    const newSheets = {
      ...sheets,
      [newName]: [
        ['', '', '', '', ''],
        ['', '', '', '', ''],
        ['', '', '', '', ''],
      ],
    };

    setSheetNames(newSheetNames);
    setSheets(newSheets);
    setActiveSheetName(newName);
    pushHistory(newSheets);
  };

  // 시트 이름 변경
  const handleRenameSheet = (oldName: string) => {
    const newName = prompt('새 시트 이름을 입력하세요:', oldName);
    if (!newName || !newName.trim() || newName.trim() === oldName) return;
    const safeName = newName.trim().slice(0, 31);

    if (sheetNames.includes(safeName)) {
      alert('이미 같은 이름의 시트가 존재합니다.');
      return;
    }

    const newSheetNames = sheetNames.map((n) => (n === oldName ? safeName : n));
    const newSheets: Record<string, (string | number)[][]> = {};
    sheetNames.forEach((n) => {
      if (n === oldName) {
        newSheets[safeName] = sheets[oldName] || [];
      } else {
        newSheets[n] = sheets[n];
      }
    });

    setSheetNames(newSheetNames);
    setSheets(newSheets);
    if (activeSheetName === oldName) {
      setActiveSheetName(safeName);
    }
    pushHistory(newSheets);
  };

  // 시트 삭제
  const handleDeleteSheet = (sheetNameToDelete: string) => {
    if (sheetNames.length <= 1) {
      alert('최소 1개의 시트는 유지되어야 합니다.');
      return;
    }
    if (!confirm(`'${sheetNameToDelete}' 시트를 정말 삭제하시겠습니까?`)) return;

    const newSheetNames = sheetNames.filter((n) => n !== sheetNameToDelete);
    const newSheets = { ...sheets };
    delete newSheets[sheetNameToDelete];

    setSheetNames(newSheetNames);
    setSheets(newSheets);
    if (activeSheetName === sheetNameToDelete) {
      setActiveSheetName(newSheetNames[0]);
    }
    pushHistory(newSheets);
  };

  // 키보드 네비게이션 및 단축키 핸들러
  const handleKeyDown = (e: React.KeyboardEvent) => {
    // Ctrl + Z / Ctrl + Y
    if ((e.ctrlKey || e.metaKey) && e.key === 'z') {
      e.preventDefault();
      if (e.shiftKey) {
        handleRedo();
      } else {
        handleUndo();
      }
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

    // 편집 모드 아닐 때
    if (!editingCell && selectedCell) {
      if (e.key === 'Enter' || e.key === 'F2') {
        e.preventDefault();
        startEditing(selectedCell.r, selectedCell.c);
        return;
      }
      if (e.key === 'ArrowUp') {
        e.preventDefault();
        setSelectedCell({ r: Math.max(0, selectedCell.r - 1), c: selectedCell.c });
      } else if (e.key === 'ArrowDown') {
        e.preventDefault();
        setSelectedCell({ r: Math.min(rowCount - 1, selectedCell.r + 1), c: selectedCell.c });
      } else if (e.key === 'ArrowLeft') {
        e.preventDefault();
        setSelectedCell({ r: selectedCell.r, c: Math.max(0, selectedCell.c - 1) });
      } else if (e.key === 'ArrowRight' || e.key === 'Tab') {
        e.preventDefault();
        setSelectedCell({ r: selectedCell.r, c: Math.min(colCount - 1, selectedCell.c + 1) });
      } else if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        // 셀 내용 지우기
        const newRows = currentRows.map((row) => [...row]);
        if (newRows[selectedCell.r]) {
          newRows[selectedCell.r][selectedCell.c] = '';
          const newSheets = { ...sheets, [activeSheetName]: newRows };
          setSheets(newSheets);
          pushHistory(newSheets);
        }
      } else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        // 일반 타이핑 시작 시 바로 편집 모드로 진입
        setSelectedCell({ r: selectedCell.r, c: selectedCell.c });
        setEditingCell({ r: selectedCell.r, c: selectedCell.c });
        setEditInputVal(e.key);
      }
    }
  };

  // 편집 인풋 키 핸들러
  const handleEditInputKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      commitEditing();
      // 아래 셀로 이동
      if (editingCell) {
        setSelectedCell({ r: Math.min(rowCount - 1, editingCell.r + 1), c: editingCell.c });
      }
    } else if (e.key === 'Tab') {
      e.preventDefault();
      commitEditing();
      if (editingCell) {
        if (e.shiftKey) {
          setSelectedCell({ r: editingCell.r, c: Math.max(0, editingCell.c - 1) });
        } else {
          setSelectedCell({ r: editingCell.r, c: Math.min(colCount - 1, editingCell.c + 1) });
        }
      }
    } else if (e.key === 'Escape') {
      e.preventDefault();
      setEditingCell(null);
    }
  };

  // 저장 및 메모 반영 핸들러
  const handleSave = async () => {
    commitEditing();
    setIsSaving(true);
    setSaveToast(null);
    setErrorToast(null);

    try {
      const updatedDataUrl = generateExcelDataUrl(sheets, format);
      if (onSave) {
        await onSave(updatedDataUrl);
      }
      setHasChanges(false);
      setSaveToast('스프레드시트가 성공적으로 저장 및 메모에 반영되었습니다!');
      setTimeout(() => setSaveToast(null), 3000);
    } catch (err) {
      console.error('스프레드시트 저장 실패:', err);
      setErrorToast('스프레드시트를 저장하는 중 오류가 발생했습니다.');
    } finally {
      setIsSaving(false);
    }
  };

  // 다운로드 핸들러
  const handleDownload = () => {
    commitEditing();
    try {
      const exportUrl = generateExcelDataUrl(sheets, format);
      const link = document.createElement('a');
      link.href = exportUrl;
      link.download = title.endsWith(`.${format}`) ? title : `${title}.${format}`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (err) {
      console.error('다운로드 생성 실패:', err);
      alert('파일을 생성하지 못했습니다.');
    }
  };

  // 닫기 핸들러 (저장되지 않은 변경사항 확인)
  const handleClose = () => {
    if (hasChanges) {
      if (!confirm('저장되지 않은 변경 내용이 있습니다. 저장하지 않고 편집기를 닫으시겠습니까?')) {
        return;
      }
    }
    onClose();
  };

  if (!isOpen) return null;

  return (
    <div
      className="fixed inset-0 z-50 bg-slate-950/90 backdrop-blur-md flex flex-col items-center justify-center p-2 md:p-4 select-none animate-fadeIn"
      onKeyDown={handleKeyDown}
      tabIndex={0}
    >
      {/* 팝업 컨테이너 */}
      <div className="w-full max-w-7xl h-[95vh] bg-slate-900 border border-slate-800 rounded-2xl shadow-2xl flex flex-col overflow-hidden">
        {/* 상단 1단: 제목 & 저장 / 다운로드 / 닫기 헤더 */}
        <div className="px-4 py-2.5 bg-slate-900/95 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3 shrink-0">
          <div className="flex items-center gap-2.5 min-w-0">
            <span className="p-1.5 rounded-lg bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
              <FileSpreadsheet size={18} />
            </span>
            <div className="truncate">
              <div className="flex items-center gap-2">
                <h2 className="text-sm md:text-base font-bold text-slate-100 truncate" title={title}>
                  {title}
                </h2>
                <span className="text-[10px] px-1.5 py-0.5 rounded bg-emerald-500/10 text-emerald-300 border border-emerald-500/20 font-bold uppercase shrink-0">
                  {format} 편집기
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

          {/* 우측 저장 및 닫기 버튼 */}
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

            {/* 저장 및 반영 버튼 */}
            <button
              type="button"
              onClick={handleSave}
              disabled={isSaving}
              className="px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-semibold shadow-md flex items-center gap-1.5 transition-all active:scale-95 disabled:opacity-50"
              title="편집된 표를 메모 첨부파일 원본에 저장 및 본문 뷰어에 즉시 반영 (Ctrl + S)"
            >
              {isSaving ? <Loader2 size={13} className="animate-spin" /> : <Save size={13} />}
              <span>{isSaving ? '저장 중...' : '💾 저장 및 반영'}</span>
            </button>

            {/* 다운로드 버튼 */}
            <button
              type="button"
              onClick={handleDownload}
              className="p-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white rounded-lg border border-slate-700 transition-colors"
              title="편집된 엑셀/CSV 파일 다운로드"
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

        {/* 상단 2단: 스프레드시트 툴바 (행/열 조작, 정렬, 되돌리기) */}
        <div className="px-3 py-1.5 bg-slate-950/70 border-b border-slate-800 flex flex-wrap items-center justify-between gap-2 shrink-0 text-xs">
          {/* 좌측: 행/열 및 정렬 도구 */}
          <div className="flex items-center gap-1.5 overflow-x-auto py-0.5">
            {/* 행 조작 */}
            <div className="flex items-center gap-1 bg-slate-900 border border-slate-800 rounded-lg p-0.5">
              <button
                type="button"
                onClick={() => handleAddRow(true)}
                className="px-2 py-1 hover:bg-slate-800 text-slate-200 hover:text-emerald-300 rounded flex items-center gap-1 font-medium"
                title="선택된 행 아래에 빈 행 추가"
              >
                <Plus size={12} className="text-emerald-400" />
                <span>행 추가</span>
              </button>
              <button
                type="button"
                onClick={handleDeleteRow}
                className="px-2 py-1 hover:bg-slate-800 text-slate-200 hover:text-rose-300 rounded flex items-center gap-1 font-medium"
                title="선택된 행 삭제"
              >
                <Trash2 size={12} className="text-rose-400" />
                <span>행 삭제</span>
              </button>
            </div>

            {/* 열 조작 */}
            <div className="flex items-center gap-1 bg-slate-900 border border-slate-800 rounded-lg p-0.5">
              <button
                type="button"
                onClick={() => handleAddCol(true)}
                className="px-2 py-1 hover:bg-slate-800 text-slate-200 hover:text-emerald-300 rounded flex items-center gap-1 font-medium"
                title="선택된 열 오른쪽에 빈 열 추가"
              >
                <Plus size={12} className="text-emerald-400" />
                <span>열 추가</span>
              </button>
              <button
                type="button"
                onClick={handleDeleteCol}
                className="px-2 py-1 hover:bg-slate-800 text-slate-200 hover:text-rose-300 rounded flex items-center gap-1 font-medium"
                title="선택된 열 삭제"
              >
                <Trash2 size={12} className="text-rose-400" />
                <span>열 삭제</span>
              </button>
            </div>

            {/* 정렬 조작 */}
            <div className="flex items-center gap-1 bg-slate-900 border border-slate-800 rounded-lg p-0.5">
              <button
                type="button"
                onClick={() => handleSort(true)}
                className="px-2 py-1 hover:bg-slate-800 text-slate-200 hover:text-indigo-300 rounded flex items-center gap-1 font-medium"
                title={`선택된 열(${getColumnLabel(selectedCell?.c || 0)}) 기준 오름차순 정렬 (헤더 제외)`}
              >
                <ArrowUp size={12} className="text-indigo-400" />
                <span>오름차순</span>
              </button>
              <button
                type="button"
                onClick={() => handleSort(false)}
                className="px-2 py-1 hover:bg-slate-800 text-slate-200 hover:text-indigo-300 rounded flex items-center gap-1 font-medium"
                title={`선택된 열(${getColumnLabel(selectedCell?.c || 0)}) 기준 내림차순 정렬 (헤더 제외)`}
              >
                <ArrowDown size={12} className="text-indigo-400" />
                <span>내림차순</span>
              </button>
            </div>

            {/* 되돌리기 / 다시 실행 */}
            <div className="flex items-center gap-1 bg-slate-900 border border-slate-800 rounded-lg p-0.5">
              <button
                type="button"
                onClick={handleUndo}
                disabled={historyIndexRef.current <= 0}
                className="p-1 hover:bg-slate-800 text-slate-300 hover:text-white rounded disabled:opacity-30"
                title="되돌리기 (Ctrl + Z)"
              >
                <Undo2 size={13} />
              </button>
              <button
                type="button"
                onClick={handleRedo}
                disabled={historyIndexRef.current >= historyRef.current.length - 1}
                className="p-1 hover:bg-slate-800 text-slate-300 hover:text-white rounded disabled:opacity-30"
                title="다시 실행 (Ctrl + Y)"
              >
                <Redo2 size={13} />
              </button>
            </div>
          </div>

          {/* 우측 선택된 셀 정보 */}
          <div className="flex items-center gap-2 text-slate-400 font-mono text-[11px] shrink-0">
            {selectedCell && (
              <span className="bg-slate-900 px-2 py-0.5 rounded border border-slate-800 text-indigo-300 font-bold">
                {getColumnLabel(selectedCell.c)}
                {selectedCell.r + 1}
              </span>
            )}
            <span className="hidden md:inline text-slate-500">
              더블 클릭 또는 Enter: 셀 수정 | 방향키: 이동
            </span>
          </div>
        </div>

        {/* 상단 3단: 수식 입력줄 (Formula Bar) */}
        <div className="px-3 py-1.5 bg-slate-900 border-b border-slate-800 flex items-center gap-2 shrink-0">
          <span className="text-[11px] font-mono font-bold text-slate-400 bg-slate-950 px-2 py-1 rounded border border-slate-800 min-w-[48px] text-center">
            {selectedCell ? `${getColumnLabel(selectedCell.c)}${selectedCell.r + 1}` : 'fx'}
          </span>
          <input
            type="text"
            value={
              editingCell
                ? editInputVal
                : selectedCell
                ? getCellValue(selectedCell.r, selectedCell.c)
                : ''
            }
            onChange={(e) => {
              if (selectedCell) {
                if (!editingCell) {
                  setEditingCell(selectedCell);
                }
                setEditInputVal(e.target.value);
              }
            }}
            onFocus={() => {
              if (selectedCell && !editingCell) {
                setEditingCell(selectedCell);
                setEditInputVal(getCellValue(selectedCell.r, selectedCell.c));
              }
            }}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                commitEditing();
              }
            }}
            placeholder="수식 또는 셀 값을 입력하세요..."
            className="flex-1 px-3 py-1 bg-slate-950 border border-slate-800 rounded-lg text-xs text-slate-200 placeholder-slate-600 focus:outline-none focus:border-indigo-500 font-sans"
          />
        </div>

        {/* 메인 스프레드시트 편집 그리드 */}
        <div
          ref={gridContainerRef}
          className="flex-1 overflow-auto bg-slate-950 p-2 select-none"
        >
          <div className="inline-block min-w-full align-middle border border-slate-800 rounded-xl overflow-hidden bg-slate-900/60 shadow-inner">
            <table className="min-w-full border-collapse text-left">
              {/* 상단 열 헤더 (A, B, C...) */}
              <thead className="bg-slate-900 sticky top-0 z-20 border-b border-slate-800">
                <tr>
                  <th className="w-12 min-w-[48px] px-2 py-1.5 text-center text-[11px] font-bold text-slate-500 bg-slate-950 border-r border-slate-800 sticky left-0 z-30">
                    #
                  </th>
                  {Array.from({ length: colCount }).map((_, cIdx) => {
                    const isColSelected = selectedCell?.c === cIdx;
                    return (
                      <th
                        key={cIdx}
                        onClick={() => setSelectedCell({ r: selectedCell?.r || 0, c: cIdx })}
                        className={`px-3 py-1.5 text-[11px] font-mono font-bold border-r border-slate-800/80 cursor-pointer min-w-[100px] transition-colors ${
                          isColSelected
                            ? 'bg-indigo-950/60 text-indigo-300'
                            : 'bg-slate-900 text-slate-400 hover:bg-slate-850'
                        }`}
                      >
                        {getColumnLabel(cIdx)}
                      </th>
                    );
                  })}
                </tr>
              </thead>

              {/* 그리드 행들 */}
              <tbody className="divide-y divide-slate-800/60 font-sans">
                {Array.from({ length: rowCount }).map((_, rIdx) => {
                  const isRowSelected = selectedCell?.r === rIdx;
                  return (
                    <tr
                      key={rIdx}
                      className={isRowSelected ? 'bg-indigo-950/15' : 'hover:bg-slate-900/40'}
                    >
                      {/* 행 번호 헤더 */}
                      <td
                        onClick={() => setSelectedCell({ r: rIdx, c: selectedCell?.c || 0 })}
                        className={`px-2 py-1 text-center text-[11px] font-mono font-bold border-r border-slate-800 sticky left-0 z-10 cursor-pointer transition-colors ${
                          isRowSelected
                            ? 'bg-indigo-950 text-indigo-300'
                            : 'bg-slate-950/90 text-slate-500 hover:bg-slate-900'
                        }`}
                      >
                        {rIdx + 1}
                      </td>

                      {/* 열 셀들 */}
                      {Array.from({ length: colCount }).map((_, cIdx) => {
                        const isCellSelected = selectedCell?.r === rIdx && selectedCell?.c === cIdx;
                        const isCellEditing = editingCell?.r === rIdx && editingCell?.c === cIdx;
                        const val = getCellValue(rIdx, cIdx);

                        return (
                          <td
                            key={cIdx}
                            onClick={() => {
                              if (isCellEditing) return;
                              commitEditing();
                              setSelectedCell({ r: rIdx, c: cIdx });
                            }}
                            onDoubleClick={() => startEditing(rIdx, cIdx)}
                            className={`relative px-2.5 py-1 text-xs border-r border-slate-800/50 break-words whitespace-pre-wrap min-w-[100px] max-w-sm transition-all cursor-cell ${
                              isCellSelected
                                ? 'bg-indigo-900/25 ring-2 ring-indigo-500 ring-inset z-10 font-medium text-white'
                                : 'text-slate-200'
                            }`}
                          >
                            {isCellEditing ? (
                              <input
                                ref={editInputRef}
                                type="text"
                                value={editInputVal}
                                onChange={(e) => setEditInputVal(e.target.value)}
                                onBlur={commitEditing}
                                onKeyDown={handleEditInputKeyDown}
                                className="w-full bg-slate-950 text-white text-xs px-1.5 py-0.5 rounded border border-indigo-400 focus:outline-none shadow-md font-sans"
                              />
                            ) : (
                              <div className="min-h-[20px] select-text">{val}</div>
                            )}
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>

        {/* 하단 시트 탭 관리 바 */}
        <div className="px-3 py-2 bg-slate-900 border-t border-slate-800 flex items-center justify-between gap-3 shrink-0">
          <div className="flex items-center gap-1.5 overflow-x-auto py-0.5">
            <span className="text-[11px] font-bold text-slate-400 px-1 shrink-0">시트:</span>
            {sheetNames.map((sheetName) => {
              const isActive = activeSheetName === sheetName;
              return (
                <div
                  key={sheetName}
                  className={`flex items-center rounded-lg border transition-all text-xs font-semibold shrink-0 overflow-hidden ${
                    isActive
                      ? 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40 shadow-xs'
                      : 'bg-slate-800 hover:bg-slate-750 text-slate-400 hover:text-slate-200 border-slate-700/60'
                  }`}
                >
                  <button
                    type="button"
                    onClick={() => {
                      commitEditing();
                      setActiveSheetName(sheetName);
                      setSelectedCell({ r: 0, c: 0 });
                    }}
                    onDoubleClick={() => handleRenameSheet(sheetName)}
                    className="px-2.5 py-1"
                    title="클릭하여 시트 전환 (더블 클릭: 이름 변경)"
                  >
                    {sheetName}
                  </button>
                  {isActive && sheetNames.length > 1 && (
                    <button
                      type="button"
                      onClick={() => handleDeleteSheet(sheetName)}
                      className="pr-1.5 pl-0.5 hover:text-rose-400 transition-colors"
                      title="시트 삭제"
                    >
                      <X size={12} />
                    </button>
                  )}
                </div>
              );
            })}

            {/* + 새 시트 추가 버튼 */}
            {!isCsv && (
              <button
                type="button"
                onClick={handleAddSheet}
                className="px-2 py-1 bg-slate-800 hover:bg-emerald-600/30 text-slate-300 hover:text-emerald-300 border border-slate-700 rounded-lg text-xs font-medium flex items-center gap-1 transition-colors shrink-0"
                title="새 시트 추가"
              >
                <Plus size={12} />
                <span>시트 추가</span>
              </button>
            )}
          </div>

          <div className="text-[11px] text-slate-400 hidden sm:flex items-center gap-2 shrink-0">
            <span>💾 [저장 및 반영]을 누르면 메모 본문 뷰어와 동기화됩니다.</span>
          </div>
        </div>
      </div>
    </div>
  );
};
