import React, { useState, useEffect, useMemo, useRef } from 'react';
import { 
  X, 
  Search, 
  Download, 
  Edit3, 
  FileSpreadsheet, 
  Copy, 
  Check, 
  ZoomIn, 
  ZoomOut 
} from 'lucide-react';
import { parseExcelWorkbook, sheetRowsToMarkdownTable, type ParsedWorkbook } from '../utils/excelViewer';

interface ExcelViewerModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: string;
  title: string;
  attId?: string;
  onOpenEditor?: () => void;
  onImportTable?: (markdownTable: string) => void;
}

// 열 인덱스(0, 1, 2...)를 엑셀 열 문자('A', 'B', 'Z', 'AA'...)로 변환
function getColumnLabel(index: number): string {
  let label = '';
  let num = index;
  while (num >= 0) {
    label = String.fromCharCode((num % 26) + 65) + label;
    num = Math.floor(num / 26) - 1;
  }
  return label;
}

export const ExcelViewerModal: React.FC<ExcelViewerModalProps> = ({
  isOpen,
  onClose,
  data,
  title,
  attId,
  onOpenEditor,
  onImportTable,
}) => {
  const [workbook, setWorkbook] = useState<ParsedWorkbook | null>(null);
  const [activeSheetName, setActiveSheetName] = useState<string>('');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [zoom, setZoom] = useState<number>(1.0);
  const [copied, setCopied] = useState<boolean>(false);

  const containerRef = useRef<HTMLDivElement>(null);

  // 워크북 로드
  useEffect(() => {
    if (!isOpen || !data) return;
    try {
      const parsed = parseExcelWorkbook(data, attId);
      setWorkbook(parsed);
      setActiveSheetName(parsed.sheetNames[0] || 'Sheet1');
      setSearchQuery('');
      setZoom(1.0);
    } catch (err) {
      console.error('엑셀 뷰어 로드 실패:', err);
    }
  }, [isOpen, data, attId]);

  // 현재 활성화된 시트 데이터
  const activeRows = useMemo(() => {
    if (!workbook || !activeSheetName) return [];
    return workbook.sheets[activeSheetName] || [];
  }, [workbook, activeSheetName]);

  // 최대 열 수 계산
  const colCount = useMemo(() => {
    let max = 0;
    activeRows.forEach((r) => {
      max = Math.max(max, r.length);
    });
    return Math.max(max, 1);
  }, [activeRows]);

  // 검색 필터링된 행 목록 및 매칭 셀 좌표 계산
  const { filteredRows, matchCount } = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    if (!q) {
      return {
        filteredRows: activeRows.map((row, idx) => ({ row, originalIndex: idx })),
        matchCount: 0,
      };
    }

    let count = 0;
    const filtered: { row: (string | number)[]; originalIndex: number }[] = [];

    activeRows.forEach((row, originalIndex) => {
      let rowMatches = false;
      row.forEach((cell) => {
        if (cell !== undefined && cell !== null && String(cell).toLowerCase().includes(q)) {
          rowMatches = true;
          count++;
        }
      });
      if (rowMatches) {
        filtered.push({ row, originalIndex });
      }
    });

    return { filteredRows: filtered, matchCount: count };
  }, [activeRows, searchQuery]);

  // 마크다운 표 복사 핸들러
  const handleCopyMarkdown = () => {
    if (activeRows.length === 0) return;
    const md = sheetRowsToMarkdownTable(activeRows);
    navigator.clipboard.writeText(md).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    });
  };

  // 마크다운 표 메모 본문으로 삽입 핸들러
  const handleImport = () => {
    if (activeRows.length === 0) return;
    const md = sheetRowsToMarkdownTable(activeRows);
    if (onImportTable) {
      onImportTable(md);
      onClose();
    }
  };

  // 다운로드 핸들러
  const handleDownload = () => {
    const link = document.createElement('a');
    link.href = data;
    link.download = title.endsWith('.xlsx') || title.endsWith('.csv') || title.endsWith('.xls')
      ? title
      : `${title}.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 bg-slate-950/90 backdrop-blur-md flex flex-col items-center justify-center p-2 md:p-4 select-none animate-fadeIn">
      {/* 뷰어 메인 팝업 컨테이너 */}
      <div className="w-full max-w-7xl h-[94vh] bg-slate-900 border border-slate-800 rounded-2xl shadow-2xl flex flex-col overflow-hidden">
        {/* 상단 툴바 헤더 */}
        <div className="px-4 py-3 bg-slate-900/95 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3 shrink-0">
          <div className="flex items-center gap-2.5 min-w-0">
            <span className="p-1.5 rounded-lg bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
              <FileSpreadsheet size={18} />
            </span>
            <div className="truncate">
              <h2 className="text-sm md:text-base font-bold text-slate-100 truncate" title={title}>
                {title}
              </h2>
              <div className="flex items-center gap-2 text-[11px] text-slate-400 font-mono">
                <span>시트: {workbook?.sheetNames.length || 1}개</span>
                <span>•</span>
                <span>행: {activeRows.length}개</span>
                <span>•</span>
                <span>열: {colCount}개</span>
              </div>
            </div>
          </div>

          {/* 중앙 검색창 & 줌 컨트롤 */}
          <div className="flex items-center gap-2 flex-1 max-w-md justify-center">
            <div className="relative w-full max-w-xs">
              <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="표 내용 검색..."
                className="w-full pl-8 pr-7 py-1 bg-slate-950 border border-slate-800 rounded-lg text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:border-indigo-500"
              />
              {searchQuery && (
                <button
                  type="button"
                  onClick={() => setSearchQuery('')}
                  className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-500 hover:text-slate-300"
                >
                  <X size={12} />
                </button>
              )}
            </div>

            {searchQuery && (
              <span className="text-[11px] text-indigo-400 whitespace-nowrap">
                {matchCount}건 발견 ({filteredRows.length}행)
              </span>
            )}

            {/* 확대 / 축소 */}
            <div className="hidden sm:flex items-center gap-1 bg-slate-950 border border-slate-800 rounded-lg p-0.5 text-xs text-slate-300">
              <button
                type="button"
                onClick={() => setZoom((prev) => Math.max(0.7, prev - 0.1))}
                className="p-1 hover:text-white hover:bg-slate-800 rounded"
                title="축소"
              >
                <ZoomOut size={13} />
              </button>
              <span className="px-1 text-[11px] font-mono min-w-[42px] text-center">
                {Math.round(zoom * 100)}%
              </span>
              <button
                type="button"
                onClick={() => setZoom((prev) => Math.min(1.6, prev + 0.1))}
                className="p-1 hover:text-white hover:bg-slate-800 rounded"
                title="확대"
              >
                <ZoomIn size={13} />
              </button>
            </div>
          </div>

          {/* 우측 액션 버튼들 */}
          <div className="flex items-center gap-1.5 shrink-0">
            {onOpenEditor && (
              <button
                type="button"
                onClick={() => {
                  onClose();
                  onOpenEditor();
                }}
                className="px-2.5 py-1.5 bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-300 border border-emerald-500/30 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-colors active:scale-95"
                title="웹에서 스프레드시트 직접 편집 및 행/열 수정"
              >
                <Edit3 size={13} />
                <span>표 직접 편집</span>
              </button>
            )}

            <button
              type="button"
              onClick={handleCopyMarkdown}
              className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-750 text-slate-200 rounded-lg text-xs font-medium border border-slate-700 flex items-center gap-1.5 transition-colors active:scale-95"
              title="현재 시트를 마크다운 표 형식으로 클립보드에 복사"
            >
              {copied ? <Check size={13} className="text-emerald-400" /> : <Copy size={13} />}
              <span>{copied ? '복사됨!' : '마크다운 복사'}</span>
            </button>

            {onImportTable && (
              <button
                type="button"
                onClick={handleImport}
                className="px-2.5 py-1.5 bg-indigo-600/30 hover:bg-indigo-600 text-indigo-200 hover:text-white rounded-lg text-xs font-semibold border border-indigo-500/40 flex items-center gap-1.5 transition-colors active:scale-95"
                title="현재 시트 표를 메모 본문으로 가져오기"
              >
                <span>📝 메모로 가져오기</span>
              </button>
            )}

            <button
              type="button"
              onClick={handleDownload}
              className="p-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white rounded-lg border border-slate-700 transition-colors"
              title="파일 다운로드"
            >
              <Download size={15} />
            </button>

            <button
              type="button"
              onClick={onClose}
              className="p-1.5 bg-slate-800 hover:bg-rose-900/40 text-slate-400 hover:text-rose-300 rounded-lg border border-slate-700 transition-colors"
              title="닫기 (Esc)"
            >
              <X size={16} />
            </button>
          </div>
        </div>

        {/* 메인 스프레드시트 표 영역 */}
        <div
          ref={containerRef}
          className="flex-1 overflow-auto bg-slate-950 p-3 select-text"
          style={{ fontSize: `${Math.round(13 * zoom)}px` }}
        >
          {filteredRows.length === 0 ? (
            <div className="h-full flex flex-col items-center justify-center text-slate-500 gap-2">
              <FileSpreadsheet size={36} className="text-slate-600" />
              <p className="text-sm">
                {searchQuery ? '검색된 셀 데이터가 없습니다.' : '시트에 데이터가 없습니다.'}
              </p>
            </div>
          ) : (
            <div className="inline-block min-w-full align-middle border border-slate-800 rounded-xl overflow-hidden bg-slate-900/60 shadow-inner">
              <table className="min-w-full border-collapse text-left">
                {/* 상단 열 헤더 (A, B, C...) */}
                <thead className="bg-slate-900 sticky top-0 z-20 border-b border-slate-800 shadow-xs">
                  <tr>
                    <th className="w-12 min-w-[48px] px-2 py-2 text-center text-[11px] font-bold text-slate-500 bg-slate-950/80 border-r border-slate-800 sticky left-0 z-30">
                      #
                    </th>
                    {Array.from({ length: colCount }).map((_, cIdx) => (
                      <th
                        key={cIdx}
                        className="px-3 py-2 text-[11px] font-semibold text-slate-400 border-r border-slate-800/80 bg-slate-900 min-w-[90px]"
                      >
                        {getColumnLabel(cIdx)}
                      </th>
                    ))}
                  </tr>
                </thead>

                {/* 행 및 셀 데이터 */}
                <tbody className="divide-y divide-slate-800/60 font-sans">
                  {filteredRows.map(({ row, originalIndex }) => (
                    <tr
                      key={originalIndex}
                      className="hover:bg-indigo-950/20 transition-colors group"
                    >
                      {/* 행 번호 (1, 2, 3...) */}
                      <td className="px-2 py-1.5 text-center text-[11px] font-mono font-bold text-slate-500 bg-slate-950/90 border-r border-slate-800 sticky left-0 z-10 group-hover:text-indigo-400 group-hover:bg-slate-950">
                        {originalIndex + 1}
                      </td>

                      {/* 열 셀들 */}
                      {Array.from({ length: colCount }).map((_, cIdx) => {
                        const cellVal = row[cIdx];
                        const cellStr = cellVal !== undefined && cellVal !== null ? String(cellVal) : '';
                        const isMatch =
                          searchQuery.trim() !== '' &&
                          cellStr.toLowerCase().includes(searchQuery.trim().toLowerCase());

                        return (
                          <td
                            key={cIdx}
                            className={`px-3 py-1.5 text-xs text-slate-200 border-r border-slate-800/50 break-words whitespace-pre-wrap max-w-sm ${
                              isMatch ? 'bg-amber-500/25 text-amber-200 font-semibold' : ''
                            }`}
                          >
                            {cellStr}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>

        {/* 하단 시트 탭 바 */}
        <div className="px-3 py-2 bg-slate-900 border-t border-slate-800 flex items-center justify-between gap-2 shrink-0">
          <div className="flex items-center gap-1.5 overflow-x-auto py-0.5">
            <span className="text-[11px] font-bold text-slate-400 px-1 shrink-0">시트:</span>
            {workbook?.sheetNames.map((sheetName) => (
              <button
                key={sheetName}
                type="button"
                onClick={() => setActiveSheetName(sheetName)}
                className={`px-3 py-1 rounded-lg text-xs font-semibold shrink-0 transition-all ${
                  activeSheetName === sheetName
                    ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40 shadow-xs'
                    : 'bg-slate-800 hover:bg-slate-750 text-slate-400 hover:text-slate-200 border border-slate-700/60'
                }`}
              >
                {sheetName}
              </button>
            ))}
          </div>

          <div className="text-[11px] text-slate-500 hidden sm:block shrink-0">
            마우스 휠이나 터치로 상하좌우 부드럽게 스크롤할 수 있습니다.
          </div>
        </div>
      </div>
    </div>
  );
};
