import React, { useState, useMemo, useRef, useEffect } from 'react';
import type { Note, Project } from '../types';
import { 
  ChevronLeft, 
  ChevronRight, 
  ChevronDown,
  Plus, 
  X, 
  Check, 
  FileText, 
  Folder, 
  Lock,
  Calendar as CalendarIcon
} from 'lucide-react';

interface CalendarViewProps {
  notes: Note[];
  projects: Project[];
  activeProjectId: string | null;
  onSelectNote: (noteId: string) => void;
  onCreateNoteForDate: (dateStr: string, targetProjectId?: string) => void;
  onClose: () => void;
}

export const CalendarView: React.FC<CalendarViewProps> = ({
  notes,
  projects,
  activeProjectId,
  onSelectNote,
  onCreateNoteForDate,
  onClose,
}) => {
  const today = useMemo(() => new Date(), []);
  const [currentYear, setCurrentYear] = useState<number>(today.getFullYear());
  const [currentMonth, setCurrentMonth] = useState<number>(today.getMonth()); // 0-indexed

  // 선택된 날짜 (기본값: 오늘)
  const [selectedDateStr, setSelectedDateStr] = useState<string>(() => {
    const y = today.getFullYear();
    const m = String(today.getMonth() + 1).padStart(2, '0');
    const d = String(today.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  });

  // 폴더 필터 팝오버 열림 상태
  const [isFolderFilterOpen, setIsFolderFilterOpen] = useState(false);
  const filterDropdownRef = useRef<HTMLDivElement>(null);

  // 선택된 폴더 ID 집합 (체크박스 다중 선택)
  const [selectedFolderIds, setSelectedFolderIds] = useState<Set<string>>(() => {
    if (activeProjectId && activeProjectId !== 'ALL') {
      return new Set([activeProjectId]);
    }
    return new Set(projects.map((p) => p.id));
  });

  // 팝오버 외부 클릭 시 닫기
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (filterDropdownRef.current && !filterDropdownRef.current.contains(e.target as Node)) {
        setIsFolderFilterOpen(false);
      }
    };
    if (isFolderFilterOpen) {
      document.addEventListener('mousedown', handleClickOutside);
    }
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [isFolderFilterOpen]);

  // 폴더 ID -> Project 맵
  const projectMap = useMemo(() => {
    const map: Record<string, Project> = {};
    projects.forEach((p) => { map[p.id] = p; });
    return map;
  }, [projects]);

  // 사이드바와 100% 동일한 계층 트리 구조(depth 포함) 계산
  const hierarchicalProjects = useMemo(() => {
    const rootProjects: Project[] = [];
    const childrenMap: Record<string, Project[]> = {};

    projects.forEach((p) => {
      if (!p.parentId) {
        rootProjects.push(p);
      } else {
        if (!childrenMap[p.parentId]) childrenMap[p.parentId] = [];
        childrenMap[p.parentId].push(p);
      }
    });

    const result: { project: Project; depth: number }[] = [];
    const traverse = (proj: Project, depth: number) => {
      result.push({ project: proj, depth });
      const children = childrenMap[proj.id] || [];
      children.forEach((child) => traverse(child, depth + 1));
    };

    rootProjects.forEach((root) => traverse(root, 0));

    projects.forEach((p) => {
      if (!result.some((r) => r.project.id === p.id)) {
        result.push({ project: p, depth: 0 });
      }
    });

    return result;
  }, [projects]);

  // 이전 달 / 다음 달 / 오늘 이동
  const handlePrevMonth = () => {
    if (currentMonth === 0) {
      setCurrentYear((y) => y - 1);
      setCurrentMonth(11);
    } else {
      setCurrentMonth((m) => m - 1);
    }
  };

  const handleNextMonth = () => {
    if (currentMonth === 11) {
      setCurrentYear((y) => y + 1);
      setCurrentMonth(0);
    } else {
      setCurrentMonth((m) => m + 1);
    }
  };

  const handleGoToday = () => {
    const now = new Date();
    setCurrentYear(now.getFullYear());
    setCurrentMonth(now.getMonth());
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const d = String(now.getDate()).padStart(2, '0');
    setSelectedDateStr(`${y}-${m}-${d}`);
  };

  // 메모의 기준 날짜 (YYYY-MM-DD) 추출
  const getNoteDateStr = (note: Note): string => {
    const match = note.title.match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
    if (match) {
      const year = match[1];
      const month = match[2].padStart(2, '0');
      const day = match[3].padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
    const d = new Date(note.createdAt);
    if (isNaN(d.getTime())) return '';
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  };

  // 체크된 폴더들의 메모만 필터링하여 날짜별 맵핑
  const dateNotesMap = useMemo(() => {
    const map: Record<string, Note[]> = {};

    const filtered = notes.filter((n) => selectedFolderIds.has(n.projectId));

    filtered.forEach((note) => {
      const dateStr = getNoteDateStr(note);
      if (dateStr) {
        if (!map[dateStr]) map[dateStr] = [];
        map[dateStr].push(note);
      }
    });

    Object.keys(map).forEach((k) => {
      map[k].sort((a, b) => new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime());
    });

    return map;
  }, [notes, selectedFolderIds]);

  // 달력 일자 그리드 계산
  const calendarDays = useMemo(() => {
    const firstDayOfMonth = new Date(currentYear, currentMonth, 1);
    const lastDayOfMonth = new Date(currentYear, currentMonth + 1, 0);

    const startDayOfWeek = firstDayOfMonth.getDay();
    const daysInMonth = lastDayOfMonth.getDate();

    const prevMonthLastDay = new Date(currentYear, currentMonth, 0).getDate();
    const prevDays: { dayNum: number; dateStr: string; isCurrentMonth: boolean }[] = [];
    for (let i = startDayOfWeek - 1; i >= 0; i--) {
      const dayNum = prevMonthLastDay - i;
      const prevDate = new Date(currentYear, currentMonth - 1, dayNum);
      const y = prevDate.getFullYear();
      const m = String(prevDate.getMonth() + 1).padStart(2, '0');
      const d = String(dayNum).padStart(2, '0');
      prevDays.push({ dayNum, dateStr: `${y}-${m}-${d}`, isCurrentMonth: false });
    }

    const currentDays: { dayNum: number; dateStr: string; isCurrentMonth: boolean }[] = [];
    for (let dayNum = 1; dayNum <= daysInMonth; dayNum++) {
      const y = currentYear;
      const m = String(currentMonth + 1).padStart(2, '0');
      const d = String(dayNum).padStart(2, '0');
      currentDays.push({ dayNum, dateStr: `${y}-${m}-${d}`, isCurrentMonth: true });
    }

    const totalDaysSoFar = prevDays.length + currentDays.length;
    const remainingDays = totalDaysSoFar % 7 === 0 ? 0 : 7 - (totalDaysSoFar % 7);
    const nextDays: { dayNum: number; dateStr: string; isCurrentMonth: boolean }[] = [];
    for (let dayNum = 1; dayNum <= remainingDays; dayNum++) {
      const nextDate = new Date(currentYear, currentMonth + 1, dayNum);
      const y = nextDate.getFullYear();
      const m = String(nextDate.getMonth() + 1).padStart(2, '0');
      const d = String(dayNum).padStart(2, '0');
      nextDays.push({ dayNum, dateStr: `${y}-${m}-${d}`, isCurrentMonth: false });
    }

    return [...prevDays, ...currentDays, ...nextDays];
  }, [currentYear, currentMonth]);

  const todayStr = useMemo(() => {
    const y = today.getFullYear();
    const m = String(today.getMonth() + 1).padStart(2, '0');
    const d = String(today.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }, [today]);

  const selectedDateNotes = dateNotesMap[selectedDateStr] || [];

  // 메모 작성 시 대상이 될 폴더 ID 계산:
  // 1) 체크된 폴더가 1개인 경우 -> 그 폴더에 바로 생성
  // 2) activeProjectId가 체크된 폴더 목록에 속해 있는 경우 -> activeProjectId 사용
  // 3) 그 외: 첫 번째 체크된 폴더 사용
  const targetFolderId = useMemo(() => {
    if (selectedFolderIds.size === 1) {
      return Array.from(selectedFolderIds)[0];
    }
    if (activeProjectId && activeProjectId !== 'ALL' && selectedFolderIds.has(activeProjectId)) {
      return activeProjectId;
    }
    return Array.from(selectedFolderIds)[0] || undefined;
  }, [selectedFolderIds, activeProjectId]);

  const targetFolderName = targetFolderId ? projectMap[targetFolderId]?.name : null;

  return (
    <div className="flex-1 flex flex-col h-full bg-slate-950 text-slate-100 overflow-hidden select-none">
      {/* 상단 헤더: 달력 내비게이션 & 계층형 트리 체크박스 필터 & 닫기 */}
      <div className="p-3 md:px-6 border-b border-slate-800 bg-slate-900/60 flex flex-wrap items-center justify-between gap-3 shrink-0">
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-1.5 text-indigo-400">
            <CalendarIcon size={20} />
            <h2 className="font-bold text-base md:text-lg text-slate-100 tracking-tight">
              {currentYear}년 {currentMonth + 1}월
            </h2>
          </div>

          <div className="flex items-center gap-1 bg-slate-800/80 border border-slate-700/80 rounded-lg p-0.5">
            <button
              onClick={handlePrevMonth}
              className="p-1 hover:bg-slate-700 text-slate-300 hover:text-white rounded transition-colors"
              title="이전 달"
            >
              <ChevronLeft size={16} />
            </button>
            <button
              onClick={handleGoToday}
              className="px-2 py-0.5 text-xs text-slate-300 hover:text-white font-medium hover:bg-slate-700 rounded transition-colors"
            >
              오늘
            </button>
            <button
              onClick={handleNextMonth}
              className="p-1 hover:bg-slate-700 text-slate-300 hover:text-white rounded transition-colors"
              title="다음 달"
            >
              <ChevronRight size={16} />
            </button>
          </div>
        </div>

        {/* 오른쪽: 계층 트리 체크박스 필터 드롭다운 & 닫기 버튼 */}
        <div className="flex items-center gap-2 relative" ref={filterDropdownRef}>
          {/* 트리구조 체크박스 필터 토글 버튼 */}
          <button
            type="button"
            onClick={() => setIsFolderFilterOpen((prev) => !prev)}
            className={`flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-medium border transition-all ${
              selectedFolderIds.size < projects.length
                ? 'bg-indigo-600/20 text-indigo-300 border-indigo-500/50 shadow-sm'
                : 'bg-slate-900 text-slate-300 border-slate-800 hover:bg-slate-800 hover:text-white'
            }`}
            title="클릭하여 캘린더에 표시할 폴더들을 체크박스로 선택"
          >
            <Folder size={13} className="text-indigo-400 shrink-0" />
            <span className="font-medium truncate max-w-[130px] md:max-w-[200px]">
              {selectedFolderIds.size === projects.length
                ? '전체 폴더 표시'
                : selectedFolderIds.size === 1 && targetFolderName
                  ? `폴더: ${targetFolderName}`
                  : `${selectedFolderIds.size}개 폴더 선택됨`}
            </span>
            <ChevronDown size={14} className="text-slate-400 shrink-0" />
          </button>

          {/* 닫기 버튼 */}
          <button
            onClick={onClose}
            className="flex items-center gap-1 px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-lg text-xs font-semibold transition-all border border-slate-700/80 shadow-sm"
          >
            <X size={14} />
            <span>메모 닫기</span>
          </button>

          {/* 계층형 트리 체크박스 팝오버 창 */}
          {isFolderFilterOpen && (
            <div 
              className="absolute right-0 top-full mt-1.5 w-72 md:w-80 bg-slate-900 border border-slate-700/90 rounded-2xl shadow-2xl z-50 p-3 space-y-2 animate-in fade-in zoom-in-95 duration-150"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="flex items-center justify-between pb-2 border-b border-slate-800 px-1">
                <span className="text-xs font-bold text-slate-200">표시할 폴더 선택</span>
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setSelectedFolderIds(new Set(projects.map((p) => p.id)))}
                    className="text-[11px] text-indigo-400 hover:text-indigo-300 font-semibold"
                  >
                    전체 선택
                  </button>
                  <span className="text-slate-700 text-xs">|</span>
                  <button
                    type="button"
                    onClick={() => setSelectedFolderIds(new Set())}
                    className="text-[11px] text-slate-400 hover:text-slate-200"
                  >
                    전체 해제
                  </button>
                </div>
              </div>

              {/* 계층형 트리 목록 (들여쓰기 └ 구조) */}
              <div className="max-h-64 overflow-y-auto space-y-1 pr-1 scrollbar-thin">
                {hierarchicalProjects.map(({ project: p, depth }) => {
                  const isChecked = selectedFolderIds.has(p.id);
                  const count = notes.filter((n) => n.projectId === p.id).length;

                  return (
                    <div
                      key={p.id}
                      className={`flex items-center justify-between py-1.5 px-2 rounded-lg hover:bg-slate-800/90 transition-colors text-xs cursor-pointer select-none group ${
                        isChecked ? 'text-slate-100 font-medium' : 'text-slate-500'
                      }`}
                      style={{ paddingLeft: `${depth * 14 + 6}px` }}
                      onClick={() => {
                        const next = new Set(selectedFolderIds);
                        if (next.has(p.id)) {
                          next.delete(p.id);
                        } else {
                          next.add(p.id);
                        }
                        setSelectedFolderIds(next);
                      }}
                    >
                      <div className="flex items-center gap-2 truncate">
                        <input
                          type="checkbox"
                          checked={isChecked}
                          onChange={() => {}} // 부모 div onClick으로 처리
                          className="accent-indigo-500 rounded cursor-pointer shrink-0"
                        />
                        {depth > 0 && (
                          <span className="text-slate-500 font-mono text-[11px] shrink-0">└</span>
                        )}
                        {p.isLocalOnly ? (
                          <Lock size={12} className="text-amber-400 shrink-0" />
                        ) : (
                          <Folder size={12} className="text-indigo-400 shrink-0" />
                        )}
                        <span className="truncate">{p.name}</span>
                      </div>

                      <div className="flex items-center gap-1.5 shrink-0 pl-1">
                        <span className="text-[10px] text-slate-500 font-mono">
                          {count}
                        </span>
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            setSelectedFolderIds(new Set([p.id]));
                          }}
                          className="text-[10px] text-indigo-400 hover:text-indigo-300 hover:bg-indigo-950/60 border border-indigo-500/20 px-1 py-0.5 rounded opacity-0 group-hover:opacity-100 transition-opacity"
                          title="이 폴더만 단독 보기"
                        >
                          단독
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>

              <div className="pt-2 border-t border-slate-800 flex items-center justify-between text-xs">
                <span className="text-[11px] text-slate-400">
                  {selectedFolderIds.size}개 폴더 활성화
                </span>
                <button
                  type="button"
                  onClick={() => setIsFolderFilterOpen(false)}
                  className="px-3 py-1 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-semibold"
                >
                  적용
                </button>
              </div>
            </div>
          )}
        </div>
      </div>

      {/* 요일 헤더 */}
      <div className="grid grid-cols-7 border-b border-slate-800 bg-slate-900/40 text-center text-xs font-semibold py-1.5 shrink-0">
        <div className="text-rose-400">일</div>
        <div className="text-slate-400">월</div>
        <div className="text-slate-400">화</div>
        <div className="text-slate-400">수</div>
        <div className="text-slate-400">목</div>
        <div className="text-slate-400">금</div>
        <div className="text-sky-400">토</div>
      </div>

      {/* 메인 캘린더 그리드 */}
      <div className="flex-1 grid grid-cols-7 auto-rows-fr overflow-y-auto bg-slate-950 divide-x divide-y divide-slate-800/60 border-b border-slate-800">
        {calendarDays.map((item, idx) => {
          const dayNotes = dateNotesMap[item.dateStr] || [];
          const isToday = item.dateStr === todayStr;
          const isSelected = item.dateStr === selectedDateStr;
          const isSunday = idx % 7 === 0;
          const isSaturday = idx % 7 === 6;

          return (
            <div
              key={item.dateStr}
              onClick={() => setSelectedDateStr(item.dateStr)}
              className={`p-1 md:p-1.5 flex flex-col min-h-[90px] md:min-h-[110px] transition-colors relative cursor-pointer group/cell ${
                !item.isCurrentMonth ? 'bg-slate-950/40 opacity-40' : 'bg-slate-950 hover:bg-slate-900/50'
              } ${isSelected ? 'ring-2 ring-indigo-500/80 ring-inset bg-indigo-950/20' : ''}`}
            >
              {/* 날짜 상단 바 (숫자 + 추가 버튼) */}
              <div className="flex items-center justify-between mb-1">
                <span
                  className={`text-xs font-bold w-6 h-6 flex items-center justify-center rounded-full transition-all ${
                    isToday
                      ? 'bg-indigo-600 text-white font-black shadow-md shadow-indigo-600/30'
                      : isSunday
                        ? 'text-rose-400'
                        : isSaturday
                          ? 'text-sky-400'
                          : item.isCurrentMonth
                            ? 'text-slate-300'
                            : 'text-slate-600'
                  }`}
                >
                  {item.dayNum}
                </span>

                {/* 해당 날짜에 새 메모 즉시 작성 (+) */}
                <button
                  type="button"
                  onClick={(e) => {
                    e.stopPropagation();
                    onCreateNoteForDate(item.dateStr, targetFolderId);
                  }}
                  className="opacity-0 group-hover/cell:opacity-100 p-1 text-slate-500 hover:text-indigo-400 hover:bg-slate-800 rounded transition-all"
                  title={`${item.dateStr} 새 메모 작성 (${targetFolderName || '현재 폴더'})`}
                >
                  <Plus size={13} />
                </button>
              </div>

              {/* 해당 날짜의 메모 목록 (최대 3~4개 직접 노출) */}
              <div className="flex-1 flex flex-col gap-1 overflow-hidden">
                {dayNotes.slice(0, 3).map((note) => {
                  const hasUndoneTask = note.content.includes('- [ ]');
                  const hasDoneTask = note.content.includes('- [x]') || note.content.includes('- [X]');
                  const folder = projectMap[note.projectId];

                  return (
                    <button
                      key={note.id}
                      type="button"
                      onClick={(e) => {
                        e.stopPropagation();
                        onSelectNote(note.id);
                      }}
                      className="w-full text-left p-1 px-1.5 rounded bg-slate-900/90 hover:bg-indigo-600/30 border border-slate-800 hover:border-indigo-500/40 transition-all flex items-center gap-1 group/item active:scale-95 shadow-xs"
                      title={`[${folder?.name || '폴더'}] ${note.title} (클릭하여 열기)`}
                    >
                      {/* 할 일 / 완료 / 일반 아이콘 구분 */}
                      {hasUndoneTask ? (
                        <span className="text-amber-400 shrink-0 text-[10px]">☐</span>
                      ) : hasDoneTask ? (
                        <Check size={10} className="text-emerald-400 shrink-0" />
                      ) : note.isLocalOnly ? (
                        <Lock size={10} className="text-amber-400 shrink-0" />
                      ) : (
                        <span className="text-indigo-400 shrink-0 text-[10px]">•</span>
                      )}

                      <span className="truncate text-[11px] text-slate-200 group-hover/item:text-indigo-200 font-medium flex-1">
                        {note.title || '제목 없음'}
                      </span>
                    </button>
                  );
                })}

                {/* 3개 초과 시 +N개 더보기 표시 */}
                {dayNotes.length > 3 && (
                  <span className="text-[10px] text-indigo-400 font-semibold px-1">
                    +{dayNotes.length - 3}개 더보기
                  </span>
                )}
              </div>
            </div>
          );
        })}
      </div>

      {/* 하단: 선택된 날짜 상세 패널 (모바일 및 날짜별 메모 한눈에 보기) */}
      <div className="p-3 md:px-6 border-t border-slate-800 bg-slate-900/90 flex flex-col md:flex-row md:items-center justify-between gap-3 shrink-0">
        <div className="flex items-center gap-2">
          <CalendarIcon size={16} className="text-indigo-400" />
          <span className="text-xs md:text-sm font-bold text-slate-100">
            {selectedDateStr} 메모 목록 ({selectedDateNotes.length}개)
          </span>
          <button
            onClick={() => onCreateNoteForDate(selectedDateStr, targetFolderId)}
            className="ml-2 flex items-center gap-1 px-2.5 py-1 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-semibold shadow-sm transition-all active:scale-95"
            title={`${targetFolderName ? `'${targetFolderName}' 폴더에 ` : ''}선택한 날짜로 새 메모 작성`}
          >
            <Plus size={13} />
            <span>
              {targetFolderName ? `'${targetFolderName}'에 새 메모` : '이 날짜에 새 메모'}
            </span>
          </button>
        </div>

        {/* 선택 날짜의 메모 태그 리스트 */}
        <div className="flex items-center gap-1.5 overflow-x-auto scrollbar-none py-1">
          {selectedDateNotes.length === 0 ? (
            <span className="text-xs text-slate-500">작성된 메모가 없습니다. [+ 새 메모]를 눌러 기록해 보세요!</span>
          ) : (
            selectedDateNotes.map((note) => {
              const folder = projectMap[note.projectId];
              return (
                <button
                  key={note.id}
                  onClick={() => onSelectNote(note.id)}
                  className="flex items-center gap-1.5 px-2.5 py-1 bg-slate-800 hover:bg-indigo-600/30 border border-slate-700/80 hover:border-indigo-500/50 rounded-lg text-xs font-medium text-slate-200 hover:text-white transition-all shrink-0 active:scale-95"
                  title={`[${folder?.name || '폴더'}] ${note.title}`}
                >
                  <FileText size={12} className="text-indigo-400 shrink-0" />
                  {folder && (
                    <span className="text-[10px] text-slate-400 font-mono">[{folder.name}]</span>
                  )}
                  <span className="max-w-[130px] truncate">{note.title}</span>
                </button>
              );
            })
          )}
        </div>
      </div>
    </div>
  );
};
