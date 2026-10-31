import React, { useState, useMemo } from 'react';
import type { Note, Project } from '../types';
import { 
  ChevronLeft, 
  ChevronRight, 
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
  onCreateNoteForDate: (dateStr: string) => void;
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
  // 현재 보고 있는 달력 년/월 상태 (기본값: 오늘 날짜 기준)
  const today = useMemo(() => new Date(), []);
  const [currentYear, setCurrentYear] = useState<number>(today.getFullYear());
  const [currentMonth, setCurrentMonth] = useState<number>(today.getMonth()); // 0-indexed (0=1월, 9=10월)

  // 선택된 날짜 (모바일 및 상세 패널용)
  const [selectedDateStr, setSelectedDateStr] = useState<string>(() => {
    const y = today.getFullYear();
    const m = String(today.getMonth() + 1).padStart(2, '0');
    const d = String(today.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  });

  // 폴더 필터 (기본값: 'ALL')
  const [filterProjectId, setFilterProjectId] = useState<string>(activeProjectId || 'ALL');

  // 이전 달 이동
  const handlePrevMonth = () => {
    if (currentMonth === 0) {
      setCurrentYear((y) => y - 1);
      setCurrentMonth(11);
    } else {
      setCurrentMonth((m) => m - 1);
    }
  };

  // 다음 달 이동
  const handleNextMonth = () => {
    if (currentMonth === 11) {
      setCurrentYear((y) => y + 1);
      setCurrentMonth(0);
    } else {
      setCurrentMonth((m) => m + 1);
    }
  };

  // 오늘 날짜로 즉시 이동
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
  // 1순위: 제목에 YYYY-MM-DD 형식의 날짜가 명시되어 있으면 해당 날짜 사용
  // 2순위: 메모 생성일(createdAt) 사용
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

  // 날짜별 메모 맵핑 (YYYY-MM-DD -> Note[])
  const dateNotesMap = useMemo(() => {
    const map: Record<string, Note[]> = {};

    const filtered = filterProjectId === 'ALL'
      ? notes
      : notes.filter((n) => n.projectId === filterProjectId);

    filtered.forEach((note) => {
      const dateStr = getNoteDateStr(note);
      if (dateStr) {
        if (!map[dateStr]) map[dateStr] = [];
        map[dateStr].push(note);
      }
    });

    // 각 날짜별 메모들을 최신순(업데이트순)으로 정렬
    Object.keys(map).forEach((k) => {
      map[k].sort((a, b) => new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime());
    });

    return map;
  }, [notes, filterProjectId]);

  // 달력 그리드 계산 (해당 월의 날짜들 + 앞뒤 채우기 날짜)
  const calendarDays = useMemo(() => {
    const firstDayOfMonth = new Date(currentYear, currentMonth, 1);
    const lastDayOfMonth = new Date(currentYear, currentMonth + 1, 0);

    const startDayOfWeek = firstDayOfMonth.getDay(); // 0(일) ~ 6(토)
    const daysInMonth = lastDayOfMonth.getDate();

    // 이전 달 마지막 날들
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

    // 이번 달 날들
    const currentDays: { dayNum: number; dateStr: string; isCurrentMonth: boolean }[] = [];
    for (let dayNum = 1; dayNum <= daysInMonth; dayNum++) {
      const y = currentYear;
      const m = String(currentMonth + 1).padStart(2, '0');
      const d = String(dayNum).padStart(2, '0');
      currentDays.push({ dayNum, dateStr: `${y}-${m}-${d}`, isCurrentMonth: true });
    }

    // 다음 달 앞선 날들 (7의 배수로 그리드 완성)
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

  // 오늘 날짜 문자열
  const todayStr = useMemo(() => {
    const y = today.getFullYear();
    const m = String(today.getMonth() + 1).padStart(2, '0');
    const d = String(today.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }, [today]);

  // 선택된 날짜의 메모들
  const selectedDateNotes = dateNotesMap[selectedDateStr] || [];

  return (
    <div className="flex-1 flex flex-col h-full bg-slate-950 text-slate-100 overflow-hidden select-none">
      {/* 상단 헤더: 달력 내비게이션 & 폴더 필터 & 닫기 */}
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

        {/* 폴더 필터 & 닫기 버튼 */}
        <div className="flex items-center gap-2">
          {/* 폴더 필터 드롭다운 */}
          <div className="flex items-center gap-1.5 bg-slate-900 border border-slate-800 rounded-lg px-2.5 py-1 text-xs text-slate-300">
            <Folder size={12} className="text-indigo-400 shrink-0" />
            <select
              value={filterProjectId}
              onChange={(e) => setFilterProjectId(e.target.value)}
              className="bg-transparent text-xs text-slate-200 focus:outline-none cursor-pointer"
            >
              <option value="ALL" className="bg-slate-900">전체 폴더 메모</option>
              {projects.map((p) => (
                <option key={p.id} value={p.id} className="bg-slate-900">
                  {p.isLocalOnly ? `🔒 ${p.name}` : p.name}
                </option>
              ))}
            </select>
          </div>

          <button
            onClick={onClose}
            className="flex items-center gap-1 px-3 py-1 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-lg text-xs font-semibold transition-all border border-slate-700/80 shadow-sm"
          >
            <X size={14} />
            <span>메모 닫기</span>
          </button>
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
              className={`p-1 md:p-1.5 flex flex-col min-h-[90px] md:min-h-[110px] transition-colors relative cursor-pointer ${
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
                    onCreateNoteForDate(item.dateStr);
                  }}
                  className="opacity-0 hover:opacity-100 group-hover:opacity-100 p-1 text-slate-500 hover:text-indigo-400 rounded transition-opacity"
                  title={`${item.dateStr} 새 메모 작성`}
                >
                  <Plus size={12} />
                </button>
              </div>

              {/* 해당 날짜의 메모 목록 (최대 3~4개 직접 노출) */}
              <div className="flex-1 flex flex-col gap-1 overflow-hidden">
                {dayNotes.slice(0, 3).map((note) => {
                  const hasUndoneTask = note.content.includes('- [ ]');
                  const hasDoneTask = note.content.includes('- [x]') || note.content.includes('- [X]');

                  return (
                    <button
                      key={note.id}
                      type="button"
                      onClick={(e) => {
                        e.stopPropagation();
                        onSelectNote(note.id);
                      }}
                      className="w-full text-left p-1 px-1.5 rounded bg-slate-900/90 hover:bg-indigo-600/30 border border-slate-800 hover:border-indigo-500/40 transition-all flex items-center gap-1 group active:scale-95 shadow-xs"
                      title={`[${note.title}] 클릭하여 메모 열기`}
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

                      <span className="truncate text-[11px] text-slate-200 group-hover:text-indigo-200 font-medium">
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
            onClick={() => onCreateNoteForDate(selectedDateStr)}
            className="ml-2 flex items-center gap-1 px-2.5 py-1 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-semibold shadow-sm transition-all active:scale-95"
            title="선택한 날짜에 새 메모 작성"
          >
            <Plus size={13} />
            <span>이 날짜에 새 메모</span>
          </button>
        </div>

        {/* 선택 날짜의 메모 태그 리스트 */}
        <div className="flex items-center gap-1.5 overflow-x-auto scrollbar-none py-1">
          {selectedDateNotes.length === 0 ? (
            <span className="text-xs text-slate-500">작성된 메모가 없습니다. [+ 이 날짜에 새 메모]를 눌러 기록해 보세요!</span>
          ) : (
            selectedDateNotes.map((note) => (
              <button
                key={note.id}
                onClick={() => onSelectNote(note.id)}
                className="flex items-center gap-1.5 px-2.5 py-1 bg-slate-800 hover:bg-indigo-600/30 border border-slate-700/80 hover:border-indigo-500/50 rounded-lg text-xs font-medium text-slate-200 hover:text-white transition-all shrink-0 active:scale-95"
              >
                <FileText size={12} className="text-indigo-400" />
                <span className="max-w-[150px] truncate">{note.title}</span>
              </button>
            ))
          )}
        </div>
      </div>
    </div>
  );
};
