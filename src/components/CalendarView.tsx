import React, { useState, useMemo, useRef, useEffect, useCallback } from 'react';
import type { Note, Project } from '../types';
import { 
  ChevronLeft, 
  ChevronRight, 
  ChevronDown,
  ChevronUp,
  Plus, 
  X, 
  Check, 
  FileText, 
  Folder, 
  Lock,
  Calendar as CalendarIcon,
  Menu
} from 'lucide-react';

interface CalendarViewProps {
  notes: Note[];
  projects: Project[];
  activeProjectId: string | null;
  onSelectNote: (noteId: string) => void;
  onCreateNoteForDate: (dateStr: string, targetProjectId?: string) => void;
  onClose: () => void;
  onMoveNoteDate?: (noteId: string, targetDateStr: string) => void;
  onOpenSidebar?: () => void;
}

export const CalendarView: React.FC<CalendarViewProps> = ({
  notes,
  projects,
  activeProjectId,
  onSelectNote,
  onCreateNoteForDate,
  onClose,
  onMoveNoteDate,
  onOpenSidebar,
}) => {
  const today = useMemo(() => new Date(), []);
  const [currentYear, setCurrentYear] = useState<number>(today.getFullYear());
  const [currentMonth, setCurrentMonth] = useState<number>(today.getMonth()); // 0-indexed

  // 드래그 앤 드롭 이동 상태 관리
  const [draggedNoteId, setDraggedNoteId] = useState<string | null>(null);
  const [dragOverDateStr, setDragOverDateStr] = useState<string | null>(null);

  // 선택된 날짜 (기본값: 오늘)
  const [selectedDateStr, setSelectedDateStr] = useState<string>(() => {
    const y = today.getFullYear();
    const m = String(today.getMonth() + 1).padStart(2, '0');
    const d = String(today.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  });

  // 특정 날짜 '+N개 더보기' 클릭 시 펼쳐서 표시할 날짜 (YYYY-MM-DD)
  const [expandedDateStr, setExpandedDateStr] = useState<string | null>(null);
  const expandedPanelRef = useRef<HTMLDivElement>(null);

  // 확대된 날짜 목록 밖을 클릭하거나 Esc를 누르면 자동으로 접기
  useEffect(() => {
    if (!expandedDateStr) return;

    const handlePointerDown = (e: MouseEvent | TouchEvent) => {
      if (expandedPanelRef.current && !expandedPanelRef.current.contains(e.target as Node)) {
        setExpandedDateStr(null);
      }
    };

    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        setExpandedDateStr(null);
      }
    };

    document.addEventListener('mousedown', handlePointerDown);
    document.addEventListener('touchstart', handlePointerDown);
    document.addEventListener('keydown', handleKeyDown);

    return () => {
      document.removeEventListener('mousedown', handlePointerDown);
      document.removeEventListener('touchstart', handlePointerDown);
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [expandedDateStr]);

  // 특정 폴더의 모든 하위(자식, 손자 등) 폴더 ID 목록 추출 헬퍼 함수
  const getFolderWithDescendants = useCallback((rootId: string): string[] => {
    const result: string[] = [rootId];
    const findChildren = (parentId: string) => {
      projects.filter((p) => p.parentId === parentId).forEach((child) => {
        result.push(child.id);
        findChildren(child.id);
      });
    };
    findChildren(rootId);
    return result;
  }, [projects]);

  // 폴더 필터 팝오버 열림 상태
  const [isFolderFilterOpen, setIsFolderFilterOpen] = useState(false);
  const filterDropdownRef = useRef<HTMLDivElement>(null);

  // 선택된 폴더 ID 집합 (상위 폴더 클릭 시 하위 폴더들까지 기본으로 자동 포함!)
  const [selectedFolderIds, setSelectedFolderIds] = useState<Set<string>>(() => {
    if (activeProjectId && activeProjectId !== 'ALL') {
      const familyIds = [activeProjectId];
      const findChildren = (parentId: string) => {
        projects.filter((p) => p.parentId === parentId).forEach((child) => {
          familyIds.push(child.id);
          findChildren(child.id);
        });
      };
      findChildren(activeProjectId);
      return new Set(familyIds);
    }
    return new Set(projects.map((p) => p.id));
  });

  // 마지막으로 열었던 폴더 또는 기본 폴더 ID 추출
  const resolveDefaultFolderId = useCallback((): string => {
    // 1. 현재 activeProjectId가 'ALL'이 아니고 유효한 폴더인 경우
    if (activeProjectId && activeProjectId !== 'ALL' && projects.some((p) => p.id === activeProjectId)) {
      return activeProjectId;
    }

    // 2. localStorage에 기억된 마지막 실제 선택 폴더가 있는 경우
    try {
      const lastFolderId = localStorage.getItem('mymemo_last_folder_id');
      if (lastFolderId && projects.some((p) => p.id === lastFolderId)) {
        return lastFolderId;
      }
    } catch {}

    // 3. 최상위(부모가 없는) 일반 폴더 우선 선택 (자식 폴더인 'ㄴ 문제점' 등 방지)
    const topLevel = projects.find((p) => !p.parentId);
    if (topLevel) return topLevel.id;

    return projects[0]?.id || '';
  }, [activeProjectId, projects]);

  // 메모 작성 시 대상이 될 폴더 ID (사용자가 드롭다운으로 언제든 변경 가능)
  const [createFolderId, setCreateFolderId] = useState<string>(resolveDefaultFolderId);

  useEffect(() => {
    if (activeProjectId && activeProjectId !== 'ALL' && projects.some((p) => p.id === activeProjectId)) {
      setCreateFolderId(activeProjectId);
      try { localStorage.setItem('mymemo_last_folder_id', activeProjectId); } catch {}
      // 활성 폴더가 변경되면 캘린더 필터도 해당 폴더 및 모든 하위 폴더로 자동 설정
      const familyIds = getFolderWithDescendants(activeProjectId);
      setSelectedFolderIds(new Set(familyIds));
    } else {
      // activeProjectId가 'ALL'이거나 없는 경우:
      // 직전에 마지막으로 열었던 실제 폴더(localStorage)를 우선 복원하여 설정!
      const fallbackId = resolveDefaultFolderId();
      if (fallbackId) {
        setCreateFolderId(fallbackId);
      }
    }
  }, [activeProjectId, projects, getFolderWithDescendants, resolveDefaultFolderId]);

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
    setExpandedDateStr(null);
    if (currentMonth === 0) {
      setCurrentYear((y) => y - 1);
      setCurrentMonth(11);
    } else {
      setCurrentMonth((m) => m - 1);
    }
  };

  const handleNextMonth = () => {
    setExpandedDateStr(null);
    if (currentMonth === 11) {
      setCurrentYear((y) => y + 1);
      setCurrentMonth(0);
    } else {
      setCurrentMonth((m) => m + 1);
    }
  };

  const handleGoToday = () => {
    setExpandedDateStr(null);
    const now = new Date();
    setCurrentYear(now.getFullYear());
    setCurrentMonth(now.getMonth());
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const d = String(now.getDate()).padStart(2, '0');
    setSelectedDateStr(`${y}-${m}-${d}`);
  };

  // 하단 메모 목록에서 마우스 휠을 위/아래로 굴릴 때 좌/우 가로 스크롤로 부드럽게 변환
  const handleBottomListWheel = (e: React.WheelEvent<HTMLDivElement>) => {
    if (e.deltaY !== 0) {
      e.currentTarget.scrollLeft += e.deltaY;
    }
  };

  // 캘린더 빈 공간 좌/우 드래그 및 터치 스와이프를 통한 월 변경 핸들러
  const touchStartXRef = useRef<number | null>(null);
  const touchStartYRef = useRef<number | null>(null);
  const mouseStartXRef = useRef<number | null>(null);
  const mouseStartYRef = useRef<number | null>(null);
  const isMouseDownOnGridRef = useRef<boolean>(false);

  // 터치 스와이프 (휴대폰 모바일)
  const handleTouchStart = (e: React.TouchEvent) => {
    const target = e.target as HTMLElement;
    if (target.closest('button') || target.closest('input') || target.closest('select')) return;
    touchStartXRef.current = e.touches[0].clientX;
    touchStartYRef.current = e.touches[0].clientY;
  };

  const handleTouchEnd = (e: React.TouchEvent) => {
    if (touchStartXRef.current === null || touchStartYRef.current === null) return;
    const endX = e.changedTouches[0].clientX;
    const endY = e.changedTouches[0].clientY;
    const deltaX = endX - touchStartXRef.current;
    const deltaY = endY - touchStartYRef.current;

    // 수평 스와이프가 수직 스크롤보다 크고 거리가 45px 이상일 때만 월 변경
    if (Math.abs(deltaX) > Math.abs(deltaY) * 1.2 && Math.abs(deltaX) > 45) {
      if (deltaX < 0) {
        handleNextMonth(); // 왼쪽으로 스와이프: 다음 달
      } else {
        handlePrevMonth(); // 오른쪽으로 스와이프: 이전 달
      }
    }

    touchStartXRef.current = null;
    touchStartYRef.current = null;
  };

  // 마우스 드래그 (데스크톱)
  const handleMouseDownOnGrid = (e: React.MouseEvent) => {
    const target = e.target as HTMLElement;
    if (target.closest('button') || target.closest('input') || target.closest('select')) return;
    mouseStartXRef.current = e.clientX;
    mouseStartYRef.current = e.clientY;
    isMouseDownOnGridRef.current = true;
  };

  const handleMouseUpOnGrid = (e: React.MouseEvent) => {
    if (!isMouseDownOnGridRef.current || mouseStartXRef.current === null || mouseStartYRef.current === null) return;
    const deltaX = e.clientX - mouseStartXRef.current;
    const deltaY = e.clientY - mouseStartYRef.current;

    // 수평 드래그 거리가 50px 이상일 때 월 변경
    if (Math.abs(deltaX) > Math.abs(deltaY) * 1.2 && Math.abs(deltaX) > 50 && !draggedNoteId) {
      if (deltaX < 0) {
        handleNextMonth(); // 마우스 왼쪽 드래그: 다음 달
      } else {
        handlePrevMonth(); // 마우스 오른쪽 드래그: 이전 달
      }
    }

    mouseStartXRef.current = null;
    mouseStartYRef.current = null;
    isMouseDownOnGridRef.current = false;
  };

  const handleMouseLeaveGrid = () => {
    isMouseDownOnGridRef.current = false;
    mouseStartXRef.current = null;
    mouseStartYRef.current = null;
  };

  // 메모의 기준 날짜 (YYYY-MM-DD) 추출
  // [1순위] 본문 내용의 날짜 (상단 헤더 # 📋 YYYY-MM-DD 등) -> 제목은 사용자가 캘린더에 보기 좋게 자유롭게 지을 수 있도록 최우선 적용!
  // [2순위] 제목 내 날짜 (YYYY-MM-DD / YYMMDD)
  // [3순위] 메모 생성일시 (createdAt)
  const getNoteDateStr = (note: Note): string => {
    // 1-1. [1순위 최우선] 본문 상단 헤더(# 📋 YYYY-MM-DD, # YYYY-MM-DD 등)에서 날짜 검색
    const headerMatch = note.content.match(/#+\s*.*?\b(\d{4})[-./](\d{1,2})[-./](\d{1,2})\b/);
    if (headerMatch) {
      const year = headerMatch[1];
      const month = headerMatch[2].padStart(2, '0');
      const day = headerMatch[3].padStart(2, '0');
      return `${year}-${month}-${day}`;
    }

    // 1-2. [1순위 보조] 본문 전체에서 첫 번째 YYYY-MM-DD (또는 YYYY.MM.DD, YYYY/MM/DD) 검색
    const contentMatch = note.content.match(/\b(\d{4})[-./](\d{1,2})[-./](\d{1,2})\b/);
    if (contentMatch) {
      const year = contentMatch[1];
      const month = contentMatch[2].padStart(2, '0');
      const day = contentMatch[3].padStart(2, '0');
      return `${year}-${month}-${day}`;
    }

    // 2-1. [2순위] 제목에서 YYYY-MM-DD, YYYY.MM.DD, YYYY/MM/DD 검색
    const titleMatch = note.title.match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
    if (titleMatch) {
      const year = titleMatch[1];
      const month = titleMatch[2].padStart(2, '0');
      const day = titleMatch[3].padStart(2, '0');
      return `${year}-${month}-${day}`;
    }

    // 2-2. [2순위 보조] 제목에서 6자리 YYMMDD 검색 (예: 261102, 261114)
    const titleShortMatch = note.title.match(/\b(2\d)(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])\b/);
    if (titleShortMatch) {
      const year = `20${titleShortMatch[1]}`;
      const month = titleShortMatch[2];
      const day = titleShortMatch[3];
      return `${year}-${month}-${day}`;
    }

    // 3. [3순위] 메모 생성일시(createdAt) 기준
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

  // 메모 작성 대상 폴더 이름
  const createFolderName = createFolderId ? projectMap[createFolderId]?.name : null;

  // 헤더 요약용 폴더명 (체크된 폴더가 1개일 때)
  const singleSelectedFolderName = selectedFolderIds.size === 1 ? projectMap[Array.from(selectedFolderIds)[0]]?.name : null;

  return (
    <div className="flex-1 flex flex-col h-full bg-slate-950 text-slate-100 overflow-hidden select-none">
      {/* 상단 헤더: 달력 내비게이션 & 계층형 트리 체크박스 필터 & 닫기 */}
      <div className="p-3 md:px-6 border-b border-slate-800 bg-slate-900/60 flex flex-wrap items-center justify-between gap-3 shrink-0">
        <div className="flex items-center gap-2 md:gap-3">
          {onOpenSidebar && (
            <button
              type="button"
              onClick={onOpenSidebar}
              className="md:hidden p-1.5 -ml-1 text-slate-400 hover:text-white rounded-lg hover:bg-slate-800 transition-colors shrink-0"
              title="메뉴(사이드바) 열기"
            >
              <Menu size={18} />
            </button>
          )}

          <div className="flex items-center gap-1.5 text-indigo-400 shrink-0">
            <CalendarIcon size={20} className="hidden xs:block" />
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
                : selectedFolderIds.size === 1 && singleSelectedFolderName
                  ? `폴더: ${singleSelectedFolderName}`
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

          {/* 계층형 트리 체크박스 팝오버 창 (모바일에서는 화면 중앙 팝업 모달, 데스크톱에서는 버튼 아래 드롭다운) */}
          {isFolderFilterOpen && (
            <div 
              className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs md:bg-transparent md:p-0 md:absolute md:inset-auto md:right-0 md:top-full md:mt-1.5"
              onClick={() => setIsFolderFilterOpen(false)}
            >
              <div 
                className="w-full max-w-sm md:w-80 bg-slate-900 border border-slate-700/90 rounded-2xl shadow-2xl p-4 md:p-3 space-y-2.5 animate-in fade-in zoom-in-95 duration-150"
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
                        className={`flex items-center justify-between py-2 px-2.5 rounded-lg hover:bg-slate-800/90 transition-colors text-xs cursor-pointer select-none group ${
                          isChecked ? 'text-slate-100 font-medium bg-slate-800/40' : 'text-slate-400'
                        }`}
                        style={{ paddingLeft: `${Math.min(depth * 14 + 10, 48)}px` }}
                        onClick={() => {
                          const familyIds = getFolderWithDescendants(p.id);
                          const next = new Set(selectedFolderIds);
                          const isCurrentlyChecked = next.has(p.id);
                          if (isCurrentlyChecked) {
                            familyIds.forEach((id) => next.delete(id));
                          } else {
                            familyIds.forEach((id) => next.add(id));
                          }
                          setSelectedFolderIds(next);
                        }}
                      >
                        <div className="flex items-center gap-2 truncate">
                          <input
                            type="checkbox"
                            checked={isChecked}
                            onChange={() => {}} // 부모 div onClick으로 처리
                            className="accent-indigo-500 rounded cursor-pointer shrink-0 w-4 h-4"
                          />
                          {depth > 0 && (
                            <span className="text-slate-500 font-mono text-[11px] shrink-0">└</span>
                          )}
                          {p.isLocalOnly ? (
                            <Lock size={13} className="text-amber-400 shrink-0" />
                          ) : (
                            <Folder size={13} className="text-indigo-400 shrink-0" />
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
                              const familyIds = getFolderWithDescendants(p.id);
                              setSelectedFolderIds(new Set(familyIds));
                            }}
                            className="text-[10px] text-indigo-400 hover:text-indigo-300 hover:bg-indigo-950/60 border border-indigo-500/20 px-1.5 py-0.5 rounded md:opacity-0 group-hover:opacity-100 transition-opacity"
                            title="이 폴더 및 하위 폴더만 보기"
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
                    className="px-4 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-semibold shadow-sm"
                  >
                    적용
                  </button>
                </div>
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

      {/* 메인 캘린더 그리드 (빈 공간 좌우 드래그/터치 스와이프로 월 이동 지원) */}
      <div 
        onTouchStart={handleTouchStart}
        onTouchEnd={handleTouchEnd}
        onMouseDown={handleMouseDownOnGrid}
        onMouseUp={handleMouseUpOnGrid}
        onMouseLeave={handleMouseLeaveGrid}
        className="flex-1 grid grid-cols-7 auto-rows-fr overflow-y-auto bg-slate-950 divide-x divide-y divide-slate-800/60 border-b border-slate-800 select-none cursor-default"
      >
        {calendarDays.map((item, idx) => {
          const dayNotes = dateNotesMap[item.dateStr] || [];
          const isToday = item.dateStr === todayStr;
          const isSelected = item.dateStr === selectedDateStr;
          const isDragOver = dragOverDateStr === item.dateStr;
          const isSunday = idx % 7 === 0;
          const isSaturday = idx % 7 === 6;
          const isExpanded = expandedDateStr === item.dateStr;

          return (
            <div
              key={item.dateStr}
              onClick={() => setSelectedDateStr(item.dateStr)}
              onDragOver={(e) => {
                if (!onMoveNoteDate) return;
                e.preventDefault();
                e.stopPropagation();
                e.dataTransfer.dropEffect = 'move';
                if (dragOverDateStr !== item.dateStr) {
                  setDragOverDateStr(item.dateStr);
                }
              }}
              onDragLeave={(e) => {
                e.stopPropagation();
                if (e.currentTarget.contains(e.relatedTarget as Node)) return;
                if (dragOverDateStr === item.dateStr) {
                  setDragOverDateStr(null);
                }
              }}
              onDrop={(e) => {
                if (!onMoveNoteDate) return;
                e.preventDefault();
                e.stopPropagation();
                const noteId = e.dataTransfer.getData('text/plain') || draggedNoteId;
                if (noteId) {
                  onMoveNoteDate(noteId, item.dateStr);
                }
                setDraggedNoteId(null);
                setDragOverDateStr(null);
              }}
              className={`p-1 md:p-1.5 flex flex-col min-h-[90px] md:min-h-[110px] transition-all relative cursor-pointer group/cell ${
                !item.isCurrentMonth ? 'bg-slate-950/40 opacity-40' : 'bg-slate-950 hover:bg-slate-900/50'
              } ${isSelected ? 'ring-2 ring-indigo-500/80 ring-inset bg-indigo-950/20' : ''} ${
                isDragOver ? 'ring-2 ring-indigo-400 bg-indigo-950/70 scale-[1.01] z-20 shadow-xl' : ''
              } ${isExpanded ? 'z-30' : ''}`}
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
                    onCreateNoteForDate(item.dateStr, createFolderId);
                  }}
                  className="opacity-0 group-hover/cell:opacity-100 p-1 text-slate-500 hover:text-indigo-400 hover:bg-slate-800 rounded transition-all"
                  title={`${item.dateStr} 새 메모 작성 (${createFolderName || '현재 폴더'})`}
                >
                  <Plus size={13} />
                </button>
              </div>

              {/* 해당 날짜의 메모 목록 (모바일은 1개, 데스크톱은 3개 표시) */}
              <div className="flex-1 flex flex-col gap-1 overflow-hidden pointer-events-auto">
                {dayNotes.slice(0, 3).map((note, nIdx) => {
                  const hasUndoneTask = note.content.includes('- [ ]');
                  const hasDoneTask = note.content.includes('- [x]') || note.content.includes('- [X]');
                  const folder = projectMap[note.projectId];
                  const isDraggingThis = draggedNoteId === note.id;

                  return (
                    <button
                      key={note.id}
                      type="button"
                      draggable={true}
                      onDragStart={(e) => {
                        e.stopPropagation();
                        e.dataTransfer.setData('text/plain', note.id);
                        e.dataTransfer.effectAllowed = 'move';
                        setDraggedNoteId(note.id);
                      }}
                      onDragEnd={() => {
                        setDraggedNoteId(null);
                        setDragOverDateStr(null);
                      }}
                      onClick={(e) => {
                        e.stopPropagation();
                        onSelectNote(note.id);
                      }}
                      className={`w-full text-left p-1 px-1.5 rounded bg-slate-900/90 hover:bg-indigo-600/30 border border-slate-800 hover:border-indigo-500/40 transition-all flex items-center gap-1 group/item active:scale-95 shadow-xs cursor-grab active:cursor-grabbing ${
                        nIdx > 0 ? 'hidden md:flex' : 'flex'
                      } ${
                        isDraggingThis ? 'opacity-40 scale-95 ring-1 ring-indigo-400' : ''
                      }`}
                      title={`[${folder?.name || '폴더'}] ${note.title} (클릭하여 열기 / 다른 날짜로 드래그 이동)`}
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

                      <span className="truncate text-[10px] md:text-[11px] text-slate-200 group-hover/item:text-indigo-200 font-medium flex-1">
                        {note.title || '제목 없음'}
                      </span>
                    </button>
                  );
                })}

                {/* 모바일에서는 1개 초과 시 '+N개' 배지 버튼, 데스크톱에서는 3개 초과 시 '+N개 더보기' 버튼 */}
                {dayNotes.length > 1 && (
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      setSelectedDateStr(item.dateStr);
                      setExpandedDateStr(isExpanded ? null : item.dateStr);
                    }}
                    className="md:hidden text-[9px] text-indigo-300 font-bold px-1 py-0.5 bg-indigo-950/80 hover:bg-indigo-900 rounded border border-indigo-500/30 w-fit shrink-0 active:scale-95 cursor-pointer"
                    title="클릭하여 전체 메모 목록 아래로 펼치기"
                  >
                    +{dayNotes.length - 1}개
                  </button>
                )}
                {dayNotes.length > 3 && (
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      setSelectedDateStr(item.dateStr);
                      setExpandedDateStr(isExpanded ? null : item.dateStr);
                    }}
                    className="hidden md:flex items-center gap-1 text-[10px] text-indigo-400 hover:text-indigo-200 font-bold px-1.5 py-0.5 rounded-md hover:bg-indigo-600/20 border border-transparent hover:border-indigo-500/30 transition-all text-left w-fit active:scale-95 cursor-pointer mt-0.5"
                    title="클릭하여 전체 메모 목록 아래로 펼쳐보기"
                  >
                    <span>+{dayNotes.length - 3}개 더보기</span>
                    <ChevronDown size={11} className="text-indigo-400" />
                  </button>
                )}
              </div>

              {/* 아래로 확대 펼침 팝업 패널 */}
              {isExpanded && (
                <div
                  ref={expandedPanelRef}
                  onClick={(e) => e.stopPropagation()}
                  className={`absolute top-0 z-40 bg-slate-900/98 backdrop-blur-md border-2 border-indigo-500 rounded-xl shadow-2xl p-2.5 flex flex-col gap-1.5 animate-in fade-in zoom-in-95 duration-150 ${
                    idx % 7 >= 5 ? 'right-0' : 'left-0'
                  }`}
                  style={{ minWidth: '100%', width: 'max(100%, 220px)' }}
                >
                  {/* 상단 헤더: 날짜 뱃지 + 요일 + 개수 + 새 메모 / 닫기 버튼 */}
                  <div className="flex items-center justify-between pb-1.5 border-b border-slate-800 shrink-0">
                    <div className="flex items-center gap-1.5 min-w-0">
                      <span
                        className={`text-xs font-bold w-6 h-6 flex items-center justify-center rounded-full shrink-0 ${
                          isToday
                            ? 'bg-indigo-600 text-white font-black shadow-md shadow-indigo-600/30'
                            : isSunday
                              ? 'text-rose-400 bg-rose-500/10'
                              : isSaturday
                                ? 'text-sky-400 bg-sky-500/10'
                                : 'text-slate-200 bg-slate-800'
                        }`}
                      >
                        {item.dayNum}
                      </span>
                      <div className="min-w-0">
                        <div className="text-[11px] font-bold text-slate-100 flex items-center gap-1">
                          <span>{item.dayNum}일 ({['일','월','화','수','목','금','토'][idx % 7]})</span>
                          <span className="text-[10px] text-indigo-400 font-semibold bg-indigo-500/10 px-1 rounded border border-indigo-500/20">
                            전체 {dayNotes.length}개
                          </span>
                        </div>
                      </div>
                    </div>

                    <div className="flex items-center gap-1 shrink-0">
                      <button
                        type="button"
                        onClick={() => onCreateNoteForDate(item.dateStr, createFolderId)}
                        className="p-1 text-slate-400 hover:text-indigo-400 hover:bg-slate-800 rounded transition-all"
                        title={`${item.dateStr} 새 메모 추가`}
                      >
                        <Plus size={13} />
                      </button>
                      <button
                        type="button"
                        onClick={() => setExpandedDateStr(null)}
                        className="p-1 text-slate-400 hover:text-rose-400 hover:bg-slate-800 rounded transition-all"
                        title="접기 (Esc)"
                      >
                        <X size={13} />
                      </button>
                    </div>
                  </div>

                  {/* 전체 메모 카드 목록 (스크롤 가능, 모든 메모 클릭 및 드래그 지원) */}
                  <div className="flex-1 flex flex-col gap-1 overflow-y-auto max-h-[260px] pr-0.5 scrollbar-thin scrollbar-thumb-slate-700 scrollbar-track-transparent">
                    {dayNotes.map((note) => {
                      const hasUndoneTask = note.content.includes('- [ ]');
                      const hasDoneTask = note.content.includes('- [x]') || note.content.includes('- [X]');
                      const folder = projectMap[note.projectId];
                      const isDraggingThis = draggedNoteId === note.id;

                      return (
                        <button
                          key={note.id}
                          type="button"
                          draggable={true}
                          onDragStart={(e) => {
                            e.stopPropagation();
                            e.dataTransfer.setData('text/plain', note.id);
                            e.dataTransfer.effectAllowed = 'move';
                            setDraggedNoteId(note.id);
                          }}
                          onDragEnd={() => {
                            setDraggedNoteId(null);
                            setDragOverDateStr(null);
                          }}
                          onClick={(e) => {
                            e.stopPropagation();
                            setExpandedDateStr(null);
                            onSelectNote(note.id);
                          }}
                          className={`w-full text-left p-1.5 px-2 rounded-lg bg-slate-800/90 hover:bg-indigo-600/35 border border-slate-700/80 hover:border-indigo-500/60 transition-all flex items-center gap-1.5 group/card active:scale-98 shadow-xs cursor-grab active:cursor-grabbing ${
                            isDraggingThis ? 'opacity-40 scale-95 ring-1 ring-indigo-400' : ''
                          }`}
                          title={`[${folder?.name || '폴더'}] ${note.title} (클릭하여 열기 / 다른 날짜로 드래그 이동)`}
                        >
                          {hasUndoneTask ? (
                            <span className="text-amber-400 shrink-0 text-xs font-bold leading-none">☐</span>
                          ) : hasDoneTask ? (
                            <Check size={12} className="text-emerald-400 shrink-0" />
                          ) : note.isLocalOnly ? (
                            <Lock size={12} className="text-amber-400 shrink-0" />
                          ) : (
                            <span className="text-indigo-400 shrink-0 text-xs font-bold leading-none">•</span>
                          )}

                          <div className="flex-1 min-w-0 flex items-center gap-1.5">
                            {folder && (
                              <span className="text-[10px] text-slate-400 font-mono truncate max-w-[80px] shrink-0">
                                [{folder.name}]
                              </span>
                            )}
                            <span className="truncate text-xs text-slate-200 group-hover/card:text-white font-medium flex-1">
                              {note.title || '제목 없음'}
                            </span>
                          </div>
                        </button>
                      );
                    })}
                  </div>

                  {/* 하단 접기 버튼 */}
                  <button
                    type="button"
                    onClick={() => setExpandedDateStr(null)}
                    className="w-full py-1 text-center text-[11px] text-indigo-400 hover:text-indigo-300 font-semibold bg-indigo-950/50 hover:bg-indigo-900/60 rounded-lg border border-indigo-500/25 transition-all flex items-center justify-center gap-1 active:scale-98 mt-0.5"
                  >
                    <ChevronUp size={12} />
                    <span>접기</span>
                  </button>
                </div>
              )}
            </div>
          );
        })}
      </div>

      {/* 하단: 선택된 날짜 상세 패널 (모바일 및 날짜별 메모 한눈에 보기) */}
      <div className="p-3 md:px-6 border-t border-slate-800 bg-slate-900/90 flex flex-col md:flex-row md:items-center justify-between gap-3 shrink-0">
        <div className="flex flex-wrap items-center gap-2.5">
          <div className="flex items-center gap-2 shrink-0">
            <CalendarIcon size={16} className="text-indigo-400" />
            <span className="text-xs md:text-sm font-bold text-slate-100 whitespace-nowrap">
              {selectedDateStr} 메모 목록 ({selectedDateNotes.length}개)
            </span>
          </div>

          {/* 추가할 대상 폴더 선택 드롭다운 & 새 메모 추가 버튼 */}
          <div className="flex items-center gap-1.5 shrink-0">
            <div className="flex items-center gap-1 bg-slate-800/90 border border-slate-700/80 rounded-lg px-2 py-1 shadow-xs">
              <Folder size={12} className="text-indigo-400 shrink-0" />
              <select
                value={createFolderId}
                onChange={(e) => {
                  const val = e.target.value;
                  setCreateFolderId(val);
                  try { localStorage.setItem('mymemo_last_folder_id', val); } catch {}
                }}
                className="bg-transparent text-xs text-slate-200 font-medium focus:outline-hidden cursor-pointer max-w-[130px] sm:max-w-[180px] truncate"
                title="새 메모를 추가할 대상 폴더 선택"
              >
                {hierarchicalProjects.map(({ project: p, depth }) => (
                  <option key={p.id} value={p.id} className="bg-slate-900 text-slate-200">
                    {depth > 0 ? `${'\u00A0'.repeat(depth * 3)}└ ` : ''}
                    {p.isLocalOnly ? '🔒 ' : '📁 '}
                    {p.name}
                  </option>
                ))}
              </select>
            </div>

            <button
              onClick={() => onCreateNoteForDate(selectedDateStr, createFolderId)}
              className="flex items-center gap-1 px-2.5 py-1 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-semibold shadow-sm transition-all active:scale-95 shrink-0"
              title={`${createFolderName ? `'${createFolderName}' 폴더에 ` : ''}선택한 날짜로 새 메모 작성`}
            >
              <Plus size={13} />
              <span>
                {createFolderName ? `'${createFolderName}'에 새 메모` : '이 날짜에 새 메모'}
              </span>
            </button>
          </div>
        </div>

        {/* 선택 날짜의 메모 태그 리스트 (마우스 휠 위/아래 회전 시 좌우로 가로 스크롤) */}
        <div 
          onWheel={handleBottomListWheel}
          className="flex-1 min-w-0 flex items-center gap-1.5 overflow-x-auto scrollbar-thin scrollbar-thumb-slate-700/60 scrollbar-track-transparent py-1"
        >
          {selectedDateNotes.length === 0 ? (
            <span className="text-xs text-slate-500">작성된 메모가 없습니다. [+ 새 메모]를 눌러 기록해 보세요!</span>
          ) : (
            selectedDateNotes.map((note) => {
              const folder = projectMap[note.projectId];
              const isDraggingThis = draggedNoteId === note.id;
              return (
                <button
                  key={note.id}
                  draggable={true}
                  onDragStart={(e) => {
                    e.stopPropagation();
                    e.dataTransfer.setData('text/plain', note.id);
                    e.dataTransfer.effectAllowed = 'move';
                    setDraggedNoteId(note.id);
                  }}
                  onDragEnd={() => {
                    setDraggedNoteId(null);
                    setDragOverDateStr(null);
                  }}
                  onClick={() => onSelectNote(note.id)}
                  className={`flex items-center gap-1.5 px-2.5 py-1 bg-slate-800 hover:bg-indigo-600/30 border border-slate-700/80 hover:border-indigo-500/50 rounded-lg text-xs font-medium text-slate-200 hover:text-white transition-all shrink-0 active:scale-95 cursor-grab active:cursor-grabbing ${
                    isDraggingThis ? 'opacity-40 scale-95 ring-1 ring-indigo-400' : ''
                  }`}
                  title={`[${folder?.name || '폴더'}] ${note.title} (달력 칸으로 드래그하여 날짜 변경 가능)`}
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
