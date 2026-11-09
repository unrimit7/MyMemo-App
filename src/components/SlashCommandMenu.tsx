import React, { useEffect, useRef } from 'react';
import {
  Table,
  ChevronRight,
  Info,
  Lightbulb,
  AlertTriangle,
  BookmarkCheck,
  ShieldAlert,
  CheckSquare,
  Calendar,
  Clock,
  Heading1,
  Heading2,
  Heading3,
  Code,
  Minus,
  Quote,
  Sparkles,
  Link2,
} from 'lucide-react';

export interface SlashCommandItem {
  id: string;
  category: '특수 블록' | '기본 서식' | '실무 도구';
  title: string;
  description: string;
  icon: React.ReactNode;
  keywords: string[];
  snippet: string;
  cursorOffset?: number; // 커서 위치 조정 (기본값: 스니펫 끝)
}

export function getSlashCommands(): SlashCommandItem[] {
  const now = new Date();
  const days = ['일', '월', '화', '수', '목', '금', '토'];
  const dateStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(
    now.getDate()
  ).padStart(2, '0')} (${days[now.getDay()]})`;
  const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(
    2,
    '0'
  )}:${String(now.getSeconds()).padStart(2, '0')}`;

  return [
    // 1. 특수 블록 서식
    {
      id: 'wikilink',
      category: '특수 블록',
      title: '메모 연결 위키링크 ([[ ]])',
      description: '다른 메모로 바로 이동하거나 새 메모를 생성하는 양방향 링크',
      icon: <Link2 size={16} className="text-indigo-400" />,
      keywords: ['링크', '위키', '위키링크', '백링크', '연결', 'link', 'flrzm'],
      snippet: '[[]]',
      cursorOffset: 2,
    },
    {
      id: 'table',
      category: '특수 블록',
      title: '표 삽입 (Table)',
      description: '3열 2행 마크다운 데이터 표 생성',
      icon: <Table size={16} className="text-orange-400" />,
      keywords: ['표', 'table', '스프레드시트', '격자', 'vhy'],
      snippet: '\n\n| 항목 | 내용 | 상태 |\n| :--- | :--- | :---: |\n| 항목 1 | 내용 1 | 완료 |\n| 항목 2 | 내용 2 | 진행중 |\n\n',
    },
    {
      id: 'toggle',
      category: '특수 블록',
      title: '접이식 토글 목록 (Toggle)',
      description: '클릭하면 펼쳐지는 세련된 접이식 내용 상자',
      icon: <ChevronRight size={16} className="text-amber-400" />,
      keywords: ['토글', 'toggle', '접기', '펼치기', 'details', 'summary', 'xhrmf'],
      snippet: '\n<details>\n<summary>여기를 클릭하여 상세 내용 펼치기</summary>\n\n숨겨둘 상세 내용이나 부가 설명을 여기에 작성하세요.\n- 상세 항목 A\n- 상세 항목 B\n</details>\n\n',
    },
    {
      id: 'callout-note',
      category: '특수 블록',
      title: '참고 콜아웃 (Note)',
      description: '파란색 안내 및 참고 사항 강조 상자',
      icon: <Info size={16} className="text-blue-400" />,
      keywords: ['노트', 'note', '참고', '안내', '콜아웃', 'ckarh'],
      snippet: '\n> [!NOTE]\n> 참고할 내용이나 일반 안내 사항을 여기에 작성하세요.\n\n',
    },
    {
      id: 'callout-tip',
      category: '특수 블록',
      title: '유용한 팁 콜아웃 (Tip)',
      description: '초록색 팁 및 권장 사항 강조 상자',
      icon: <Lightbulb size={16} className="text-emerald-400" />,
      keywords: ['팁', 'tip', '꿀팁', '권장', '추천', 'xlq'],
      snippet: '\n> [!TIP]\n> 업무에 유용한 팁이나 권장 사항을 여기에 작성하세요.\n\n',
    },
    {
      id: 'callout-important',
      category: '특수 블록',
      title: '중요 강조 콜아웃 (Important)',
      description: '보라색 핵심 및 필수 확인 강조 상자',
      icon: <BookmarkCheck size={16} className="text-purple-400" />,
      keywords: ['중요', 'important', '핵심', '필수', 'wnddy'],
      snippet: '\n> [!IMPORTANT]\n> 반드시 확인해야 할 핵심 요약 사항을 여기에 작성하세요.\n\n',
    },
    {
      id: 'callout-warning',
      category: '특수 블록',
      title: '주의 콜아웃 (Warning)',
      description: '노란색 주의 및 경고 강조 상자',
      icon: <AlertTriangle size={16} className="text-amber-400" />,
      keywords: ['주의', 'warning', '경고', '조심', 'wndml'],
      snippet: '\n> [!WARNING]\n> 주의해야 할 점이나 주의 사항을 여기에 작성하세요.\n\n',
    },
    {
      id: 'callout-caution',
      category: '특수 블록',
      title: '위험 콜아웃 (Caution)',
      description: '빨간색 위험 및 고장 주의 강조 상자',
      icon: <ShieldAlert size={16} className="text-rose-400" />,
      keywords: ['위험', 'caution', '금지', '오류', 'dnlgja'],
      snippet: '\n> [!CAUTION]\n> 치명적인 위험이나 중대한 오류 요소를 여기에 작성하세요.\n\n',
    },

    // 2. 기본 서식
    {
      id: 'todo',
      category: '기본 서식',
      title: '체크리스트 (To-Do)',
      description: '클릭으로 완료 체크 가능한 할 일 목록',
      icon: <CheckSquare size={16} className="text-emerald-400" />,
      keywords: ['체크', '할일', 'todo', 'task', 'cpzm', 'gkfdlf'],
      snippet: '- [ ] ',
    },
    {
      id: 'h1',
      category: '기본 서식',
      title: '대제목 1 (H1)',
      description: '가장 큰 핵심 섹션 제목',
      icon: <Heading1 size={16} className="text-slate-300" />,
      keywords: ['제목', 'h1', '대제목', 'wpahr'],
      snippet: '# ',
    },
    {
      id: 'h2',
      category: '기본 서식',
      title: '중제목 2 (H2)',
      description: '중간 크기의 소단락 제목',
      icon: <Heading2 size={16} className="text-slate-300" />,
      keywords: ['제목', 'h2', '중제목', 'wpahr'],
      snippet: '## ',
    },
    {
      id: 'h3',
      category: '기본 서식',
      title: '소제목 3 (H3)',
      description: '작은 하위 항목 제목',
      icon: <Heading3 size={16} className="text-slate-300" />,
      keywords: ['제목', 'h3', '소제목', 'wpahr'],
      snippet: '### ',
    },
    {
      id: 'quote',
      category: '기본 서식',
      title: '인용구 (Quote)',
      description: '들여쓰기 인용 단락',
      icon: <Quote size={16} className="text-slate-300" />,
      keywords: ['인용', 'quote', '인용구', 'dlsdyd'],
      snippet: '> ',
    },
    {
      id: 'divider',
      category: '기본 서식',
      title: '구분선 (Divider)',
      description: '단락을 나누는 가로 분할선',
      icon: <Minus size={16} className="text-slate-400" />,
      keywords: ['구분선', '선', 'divider', 'hr', 'line', 'rnqnstjs'],
      snippet: '\n\n---\n\n',
    },
    {
      id: 'code',
      category: '기본 서식',
      title: '코드 블록 (Code Block)',
      description: '프로그래밍/스크립트 문법 강조 상자',
      icon: <Code size={16} className="text-sky-400" />,
      keywords: ['코드', 'code', '소스', 'zhde'],
      snippet: '\n```\n\n```\n',
    },

    // 3. 실무 도구
    {
      id: 'date',
      category: '실무 도구',
      title: '오늘 날짜 삽입',
      description: `현재 날짜 (${dateStr})`,
      icon: <Calendar size={16} className="text-orange-400" />,
      keywords: ['날짜', 'date', 'today', '오늘', 'skfwk'],
      snippet: dateStr,
    },
    {
      id: 'time',
      category: '실무 도구',
      title: '현재 시각 삽입',
      description: `현재 시각 (${timeStr})`,
      icon: <Clock size={16} className="text-sky-400" />,
      keywords: ['시각', '시간', 'time', 'now', 'tlrkr'],
      snippet: timeStr,
    },
  ];
}

interface SlashCommandMenuProps {
  isOpen: boolean;
  query: string;
  selectedIndex: number;
  onSelectCommand: (cmd: SlashCommandItem) => void;
  onClose: () => void;
  position?: { top: number; left: number };
}

export const SlashCommandMenu: React.FC<SlashCommandMenuProps> = ({
  isOpen,
  query,
  selectedIndex,
  onSelectCommand,
  onClose,
  position,
}) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const commands = getSlashCommands();

  // 검색어 필터링 (초성 및 키워드 포함)
  const filtered = commands.filter((cmd) => {
    if (!query) return true;
    const cleanQ = query.toLowerCase().trim();
    if (cmd.title.toLowerCase().includes(cleanQ)) return true;
    if (cmd.description.toLowerCase().includes(cleanQ)) return true;
    return cmd.keywords.some((k) => k.toLowerCase().includes(cleanQ));
  });

  // 선택된 항목이 스크롤 화면 밖으로 나가면 자동 스크롤
  useEffect(() => {
    if (!containerRef.current) return;
    const activeEl = containerRef.current.querySelector<HTMLElement>(`[data-cmd-index="${selectedIndex}"]`);
    if (activeEl) {
      activeEl.scrollIntoView({ block: 'nearest' });
    }
  }, [selectedIndex]);

  // 바깥 클릭 시 닫기
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        onClose();
      }
    };
    if (isOpen) {
      document.addEventListener('mousedown', handleClickOutside);
    }
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [isOpen, onClose]);

  if (!isOpen) return null;

  return (
    <div
      ref={containerRef}
      style={{
        top: position?.top !== undefined ? `${position.top}px` : 'auto',
        left: position?.left !== undefined ? `${position.left}px` : 'auto',
        maxHeight: '340px',
      }}
      className="absolute z-50 w-72 sm:w-80 bg-slate-900 border border-slate-700/80 rounded-2xl shadow-2xl overflow-hidden flex flex-col text-slate-200 animate-fadeIn backdrop-blur-md"
    >
      {/* 상단 힌트 바 */}
      <div className="px-3.5 py-2 bg-slate-950/80 border-b border-slate-800 flex items-center justify-between text-[11px] text-slate-400 select-none">
        <span className="flex items-center gap-1.5 font-bold text-orange-400">
          <Sparkles size={12} />
          <span>빠른 서식 메뉴</span>
          {query && <span className="font-mono text-slate-300 bg-slate-800 px-1 rounded">/{query}</span>}
        </span>
        <span className="text-[10px] text-slate-500">↑↓ 이동 · Enter 선택 · Esc 닫기</span>
      </div>

      {/* 명령어 목록 */}
      <div className="flex-1 overflow-y-auto p-1.5 space-y-0.5 scrollbar-thin">
        {filtered.length === 0 ? (
          <div className="p-4 text-center text-xs text-slate-500">
            일치하는 서식이 없습니다.
          </div>
        ) : (
          filtered.map((cmd, idx) => {
            const isSelected = idx === selectedIndex;
            return (
              <button
                key={cmd.id}
                type="button"
                data-cmd-index={idx}
                onClick={() => onSelectCommand(cmd)}
                className={`w-full px-2.5 py-2 rounded-xl text-left flex items-center gap-3 transition-colors cursor-pointer ${
                  isSelected
                    ? 'bg-orange-600/25 border border-orange-500/50 text-white'
                    : 'hover:bg-slate-800/80 text-slate-300 border border-transparent'
                }`}
              >
                <div
                  className={`p-1.5 rounded-lg shrink-0 ${
                    isSelected ? 'bg-orange-500/30 text-orange-300' : 'bg-slate-800 text-slate-400'
                  }`}
                >
                  {cmd.icon}
                </div>
                <div className="flex-1 min-w-0">
                  <div className="text-xs font-semibold truncate flex items-center justify-between">
                    <span>{cmd.title}</span>
                    <span className="text-[10px] text-slate-500 font-normal shrink-0">{cmd.category}</span>
                  </div>
                  <div className="text-[11px] text-slate-400 truncate">{cmd.description}</div>
                </div>
              </button>
            );
          })
        )}
      </div>
    </div>
  );
};
