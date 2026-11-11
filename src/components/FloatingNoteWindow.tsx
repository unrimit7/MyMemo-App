import React, { useState, useEffect, useRef } from 'react';
import type { Note, Project } from '../types';
import { marked } from 'marked';
import { renderAttachmentsInHtml } from '../utils/attachment';
import { stripFolderPrefixFromTitle } from '../utils/folderPath';
import { PopoutWindow } from './PopoutWindow';
import { loadPdfDocument, renderPdfPageToCanvas } from '../utils/pdfViewer';
import { loadHwpDocument, renderHwpPageToSvg } from '../utils/hwpViewer';
import { renderDocxToContainer } from '../utils/docxViewer';
import { PdfViewerModal } from './PdfViewerModal';
import { HwpViewerModal } from './HwpViewerModal';
import { DocxViewerModal } from './DocxViewerModal';
import { ExcelViewerModal } from './ExcelViewerModal';
import { ExcelEditorModal } from './ExcelEditorModal';
import { invalidateExcelCache } from '../utils/excelViewer';
import {
  X,
  Maximize2,
  Minimize2,
  Eye,
  Edit3,
  Columns,
  FileText,
  Lock,
  Copy,
  Check,
  ExternalLink,
} from 'lucide-react';

export interface FloatingNoteWindowProps {
  note: Note;
  allNotes?: Note[];
  projects?: Project[];
  onClose: () => void;
  onUpdateNote: (updated: Partial<Note>, specificNoteId?: string) => void;
  initialX?: number;
  initialY?: number;
  zIndex?: number;
  onFocus?: () => void;
  initialPopout?: boolean;
}

export const FloatingNoteWindow: React.FC<FloatingNoteWindowProps> = ({
  note,
  allNotes = [],
  onClose,
  onUpdateNote,
  initialX = 80,
  initialY = 70,
  zIndex = 60,
  onFocus,
  initialPopout = false,
}) => {
  const [isPopout, setIsPopout] = useState<boolean>(initialPopout);
  const [localContent, setLocalContent] = useState(note.content);
  const [viewMode, setViewMode] = useState<'preview' | 'edit' | 'split'>('preview');
  const [isMaximized, setIsMaximized] = useState(false);
  const [copied, setCopied] = useState(false);

  // 모달 상태 (PDF, HWP, 워드, 엑셀 크게 보기 및 직접 편집)
  const [activePdfModal, setActivePdfModal] = useState<{ data: string; title: string; initialPage?: number } | null>(null);
  const [activeHwpModal, setActiveHwpModal] = useState<{ data: string; title: string; initialPage?: number } | null>(null);
  const [activeDocxModal, setActiveDocxModal] = useState<{ data: string; title: string } | null>(null);
  const [activeExcelModal, setActiveExcelModal] = useState<{ data: string; title: string; attId?: string } | null>(null);
  const [activeExcelEditorModal, setActiveExcelEditorModal] = useState<{ data: string; title: string; attId?: string } | null>(null);

  // 미리보기 컨테이너 및 문서 페이지 번호 관리
  const previewContainerRef = useRef<HTMLDivElement>(null);
  const pdfPagesRef = useRef<Map<string, number>>(new Map());
  const hwpPagesRef = useRef<Map<string, number>>(new Map());

  // 위치 및 크기 상태 (인페이지 플로팅 창 전용)
  const [position, setPosition] = useState({ x: initialX, y: initialY });
  const [size, setSize] = useState({
    width: Math.min(760, typeof window !== 'undefined' ? window.innerWidth - 60 : 760),
    height: Math.min(680, typeof window !== 'undefined' ? window.innerHeight - 100 : 680),
  });

  const [isDragging, setIsDragging] = useState(false);
  const [isResizing, setIsResizing] = useState(false);

  const dragRef = useRef({ startX: 0, startY: 0, initialPosX: 0, initialPosY: 0 });
  const resizeRef = useRef({ startX: 0, startY: 0, initialW: 0, initialH: 0 });
  const syncTimerRef = useRef<any>(null);

  // 외부(메인 에디터)에서 note.content가 변경되었을 때 동기화
  useEffect(() => {
    setLocalContent(note.content);
  }, [note.content]);

  // 상위 메모 업데이트 디바운스
  const handleContentChange = (newContent: string) => {
    setLocalContent(newContent);
    if (syncTimerRef.current) clearTimeout(syncTimerRef.current);
    syncTimerRef.current = setTimeout(() => {
      onUpdateNote(
        {
          content: newContent,
          updatedAt: new Date().toISOString(),
        },
        note.id
      );
    }, 400);
  };

  // 인라인 PDF 페이지 렌더러
  const renderInlinePdf = (attId: string) => {
    const preview = previewContainerRef.current;
    if (!preview) return;

    const viewer = preview.querySelector<HTMLElement>(`.pdf-inline-viewer[data-att-id="${attId}"]`);
    if (!viewer) return;

    const att = note.attachments?.[attId];
    if (!att || !att.data) return;

    const canvas = viewer.querySelector<HTMLCanvasElement>(`#pdf-canvas-${attId}`);
    const loadingBox = viewer.querySelector<HTMLElement>(`#pdf-loading-${attId}`);
    const pageIndicator = viewer.querySelector<HTMLElement>(`#pdf-page-indicator-${attId}`);
    const prevBtn = viewer.querySelector<HTMLButtonElement>('.pdf-btn-prev');
    const nextBtn = viewer.querySelector<HTMLButtonElement>('.pdf-btn-next');

    const currentPage = pdfPagesRef.current.get(attId) || 1;

    loadPdfDocument(att.data, attId)
      .then((doc) => {
        const total = doc.numPages;
        if (pageIndicator) pageIndicator.textContent = `${currentPage} / ${total}`;
        if (prevBtn) prevBtn.disabled = currentPage <= 1;
        if (nextBtn) nextBtn.disabled = currentPage >= total;

        if (canvas) {
          const containerWidth = Math.max(300, viewer.clientWidth || preview.clientWidth || 600);
          renderPdfPageToCanvas(doc, currentPage, canvas, Math.min(containerWidth, 800))
            .then(() => {
              if (loadingBox) loadingBox.classList.add('hidden');
              canvas.classList.remove('hidden');
            })
            .catch((err) => {
              console.error('인라인 PDF 페이지 렌더링 실패:', err);
              if (loadingBox) {
                loadingBox.innerHTML = '<span class="text-rose-400 font-semibold text-xs">⚠️ PDF 렌더링 오류</span>';
              }
            });
        }
      })
      .catch((err) => {
        console.error('인라인 PDF 렌더링 실패:', err);
        if (loadingBox) {
          loadingBox.innerHTML = '<span class="text-rose-400 font-semibold text-xs">⚠️ PDF 로드 오류</span>';
        }
      });
  };

  // 인라인 HWP 페이지 렌더러
  const renderInlineHwp = (attId: string) => {
    const preview = previewContainerRef.current;
    if (!preview) return;

    const viewer = preview.querySelector<HTMLElement>(`.hwp-inline-viewer[data-att-id="${attId}"]`);
    if (!viewer) return;

    const att = note.attachments?.[attId];
    if (!att || !att.data) return;

    const svgContainer = viewer.querySelector<HTMLElement>(`#hwp-svg-container-${attId}`);
    const loadingBox = viewer.querySelector<HTMLElement>(`#hwp-loading-${attId}`);
    const pageIndicator = viewer.querySelector<HTMLElement>(`#hwp-page-indicator-${attId}`);
    const prevBtn = viewer.querySelector<HTMLButtonElement>('.hwp-btn-prev');
    const nextBtn = viewer.querySelector<HTMLButtonElement>('.hwp-btn-next');

    const currentPage = hwpPagesRef.current.get(attId) || 1;

    loadHwpDocument(att.data, attId)
      .then((doc) => {
        const total = doc.pageCount();
        if (pageIndicator) pageIndicator.textContent = `${currentPage} / ${total}`;
        if (prevBtn) prevBtn.disabled = currentPage <= 1;
        if (nextBtn) nextBtn.disabled = currentPage >= total;

        if (svgContainer) {
          const pageSvg = renderHwpPageToSvg(doc, currentPage - 1);
          svgContainer.innerHTML = pageSvg;
          if (loadingBox) loadingBox.classList.add('hidden');
          svgContainer.classList.remove('hidden');
        }
      })
      .catch((err) => {
        console.error('인라인 HWP 렌더링 실패:', err);
        if (loadingBox) {
          loadingBox.innerHTML = '<span class="text-rose-400 font-semibold text-xs">⚠️ HWP 로드 실패</span>';
        }
      });
  };

  // 인라인 Word(DOCX) 페이지 렌더러
  const renderInlineDocx = (attId: string) => {
    const preview = previewContainerRef.current;
    if (!preview) return;

    const viewer = preview.querySelector<HTMLElement>(`.docx-inline-viewer[data-att-id="${attId}"]`);
    if (!viewer) return;

    const att = note.attachments?.[attId];
    if (!att || !att.data) return;

    const docxContainer = viewer.querySelector<HTMLElement>(`#docx-container-${attId}`);
    const loadingBox = viewer.querySelector<HTMLElement>(`#docx-loading-${attId}`);

    if (docxContainer) {
      renderDocxToContainer(att.data, docxContainer)
        .then(() => {
          if (loadingBox) loadingBox.classList.add('hidden');
          docxContainer.classList.remove('hidden');
        })
        .catch((err) => {
          console.error('인라인 워드 렌더링 실패:', err);
          if (loadingBox) {
            loadingBox.innerHTML = '<span class="text-rose-400 font-semibold text-xs">⚠️ 워드 문서 로드 실패</span>';
          }
        });
    }
  };

  // 마크다운 HTML 렌더링
  const renderedHtml = React.useMemo(() => {
    try {
      let normalized = (localContent || '').replace(/^(\s*-\s*\[[ xX]\])([^\s\n])/gm, '$1 $2');
      normalized = normalized.replace(/==([^=\n]+)==/g, '<mark style="background-color: #fef08a; color: #0f172a; padding: 2px 5px; border-radius: 4px;">$1</mark>');
      normalized = normalized.replace(/(<summary>[\s\S]*?<\/summary>)\s*([^\s\n<])/gi, '$1\n\n$2');
      let raw = marked.parse(normalized, { breaks: true, gfm: true }) as string;

      // 위키링크 치환 (내부 [폴더명] 대괄호 지원)
      raw = raw.replace(/\[\[((?:(?!\[\[|\]\])[^\n\r])+?)\]\]/g, (match, inner) => {
        const parts = inner.split('|');
        const targetTitle = parts[0].trim();
        if (!targetTitle) return match;
        const alias = (parts[1] !== undefined ? parts[1].trim() : targetTitle) || targetTitle;
        const cleanTarget = targetTitle.toLowerCase();
        const targetNote = allNotes.find((n) => {
          const rawTitle = n.title.trim().toLowerCase();
          const strippedTitle = stripFolderPrefixFromTitle(n.title).toLowerCase();
          return rawTitle === cleanTarget || strippedTitle === cleanTarget;
        });
        const safeTarget = targetTitle.replace(/"/g, '&quot;');
        const safeAlias = alias.replace(/</g, '&lt;').replace(/>/g, '&gt;');
        if (targetNote) {
          return `<button type="button" class="wikilink-pill wikilink-exists" data-wikilink-id="${targetNote.id}" data-wikilink-title="${safeTarget}" title="메모 바로가기: ${safeTarget}">🔗 ${safeAlias}</button>`;
        } else {
          return `<button type="button" class="wikilink-pill wikilink-missing" data-wikilink-create="true" data-wikilink-title="${safeTarget}" title="새 메모: ${safeTarget}">📄 ${safeAlias} <span class="text-[10px] text-indigo-400 opacity-90 ml-0.5 font-normal">+생성</span></button>`;
        }
      });

      return renderAttachmentsInHtml(raw, note.attachments || {});
    } catch {
      return '<p class="text-rose-400">마크다운을 렌더링할 수 없습니다.</p>';
    }
  }, [localContent, note.attachments, allNotes]);

  // 모든 인라인 문서 일괄 렌더링 함수
  const renderAllInlineDocuments = () => {
    const preview = previewContainerRef.current;
    if (!preview) return;

    const pdfViewers = preview.querySelectorAll<HTMLElement>('.pdf-inline-viewer');
    pdfViewers.forEach((v) => {
      const attId = v.getAttribute('data-att-id');
      if (attId) {
        if (!pdfPagesRef.current.has(attId)) {
          pdfPagesRef.current.set(attId, 1);
        }
        renderInlinePdf(attId);
      }
    });

    const hwpViewers = preview.querySelectorAll<HTMLElement>('.hwp-inline-viewer');
    hwpViewers.forEach((v) => {
      const attId = v.getAttribute('data-att-id');
      if (attId) {
        if (!hwpPagesRef.current.has(attId)) {
          hwpPagesRef.current.set(attId, 1);
        }
        renderInlineHwp(attId);
      }
    });

    const docxViewers = preview.querySelectorAll<HTMLElement>('.docx-inline-viewer');
    docxViewers.forEach((v) => {
      const attId = v.getAttribute('data-att-id');
      if (attId) {
        renderInlineDocx(attId);
      }
    });
  };

  // 미리보기 화면 갱신 또는 모드 전환 시 인라인 문서 자동 렌더링
  useEffect(() => {
    if (viewMode !== 'preview' && viewMode !== 'split') return;

    // 초기 마운트 직후 DOM 부착 대기
    const timer = setTimeout(() => {
      renderAllInlineDocuments();
    }, 60);

    const preview = previewContainerRef.current;
    if (!preview) return () => clearTimeout(timer);

    // <details> 토글을 펼칠 때 숨겨져 있던 문서 자동 재렌더링
    const handleDetailsToggle = (e: Event) => {
      const details = e.target as HTMLElement;
      if (!details || details.tagName !== 'DETAILS') return;
      if (!(details as HTMLDetailsElement).open) return;

      const dViewers = details.querySelectorAll<HTMLElement>('.docx-inline-viewer');
      dViewers.forEach((v) => {
        const attId = v.getAttribute('data-att-id');
        if (attId) renderInlineDocx(attId);
      });
      const hViewers = details.querySelectorAll<HTMLElement>('.hwp-inline-viewer');
      hViewers.forEach((v) => {
        const attId = v.getAttribute('data-att-id');
        if (attId) renderInlineHwp(attId);
      });
      const pViewers = details.querySelectorAll<HTMLElement>('.pdf-inline-viewer');
      pViewers.forEach((v) => {
        const attId = v.getAttribute('data-att-id');
        if (attId) renderInlinePdf(attId);
      });
    };

    preview.addEventListener('toggle', handleDetailsToggle, true);

    return () => {
      clearTimeout(timer);
      preview.removeEventListener('toggle', handleDetailsToggle, true);
    };
  }, [renderedHtml, viewMode, isPopout]);

  // 창 드래그 시작 (인페이지 플로팅 전용)
  const handleHeaderPointerDown = (e: React.PointerEvent) => {
    if (isMaximized || isPopout) return;
    const target = e.target as HTMLElement;
    if (target.closest('button, input, select, textarea, a, .no-drag')) return;
    if (onFocus) onFocus();

    e.preventDefault();
    setIsDragging(true);
    dragRef.current = {
      startX: e.clientX,
      startY: e.clientY,
      initialPosX: position.x,
      initialPosY: position.y,
    };
  };

  // 창 크기 조절 시작 (인페이지 플로팅 전용)
  const handleResizePointerDown = (e: React.PointerEvent) => {
    if (isMaximized || isPopout) return;
    if (onFocus) onFocus();

    e.preventDefault();
    e.stopPropagation();
    setIsResizing(true);
    resizeRef.current = {
      startX: e.clientX,
      startY: e.clientY,
      initialW: size.width,
      initialH: size.height,
    };
  };

  // 전역 마우스/터치 리스너
  useEffect(() => {
    if (!isDragging && !isResizing) return;

    const handlePointerMove = (e: PointerEvent) => {
      if (isDragging) {
        const deltaX = e.clientX - dragRef.current.startX;
        const deltaY = e.clientY - dragRef.current.startY;
        const newX = dragRef.current.initialPosX + deltaX;
        const newY = dragRef.current.initialPosY + deltaY;

        const maxW = typeof window !== 'undefined' ? window.innerWidth : 1200;
        const maxH = typeof window !== 'undefined' ? window.innerHeight : 900;
        const clampedX = Math.max(-size.width + 100, Math.min(maxW - 80, newX));
        const clampedY = Math.max(0, Math.min(maxH - 50, newY));

        setPosition({ x: clampedX, y: clampedY });
      }

      if (isResizing) {
        const deltaX = e.clientX - resizeRef.current.startX;
        const deltaY = e.clientY - resizeRef.current.startY;
        const newW = resizeRef.current.initialW + deltaX;
        const newH = resizeRef.current.initialH + deltaY;

        const maxW = typeof window !== 'undefined' ? window.innerWidth - position.x - 10 : 1200;
        const maxH = typeof window !== 'undefined' ? window.innerHeight - position.y - 10 : 900;

        setSize({
          width: Math.max(360, Math.min(maxW, newW)),
          height: Math.max(280, Math.min(maxH, newH)),
        });
      }
    };

    const handlePointerUp = () => {
      setIsDragging(false);
      setIsResizing(false);
    };

    window.addEventListener('pointermove', handlePointerMove);
    window.addEventListener('pointerup', handlePointerUp);

    return () => {
      window.removeEventListener('pointermove', handlePointerMove);
      window.removeEventListener('pointerup', handlePointerUp);
    };
  }, [isDragging, isResizing, size.width, position.x, position.y]);

  // 미리보기 영역 내 클릭 이벤트 핸들러 (체크박스, PDF/HWP 페이지 이동, 토글 접기/해제, 크게 보기 등)
  const handlePreviewClick = (e: React.MouseEvent<HTMLDivElement>) => {
    const target = e.target as HTMLElement;

    // 1. 체크박스 원클릭 토글
    if (target.tagName === 'INPUT' && (target as HTMLInputElement).type === 'checkbox') {
      const checkbox = target as HTMLInputElement;
      const isChecked = checkbox.checked;
      const checkboxes = Array.from(e.currentTarget.querySelectorAll('input[type="checkbox"]'));
      const index = checkboxes.indexOf(checkbox);
      if (index >= 0) {
        let count = 0;
        const newContent = localContent.replace(/^(\s*-\s*\[)([ xX])(\])/gm, (match, prefix, _mark, suffix) => {
          if (count === index) {
            count++;
            return `${prefix}${isChecked ? 'x' : ' '}${suffix}`;
          }
          count++;
          return match;
        });
        handleContentChange(newContent);
      }
      return;
    }

    // 2. PDF 이전/다음 페이지 버튼
    const prevPdfBtn = target.closest<HTMLButtonElement>('.pdf-btn-prev');
    if (prevPdfBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = prevPdfBtn.getAttribute('data-att-id');
      if (attId) {
        const curr = pdfPagesRef.current.get(attId) || 1;
        if (curr > 1) {
          pdfPagesRef.current.set(attId, curr - 1);
          renderInlinePdf(attId);
        }
      }
      return;
    }

    const nextPdfBtn = target.closest<HTMLButtonElement>('.pdf-btn-next');
    if (nextPdfBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = nextPdfBtn.getAttribute('data-att-id');
      if (attId) {
        const curr = pdfPagesRef.current.get(attId) || 1;
        pdfPagesRef.current.set(attId, curr + 1);
        renderInlinePdf(attId);
      }
      return;
    }

    // 3. HWP 이전/다음 페이지 버튼
    const prevHwpBtn = target.closest<HTMLButtonElement>('.hwp-btn-prev');
    if (prevHwpBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = prevHwpBtn.getAttribute('data-att-id');
      if (attId) {
        const curr = hwpPagesRef.current.get(attId) || 1;
        if (curr > 1) {
          hwpPagesRef.current.set(attId, curr - 1);
          renderInlineHwp(attId);
        }
      }
      return;
    }

    const nextHwpBtn = target.closest<HTMLButtonElement>('.hwp-btn-next');
    if (nextHwpBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = nextHwpBtn.getAttribute('data-att-id');
      if (attId) {
        const curr = hwpPagesRef.current.get(attId) || 1;
        hwpPagesRef.current.set(attId, curr + 1);
        renderInlineHwp(attId);
      }
      return;
    }

    // 4. 첨부파일 [📁 접기 / 📂 접기 해제] 버튼
    const toggleAttBtn = target.closest<HTMLElement>('.att-btn-toggle');
    if (toggleAttBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = toggleAttBtn.getAttribute('data-att-id');
      const fileName = toggleAttBtn.getAttribute('data-file-name') || '문서';
      if (!attId) return;

      const content = localContent;
      const wrappedRegex = new RegExp(
        `[\\r\\n]*<details[^>]*>\\s*<summary>[\\s\\S]*?<\\/summary>\\s*([\\s\\S]*?(!?\\[[^\\]]*\\]\\((?:attachment:)?${attId}\\))[\\s\\S]*?)\\s*<\\/details>[\\r\\n]*`,
        'i'
      );
      const match = content.match(wrappedRegex);

      if (match) {
        const innerContent = match[1].trim();
        const newContent = content.replace(wrappedRegex, `\n\n${innerContent}\n\n`);
        handleContentChange(newContent);
      } else {
        const plainAttRegex = new RegExp(
          `[\\r\\n]*(!?\\[[^\\]]*\\]\\((?:attachment:)?${attId}\\))[\\r\\n]*`,
          'i'
        );
        const plainMatch = content.match(plainAttRegex);
        if (plainMatch) {
          const rawTag = plainMatch[1].trim();
          const wrappedBlock = `\n\n<details>\n<summary>📄 ${fileName}</summary>\n\n${rawTag}\n\n</details>\n\n`;
          const newContent = content.replace(plainAttRegex, wrappedBlock);
          handleContentChange(newContent);
        }
      }
      return;
    }

    // 5. [🔍 크게 보기] 모달 열기
    const fullscreenPdfBtn = target.closest<HTMLElement>('.pdf-btn-fullscreen') || target.closest<HTMLCanvasElement>('canvas.pdf-canvas');
    if (fullscreenPdfBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = fullscreenPdfBtn.getAttribute('data-att-id');
      const title = fullscreenPdfBtn.getAttribute('data-file-name') || 'PDF 문서';
      const att = attId ? note.attachments?.[attId] : null;
      if (att && att.data) {
        const curr = pdfPagesRef.current.get(attId!) || 1;
        setActivePdfModal({ data: att.data, title, initialPage: curr });
      }
      return;
    }

    const fullscreenHwpBtn = target.closest<HTMLElement>('.hwp-btn-fullscreen') || target.closest<HTMLElement>('.hwp-svg-container');
    if (fullscreenHwpBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = fullscreenHwpBtn.getAttribute('data-att-id');
      const title = fullscreenHwpBtn.getAttribute('data-file-name') || '한글 문서';
      const att = attId ? note.attachments?.[attId] : null;
      if (att && att.data) {
        const curr = hwpPagesRef.current.get(attId!) || 1;
        setActiveHwpModal({ data: att.data, title, initialPage: curr });
      }
      return;
    }

    const fullscreenDocxBtn = target.closest<HTMLElement>('.docx-btn-fullscreen');
    if (fullscreenDocxBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = fullscreenDocxBtn.getAttribute('data-att-id');
      const title = fullscreenDocxBtn.getAttribute('data-file-name') || '워드 문서';
      const att = attId ? note.attachments?.[attId] : null;
      if (att && att.data) {
        setActiveDocxModal({ data: att.data, title });
      }
      return;
    }

    // 엑셀 [✏️ 직접 편집] 버튼 클릭 처리
    const editExcelBtn = target.closest<HTMLElement>('.excel-btn-edit');
    if (editExcelBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = editExcelBtn.getAttribute('data-att-id') || '';
      const title = editExcelBtn.getAttribute('data-file-name') || '엑셀 문서';
      const atts = note.attachments || {};
      const targetAtt = (attId && atts[attId])
        || Object.values(atts).find((a) => a.name === title)
        || Object.values(atts).find((a) => {
          const n = (a.name || '').toLowerCase();
          return n.endsWith('.xlsx') || n.endsWith('.xls') || n.endsWith('.csv');
        });

      if (targetAtt && targetAtt.data) {
        const resolvedAttId = targetAtt.id || attId || Object.keys(atts).find((k) => atts[k] === targetAtt) || '';
        setActiveExcelEditorModal({
          data: targetAtt.data,
          title: targetAtt.name || title,
          attId: resolvedAttId,
        });
      }
      return;
    }

    const fullscreenExcelBtn = target.closest<HTMLElement>('.excel-btn-fullscreen');
    if (fullscreenExcelBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = fullscreenExcelBtn.getAttribute('data-att-id') || '';
      const title = fullscreenExcelBtn.getAttribute('data-file-name') || '스프레드시트';
      const atts = note.attachments || {};
      const targetAtt = (attId && atts[attId])
        || Object.values(atts).find((a) => a.name === title)
        || Object.values(atts).find((a) => {
          const n = (a.name || '').toLowerCase();
          return n.endsWith('.xlsx') || n.endsWith('.xls') || n.endsWith('.csv');
        });

      if (targetAtt && targetAtt.data) {
        const resolvedAttId = targetAtt.id || attId || Object.keys(atts).find((k) => atts[k] === targetAtt) || '';
        setActiveExcelModal({ data: targetAtt.data, title: targetAtt.name || title, attId: resolvedAttId });
      }
      return;
    }
  };

  const handleCopy = () => {
    navigator.clipboard.writeText(localContent);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  // 본문 공통 렌더러 (에디터 / 미리보기 / 분할)
  const renderEditorBody = () => (
    <div className="flex-1 min-h-0 flex overflow-hidden bg-slate-950">
      {/* 편집기 (에디터) */}
      {(viewMode === 'edit' || viewMode === 'split') && (
        <div className="flex-1 h-full p-3 sm:p-4 border-r border-slate-800/80 flex flex-col min-w-0">
          <textarea
            value={localContent}
            onChange={(e) => handleContentChange(e.target.value)}
            placeholder="마크다운으로 자유롭게 작성해 보세요..."
            className="w-full flex-1 bg-transparent text-slate-100 text-xs sm:text-sm font-mono leading-relaxed outline-none resize-none placeholder-slate-600 select-text"
            spellCheck={false}
          />
        </div>
      )}

      {/* 마크다운 미리보기 */}
      {(viewMode === 'preview' || viewMode === 'split') && (
        <div
          ref={previewContainerRef}
          onClick={handlePreviewClick}
          className="flex-1 h-full p-4 overflow-y-auto select-text min-w-0 markdown-preview"
        >
          <div
            className="prose prose-invert prose-slate max-w-none text-xs sm:text-sm leading-relaxed"
            dangerouslySetInnerHTML={{ __html: renderedHtml }}
          />
        </div>
      )}
    </div>
  );

  // 문서 크게 보기 모달들 렌더링
  const renderDocumentModals = () => (
    <>
      {activePdfModal && (
        <PdfViewerModal
          isOpen={true}
          data={activePdfModal.data}
          title={activePdfModal.title}
          initialPage={activePdfModal.initialPage}
          onClose={() => setActivePdfModal(null)}
        />
      )}
      {activeHwpModal && (
        <HwpViewerModal
          isOpen={true}
          data={activeHwpModal.data}
          title={activeHwpModal.title}
          initialPage={activeHwpModal.initialPage}
          onClose={() => setActiveHwpModal(null)}
        />
      )}
      {activeDocxModal && (
        <DocxViewerModal
          isOpen={true}
          data={activeDocxModal.data}
          title={activeDocxModal.title}
          onClose={() => setActiveDocxModal(null)}
        />
      )}
      {activeExcelModal && (
        <ExcelViewerModal
          isOpen={true}
          data={activeExcelModal.data}
          title={activeExcelModal.title}
          attId={activeExcelModal.attId || undefined}
          onClose={() => setActiveExcelModal(null)}
          onOpenEditor={() => {
            setActiveExcelEditorModal({
              data: activeExcelModal.data,
              title: activeExcelModal.title,
              attId: activeExcelModal.attId,
            });
            setActiveExcelModal(null);
          }}
        />
      )}
      {activeExcelEditorModal && (
        <ExcelEditorModal
          isOpen={true}
          data={activeExcelEditorModal.data}
          title={activeExcelEditorModal.title}
          attId={activeExcelEditorModal.attId}
          onClose={() => setActiveExcelEditorModal(null)}
          onSave={async (updatedDataUrl) => {
            let attId = activeExcelEditorModal.attId;
            const atts = note.attachments || {};
            if (!attId || !atts[attId]) {
              attId = Object.keys(atts).find((id) => atts[id].name === activeExcelEditorModal.title)
                || Object.keys(atts).find((id) => {
                  const n = (atts[id].name || '').toLowerCase();
                  return n.endsWith('.xlsx') || n.endsWith('.xls') || n.endsWith('.csv');
                }) || '';
            }
            if (attId) {
              invalidateExcelCache(attId);
              invalidateExcelCache();
              const updatedAtts = {
                ...atts,
                [attId]: {
                  ...(atts[attId] || {}),
                  data: updatedDataUrl,
                },
              };
              onUpdateNote({ attachments: updatedAtts }, note.id);
              setActiveExcelEditorModal((prev) => prev ? { ...prev, data: updatedDataUrl } : null);
              setActiveExcelModal((prev) => prev && (prev.attId === attId || prev.title === activeExcelEditorModal.title) ? { ...prev, data: updatedDataUrl } : prev);
            }
          }}
        />
      )}
    </>
  );

  // 1. 브라우저 밖 독립 OS 팝업 창 모드
  if (isPopout) {
    return (
      <PopoutWindow
        title={note.title || '메모'}
        onClose={onClose}
      >
        <div className="w-full h-full flex flex-col bg-slate-950 text-slate-100 overflow-hidden">
          {/* 상단 툴바 헤더 */}
          <div className="px-3.5 py-2.5 bg-slate-900 border-b border-slate-800 flex items-center justify-between gap-2 shrink-0 select-none">
            {/* 제목 & 뱃지 */}
            <div className="flex items-center gap-2 min-w-0 max-w-[50%]">
              <span className="p-1 rounded-md bg-indigo-500/10 text-indigo-400 border border-indigo-500/20 shrink-0">
                <FileText size={15} />
              </span>
              <span className="font-bold text-xs sm:text-sm text-slate-100 truncate" title={note.title}>
                {note.title || '제목 없음'}
              </span>
              {note.isLocalOnly && (
                <span className="text-[10px] bg-amber-500/10 text-amber-400 px-1.5 py-0.5 rounded border border-amber-500/20 shrink-0 flex items-center gap-0.5">
                  <Lock size={9} /> 로컬
                </span>
              )}
              <span className="text-[10px] bg-indigo-500/10 text-indigo-300 px-1.5 py-0.5 rounded border border-indigo-500/20 shrink-0 hidden xs:inline-block font-semibold">
                독립 새 창 (OS 팝업)
              </span>
            </div>

            {/* 컨트롤 버튼 그룹 */}
            <div className="flex items-center gap-1 sm:gap-1.5 shrink-0">
              {/* 뷰 모드 토글 */}
              <div className="flex items-center bg-slate-950 p-0.5 rounded-lg border border-slate-800 text-xs">
                <button
                  type="button"
                  onClick={() => setViewMode('preview')}
                  className={`p-1 sm:px-2 rounded flex items-center gap-1 transition-colors ${
                    viewMode === 'preview' ? 'bg-indigo-600 text-white shadow-xs' : 'text-slate-400 hover:text-slate-200'
                  }`}
                  title="미리보기 모드"
                >
                  <Eye size={13} />
                  <span className="hidden sm:inline text-[11px]">보기</span>
                </button>
                <button
                  type="button"
                  onClick={() => setViewMode('edit')}
                  className={`p-1 sm:px-2 rounded flex items-center gap-1 transition-colors ${
                    viewMode === 'edit' ? 'bg-indigo-600 text-white shadow-xs' : 'text-slate-400 hover:text-slate-200'
                  }`}
                  title="편집 모드"
                >
                  <Edit3 size={13} />
                  <span className="hidden sm:inline text-[11px]">편집</span>
                </button>
                <button
                  type="button"
                  onClick={() => setViewMode('split')}
                  className={`p-1 sm:px-2 rounded flex items-center gap-1 transition-colors ${
                    viewMode === 'split' ? 'bg-indigo-600 text-white shadow-xs' : 'text-slate-400 hover:text-slate-200'
                  }`}
                  title="양방향 분할 모드"
                >
                  <Columns size={13} />
                  <span className="hidden sm:inline text-[11px]">분할</span>
                </button>
              </div>

              {/* 본문 복사 */}
              <button
                type="button"
                onClick={handleCopy}
                className="p-1.5 bg-slate-800 hover:bg-slate-750 text-slate-400 hover:text-slate-200 rounded-lg border border-slate-700/80 transition-colors"
                title="본문 전체 복사"
              >
                {copied ? <Check size={13} className="text-emerald-400" /> : <Copy size={13} />}
              </button>

              {/* 브라우저 내부 창으로 복귀 버튼 */}
              <button
                type="button"
                onClick={() => setIsPopout(false)}
                className="p-1.5 sm:px-2 bg-slate-800 hover:bg-slate-700 text-indigo-300 hover:text-indigo-200 rounded-lg border border-slate-700/80 transition-colors flex items-center gap-1 text-xs"
                title="브라우저 화면 안으로 다시 넣기"
              >
                <Minimize2 size={13} />
                <span className="hidden sm:inline text-[11px]">내부 창으로</span>
              </button>

              {/* 닫기 버튼 */}
              <button
                type="button"
                onClick={onClose}
                className="p-1.5 bg-slate-800 hover:bg-rose-900/40 text-slate-400 hover:text-rose-300 rounded-lg border border-slate-700/80 transition-colors"
                title="창 닫기"
              >
                <X size={15} />
              </button>
            </div>
          </div>

          {/* 본문 에디터/미리보기 */}
          {renderEditorBody()}

          {/* 크게 보기 전면 모달 */}
          {renderDocumentModals()}
        </div>
      </PopoutWindow>
    );
  }

  // 2. 브라우저 화면 내 가상 플로팅 창 모드
  return (
    <div
      onClick={onFocus}
      style={{
        position: 'fixed',
        left: isMaximized ? '0px' : `${position.x}px`,
        top: isMaximized ? '0px' : `${position.y}px`,
        width: isMaximized ? '100vw' : `${size.width}px`,
        height: isMaximized ? '100vh' : `${size.height}px`,
        zIndex,
      }}
      className={`pointer-events-auto flex flex-col bg-slate-900 border border-slate-700/80 shadow-[0_25px_60px_-15px_rgba(0,0,0,0.85)] ring-1 ring-white/10 select-none overflow-hidden transition-[box-shadow] duration-150 ${
        isMaximized ? 'rounded-none' : 'rounded-2xl'
      }`}
    >
      {/* 플로팅 헤더 바 */}
      <div
        onPointerDown={handleHeaderPointerDown}
        style={{ cursor: isMaximized ? 'default' : isDragging ? 'grabbing' : 'grab' }}
        className="px-3.5 py-2.5 bg-slate-900/95 border-b border-slate-800 flex items-center justify-between gap-2 shrink-0 select-none"
      >
        {/* 좌측: 메모 제목 & 로컬 배지 */}
        <div className="flex items-center gap-2 min-w-0 max-w-[45%]">
          <span className="p-1 rounded-md bg-indigo-500/10 text-indigo-400 border border-indigo-500/20 shrink-0">
            <FileText size={15} />
          </span>
          <span className="font-bold text-xs sm:text-sm text-slate-100 truncate" title={note.title}>
            {note.title || '제목 없음'}
          </span>
          {note.isLocalOnly && (
            <span className="text-[10px] bg-amber-500/10 text-amber-400 px-1.5 py-0.5 rounded border border-amber-500/20 shrink-0 flex items-center gap-0.5">
              <Lock size={9} /> 로컬
            </span>
          )}
        </div>

        {/* 우측: 보기/편집 모드 전환 & 새 창 분리 & 전체화면 & 닫기 */}
        <div className="flex items-center gap-1 shrink-0 no-drag">
          {/* 뷰 모드 토글 */}
          <div className="flex items-center bg-slate-950 p-0.5 rounded-lg border border-slate-800 text-xs">
            <button
              type="button"
              onClick={() => setViewMode('preview')}
              className={`p-1 rounded transition-colors ${
                viewMode === 'preview' ? 'bg-indigo-600 text-white shadow-xs' : 'text-slate-400 hover:text-slate-200'
              }`}
              title="미리보기 모드"
            >
              <Eye size={13} />
            </button>
            <button
              type="button"
              onClick={() => setViewMode('edit')}
              className={`p-1 rounded transition-colors ${
                viewMode === 'edit' ? 'bg-indigo-600 text-white shadow-xs' : 'text-slate-400 hover:text-slate-200'
              }`}
              title="편집 모드"
            >
              <Edit3 size={13} />
            </button>
            <button
              type="button"
              onClick={() => setViewMode('split')}
              className={`p-1 rounded transition-colors ${
                viewMode === 'split' ? 'bg-indigo-600 text-white shadow-xs' : 'text-slate-400 hover:text-slate-200'
              }`}
              title="양방향 분할 모드"
            >
              <Columns size={13} />
            </button>
          </div>

          <button
            type="button"
            onClick={handleCopy}
            className="p-1.5 bg-slate-800 hover:bg-slate-750 text-slate-400 hover:text-slate-200 rounded-lg border border-slate-700/80 transition-colors"
            title="본문 전체 복사"
          >
            {copied ? <Check size={13} className="text-emerald-400" /> : <Copy size={13} />}
          </button>

          {/* 브라우저 밖 독립 새 창(OS 팝업)으로 분리 버튼 */}
          <button
            type="button"
            onClick={() => setIsPopout(true)}
            className="p-1.5 px-2 bg-indigo-950/60 hover:bg-indigo-900 text-indigo-300 hover:text-white rounded-lg border border-indigo-500/40 transition-colors flex items-center gap-1"
            title="브라우저 밖 독립 새 창(OS 팝업)으로 분리 (바탕화면/다른 모니터로 이동 가능)"
          >
            <ExternalLink size={13} className="text-indigo-400" />
            <span className="hidden sm:inline text-[11px] font-semibold">새 창</span>
          </button>

          <button
            type="button"
            onClick={() => setIsMaximized((prev) => !prev)}
            className="p-1.5 bg-slate-800 hover:bg-slate-750 text-slate-400 hover:text-slate-200 rounded-lg border border-slate-700/80 transition-colors"
            title={isMaximized ? '원래 크기로 복원' : '전체화면'}
          >
            {isMaximized ? <Minimize2 size={13} /> : <Maximize2 size={13} />}
          </button>

          <button
            type="button"
            onClick={onClose}
            className="p-1.5 bg-slate-800 hover:bg-rose-900/40 text-slate-400 hover:text-rose-300 rounded-lg border border-slate-700/80 transition-colors"
            title="플로팅 창 닫기"
          >
            <X size={15} />
          </button>
        </div>
      </div>

      {/* 본문 콘텐츠 영역 */}
      {renderEditorBody()}

      {/* 우측 하단 크기 조절 손잡이 */}
      {!isMaximized && (
        <div
          onPointerDown={handleResizePointerDown}
          className="absolute bottom-0 right-0 w-5 h-5 cursor-nwse-resize z-30 flex items-center justify-center text-slate-500 hover:text-indigo-400 pointer-events-auto select-none"
          title="드래그하여 창 크기 조절"
        >
          <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor">
            <circle cx="10" cy="10" r="1.2" />
            <circle cx="6" cy="10" r="1.2" />
            <circle cx="10" cy="6" r="1.2" />
            <circle cx="2" cy="10" r="1.2" />
            <circle cx="6" cy="6" r="1.2" />
            <circle cx="10" cy="2" r="1.2" />
          </svg>
        </div>
      )}

      {/* 크게 보기 전면 모달 */}
      {renderDocumentModals()}
    </div>
  );
};
