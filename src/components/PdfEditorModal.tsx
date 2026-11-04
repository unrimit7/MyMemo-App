import React, { useState, useEffect, useRef, useCallback } from 'react';
import { 
  X, 
  Save, 
  Download, 
  ChevronLeft, 
  ChevronRight, 
  PenTool, 
  Highlighter, 
  Type, 
  Stamp, 
  RotateCcw, 
  Trash2, 
  ZoomIn, 
  ZoomOut, 
  Loader2, 
  CheckCircle,
  Move
} from 'lucide-react';
import { PDFDocument } from 'pdf-lib';
import { loadPdfDocument, renderPdfPageToCanvas, invalidatePdfCache } from '../utils/pdfViewer';
import type { PDFDocumentProxy } from 'pdfjs-dist';

interface Stroke {
  id: string;
  type: 'pen' | 'highlighter';
  color: string;
  width: number;
  opacity: number;
  points: { x: number; y: number }[];
}

interface TextAnnotation {
  id: string;
  x: number; // 0~1 normalized coordinate
  y: number; // 0~1 normalized coordinate
  text: string;
  fontSize: number;
  color: string;
  hasBox: boolean;
  boxBorderColor: string;
  boxBgColor: string;
}

interface StampAnnotation {
  id: string;
  x: number; // 0~1 normalized coordinate
  y: number; // 0~1 normalized coordinate
  width: number; // normalized width
  height: number; // normalized height
  dataUrl: string;
}

interface HistoryAction {
  type: 'stroke' | 'text' | 'stamp';
  id: string;
}

interface PageAnnotations {
  strokes: Stroke[];
  texts: TextAnnotation[];
  stamps: StampAnnotation[];
  history: HistoryAction[];
}

interface PdfEditorModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: string;
  title: string;
  attId?: string;
  initialPage?: number;
  onSave?: (updatedDataUrl: string) => Promise<void> | void;
}

// Data URL 또는 Base64 문자열을 안전하게 Uint8Array로 변환
function dataUrlToUint8Array(dataOrUrl: string): Uint8Array {
  const commaIdx = dataOrUrl.indexOf(',');
  const base64 = commaIdx !== -1 ? dataOrUrl.slice(commaIdx + 1) : dataOrUrl;
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function uint8ArrayToDataUrl(bytes: Uint8Array, mimeType: string = 'application/pdf'): Promise<string> {
  return new Promise((resolve, reject) => {
    const blob = new Blob([bytes as any], { type: mimeType });
    const reader = new FileReader();
    reader.onloadend = () => {
      resolve(reader.result as string);
    };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

const PEN_COLORS = ['#dc2626', '#2563eb', '#16a34a', '#000000', '#9333ea'];
const HIGHLIGHTER_COLORS = ['#fde047', '#86efac', '#f472b6', '#67e8f9'];
const TEXT_COLORS = ['#000000', '#dc2626', '#2563eb', '#16a34a', '#9333ea'];
const BORDER_COLORS = ['#000000', '#dc2626', '#2563eb', '#16a34a', '#9333ea', '#64748b'];

export const PdfEditorModal: React.FC<PdfEditorModalProps> = ({
  isOpen,
  onClose,
  data,
  title,
  attId,
  initialPage = 1,
  onSave,
}) => {
  const [currentPdfData, setCurrentPdfData] = useState<string>(data);
  const [pdfDoc, setPdfDoc] = useState<PDFDocumentProxy | null>(null);
  const [currentPage, setCurrentPage] = useState<number>(initialPage);
  const [totalPages, setTotalPages] = useState<number>(1);
  const [zoom, setZoom] = useState<number>(1.0);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [saveSuccess, setSaveSuccess] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  // 툴 선택 ('pan' | 'pen' | 'highlighter' | 'text' | 'stamp')
  const [activeTool, setActiveTool] = useState<'pan' | 'pen' | 'highlighter' | 'text' | 'stamp'>('pen');
  const [penColor, setPenColor] = useState<string>('#dc2626');
  const [penWidth, setPenWidth] = useState<number>(3);
  const [highlighterColor, setHighlighterColor] = useState<string>('#fde047');
  const [highlighterWidth, setHighlighterWidth] = useState<number>(20);
  
  // 텍스트 & 사각형 테두리 설정
  const [textFontSize, setTextFontSize] = useState<number>(16);
  const [textColor, setTextColor] = useState<string>('#000000');
  const [textHasBox, setTextHasBox] = useState<boolean>(true);
  const [textBoxBorderColor, setTextBoxBorderColor] = useState<string>('#000000');
  const [textBoxBgColor, setTextBoxBgColor] = useState<string>('#ffffff');

  // 텍스트 입력 팝업 모달 상태
  const [showTextModal, setShowTextModal] = useState<boolean>(false);
  const [pendingTextPos, setPendingTextPos] = useState<{ x: number; y: number } | null>(null);
  const [inputModalText, setInputModalText] = useState<string>('');

  // 서명/도장 모달 및 클릭 위치 상태
  const [showSignPad, setShowSignPad] = useState<boolean>(false);
  const [pendingStampPos, setPendingStampPos] = useState<{ x: number; y: number } | null>(null);

  // 페이지별 주석 데이터 (1-based page index)
  const [annotations, setAnnotations] = useState<Record<number, PageAnnotations>>({});

  const baseCanvasRef = useRef<HTMLCanvasElement>(null);
  const overlayCanvasRef = useRef<HTMLCanvasElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const signCanvasRef = useRef<HTMLCanvasElement>(null);

  // 진행 중인 드로잉 스트로크
  const currentStrokeRef = useRef<Stroke | null>(null);

  // 드래그 중인 도장/텍스트 아이템 추적
  const draggingItemRef = useRef<{
    type: 'stamp' | 'text';
    id: string;
    offsetX: number;
    offsetY: number;
  } | null>(null);

  // [이동] 도구: 화면 팬(Pan Hand Tool) 마우스 드래그 추적
  const [isPanDragging, setIsPanDragging] = useState<boolean>(false);
  const panDragStartRef = useRef<{ x: number; y: number; scrollLeft: number; scrollTop: number } | null>(null);

  // props data 변경 시 로컬 데이터 동기화
  useEffect(() => {
    setCurrentPdfData(data);
  }, [data]);

  // PDF 문서 로딩 (항상 최신 dataSignature 및 attId 반영)
  useEffect(() => {
    if (!isOpen || !currentPdfData) return;

    let isMounted = true;
    setIsLoading(true);
    setErrorMessage(null);
    setCurrentPage(initialPage);
    setZoom(1.0);
    setAnnotations({});

    loadPdfDocument(currentPdfData, attId)
      .then((doc) => {
        if (!isMounted) return;
        setPdfDoc(doc);
        setTotalPages(doc.numPages);
        setIsLoading(false);
      })
      .catch((err) => {
        if (!isMounted) return;
        console.error('PDF 로드 실패:', err);
        setErrorMessage('PDF 문서를 불러오지 못했습니다.');
        setIsLoading(false);
      });

    return () => {
      isMounted = false;
    };
  }, [isOpen, currentPdfData, initialPage, attId]);

  // 베이스 PDF 캔버스 렌더링
  useEffect(() => {
    if (!isOpen || !pdfDoc || !baseCanvasRef.current) return;

    const canvas = baseCanvasRef.current;
    const containerWidth = containerRef.current?.clientWidth || window.innerWidth;
    const baseWidth = Math.min(Math.max(300, containerWidth - 32), 900);

    renderPdfPageToCanvas(pdfDoc, currentPage, canvas, baseWidth, zoom, 0)
      .then(() => {
        // 베이스 캔버스 렌더링 완료 후 오버레이 캔버스 크기 맞춤 및 주석 다시 그리기
        const overlay = overlayCanvasRef.current;
        if (overlay) {
          overlay.width = canvas.width;
          overlay.height = canvas.height;
          overlay.style.width = `${canvas.clientWidth}px`;
          overlay.style.height = `${canvas.clientHeight}px`;
          redrawOverlay();
        }
      })
      .catch((err) => {
        console.error('PDF 페이지 렌더링 오류:', err);
      });
  }, [isOpen, pdfDoc, currentPage, zoom]);

  // [이동] 도구: 마우스 드래그로 화면 스크롤 팬(Pan) 이동
  const handleContainerMouseDown = (e: React.MouseEvent<HTMLDivElement>) => {
    if (e.button !== 0) return; // 좌클릭만
    if (activeTool !== 'pan') return;
    const container = containerRef.current;
    if (!container) return;

    e.preventDefault();
    panDragStartRef.current = {
      x: e.clientX,
      y: e.clientY,
      scrollLeft: container.scrollLeft,
      scrollTop: container.scrollTop,
    };
    setIsPanDragging(true);
  };

  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (!panDragStartRef.current || !containerRef.current) return;
      const dx = e.clientX - panDragStartRef.current.x;
      const dy = e.clientY - panDragStartRef.current.y;
      containerRef.current.scrollLeft = panDragStartRef.current.scrollLeft - dx;
      containerRef.current.scrollTop = panDragStartRef.current.scrollTop - dy;
    };

    const handleMouseUp = () => {
      panDragStartRef.current = null;
      setIsPanDragging(false);
    };

    if (isPanDragging) {
      window.addEventListener('mousemove', handleMouseMove);
      window.addEventListener('mouseup', handleMouseUp);
    }

    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [isPanDragging]);

  // 오버레이 캔버스 재렌더링
  const redrawOverlay = useCallback(() => {
    const overlay = overlayCanvasRef.current;
    if (!overlay) return;
    const ctx = overlay.getContext('2d');
    if (!ctx) return;

    ctx.clearRect(0, 0, overlay.width, overlay.height);

    const pageData = annotations[currentPage];
    if (!pageData) return;

    const w = overlay.width;
    const h = overlay.height;

    // 1. 스트로크 (펜 & 형광펜)
    pageData.strokes.forEach((stroke) => {
      if (stroke.points.length < 2) return;
      ctx.save();
      ctx.beginPath();
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.strokeStyle = stroke.color;
      ctx.lineWidth = stroke.width * (w / 800); // 뷰 크기에 비례 스케일링
      ctx.globalAlpha = stroke.opacity;

      ctx.moveTo(stroke.points[0].x * w, stroke.points[0].y * h);
      for (let i = 1; i < stroke.points.length; i++) {
        ctx.lineTo(stroke.points[i].x * w, stroke.points[i].y * h);
      }
      ctx.stroke();
      ctx.restore();
    });

    // 2. 도장/서명 이미지 (위치 지정 및 드래그 대상)
    pageData.stamps.forEach((stamp) => {
      const img = new Image();
      img.src = stamp.dataUrl;
      const sx = stamp.x * w;
      const sy = stamp.y * h;
      const sw = stamp.width * w;
      const sh = stamp.height * h;

      if (img.complete) {
        ctx.drawImage(img, sx, sy, sw, sh);
      } else {
        img.onload = () => {
          ctx.drawImage(img, sx, sy, sw, sh);
        };
      }

      // 호버/선택 힌트 얇은 외곽선
      ctx.save();
      ctx.strokeStyle = 'rgba(99, 102, 241, 0.4)';
      ctx.setLineDash([4, 4]);
      ctx.lineWidth = 1;
      ctx.strokeRect(sx - 2, sy - 2, sw + 4, sh + 4);
      ctx.restore();
    });

    // 3. 텍스트 주석 (사각형 테두리 및 배경 채우기 포함)
    pageData.texts.forEach((textItem) => {
      ctx.save();
      const scaledSize = Math.max(10, textItem.fontSize * (w / 800));
      ctx.font = `bold ${scaledSize}px Pretendard, -apple-system, BlinkMacSystemFont, sans-serif`;
      const x = textItem.x * w;
      const y = textItem.y * h;

      const lines = textItem.text.split('\n');
      const lineHeight = scaledSize * 1.3;

      let maxLineWidth = 0;
      lines.forEach((l) => {
        maxLineWidth = Math.max(maxLineWidth, ctx.measureText(l).width);
      });
      if (maxLineWidth === 0) {
        maxLineWidth = scaledSize * 3;
      }

      const padX = Math.max(5, scaledSize * 0.4);
      const padY = Math.max(4, scaledSize * 0.3);
      const boxX = x - padX;
      const boxY = y - scaledSize - padY;
      const boxW = maxLineWidth + padX * 2;
      const boxH = lines.length * lineHeight + padY * 2;

      // 3-1. 배경 채우기 (투명이 아닌 경우)
      if (textItem.boxBgColor && textItem.boxBgColor !== 'transparent') {
        ctx.fillStyle = textItem.boxBgColor;
        ctx.fillRect(boxX, boxY, boxW, boxH);
      }

      // 3-2. 사각형 테두리 그리기
      if (textItem.hasBox) {
        ctx.strokeStyle = textItem.boxBorderColor || textItem.color || '#000000';
        ctx.lineWidth = Math.max(1.5, 2 * (w / 800));
        ctx.strokeRect(boxX, boxY, boxW, boxH);
      }

      // 3-3. 텍스트 글자 그리기
      ctx.fillStyle = textItem.color;
      lines.forEach((line, idx) => {
        ctx.fillText(line, x, y + idx * lineHeight);
      });
      ctx.restore();
    });
  }, [annotations, currentPage]);

  useEffect(() => {
    redrawOverlay();
  }, [redrawOverlay]);

  // 마우스/터치 드로잉 및 아이템 드래그 이벤트
  const handlePointerDown = (e: React.PointerEvent<HTMLCanvasElement>) => {
    if (activeTool === 'pan') return;

    const overlay = overlayCanvasRef.current;
    if (!overlay) return;
    const rect = overlay.getBoundingClientRect();

    const normX = Math.max(0, Math.min(1, (e.clientX - rect.left) / rect.width));
    const normY = Math.max(0, Math.min(1, (e.clientY - rect.top) / rect.height));

    const pageData = annotations[currentPage];

    // 1. 기존 도장/서명 위를 클릭했는지 확인 -> 마우스 드래그로 위치 이동 시작
    if (pageData && pageData.stamps.length > 0) {
      const clickedStamp = [...pageData.stamps].reverse().find((s) => {
        return normX >= s.x && normX <= s.x + s.width && normY >= s.y && normY <= s.y + s.height;
      });
      if (clickedStamp) {
        overlay.setPointerCapture(e.pointerId);
        draggingItemRef.current = {
          type: 'stamp',
          id: clickedStamp.id,
          offsetX: normX - clickedStamp.x,
          offsetY: normY - clickedStamp.y,
        };
        return;
      }
    }

    // 2. 기존 텍스트 박스 위를 클릭했는지 확인 -> 마우스 드래그로 위치 이동 시작
    if (pageData && pageData.texts.length > 0) {
      const w = overlay.width;
      const clickedText = [...pageData.texts].reverse().find((t) => {
        const scaledSize = Math.max(10, t.fontSize * (w / 800));
        const lineHeight = scaledSize * 1.3;
        const lines = t.text.split('\n');
        const padX = Math.max(5, scaledSize * 0.4) / w;
        const padY = Math.max(4, scaledSize * 0.3) / overlay.height;
        const boxX = t.x - padX;
        const boxY = t.y - (scaledSize / overlay.height) - padY;
        const boxW = (t.fontSize * 0.7 * t.text.length + 30) / w;
        const boxH = (lines.length * lineHeight) / overlay.height + padY * 2;
        return normX >= boxX && normX <= boxX + boxW && normY >= boxY && normY <= boxY + boxH;
      });
      if (clickedText) {
        overlay.setPointerCapture(e.pointerId);
        draggingItemRef.current = {
          type: 'text',
          id: clickedText.id,
          offsetX: normX - clickedText.x,
          offsetY: normY - clickedText.y,
        };
        return;
      }
    }

    // 3. 서명/도장 도구일 때 빈 공간 클릭 시: 클릭한 위치에 도장 생성 모달 열기
    if (activeTool === 'stamp') {
      setPendingStampPos({ x: normX, y: normY });
      setShowSignPad(true);
      return;
    }

    // 4. 텍스트 도구일 때 빈 공간 클릭 시: 클릭한 위치에 텍스트 생성 모달 열기
    if (activeTool === 'text') {
      setPendingTextPos({ x: normX, y: normY });
      setShowTextModal(true);
      return;
    }

    // 5. 펜 또는 형광펜 드로잉 시작
    if (activeTool === 'pen' || activeTool === 'highlighter') {
      overlay.setPointerCapture(e.pointerId);
      const isHighlighter = activeTool === 'highlighter';
      currentStrokeRef.current = {
        id: `stroke_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
        type: isHighlighter ? 'highlighter' : 'pen',
        color: isHighlighter ? highlighterColor : penColor,
        width: isHighlighter ? highlighterWidth : penWidth,
        opacity: isHighlighter ? 0.4 : 1.0,
        points: [{ x: normX, y: normY }],
      };
    }
  };

  const handlePointerMove = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const overlay = overlayCanvasRef.current;
    if (!overlay) return;
    const rect = overlay.getBoundingClientRect();

    const normX = Math.max(0, Math.min(1, (e.clientX - rect.left) / rect.width));
    const normY = Math.max(0, Math.min(1, (e.clientY - rect.top) / rect.height));

    // A. 도장 또는 텍스트 아이템 드래그 이동 처리
    if (draggingItemRef.current) {
      const drag = draggingItemRef.current;
      setAnnotations((prev) => {
        const curr = prev[currentPage];
        if (!curr) return prev;

        if (drag.type === 'stamp') {
          return {
            ...prev,
            [currentPage]: {
              ...curr,
              stamps: curr.stamps.map((s) =>
                s.id === drag.id
                  ? {
                      ...s,
                      x: Math.max(0, Math.min(1 - s.width, normX - drag.offsetX)),
                      y: Math.max(0, Math.min(1 - s.height, normY - drag.offsetY)),
                    }
                  : s
              ),
            },
          };
        } else if (drag.type === 'text') {
          return {
            ...prev,
            [currentPage]: {
              ...curr,
              texts: curr.texts.map((t) =>
                t.id === drag.id
                  ? {
                      ...t,
                      x: Math.max(0, Math.min(0.9, normX - drag.offsetX)),
                      y: Math.max(0.04, Math.min(0.96, normY - drag.offsetY)),
                    }
                  : t
              ),
            },
          };
        }
        return prev;
      });
      return;
    }

    // B. 실시간 펜/형광펜 스트로크 드로잉 처리
    if (!currentStrokeRef.current || (activeTool !== 'pen' && activeTool !== 'highlighter')) return;

    currentStrokeRef.current.points.push({ x: normX, y: normY });

    const ctx = overlay.getContext('2d');
    if (!ctx) return;
    const stroke = currentStrokeRef.current;
    const len = stroke.points.length;
    if (len < 2) return;

    const p1 = stroke.points[len - 2];
    const p2 = stroke.points[len - 1];

    ctx.save();
    ctx.beginPath();
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.strokeStyle = stroke.color;
    ctx.lineWidth = stroke.width * (overlay.width / 800);
    ctx.globalAlpha = stroke.opacity;
    ctx.moveTo(p1.x * overlay.width, p1.y * overlay.height);
    ctx.lineTo(p2.x * overlay.width, p2.y * overlay.height);
    ctx.stroke();
    ctx.restore();
  };

  const handlePointerUp = (e: React.PointerEvent<HTMLCanvasElement>) => {
    if (draggingItemRef.current) {
      draggingItemRef.current = null;
      try {
        e.currentTarget.releasePointerCapture(e.pointerId);
      } catch {}
      return;
    }

    if (!currentStrokeRef.current) return;
    try {
      e.currentTarget.releasePointerCapture(e.pointerId);
    } catch {
      // ignore
    }

    const finishedStroke = currentStrokeRef.current;
    currentStrokeRef.current = null;

    if (finishedStroke.points.length > 1) {
      setAnnotations((prev) => {
        const curr = prev[currentPage] || { strokes: [], texts: [], stamps: [], history: [] };
        return {
          ...prev,
          [currentPage]: {
            ...curr,
            strokes: [...curr.strokes, finishedStroke],
            history: [...(curr.history || []), { type: 'stroke', id: finishedStroke.id }],
          },
        };
      });
    }
  };

  // 마지막 주석 1개 실행 취소 (Undo) - 정확한 역순(LIFO)으로 마지막 작업부터 취소
  const handleUndo = () => {
    setAnnotations((prev) => {
      const curr = prev[currentPage];
      if (!curr) return prev;

      // 1. unified history 스택이 있으면 가장 최근(마지막) 추가된 액션부터 취소
      if (curr.history && curr.history.length > 0) {
        const lastAction = curr.history[curr.history.length - 1];
        const newHistory = curr.history.slice(0, -1);

        if (lastAction.type === 'stroke') {
          return {
            ...prev,
            [currentPage]: {
              ...curr,
              strokes: curr.strokes.filter((s) => s.id !== lastAction.id),
              history: newHistory,
            },
          };
        } else if (lastAction.type === 'text') {
          return {
            ...prev,
            [currentPage]: {
              ...curr,
              texts: curr.texts.filter((t) => t.id !== lastAction.id),
              history: newHistory,
            },
          };
        } else if (lastAction.type === 'stamp') {
          return {
            ...prev,
            [currentPage]: {
              ...curr,
              stamps: curr.stamps.filter((s) => s.id !== lastAction.id),
              history: newHistory,
            },
          };
        }
      }

      // 2. 만약 history가 비어있는 경우 fallback (이전 호환성)
      if (curr.strokes.length > 0) {
        return {
          ...prev,
          [currentPage]: { ...curr, strokes: curr.strokes.slice(0, -1) },
        };
      }
      if (curr.texts.length > 0) {
        return {
          ...prev,
          [currentPage]: { ...curr, texts: curr.texts.slice(0, -1) },
        };
      }
      if (curr.stamps.length > 0) {
        return {
          ...prev,
          [currentPage]: { ...curr, stamps: curr.stamps.slice(0, -1) },
        };
      }
      return prev;
    });
  };

  // 현재 페이지 주석 전체 삭제
  const handleClearPage = () => {
    if (!window.confirm('현재 페이지의 모든 주석을 지우시겠습니까?')) return;
    setAnnotations((prev) => ({
      ...prev,
      [currentPage]: { strokes: [], texts: [], stamps: [], history: [] },
    }));
  };

  // 모달을 통한 텍스트 및 사각형 박스 삽입 확인
  const handleConfirmInsertText = () => {
    const trimmed = inputModalText.trim();
    // 투명 배경인데 글자가 비어있으면 삽입 취소. 단, 흰색/유색 배경(화이트 패치)인 경우 빈칸('  ')으로 수정테이프처럼 삽입 허용
    if (!trimmed && textBoxBgColor === 'transparent') return;

    const finalText = trimmed || '   ';
    const pos = pendingTextPos || { x: 0.5, y: 0.3 };
    const newId = `text_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
    const newText: TextAnnotation = {
      id: newId,
      x: pos.x,
      y: pos.y,
      text: finalText,
      fontSize: textFontSize,
      color: textColor,
      hasBox: textHasBox,
      boxBorderColor: textBoxBorderColor,
      boxBgColor: textBoxBgColor,
    };

    setAnnotations((prev) => {
      const curr = prev[currentPage] || { strokes: [], texts: [], stamps: [], history: [] };
      return {
        ...prev,
        [currentPage]: {
          ...curr,
          texts: [...curr.texts, newText],
          history: [...(curr.history || []), { type: 'text', id: newId }],
        },
      };
    });

    setInputModalText('');
    setShowTextModal(false);
    setPendingTextPos(null);
  };

  // 서명 패드 드로잉 로직
  const isSigningRef = useRef(false);
  const handleSignPointerDown = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const canvas = signCanvasRef.current;
    if (!canvas) return;
    canvas.setPointerCapture(e.pointerId);
    isSigningRef.current = true;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const rect = canvas.getBoundingClientRect();
    ctx.beginPath();
    ctx.moveTo(e.clientX - rect.left, e.clientY - rect.top);
  };

  const handleSignPointerMove = (e: React.PointerEvent<HTMLCanvasElement>) => {
    if (!isSigningRef.current) return;
    const canvas = signCanvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const rect = canvas.getBoundingClientRect();
    ctx.lineWidth = 3;
    ctx.lineCap = 'round';
    ctx.strokeStyle = '#dc2626'; // 빨간색 인장/서명 색상
    ctx.lineTo(e.clientX - rect.left, e.clientY - rect.top);
    ctx.stroke();
  };

  const handleSignPointerUp = (e: React.PointerEvent<HTMLCanvasElement>) => {
    isSigningRef.current = false;
    try {
      e.currentTarget.releasePointerCapture(e.pointerId);
    } catch {}
  };

  const handleClearSignCanvas = () => {
    const canvas = signCanvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (ctx) ctx.clearRect(0, 0, canvas.width, canvas.height);
  };

  // 완성된 손글씨 서명을 지정된 위치(클릭한 위치)에 삽입
  const handleInsertSignature = () => {
    const canvas = signCanvasRef.current;
    if (!canvas) return;
    const signDataUrl = canvas.toDataURL('image/png');

    const pos = pendingStampPos || { x: 0.5, y: 0.5 };
    const stampId = `stamp_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
    const newStamp: StampAnnotation = {
      id: stampId,
      x: Math.max(0, Math.min(0.8, pos.x - 0.1)),
      y: Math.max(0, Math.min(0.88, pos.y - 0.05)),
      width: 0.2, // 페이지 너비의 20%
      height: 0.1, // 페이지 높이의 10%
      dataUrl: signDataUrl,
    };

    setAnnotations((prev) => {
      const curr = prev[currentPage] || { strokes: [], texts: [], stamps: [], history: [] };
      return {
        ...prev,
        [currentPage]: {
          ...curr,
          stamps: [...curr.stamps, newStamp],
          history: [...(curr.history || []), { type: 'stamp', id: stampId }],
        },
      };
    });

    setShowSignPad(false);
    setPendingStampPos(null);
  };

  // 사전 프리셋 도장(확인/서명) 지정된 위치(클릭한 위치)에 삽입
  const handleInsertPresetStamp = (presetText: string, color: string = '#dc2626') => {
    const offCanvas = document.createElement('canvas');
    offCanvas.width = 160;
    offCanvas.height = 160;
    const ctx = offCanvas.getContext('2d');
    if (!ctx) return;

    // 둥근 도장 원형 테두리
    ctx.strokeStyle = color;
    ctx.lineWidth = 6;
    ctx.beginPath();
    ctx.arc(80, 80, 72, 0, Math.PI * 2);
    ctx.stroke();

    // 도장 글자
    ctx.fillStyle = color;
    ctx.font = 'bold 36px sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(presetText, 80, 80);

    const stampDataUrl = offCanvas.toDataURL('image/png');
    const pos = pendingStampPos || { x: 0.5, y: 0.5 };
    const stampId = `stamp_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
    const newStamp: StampAnnotation = {
      id: stampId,
      x: Math.max(0, Math.min(0.85, pos.x - 0.075)),
      y: Math.max(0, Math.min(0.88, pos.y - 0.05)),
      width: 0.15,
      height: 0.1,
      dataUrl: stampDataUrl,
    };

    setAnnotations((prev) => {
      const curr = prev[currentPage] || { strokes: [], texts: [], stamps: [], history: [] };
      return {
        ...prev,
        [currentPage]: {
          ...curr,
          stamps: [...curr.stamps, newStamp],
          history: [...(curr.history || []), { type: 'stamp', id: stampId }],
        },
      };
    });

    setShowSignPad(false);
    setPendingStampPos(null);
  };

  // 주석이 포함된 수정 PDF 바이너리 생성 (pdf-lib 활용)
  const generateAnnotatedPdfBytes = async (): Promise<Uint8Array> => {
    const originalBytes = dataUrlToUint8Array(currentPdfData);
    const pdfDocument = await PDFDocument.load(originalBytes);
    const numPages = pdfDocument.getPageCount();

    for (let p = 1; p <= numPages; p++) {
      const pageData = annotations[p];
      if (!pageData || (pageData.strokes.length === 0 && pageData.texts.length === 0 && pageData.stamps.length === 0)) {
        continue;
      }

      const pdfPage = pdfDocument.getPage(p - 1);
      const pageWidth = pdfPage.getWidth();
      const pageHeight = pdfPage.getHeight();

      // 고해상도 오프스크린 캔버스에 주석 렌더링
      const offCanvas = document.createElement('canvas');
      const renderScale = 2; // 선명한 벡터급 해상도 보장
      offCanvas.width = pageWidth * renderScale;
      offCanvas.height = pageHeight * renderScale;

      const ctx = offCanvas.getContext('2d');
      if (!ctx) continue;

      ctx.scale(renderScale, renderScale);

      // 1. 스트로크 렌더링
      pageData.strokes.forEach((stroke) => {
        if (stroke.points.length < 2) return;
        ctx.save();
        ctx.beginPath();
        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';
        ctx.strokeStyle = stroke.color;
        ctx.lineWidth = stroke.width;
        ctx.globalAlpha = stroke.opacity;

        ctx.moveTo(stroke.points[0].x * pageWidth, stroke.points[0].y * pageHeight);
        for (let i = 1; i < stroke.points.length; i++) {
          ctx.lineTo(stroke.points[i].x * pageWidth, stroke.points[i].y * pageHeight);
        }
        ctx.stroke();
        ctx.restore();
      });

      // 2. 도장/서명 렌더링
      for (const stamp of pageData.stamps) {
        await new Promise<void>((res) => {
          const img = new Image();
          img.onload = () => {
            ctx.drawImage(
              img,
              stamp.x * pageWidth,
              stamp.y * pageHeight,
              stamp.width * pageWidth,
              stamp.height * pageHeight
            );
            res();
          };
          img.onerror = () => res();
          img.src = stamp.dataUrl;
        });
      }

      // 3. 텍스트 렌더링 (사각형 테두리 및 배경 포함)
      pageData.texts.forEach((textItem) => {
        ctx.save();
        ctx.font = `bold ${textItem.fontSize}px Pretendard, -apple-system, BlinkMacSystemFont, sans-serif`;
        const x = textItem.x * pageWidth;
        const y = textItem.y * pageHeight;
        const lines = textItem.text.split('\n');
        const lineHeight = textItem.fontSize * 1.3;

        let maxLineWidth = 0;
        lines.forEach((l) => {
          maxLineWidth = Math.max(maxLineWidth, ctx.measureText(l).width);
        });
        if (maxLineWidth === 0) {
          maxLineWidth = textItem.fontSize * 3;
        }

        const padX = Math.max(5, textItem.fontSize * 0.4);
        const padY = Math.max(4, textItem.fontSize * 0.3);
        const boxX = x - padX;
        const boxY = y - textItem.fontSize - padY;
        const boxW = maxLineWidth + padX * 2;
        const boxH = lines.length * lineHeight + padY * 2;

        // 배경 채우기
        if (textItem.boxBgColor && textItem.boxBgColor !== 'transparent') {
          ctx.fillStyle = textItem.boxBgColor;
          ctx.fillRect(boxX, boxY, boxW, boxH);
        }

        // 사각형 테두리 그리기
        if (textItem.hasBox) {
          ctx.strokeStyle = textItem.boxBorderColor || textItem.color || '#000000';
          ctx.lineWidth = 2;
          ctx.strokeRect(boxX, boxY, boxW, boxH);
        }

        ctx.fillStyle = textItem.color;
        lines.forEach((line, idx) => {
          ctx.fillText(line, x, y + idx * lineHeight);
        });
        ctx.restore();
      });

      // PNG 이미지로 인코딩 후 PDF 페이지에 합성
      const overlayPngUrl = offCanvas.toDataURL('image/png');
      const pngImage = await pdfDocument.embedPng(overlayPngUrl);
      pdfPage.drawImage(pngImage, {
        x: 0,
        y: 0,
        width: pageWidth,
        height: pageHeight,
      });
    }

    return await pdfDocument.save();
  };

  // 주석 저장 및 메모 반영
  const handleSave = async () => {
    if (isSaving) return;
    setIsSaving(true);
    setErrorMessage(null);

    try {
      const modifiedBytes = await generateAnnotatedPdfBytes();
      const updatedDataUrl = await uint8ArrayToDataUrl(modifiedBytes);

      // 1. 기존 캐시 무효화
      invalidatePdfCache(attId);

      // 2. 상위 메모 상태에 저장 및 본문 뷰어 동기화
      if (onSave) {
        await onSave(updatedDataUrl);
      }

      // 3. 방금 저장된 수정본 PDF를 현재 에디터의 기본 문서로 승격
      setCurrentPdfData(updatedDataUrl);
      // 저장된 주석들은 이제 PDF 본문에 합성되었으므로 오버레이 목록 초기화
      setAnnotations({});

      // 4. 에디터 뷰어 재로드
      const newDoc = await loadPdfDocument(updatedDataUrl, attId);
      setPdfDoc(newDoc);
      setTotalPages(newDoc.numPages);

      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 3000);
    } catch (err: any) {
      console.error('PDF 주석 저장 실패:', err);
      setErrorMessage(`PDF 저장 실패: ${err.message || '문서를 저장할 수 없습니다.'}`);
    } finally {
      setIsSaving(false);
    }
  };

  // 수정된 PDF 바로 다운로드
  const handleDownload = async () => {
    try {
      const modifiedBytes = await generateAnnotatedPdfBytes();
      const blob = new Blob([modifiedBytes as any], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = title.endsWith('.pdf') ? title : `${title}.pdf`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err: any) {
      alert(`다운로드 실패: ${err.message || '오류가 발생했습니다.'}`);
    }
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex flex-col bg-slate-950/95 backdrop-blur-md text-slate-100 animate-in fade-in select-none">
      {/* 상단 툴바: 제목, 도구 선택, 저장 및 닫기 */}
      <div className="flex items-center justify-between px-3 py-2 bg-slate-900 border-b border-slate-800 gap-2 shrink-0 overflow-x-auto">
        <div className="flex items-center gap-2 shrink-0">
          <span className="text-emerald-400 font-bold text-xs sm:text-sm">✏️ PDF 주석/서명</span>
          <span className="text-xs sm:text-sm font-semibold truncate text-slate-200 max-w-[120px] sm:max-w-[200px]" title={title}>
            {title}
          </span>
          {saveSuccess && (
            <span className="flex items-center gap-1 text-[11px] text-emerald-400 bg-emerald-500/10 px-2 py-0.5 rounded-full border border-emerald-500/20 font-bold">
              <CheckCircle size={12} /> 저장됨
            </span>
          )}
        </div>

        {/* 중앙: 주석 도구 팔레트 */}
        <div className="flex items-center gap-1 bg-slate-950 p-1 rounded-xl border border-slate-800 shrink-0">
          {/* 손/이동 툴 */}
          <button
            type="button"
            onClick={() => setActiveTool('pan')}
            className={`p-1.5 rounded-lg text-xs font-semibold flex items-center gap-1 transition-all ${
              activeTool === 'pan' ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
            title="화면 이동 / 보기 (마우스로 잡고 드래그하여 화면 이동)"
          >
            <Move size={14} />
            <span className="hidden md:inline">화면 이동</span>
          </button>

          {/* 펜 드로잉 툴 */}
          <button
            type="button"
            onClick={() => setActiveTool('pen')}
            className={`p-1.5 rounded-lg text-xs font-semibold flex items-center gap-1 transition-all ${
              activeTool === 'pen' ? 'bg-indigo-600 text-white' : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
            title="펜 드로잉 (메모, 체크, 서명)"
          >
            <PenTool size={14} />
            <span className="hidden md:inline">펜</span>
          </button>

          {/* 형광펜 툴 */}
          <button
            type="button"
            onClick={() => setActiveTool('highlighter')}
            className={`p-1.5 rounded-lg text-xs font-semibold flex items-center gap-1 transition-all ${
              activeTool === 'highlighter' ? 'bg-indigo-600 text-white' : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
            title="형광펜 (반투명 강조)"
          >
            <Highlighter size={14} />
            <span className="hidden md:inline">형광펜</span>
          </button>

          {/* 텍스트 툴 */}
          <button
            type="button"
            onClick={() => setActiveTool('text')}
            className={`p-1.5 rounded-lg text-xs font-semibold flex items-center gap-1 transition-all ${
              activeTool === 'text' ? 'bg-indigo-600 text-white' : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
            title="텍스트 박스 추가 (원하는 위치 클릭하여 글자 및 사각형 박스 삽입)"
          >
            <Type size={14} />
            <span className="hidden md:inline">글자/박스</span>
          </button>

          {/* 서명 / 도장 툴 */}
          <button
            type="button"
            onClick={() => setActiveTool('stamp')}
            className={`p-1.5 rounded-lg text-xs font-semibold flex items-center gap-1 transition-all ${
              activeTool === 'stamp' ? 'bg-indigo-600 text-white' : 'text-slate-300 hover:text-white hover:bg-slate-800'
            }`}
            title="서명/도장 툴 (원하는 위치 클릭하여 서명/도장 날인)"
          >
            <Stamp size={14} className={activeTool === 'stamp' ? 'text-white' : 'text-rose-400'} />
            <span className="hidden md:inline">서명/도장</span>
          </button>

          <div className="h-4 w-px bg-slate-800 mx-1" />

          {/* 실행 취소 */}
          <button
            type="button"
            onClick={handleUndo}
            className="p-1.5 rounded-lg text-xs text-slate-400 hover:text-white hover:bg-slate-800 transition-all"
            title="마지막 작업 취소 (Undo)"
          >
            <RotateCcw size={14} />
          </button>

          {/* 전체 지우기 */}
          <button
            type="button"
            onClick={handleClearPage}
            className="p-1.5 rounded-lg text-xs text-slate-400 hover:text-rose-400 hover:bg-slate-800 transition-all"
            title="현재 페이지 주석 모두 지우기"
          >
            <Trash2 size={14} />
          </button>
        </div>

        {/* 우측: 저장, 다운로드, 닫기 */}
        <div className="flex items-center gap-1.5 shrink-0">
          <button
            type="button"
            onClick={handleSave}
            disabled={isSaving}
            className="flex items-center gap-1 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-bold transition-all shadow-md active:scale-95 disabled:opacity-50"
          >
            {isSaving ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />}
            <span>저장 및 반영</span>
          </button>

          <button
            type="button"
            onClick={handleDownload}
            className="flex items-center gap-1 px-2.5 py-1.5 bg-slate-800 hover:bg-slate-750 text-slate-200 rounded-lg text-xs font-semibold border border-slate-700 transition-all active:scale-95"
            title="수정된 PDF 다운로드"
          >
            <Download size={14} />
            <span className="hidden sm:inline">다운로드</span>
          </button>

          <button
            type="button"
            onClick={onClose}
            className="p-1.5 hover:bg-rose-500/20 text-slate-400 hover:text-rose-400 rounded-lg transition-colors"
            title="닫기"
          >
            <X size={18} />
          </button>
        </div>
      </div>

      {/* 보조 옵션 툴바 (색상 팔레트 및 페이지 이동 컨트롤) */}
      <div className="flex items-center justify-between px-3 py-1.5 bg-slate-900/60 border-b border-slate-800/80 gap-2 text-xs overflow-x-auto shrink-0">
        {/* 색상 및 도구별 상세 옵션 */}
        <div className="flex items-center gap-2 shrink-0">
          {activeTool === 'pen' && (
            <div className="flex items-center gap-1.5">
              <span className="text-slate-400 text-[11px]">펜 색상:</span>
              {PEN_COLORS.map((c) => (
                <button
                  key={c}
                  type="button"
                  onClick={() => setPenColor(c)}
                  style={{ backgroundColor: c }}
                  className={`w-4 h-4 rounded-full border ${penColor === c ? 'border-white scale-125 shadow-sm' : 'border-transparent opacity-70'}`}
                />
              ))}
              <span className="text-slate-400 text-[11px] ml-2">두께:</span>
              {[2, 4, 7].map((w) => (
                <button
                  key={w}
                  type="button"
                  onClick={() => setPenWidth(w)}
                  className={`px-1.5 py-0.5 rounded text-[11px] font-bold ${penWidth === w ? 'bg-indigo-600 text-white' : 'bg-slate-800 text-slate-400'}`}
                >
                  {w}px
                </button>
              ))}
            </div>
          )}

          {activeTool === 'highlighter' && (
            <div className="flex items-center gap-1.5">
              <span className="text-slate-400 text-[11px]">형광 색상:</span>
              {HIGHLIGHTER_COLORS.map((c) => (
                <button
                  key={c}
                  type="button"
                  onClick={() => setHighlighterColor(c)}
                  style={{ backgroundColor: c }}
                  className={`w-4 h-4 rounded-full border ${highlighterColor === c ? 'border-white scale-125 shadow-sm' : 'border-transparent opacity-70'}`}
                />
              ))}
              <span className="text-slate-400 text-[11px] ml-2">굵기:</span>
              {[15, 24, 32].map((w) => (
                <button
                  key={w}
                  type="button"
                  onClick={() => setHighlighterWidth(w)}
                  className={`px-1.5 py-0.5 rounded text-[11px] font-bold ${highlighterWidth === w ? 'bg-indigo-600 text-white' : 'bg-slate-800 text-slate-400'}`}
                >
                  {w}px
                </button>
              ))}
            </div>
          )}

          {activeTool === 'text' && (
            <div className="flex items-center gap-2 flex-wrap sm:flex-nowrap">
              <span className="text-slate-400 text-[11px]">크기:</span>
              {[12, 16, 20, 28].map((s) => (
                <button
                  key={s}
                  type="button"
                  onClick={() => setTextFontSize(s)}
                  className={`px-1.5 py-0.5 rounded text-[11px] font-bold ${textFontSize === s ? 'bg-indigo-600 text-white' : 'bg-slate-800 text-slate-400'}`}
                >
                  {s}pt
                </button>
              ))}

              <span className="text-slate-400 text-[11px] ml-1">글자색:</span>
              <div className="flex items-center gap-1">
                {TEXT_COLORS.map((c) => (
                  <button
                    key={c}
                    type="button"
                    onClick={() => setTextColor(c)}
                    style={{ backgroundColor: c }}
                    className={`w-3.5 h-3.5 rounded-full border ${textColor === c ? 'border-white scale-125' : 'border-slate-700 opacity-70'}`}
                  />
                ))}
              </div>

              <div className="h-3 w-px bg-slate-800 mx-0.5 hidden sm:block" />

              {/* 사각형 테두리 토글 */}
              <label className="flex items-center gap-1 text-[11px] font-semibold text-slate-200 cursor-pointer">
                <input
                  type="checkbox"
                  checked={textHasBox}
                  onChange={(e) => setTextHasBox(e.target.checked)}
                  className="rounded text-indigo-600 cursor-pointer"
                />
                <span>🔲 테두리</span>
              </label>

              {textHasBox && (
                <div className="flex items-center gap-1">
                  <span className="text-slate-400 text-[11px]">선색:</span>
                  {BORDER_COLORS.map((c) => (
                    <button
                      key={c}
                      type="button"
                      onClick={() => setTextBoxBorderColor(c)}
                      style={{ backgroundColor: c }}
                      className={`w-3.5 h-3.5 rounded-full border ${textBoxBorderColor === c ? 'border-white scale-125' : 'border-slate-700 opacity-70'}`}
                    />
                  ))}
                </div>
              )}

              {/* 배경 채우기 */}
              <span className="text-slate-400 text-[11px] ml-1">배경:</span>
              <div className="flex items-center gap-0.5">
                {[
                  { id: 'transparent', label: '투명' },
                  { id: '#ffffff', label: '흰색' },
                  { id: '#fef9c3', label: '연노랑' },
                ].map((bg) => (
                  <button
                    key={bg.id}
                    type="button"
                    onClick={() => setTextBoxBgColor(bg.id)}
                    className={`px-1.5 py-0.5 rounded text-[10px] font-semibold ${textBoxBgColor === bg.id ? 'bg-indigo-600 text-white' : 'bg-slate-800 text-slate-400'}`}
                  >
                    {bg.label}
                  </button>
                ))}
              </div>

              {/* 화이트 덮어쓰기 프리셋 (기존 글자 수정용) */}
              <button
                type="button"
                onClick={() => {
                  setTextHasBox(false);
                  setTextBoxBgColor('#ffffff');
                  setTextColor('#000000');
                  setPendingTextPos({ x: 0.5, y: 0.25 });
                  setShowTextModal(true);
                }}
                className="ml-1 px-2 py-0.5 bg-amber-600/30 hover:bg-amber-600 text-amber-200 hover:text-white rounded text-[11px] font-bold border border-amber-500/30 flex items-center gap-1 transition-all"
                title="기존 글자 위에 흰색 배경을 덮어 가린 뒤 새 글자를 입력합니다"
              >
                <span>🧼 화이트 덮어쓰기(수정)</span>
              </button>

              {/* 글자 직접 입력 팝업 열기 */}
              <button
                type="button"
                onClick={() => {
                  setPendingTextPos({ x: 0.5, y: 0.25 });
                  setShowTextModal(true);
                }}
                className="ml-1 px-2 py-0.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded text-[11px] font-bold transition-all shadow"
              >
                + 글자 작성
              </button>
            </div>
          )}

          {activeTool === 'stamp' && (
            <div className="flex items-center gap-2">
              <span className="text-emerald-400 text-[11px] font-semibold">📍 서명이나 도장을 찍을 PDF 위치를 클릭하세요 (삽입 후 드래그로 자유롭게 이동 가능)</span>
              <button
                type="button"
                onClick={() => {
                  setPendingStampPos({ x: 0.5, y: 0.5 });
                  setShowSignPad(true);
                }}
                className="px-2 py-0.5 bg-rose-600/30 hover:bg-rose-600 text-rose-200 hover:text-white rounded text-[11px] font-bold border border-rose-500/30"
              >
                + 서명/도장 만들기
              </button>
            </div>
          )}

          {activeTool === 'pan' && (
            <span className="text-indigo-300 text-[11px] font-medium flex items-center gap-1">
              <span>🖐️ 마우스로 문서를 클릭하고 드래그하여 화면을 자유롭게 이동하세요.</span>
            </span>
          )}
        </div>

        {/* 페이지 네비게이션 & 줌 */}
        <div className="flex items-center gap-2 shrink-0">
          <div className="flex items-center gap-1 bg-slate-950 px-2 py-0.5 rounded-lg border border-slate-800">
            <button
              type="button"
              disabled={currentPage <= 1}
              onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
              className="p-1 hover:text-white disabled:opacity-30"
              title="이전 페이지"
            >
              <ChevronLeft size={14} />
            </button>
            <span className="font-bold text-slate-200 px-1 text-xs">
              {currentPage} / {totalPages}
            </span>
            <button
              type="button"
              disabled={currentPage >= totalPages}
              onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
              className="p-1 hover:text-white disabled:opacity-30"
              title="다음 페이지"
            >
              <ChevronRight size={14} />
            </button>
          </div>

          <div className="flex items-center gap-1 bg-slate-950 px-1.5 py-0.5 rounded-lg border border-slate-800">
            <button
              type="button"
              onClick={() => setZoom((z) => Math.max(0.4, Math.round((z - 0.2) * 10) / 10))}
              className="p-1 hover:text-white"
              title="축소"
            >
              <ZoomOut size={13} />
            </button>
            <span className="text-[11px] font-bold w-10 text-center">{Math.round(zoom * 100)}%</span>
            <button
              type="button"
              onClick={() => setZoom((z) => Math.min(3.0, Math.round((z + 0.2) * 10) / 10))}
              className="p-1 hover:text-white"
              title="확대"
            >
              <ZoomIn size={13} />
            </button>
          </div>
        </div>
      </div>

      {/* 메인 뷰어 및 주석 캔버스 영역 (화면 팬 드래그 & 아이템 이동 지원) */}
      <div 
        ref={containerRef}
        onMouseDown={handleContainerMouseDown}
        className={`flex-1 overflow-auto bg-slate-950 p-2 sm:p-4 relative select-none ${
          isPanDragging
            ? 'cursor-grabbing'
            : activeTool === 'pan'
            ? 'cursor-grab touch-pan-x touch-pan-y'
            : ''
        }`}
      >
        <div className="w-max min-w-full min-h-full flex flex-col items-center justify-start py-4 sm:py-8 px-2">
          {isLoading && (
            <div className="flex flex-col items-center justify-center h-64 gap-3 text-slate-400 text-sm">
              <Loader2 className="w-8 h-8 animate-spin text-indigo-500" />
              <span>PDF 문서를 로드하고 있습니다...</span>
            </div>
          )}

          {errorMessage && (
            <div className="p-4 bg-rose-950/40 border border-rose-800 rounded-xl text-center text-rose-300 text-sm">
              <p className="font-bold">⚠️ {errorMessage}</p>
            </div>
          )}

          <div className={`relative shrink-0 shadow-2xl rounded-xl bg-white ${isLoading || errorMessage ? 'hidden' : 'block'}`}>
            {/* 1. 베이스 PDF 렌더링 캔버스 */}
            <canvas ref={baseCanvasRef} className="rounded-xl block" style={{ maxWidth: 'none' }} />

            {/* 2. 주석 상호작용 오버레이 캔버스 (도장/글자 드래그 이동 및 드로잉) */}
            <canvas
              ref={overlayCanvasRef}
              onPointerDown={handlePointerDown}
              onPointerMove={handlePointerMove}
              onPointerUp={handlePointerUp}
              onPointerCancel={handlePointerUp}
              className={`absolute inset-0 rounded-xl ${
                activeTool === 'pan' ? 'pointer-events-none' : 'touch-none cursor-crosshair'
              }`}
            />
          </div>
        </div>
      </div>

      {/* 텍스트 및 사각형 박스 삽입 모달 */}
      {showTextModal && (
        <div 
          className="fixed inset-0 z-60 flex items-center justify-center bg-black/70 backdrop-blur-xs p-4 animate-in fade-in"
          onClick={() => setShowTextModal(false)}
        >
          <div 
            className="bg-slate-900 border border-slate-800 rounded-2xl p-5 w-full max-w-md shadow-2xl space-y-4"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between pb-2 border-b border-slate-800">
              <span className="font-bold text-sm text-slate-200">📝 PDF 텍스트 및 사각형 박스 삽입</span>
              <button
                type="button"
                onClick={() => setShowTextModal(false)}
                className="p-1 hover:bg-slate-800 rounded text-slate-400"
              >
                <X size={16} />
              </button>
            </div>

            {/* 빠른 프리셋 모드 선택 */}
            <div className="space-y-1">
              <span className="text-[11px] text-slate-400 font-semibold block">빠른 설정 모드:</span>
              <div className="grid grid-cols-2 gap-2 text-xs">
                <button
                  type="button"
                  onClick={() => {
                    setTextHasBox(false);
                    setTextBoxBgColor('transparent');
                  }}
                  className={`p-2 rounded-xl border text-left flex flex-col gap-0.5 transition-all ${
                    !textHasBox && textBoxBgColor === 'transparent'
                      ? 'bg-indigo-600/20 border-indigo-500 text-indigo-200'
                      : 'bg-slate-950 border-slate-800 text-slate-400 hover:border-slate-700'
                  }`}
                >
                  <span className="font-bold text-slate-200">📝 일반 글자</span>
                  <span className="text-[10px] text-slate-400">배경 투명 일반 텍스트</span>
                </button>

                <button
                  type="button"
                  onClick={() => {
                    setTextHasBox(true);
                    setTextBoxBorderColor('#000000');
                  }}
                  className={`p-2 rounded-xl border text-left flex flex-col gap-0.5 transition-all ${
                    textHasBox
                      ? 'bg-indigo-600/20 border-indigo-500 text-indigo-200'
                      : 'bg-slate-950 border-slate-800 text-slate-400 hover:border-slate-700'
                  }`}
                >
                  <span className="font-bold text-slate-200">🔲 테두리 상자</span>
                  <span className="text-[10px] text-slate-400">사각형 박스 안 글자</span>
                </button>

                <button
                  type="button"
                  onClick={() => {
                    setTextHasBox(false);
                    setTextBoxBgColor('#ffffff');
                    setTextColor('#000000');
                  }}
                  className={`p-2 rounded-xl border text-left flex flex-col gap-0.5 transition-all col-span-2 ${
                    !textHasBox && textBoxBgColor === '#ffffff'
                      ? 'bg-amber-600/20 border-amber-500 text-amber-200'
                      : 'bg-slate-950 border-slate-800 text-slate-400 hover:border-slate-700'
                  }`}
                >
                  <div className="flex items-center justify-between">
                    <span className="font-bold text-amber-300">🧼 화이트 덮어쓰기 (기존 글자 수정)</span>
                    <span className="text-[10px] bg-amber-500/20 text-amber-300 px-1.5 py-0.5 rounded font-bold">오타 수정 추천</span>
                  </div>
                  <span className="text-[10px] text-slate-300">기존 글자 위에 흰색 배경을 덮어 가린 뒤 새 글자로 대체합니다.</span>
                </button>
              </div>
            </div>

            {/* 글자 내용 입력 */}
            <div className="space-y-1.5">
              <div className="flex items-center justify-between">
                <label className="text-xs font-semibold text-slate-300">삽입할 글자 내용:</label>
                {textBoxBgColor !== 'transparent' && !inputModalText.trim() && (
                  <span className="text-[10px] text-emerald-400 font-medium">※ 빈칸으로 삽입 시 수정테이프(화이트)로 동작</span>
                )}
              </div>
              <textarea
                value={inputModalText}
                onChange={(e) => setInputModalText(e.target.value)}
                placeholder={
                  textBoxBgColor === '#ffffff'
                    ? "수정할 새 글자를 입력하세요 (빈칸 시 흰색 가림막/수정테이프)"
                    : "글자를 입력하세요 (엔터로 줄바꿈 가능)"
                }
                rows={2}
                autoFocus
                className="w-full bg-slate-950 border border-slate-700 rounded-xl p-3 text-sm text-slate-100 focus:outline-none focus:border-indigo-500 resize-none font-medium"
              />
            </div>

            {/* 글자 크기 & 색상 */}
            <div className="grid grid-cols-2 gap-3 bg-slate-950/60 p-3 rounded-xl border border-slate-800/80">
              <div className="space-y-1">
                <span className="text-[11px] text-slate-400 font-semibold block">글자 크기</span>
                <div className="flex items-center gap-1">
                  {[12, 16, 20, 28].map((s) => (
                    <button
                      key={s}
                      type="button"
                      onClick={() => setTextFontSize(s)}
                      className={`px-2 py-0.5 rounded text-xs font-bold ${
                        textFontSize === s ? 'bg-indigo-600 text-white' : 'bg-slate-800 text-slate-400 hover:text-white'
                      }`}
                    >
                      {s}pt
                    </button>
                  ))}
                </div>
              </div>

              <div className="space-y-1">
                <span className="text-[11px] text-slate-400 font-semibold block">글자 색상</span>
                <div className="flex items-center gap-1.5 pt-0.5">
                  {TEXT_COLORS.map((c) => (
                    <button
                      key={c}
                      type="button"
                      onClick={() => setTextColor(c)}
                      style={{ backgroundColor: c }}
                      className={`w-5 h-5 rounded-full border ${
                        textColor === c ? 'border-white scale-125 shadow-sm' : 'border-slate-700 opacity-70'
                      }`}
                    />
                  ))}
                </div>
              </div>
            </div>

            {/* 사각형 테두리 박스 설정 */}
            <div className="bg-slate-950/60 p-3 rounded-xl border border-slate-800/80 space-y-2.5">
              <div className="flex items-center justify-between">
                <label className="flex items-center gap-2 text-xs font-bold text-slate-200 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={textHasBox}
                    onChange={(e) => setTextHasBox(e.target.checked)}
                    className="w-4 h-4 rounded text-indigo-600 focus:ring-0 cursor-pointer"
                  />
                  <span>🔲 사각형 테두리 그리기</span>
                </label>
                {textHasBox && (
                  <span className="text-[11px] text-indigo-400 font-semibold">테두리 활성</span>
                )}
              </div>

              {textHasBox && (
                <div className="flex items-center justify-between pt-1 border-t border-slate-800/60">
                  <span className="text-[11px] text-slate-400">테두리 선 색상:</span>
                  <div className="flex items-center gap-1.5">
                    {BORDER_COLORS.map((c) => (
                      <button
                        key={c}
                        type="button"
                        onClick={() => setTextBoxBorderColor(c)}
                        style={{ backgroundColor: c }}
                        className={`w-5 h-5 rounded-full border ${
                          textBoxBorderColor === c ? 'border-white scale-125 shadow-sm' : 'border-slate-700 opacity-70'
                        }`}
                      />
                    ))}
                  </div>
                </div>
              )}

              {/* 배경 채우기 */}
              <div className="flex items-center justify-between pt-1 border-t border-slate-800/60 text-xs">
                <span className="text-[11px] text-slate-400">배경 채우기:</span>
                <div className="flex items-center gap-1">
                  {[
                    { id: 'transparent', label: '투명' },
                    { id: '#ffffff', label: '흰색' },
                    { id: '#fef9c3', label: '연노랑' },
                  ].map((bg) => (
                    <button
                      key={bg.id}
                      type="button"
                      onClick={() => setTextBoxBgColor(bg.id)}
                      className={`px-2 py-0.5 rounded text-[11px] font-semibold ${
                        textBoxBgColor === bg.id ? 'bg-indigo-600 text-white' : 'bg-slate-800 text-slate-400'
                      }`}
                    >
                      {bg.label}
                    </button>
                  ))}
                </div>
              </div>
            </div>

            {/* 실시간 미리보기 */}
            <div className="space-y-1">
              <span className="text-[11px] text-slate-400 font-semibold block">실제 삽입 모양 미리보기:</span>
              <div className="p-3 bg-slate-950 rounded-xl border border-slate-800 flex items-center justify-center min-h-[55px]">
                <div
                  style={{
                    border: textHasBox ? `2px solid ${textBoxBorderColor}` : '1px dashed #64748b',
                    backgroundColor: textBoxBgColor === 'transparent' ? 'transparent' : textBoxBgColor,
                    color: textColor,
                    fontSize: `${textFontSize}px`,
                    fontWeight: 'bold',
                    padding: '4px 10px',
                    borderRadius: '2px',
                    display: 'inline-block',
                    maxWidth: '100%',
                    wordBreak: 'break-word',
                    whiteSpace: 'pre-wrap',
                  }}
                >
                  {inputModalText || '텍스트 미리보기'}
                </div>
              </div>
            </div>

            {/* 버튼 */}
            <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-800">
              <button
                type="button"
                onClick={() => setShowTextModal(false)}
                className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg text-xs font-semibold"
              >
                취소
              </button>
              <button
                type="button"
                onClick={handleConfirmInsertText}
                disabled={!inputModalText.trim() && textBoxBgColor === 'transparent'}
                className="px-4 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-bold transition-all shadow disabled:opacity-40"
              >
                {textBoxBgColor !== 'transparent' && !inputModalText.trim() ? '수정테이프(화이트) 삽입' : 'PDF에 삽입하기'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 서명 / 도장 입력 팝업 모달 */}
      {showSignPad && (
        <div 
          className="fixed inset-0 z-60 flex items-center justify-center bg-black/70 backdrop-blur-xs p-4 animate-in fade-in"
          onClick={() => {
            setShowSignPad(false);
            setPendingStampPos(null);
          }}
        >
          <div 
            className="bg-slate-900 border border-slate-800 rounded-2xl p-5 w-full max-w-sm shadow-2xl space-y-4"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between pb-2 border-b border-slate-800">
              <span className="font-bold text-sm text-slate-200">✍️ 서명 또는 도장 만들기</span>
              <button
                type="button"
                onClick={() => {
                  setShowSignPad(false);
                  setPendingStampPos(null);
                }}
                className="p-1 hover:bg-slate-800 rounded text-slate-400"
              >
                <X size={16} />
              </button>
            </div>

            {/* 손글씨 서명 패드 */}
            <div className="space-y-1.5">
              <span className="text-xs text-slate-400">아래 네모 칸에 손가락이나 펜으로 직접 서명하세요:</span>
              <canvas
                ref={signCanvasRef}
                width={320}
                height={140}
                onPointerDown={handleSignPointerDown}
                onPointerMove={handleSignPointerMove}
                onPointerUp={handleSignPointerUp}
                className="w-full h-[140px] bg-white rounded-xl border border-slate-700 cursor-crosshair touch-none shadow-inner"
              />
              <div className="flex items-center justify-between pt-1">
                <button
                  type="button"
                  onClick={handleClearSignCanvas}
                  className="text-xs text-slate-400 hover:text-rose-400 flex items-center gap-1"
                >
                  <Trash2 size={12} /> 다시 쓰기
                </button>
                <button
                  type="button"
                  onClick={handleInsertSignature}
                  className="px-3 py-1 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-bold transition-all shadow"
                >
                  선택한 위치에 서명 삽입
                </button>
              </div>
            </div>

            <div className="border-t border-slate-800 pt-3 space-y-2">
              <span className="text-xs text-slate-400 block">또는 인장/도장 원클릭 날인:</span>
              <div className="grid grid-cols-2 gap-2">
                <button
                  type="button"
                  onClick={() => handleInsertPresetStamp('서명')}
                  className="py-2 bg-rose-600/20 hover:bg-rose-600/30 text-rose-300 border border-rose-500/40 rounded-xl text-xs font-bold transition-all"
                >
                  🔴 [서명] 인장 날인
                </button>
                <button
                  type="button"
                  onClick={() => handleInsertPresetStamp('확인')}
                  className="py-2 bg-blue-600/20 hover:bg-blue-600/30 text-blue-300 border border-blue-500/40 rounded-xl text-xs font-bold transition-all"
                >
                  🔵 [확인] 인장 날인
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
