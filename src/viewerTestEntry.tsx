import { useState, useEffect } from 'react';
import { createRoot } from 'react-dom/client';
import './index.css';
import { PptxViewerModal } from './components/PptxViewerModal';
import { PptxEditorModal } from './components/PptxEditorModal';
import { ExcelEditorModal } from './components/ExcelEditorModal';
import { ExcelViewerModal } from './components/ExcelViewerModal';
import { DocxEditorModal } from './components/DocxEditorModal';
import { HwpEditorModal } from './components/HwpEditorModal';
import { HwpViewerModal } from './components/HwpViewerModal';
import { NoteEditor } from './components/NoteEditor';
import { FloatingNoteWindow } from './components/FloatingNoteWindow';
import { renderAttachmentsInHtml } from './utils/attachment';
import type { Note, AttachmentItem } from './types';

declare global {
  interface Window {
    __loadPptx?: (dataUrl: string, title: string, mode?: 'viewer' | 'editor') => Promise<void>;
    __setSlideIndex?: (index: number) => void;
    __SAVED_PPTX__?: string;

    __loadExcel?: (dataUrl: string, title: string) => Promise<void>;
    __loadExcelViewer?: (dataUrl: string, title: string, attId?: string) => Promise<void>;
    __loadInlineCard?: (dataUrl: string, title: string, attId?: string) => Promise<void>;
    __loadPopout?: (dataUrl: string, title: string, attId?: string) => Promise<void>;
    __SAVED_EXCEL__?: string;

    __loadDocx?: (dataUrl: string, title: string) => Promise<void>;
    __SAVED_DOCX__?: string;

    __loadHwp?: (dataUrl: string, title: string, mode?: 'viewer' | 'editor') => Promise<void>;
    __SAVED_HWP__?: string;
  }
}

export function TestApp() {
  const [data, setData] = useState<string | null>(null);
  const [title, setTitle] = useState<string>('Test Presentation');
  const [attId, setAttId] = useState<string>('test_att_1');
  const [mode, setMode] = useState<
    'viewer' | 'editor' | 'excel' | 'excel-viewer' | 'inline-card' | 'popout' | 'docx' | 'hwp-viewer' | 'hwp-editor'
  >('viewer');
  const [isOpen, setIsOpen] = useState<boolean>(false);
  const [renderKey, setRenderKey] = useState<number>(0);

  // 인라인 카드 및 팝아웃용 가상 메모 객체
  const [testNote, setTestNote] = useState<Note | null>(null);

  useEffect(() => {
    (window as any).__renderAttachmentsInHtml = renderAttachmentsInHtml;

    window.__loadPptx = async (dataUrl: string, presTitle: string, presMode: 'viewer' | 'editor' = 'viewer', newAttId?: string) => {
      setIsOpen(false);
      await new Promise(r => setTimeout(r, 30));
      setData(dataUrl);
      setTitle(presTitle);
      if (newAttId) setAttId(newAttId);
      setMode(presMode);
      setRenderKey(k => k + 1);
      setIsOpen(true);
    };

    window.__loadExcel = async (dataUrl: string, excelTitle: string, newAttId?: string) => {
      setIsOpen(false);
      await new Promise(r => setTimeout(r, 30));
      setData(dataUrl);
      setTitle(excelTitle);
      if (newAttId) setAttId(newAttId);
      setMode('excel');
      setRenderKey(k => k + 1);
      setIsOpen(true);
    };

    window.__loadExcelViewer = async (dataUrl: string, excelTitle: string, newAttId = 'test_att_1') => {
      setIsOpen(false);
      await new Promise(r => setTimeout(r, 30));
      setData(dataUrl);
      setTitle(excelTitle);
      setAttId(newAttId);
      setMode('excel-viewer');
      setRenderKey(k => k + 1);
      setIsOpen(true);
    };

    window.__loadInlineCard = async (dataUrl: string, excelTitle: string, newAttId = 'test_att_1') => {
      setIsOpen(false);
      await new Promise(r => setTimeout(r, 30));
      const attItem: AttachmentItem = {
        id: newAttId,
        name: excelTitle,
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        data: dataUrl,
      };
      const noteObj: Note = {
        id: 'test-note-1',
        projectId: 'default',
        title: '엑셀 인라인 카드 검증 메모',
        content: `# 엑셀 문서 검증\n\n[📊 ${excelTitle}](attachment:${newAttId})\n\n메모 하단 내용`,
        attachments: { [newAttId]: attItem },
        isLocalOnly: true,
        syncStatus: 'local_only',
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      };
      setData(dataUrl);
      setTitle(excelTitle);
      setAttId(newAttId);
      setTestNote(noteObj);
      setMode('inline-card');
      setRenderKey(k => k + 1);
      setIsOpen(true);
    };

    window.__loadPopout = async (dataUrl: string, excelTitle: string, newAttId = 'test_att_1') => {
      setIsOpen(false);
      await new Promise(r => setTimeout(r, 30));
      const attItem: AttachmentItem = {
        id: newAttId,
        name: excelTitle,
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        data: dataUrl,
      };
      const noteObj: Note = {
        id: 'test-note-popout',
        projectId: 'default',
        title: '엑셀 팝아웃 창 검증 메모',
        content: `# 엑셀 팝아웃 검증\n\n[📊 ${excelTitle}](attachment:${newAttId})\n\n내용 끝`,
        attachments: { [newAttId]: attItem },
        isLocalOnly: true,
        syncStatus: 'local_only',
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      };
      setData(dataUrl);
      setTitle(excelTitle);
      setAttId(newAttId);
      setTestNote(noteObj);
      setMode('popout');
      setRenderKey(k => k + 1);
      setIsOpen(true);
    };

    window.__loadDocx = async (dataUrl: string, docxTitle: string, newAttId?: string) => {
      setIsOpen(false);
      await new Promise(r => setTimeout(r, 30));
      setData(dataUrl);
      setTitle(docxTitle);
      if (newAttId) setAttId(newAttId);
      setMode('docx');
      setRenderKey(k => k + 1);
      setIsOpen(true);
    };

    window.__loadHwp = async (dataUrl: string, hwpTitle: string, hwpMode: 'viewer' | 'editor' = 'viewer', newAttId?: string) => {
      setIsOpen(false);
      await new Promise(r => setTimeout(r, 30));
      setData(dataUrl);
      setTitle(hwpTitle);
      if (newAttId) setAttId(newAttId);
      setMode(hwpMode === 'editor' ? 'hwp-editor' : 'hwp-viewer');
      setRenderKey(k => k + 1);
      setIsOpen(true);
    };
  }, []);

  if (!isOpen || !data) {
    return (
      <div className="p-8 text-center text-slate-400">
        <h1 className="text-xl font-bold mb-2">Office Test Harness Ready</h1>
        <p className="text-sm">Call window.__loadPptx(...), window.__loadExcel(...) or window.__loadHwp(...) to start.</p>
      </div>
    );
  }

  return (
    <div className="w-screen h-screen">
      {mode === 'viewer' ? (
        <PptxViewerModal
          key={renderKey}
          isOpen={isOpen}
          onClose={() => setIsOpen(false)}
          data={data}
          title={title}
        />
      ) : mode === 'editor' ? (
        <PptxEditorModal
          key={renderKey}
          isOpen={isOpen}
          onClose={() => setIsOpen(false)}
          data={data}
          title={title}
          attId={attId}
          onSave={(updated) => {
            (window as any).__SAVED_PPTX_COUNT__ = ((window as any).__SAVED_PPTX_COUNT__ || 0) + 1;
            window.__SAVED_PPTX__ = updated;
          }}
        />
      ) : mode === 'docx' ? (
        <DocxEditorModal
          key={renderKey}
          isOpen={isOpen}
          onClose={() => setIsOpen(false)}
          data={data}
          title={title}
          attId={attId}
          initialEngine="superdoc"
          onSave={(updated) => {
            (window as any).__SAVED_DOCX_COUNT__ = ((window as any).__SAVED_DOCX_COUNT__ || 0) + 1;
            window.__SAVED_DOCX__ = updated;
          }}
        />
      ) : mode === 'excel' ? (
        <ExcelEditorModal
          key={renderKey}
          isOpen={isOpen}
          onClose={() => setIsOpen(false)}
          data={data}
          title={title}
          attId={attId}
          initialEngine="exceljs"
          onSave={(updated) => {
            (window as any).__SAVED_EXCEL_COUNT__ = ((window as any).__SAVED_EXCEL_COUNT__ || 0) + 1;
            window.__SAVED_EXCEL__ = updated;
          }}
        />
      ) : mode === 'excel-viewer' ? (
        <ExcelViewerModal
          key={renderKey}
          isOpen={isOpen}
          onClose={() => setIsOpen(false)}
          data={data}
          title={title}
          attId={attId}
        />
      ) : mode === 'inline-card' && testNote ? (
        <div className="w-full h-full bg-slate-950 text-slate-100 p-4">
          <NoteEditor
            key={renderKey}
            note={testNote}
            projects={[]}
            onUpdateNote={(up) => {
              setTestNote((prev) => prev ? { ...prev, ...up } : null);
            }}
            onToggleNoteList={() => {}}
            onOpenFloatingNote={() => {}}
            onBackToList={() => {}}
          />
        </div>
      ) : mode === 'popout' && testNote ? (
        <div className="w-full h-full bg-slate-950 text-slate-100 p-4">
          <FloatingNoteWindow
            key={renderKey}
            note={testNote}
            onClose={() => setIsOpen(false)}
            onUpdateNote={(up) => {
              setTestNote((prev) => prev ? { ...prev, ...up } : null);
            }}
            initialPopout={false}
          />
        </div>
      ) : mode === 'hwp-viewer' ? (
        <HwpViewerModal
          key={renderKey}
          isOpen={isOpen}
          onClose={() => setIsOpen(false)}
          data={data}
          title={title}
        />
      ) : (
        <HwpEditorModal
          key={renderKey}
          isOpen={isOpen}
          onClose={() => setIsOpen(false)}
          data={data}
          title={title}
          attId={attId}
          onSave={(updated) => {
            (window as any).__SAVED_HWP_COUNT__ = ((window as any).__SAVED_HWP_COUNT__ || 0) + 1;
            window.__SAVED_HWP__ = updated;
          }}
        />
      )}
    </div>
  );
}

const rootEl = document.getElementById('root');
if (rootEl) {
  createRoot(rootEl).render(<TestApp />);
}
