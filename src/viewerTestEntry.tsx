import { useState, useEffect } from 'react';
import { createRoot } from 'react-dom/client';
import './index.css';
import { PptxViewerModal } from './components/PptxViewerModal';
import { PptxEditorModal } from './components/PptxEditorModal';
import { ExcelEditorModal } from './components/ExcelEditorModal';

declare global {
  interface Window {
    __loadPptx?: (dataUrl: string, title: string, mode?: 'viewer' | 'editor') => Promise<void>;
    __setSlideIndex?: (index: number) => void;
    __SAVED_PPTX__?: string;

    __loadExcel?: (dataUrl: string, title: string) => Promise<void>;
    __SAVED_EXCEL__?: string;
  }
}

export function TestApp() {
  const [data, setData] = useState<string | null>(null);
  const [title, setTitle] = useState<string>('Test Presentation');
  const [mode, setMode] = useState<'viewer' | 'editor' | 'excel'>('viewer');
  const [isOpen, setIsOpen] = useState<boolean>(false);

  useEffect(() => {
    window.__loadPptx = async (dataUrl: string, presTitle: string, presMode: 'viewer' | 'editor' = 'viewer') => {
      setData(dataUrl);
      setTitle(presTitle);
      setMode(presMode);
      setIsOpen(true);
    };

    window.__loadExcel = async (dataUrl: string, excelTitle: string) => {
      setData(dataUrl);
      setTitle(excelTitle);
      setMode('excel');
      setIsOpen(true);
    };
  }, []);

  if (!isOpen || !data) {
    return (
      <div className="p-8 text-center text-slate-400">
        <h1 className="text-xl font-bold mb-2">Office Test Harness Ready</h1>
        <p className="text-sm">Call window.__loadPptx(...) or window.__loadExcel(...) to start.</p>
      </div>
    );
  }

  return (
    <div className="w-screen h-screen">
      {mode === 'viewer' ? (
        <PptxViewerModal
          isOpen={isOpen}
          onClose={() => setIsOpen(false)}
          data={data}
          title={title}
        />
      ) : mode === 'editor' ? (
        <PptxEditorModal
          isOpen={isOpen}
          onClose={() => setIsOpen(false)}
          data={data}
          title={title}
          onSave={(updated) => {
            window.__SAVED_PPTX__ = updated;
          }}
        />
      ) : (
        <ExcelEditorModal
          isOpen={isOpen}
          onClose={() => setIsOpen(false)}
          data={data}
          title={title}
          initialEngine="exceljs"
          onSave={(updated) => {
            window.__SAVED_EXCEL__ = updated;
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
