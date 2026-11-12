import { useState, useEffect } from 'react';
import { createRoot } from 'react-dom/client';
import './index.css';
import { PptxViewerModal } from './components/PptxViewerModal';
import { PptxEditorModal } from './components/PptxEditorModal';
import { ExcelEditorModal } from './components/ExcelEditorModal';
import { HwpEditorModal } from './components/HwpEditorModal';
import { HwpViewerModal } from './components/HwpViewerModal';

declare global {
  interface Window {
    __loadPptx?: (dataUrl: string, title: string, mode?: 'viewer' | 'editor') => Promise<void>;
    __setSlideIndex?: (index: number) => void;
    __SAVED_PPTX__?: string;

    __loadExcel?: (dataUrl: string, title: string) => Promise<void>;
    __SAVED_EXCEL__?: string;

    __loadHwp?: (dataUrl: string, title: string, mode?: 'viewer' | 'editor') => Promise<void>;
    __SAVED_HWP__?: string;
  }
}

export function TestApp() {
  const [data, setData] = useState<string | null>(null);
  const [title, setTitle] = useState<string>('Test Presentation');
  const [mode, setMode] = useState<'viewer' | 'editor' | 'excel' | 'hwp-viewer' | 'hwp-editor'>('viewer');
  const [isOpen, setIsOpen] = useState<boolean>(false);
  const [renderKey, setRenderKey] = useState<number>(0);

  useEffect(() => {
    window.__loadPptx = async (dataUrl: string, presTitle: string, presMode: 'viewer' | 'editor' = 'viewer') => {
      setIsOpen(false);
      await new Promise(r => setTimeout(r, 30));
      setData(dataUrl);
      setTitle(presTitle);
      setMode(presMode);
      setRenderKey(k => k + 1);
      setIsOpen(true);
    };

    window.__loadExcel = async (dataUrl: string, excelTitle: string) => {
      setIsOpen(false);
      await new Promise(r => setTimeout(r, 30));
      setData(dataUrl);
      setTitle(excelTitle);
      setMode('excel');
      setRenderKey(k => k + 1);
      setIsOpen(true);
    };

    window.__loadHwp = async (dataUrl: string, hwpTitle: string, hwpMode: 'viewer' | 'editor' = 'viewer') => {
      setIsOpen(false);
      await new Promise(r => setTimeout(r, 30));
      setData(dataUrl);
      setTitle(hwpTitle);
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
          onSave={(updated) => {
            window.__SAVED_PPTX__ = updated;
          }}
        />
      ) : mode === 'excel' ? (
        <ExcelEditorModal
          key={renderKey}
          isOpen={isOpen}
          onClose={() => setIsOpen(false)}
          data={data}
          title={title}
          initialEngine="exceljs"
          onSave={(updated) => {
            window.__SAVED_EXCEL__ = updated;
          }}
        />
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
          onSave={(updated) => {
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
