import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { JSDOM } from 'jsdom';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, '..');

// JSDOM 환경 설정 (SuperDoc 브라우저 런타임 호환)
const dom = new JSDOM('<!DOCTYPE html><html><body><div id="editor"></div></body></html>', {
  url: 'http://localhost/',
  pretendToBeVisual: true,
});

global.window = dom.window;
global.document = dom.window.document;
try {
  Object.defineProperty(global, 'navigator', { value: dom.window.navigator, configurable: true, writable: true });
} catch (e) {
  global.navigator = dom.window.navigator;
}
global.SVGElement = dom.window.SVGElement;
global.HTMLElement = dom.window.HTMLElement;
global.EventTarget = dom.window.EventTarget;
global.Event = dom.window.Event;
global.CustomEvent = dom.window.CustomEvent;
global.AbortSignal = dom.window.AbortSignal;
global.AbortController = dom.window.AbortController;
global.Blob = globalThis.Blob;
global.File = globalThis.File || dom.window.File;
if (!dom.window.Blob.prototype.stream && globalThis.Blob.prototype.stream) {
  dom.window.Blob.prototype.stream = globalThis.Blob.prototype.stream;
}
global.FileReader = dom.window.FileReader;
global.requestAnimationFrame = (fn) => setTimeout(fn, 16);
global.cancelAnimationFrame = (id) => clearTimeout(id);

for (const key of Object.getOwnPropertyNames(dom.window)) {
  if (!(key in global)) {
    try {
      global[key] = dom.window[key];
    } catch (e) {}
  }
}

// SuperDoc 로드
const { SuperDoc } = await import('superdoc');

const targets = [
  {
    name: '워드A_사업보고서_원본',
    inputFile: 'docs/office-upgrade/samples/real/워드A_사업보고서_원본.docx',
    outputFile: 'docs/office-upgrade/out/워드A_사업보고서_원본_수정저장결과.docx',
    edits: [
      {
        findText: '2026년 하반기 사업 점검 보고서',
        replacement: '2026년 하반기 사업 점검 보고서 (수정완료)',
        description: '본문 제목 텍스트 수정'
      },
      {
        findText: '개발/디자인',
        replacement: '개발/디자인 (개편)',
        description: '표 셀 텍스트 수정'
      }
    ]
  },
  {
    name: '워드A_사업보고서_한글',
    inputFile: 'docs/office-upgrade/samples/real/워드A_사업보고서_한글.docx',
    outputFile: 'docs/office-upgrade/out/워드A_사업보고서_한글_수정저장결과.docx',
    edits: [
      {
        findText: '2026년 하반기 사업 점검 보고서',
        replacement: '2026년 하반기 사업 점검 보고서 (수정완료)',
        description: '본문 제목 텍스트 수정'
      },
      {
        findText: '개발/디자인',
        replacement: '개발/디자인 (개편)',
        description: '표 셀 텍스트 수정'
      }
    ]
  },
  {
    name: '워드B_신청서양식_원본',
    inputFile: 'docs/office-upgrade/samples/real/워드B_신청서양식_원본.docx',
    outputFile: 'docs/office-upgrade/out/워드B_신청서양식_원본_수정저장결과.docx',
    edits: [
      {
        findText: '김가상',
        replacement: '김가상 (수정완료)',
        description: '성명 칸 본문 텍스트 수정'
      },
      {
        findText: '외부 고객 시연회 진행',
        replacement: '외부 고객 시연회 진행 (개편)',
        description: '표 셀 텍스트 수정'
      }
    ]
  },
  {
    name: '워드B_신청서양식_한글',
    inputFile: 'docs/office-upgrade/samples/real/워드B_신청서양식_한글.docx',
    outputFile: 'docs/office-upgrade/out/워드B_신청서양식_한글_수정저장결과.docx',
    edits: [
      {
        findText: '김가상',
        replacement: '김가상 (수정완료)',
        description: '성명 칸 본문 텍스트 수정'
      },
      {
        findText: '외부 고객 시연회 진행',
        replacement: '외부 고객 시연회 진행 (개편)',
        description: '표 셀 텍스트 수정'
      }
    ]
  }
];

async function runSuperDocEdits() {
  console.log('====================================================');
  console.log('SuperDoc 실제 편집(본문 1곳 + 표 셀 1곳) 및 Export 실행');
  console.log('====================================================');

  for (const t of targets) {
    console.log(`\n▶ [${t.name}] 작업 시작`);
    const inputPath = path.join(rootDir, t.inputFile);
    const outputPath = path.join(rootDir, t.outputFile);

    const buf = fs.readFileSync(inputPath);
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });

    const container = dom.window.document.createElement('div');
    dom.window.document.body.appendChild(container);

    const superdoc = new SuperDoc({
      selector: container,
      document: blob,
      documentMode: 'editing',
      ui: false
    });

    try {
      await new Promise((resolve, reject) => {
        superdoc.on('ready', resolve);
        superdoc.on('error', reject);
        setTimeout(() => reject(new Error('SuperDoc init timeout (10s)')), 10000);
      });
      console.log('  ✓ SuperDoc 엔진 로드 완료');

      const runtime = superdoc.getActiveRuntime();
      const proj = runtime.getLegacyEditorProjection();
      if (!proj?.authoring?.replaceTextByText) {
        throw new Error('SuperDoc authoring.replaceTextByText API를 찾을 수 없습니다.');
      }

      for (const edit of t.edits) {
        const res = await proj.authoring.replaceTextByText({
          findText: edit.findText,
          replacement: edit.replacement
        });
        if (!res?.ok) {
          throw new Error(`텍스트 수정 실패 (${edit.description}): ${res?.reason || 'unknown'} - ${res?.detail || ''}`);
        }
        console.log(`  ✓ ${edit.description} 완료: "${edit.findText}" → "${edit.replacement}"`);
      }

      const exportedBlob = await superdoc.export({ exportType: ['docx'], triggerDownload: false });
      const arrayBuffer = await exportedBlob.arrayBuffer();
      const exportBuf = Buffer.from(arrayBuffer);

      fs.writeFileSync(outputPath, exportBuf);
      console.log(`  ✓ 수정 결과 저장 완료 (${exportBuf.length} 바이트) -> ${t.outputFile}`);
    } finally {
      try {
        superdoc.destroy();
      } catch (e) {}
      container.remove();
    }
  }

  console.log('\n====================================================');
  console.log('전체 4개 파일 SuperDoc 실제 수정 및 Export 성공!');
  console.log('====================================================');
}

runSuperDocEdits().catch((err) => {
  console.error('SuperDoc 수정 실행 중 치명적 오류:', err);
  process.exit(1);
});
