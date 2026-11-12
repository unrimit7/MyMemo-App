import fs from 'fs';
import path from 'path';
import { createServer } from 'vite';
import { chromium } from 'playwright-core';

const CHROME_PATH = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const SAMPLES_DIR = path.resolve('docs/office-upgrade/samples/real');
const OUT_DIR = path.resolve('docs/office-upgrade/out/ppt_render');

if (!fs.existsSync(OUT_DIR)) {
  fs.mkdirSync(OUT_DIR, { recursive: true });
}

const TARGETS = [
  { file: 'PPT_A_사업보고_원본.pptx', prefix: 'PPT_A_원본' },
  { file: 'PPT_A_사업보고_한쇼.pptx', prefix: 'PPT_A_한쇼' },
  { file: 'PPT_B_교육자료_원본.pptx', prefix: 'PPT_B_원본' },
  { file: 'PPT_B_교육자료_한쇼.pptx', prefix: 'PPT_B_한쇼' },
];

async function main() {
  console.log('1. Starting Vite dev server for MyMemo PptxViewerModal...');
  const server = await createServer({
    server: { port: 5182 },
    optimizeDeps: {
      exclude: ['@napi-rs/canvas'],
    },
    configFile: path.resolve('vite.config.ts'),
  });
  await server.listen();
  const address = server.httpServer.address();
  const port = address.port;
  console.log(`Vite server running at http://localhost:${port}`);

  console.log('2. Launching Chrome with Playwright...');
  const browser = await chromium.launch({
    executablePath: CHROME_PATH,
    headless: true,
  });

  const page = await browser.newPage({
    viewport: { width: 1600, height: 1000 },
  });

  await page.goto(`http://localhost:${port}/viewer_test.html`);
  console.log('Loaded viewer_test.html harness.');

  for (const target of TARGETS) {
    const filePath = path.join(SAMPLES_DIR, target.file);
    console.log(`\nProcessing presentation: ${target.file} -> ${target.prefix}_slide*.png`);
    const buf = fs.readFileSync(filePath);
    const dataUrl = `data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,${buf.toString('base64')}`;

    // Load PPTX into PptxViewerModal
    await page.evaluate(({ url, title }) => {
      if (window.__loadPptx) {
        return window.__loadPptx(url, title, 'viewer');
      }
      throw new Error('window.__loadPptx not found');
    }, { url: dataUrl, title: target.file });

    // Wait for the slide canvas to appear and loading indicator to disappear
    await page.waitForSelector('[data-testid="pptx-slide-canvas"]', { timeout: 15000 });
    await page.waitForSelector('.animate-spin', { state: 'detached', timeout: 15000 }).catch(() => {});

    // Find all thumbnail elements
    const thumbs = await page.$$('[data-testid^="slide-thumb-"]');
    console.log(`Found ${thumbs.length} slides for ${target.file}`);

    for (let i = 0; i < thumbs.length; i++) {
      // Click thumbnail to switch to slide i
      await thumbs[i].click();
      await page.waitForSelector(`[data-testid="pptx-slide-canvas"][data-slide-index="${i}"]`, { timeout: 10000 });
      // Brief wait for transitions or styles
      await page.waitForTimeout(300);

      const canvas = await page.$('[data-testid="pptx-slide-canvas"]');
      if (!canvas) throw new Error(`Canvas element missing for slide ${i + 1}`);

      const outFileName = `${target.prefix}_slide${i + 1}.png`;
      const outFilePath = path.join(OUT_DIR, outFileName);
      await canvas.screenshot({ path: outFilePath });
      const stats = fs.statSync(outFilePath);
      console.log(`  ✓ Captured ${outFileName} (${stats.size} bytes)`);
    }
  }

  console.log('\nAll 24 slides successfully captured from real PptxViewerModal!');
  await browser.close();
  await server.close();
  console.log('Cleaned up Chrome and Vite server.');
}

main().catch((err) => {
  console.error('Fatal error in viewer screenshot capture:', err);
  process.exit(1);
});
