import fs from 'fs';
import path from 'path';
import { createServer } from 'vite';
import { chromium } from 'playwright-core';

const CHROME_PATH = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const SAMPLES_DIR = path.resolve('docs/office-upgrade/samples/real');
const RENDER_DIR = path.resolve('docs/office-upgrade/out/hwp_render');

if (!fs.existsSync(RENDER_DIR)) {
  fs.mkdirSync(RENDER_DIR, { recursive: true });
}

const TARGET_FILES = [
  '한글C_업무계획_원본.hwpx',
  '한글C_업무계획_한글.hwpx',
  '한글C_업무계획_원본.hwp',
  '한글C_업무계획_한글.hwp',
];

async function main() {
  console.log('================================================================================');
  console.log('  [HWP 뷰어 페이지 렌더링 스크린샷 캡처]');
  console.log('================================================================================\n');

  const server = await createServer({
    server: { port: 5197 },
    configFile: path.resolve('vite.config.ts'),
  });
  await server.listen();
  const port = server.httpServer.address().port;

  const browser = await chromium.launch({
    executablePath: CHROME_PATH,
    headless: true,
  });
  const page = await browser.newPage({ viewport: { width: 1200, height: 1000 } });
  await page.goto(`http://localhost:${port}/viewer_test.html`);

  for (const name of TARGET_FILES) {
    const filePath = path.join(SAMPLES_DIR, name);
    const buf = fs.readFileSync(filePath);
    const fileKey = name.replace(/\./g, '_');
    const isHwpx = name.endsWith('.hwpx');
    const mimeType = isHwpx ? 'application/vnd.hancom.hwpx' : 'application/x-hwp';
    const dataUrl = `data:${mimeType};base64,${buf.toString('base64')}`;

    console.log(`[렌더링 캡처] ${name}...`);
    await page.evaluate(({ url, title }) => {
      return window.__loadHwp(url, title, 'viewer');
    }, { url: dataUrl, title: name });

    // 뷰어 SVG 렌더링 완료 대기
    await page.waitForSelector('svg', { timeout: 20000 });
    await page.waitForTimeout(600);

    // 1페이지 캡처
    const page1Path = path.join(RENDER_DIR, `${fileKey}_page1.png`);
    await page.screenshot({ path: page1Path, fullPage: false });
    console.log(`  -> 1페이지 캡처 완료: ${path.basename(page1Path)}`);

    // 2페이지로 이동 (다음 페이지 버튼 클릭)
    const nextBtn = page.locator('button[title*="다음"], button[aria-label*="다음"], button:has(svg.lucide-chevron-right)').first();
    if (await nextBtn.isVisible()) {
      await nextBtn.click();
      await page.waitForTimeout(600);
      const page2Path = path.join(RENDER_DIR, `${fileKey}_page2.png`);
      await page.screenshot({ path: page2Path, fullPage: false });
      console.log(`  -> 2페이지 캡처 완료: ${path.basename(page2Path)}`);
    }
  }

  await browser.close();
  await server.close();
  console.log('\n모든 렌더링 스크린샷 캡처가 완료되었습니다.');
}

main().catch(err => {
  console.error('[스크린샷 오류]:', err);
  process.exit(1);
});
