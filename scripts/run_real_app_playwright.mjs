import { chromium } from 'playwright-core';
import { spawn } from 'child_process';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, '..');
const outDir = path.join(projectRoot, 'docs', 'office-upgrade', 'out', 'excel_view');

// 5분 타임아웃 가드 (절대 무한 대기하지 않음)
const globalTimer = setTimeout(() => {
  console.error('\n[TIMEOUT] 5분 시간 제한 초과로 프로세스를 강제 종료합니다.');
  process.exit(1);
}, 300000);
globalTimer.unref();

function assert(condition, message) {
  if (!condition) {
    throw new Error(`[AssertionFailed] ${message}`);
  }
}

function fileToDataUrl(filePath, mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
  const buf = fs.readFileSync(filePath);
  return `data:${mime};base64,${buf.toString('base64')}`;
}

async function startViteServer() {
  const port = 5196;
  const proc = spawn('npx', ['vite', '--port', String(port), '--strictPort'], {
    cwd: projectRoot,
    shell: true,
    stdio: ['ignore', 'pipe', 'pipe'],
  });

  const url = `http://localhost:${port}/`;

  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      reject(new Error('Vite 서버 시작 30초 타임아웃'));
    }, 30000);

    const checkReady = (data) => {
      const text = data.toString();
      if (text.includes(`localhost:${port}`) || text.includes('ready in') || text.includes('Local:')) {
        clearTimeout(timer);
        resolve();
      }
    };

    proc.stdout.on('data', checkReady);
    proc.stderr.on('data', checkReady);

    proc.on('error', (err) => {
      clearTimeout(timer);
      reject(err);
    });

    proc.on('exit', (code) => {
      clearTimeout(timer);
      reject(new Error(`Vite 서버 프로세스 조기 종료 (code: ${code})`));
    });
  });

  return {
    proc,
    url,
    kill: () => {
      try {
        proc.kill('SIGTERM');
        proc.kill('SIGKILL');
      } catch {}
    },
  };
}

async function run() {
  fs.mkdirSync(outDir, { recursive: true });

  console.log('================================================================');
  console.log('실제 앱(NoteEditor + IndexedDB) Playwright E2E 검증 시작');
  console.log('================================================================');

  let server = null;
  let browser = null;
  let exitCode = 1;

  try {
    console.log('\n[1/6] Vite 테스트 서버 가동 (포트 5196)...');
    server = await startViteServer();
    console.log(`  ✓ 서버 가동 완료: ${server.url}`);

    console.log('\n[2/6] Chrome 브라우저 실행...');
    browser = await chromium.launch({
      executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
      headless: true,
    });

    const context = await browser.newContext({
      viewport: { width: 1400, height: 900 },
    });
    const page = await context.newPage();

    page.on('console', (msg) => {
      if (msg.type() === 'error') {
        console.log('  [브라우저 에러]:', msg.text());
      } else if (msg.type() === 'warn') {
        console.log('  [브라우저 경고]:', msg.text());
      }
    });

    // 샘플 파일 Base64 Data URL 준비
    const sampleMultiFile = path.join(projectRoot, 'docs', 'office-upgrade', 'samples', 'real', 'sample_다중시트_예시.xlsx');
    const sampleMultiDataUrl = fileToDataUrl(sampleMultiFile);

    // 기대값 계산 (G4 동적 남은 일수)
    const now = new Date();
    const todayUTC = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
    const targetUTC = Date.UTC(2026, 10, 27); // 2026-11-27
    const expectedDaysNum = Math.round((targetUTC - todayUTC) / 86400000);
    const expectedG4Text = `${expectedDaysNum}일`;

    console.log('\n[3/6] 실제 앱 초기화 및 IndexedDB 시딩...');
    await page.goto(server.url, { waitUntil: 'domcontentloaded' });

    // IndexedDB에 테스트 프로젝트 및 2개 메모(비용 관리표 메모, 일반 텍스트 메모) 시딩
    await page.evaluate(async ({ dataUrl }) => {
      const DB_NAME = 'MyMemo_DB';
      const DB_VERSION = 1;

      await new Promise((resolve, reject) => {
        const req = indexedDB.open(DB_NAME, DB_VERSION);
        req.onupgradeneeded = (e) => {
          const db = e.target.result;
          if (!db.objectStoreNames.contains('projects')) {
            db.createObjectStore('projects', { keyPath: 'id' });
          }
          if (!db.objectStoreNames.contains('notes')) {
            const noteStore = db.createObjectStore('notes', { keyPath: 'id' });
            noteStore.createIndex('by-project', 'projectId');
            noteStore.createIndex('by-updated', 'updatedAt');
          }
          if (!db.objectStoreNames.contains('settings')) {
            db.createObjectStore('settings');
          }
        };
        req.onsuccess = (e) => {
          const db = e.target.result;
          const tx = db.transaction(['projects', 'notes'], 'readwrite');
          const pStore = tx.objectStore('projects');
          const nStore = tx.objectStore('notes');

          pStore.put({
            id: 'proj-real-test',
            name: '업무 관리',
            createdAt: '2026-11-11T00:00:00.000Z',
            isLocalOnly: false,
          });

          // 메모 1: sample_다중시트_예시.xlsx 첨부
          nStore.put({
            id: 'note-real-excel',
            projectId: 'proj-real-test',
            title: '홈페이지 비용 관리표 메모',
            content: '# 홈페이지 비용 관리\n\n아래 첨부된 엑셀 관리표를 확인하세요:\n\n[sample_다중시트_예시.xlsx](attachment:att-real-excel)\n\n이상입니다.',
            attachments: {
              'att-real-excel': {
                id: 'att-real-excel',
                name: 'sample_다중시트_예시.xlsx',
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                data: dataUrl,
              },
            },
            isLocalOnly: false,
            syncStatus: 'synced',
            createdAt: '2026-11-11T00:00:00.000Z',
            updatedAt: '2026-11-11T00:00:00.000Z',
          });

          // 메모 2: 일반 텍스트 메모 (메모 전환 검증용)
          nStore.put({
            id: 'note-real-text',
            projectId: 'proj-real-test',
            title: '다른 일반 메모',
            content: '# 다른 일반 메모 본문\n\n이 메모는 첨부 파일이 없는 일반 마크다운 메모입니다.',
            attachments: {},
            isLocalOnly: false,
            syncStatus: 'synced',
            createdAt: '2026-11-12T00:00:00.000Z',
            updatedAt: '2026-11-12T00:00:00.000Z',
          });

          tx.oncomplete = () => resolve();
          tx.onerror = (err) => reject(err);
        };
        req.onerror = (err) => reject(err);
      });

      localStorage.setItem('mymemo_active_note_id', 'note-real-excel');
      localStorage.setItem('mymemo_active_project_id', 'ALL');
      localStorage.setItem('mymemo_theme', 'navy');
    }, { dataUrl: sampleMultiDataUrl });

    // 시딩된 데이터로 앱 새로고침
    await page.reload({ waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(1000);

    // 인라인 카드 렌더링 및 하이드레이션 완료 대기
    await page.waitForSelector('.excel-inline-viewer');
    await page.waitForTimeout(600);

    // 검증 헬퍼 함수
    async function extractAndAssertCells(caseTitle) {
      console.log(`\n--- [검증] ${caseTitle} ---`);

      // 렌더링 및 스타일 적용 대기 (최대 5초)
      await page.waitForFunction(() => {
        const tr4 = document.querySelector('.excel-inline-viewer table tbody tr:nth-child(4)');
        if (!tr4) return false;
        const tds = tr4.querySelectorAll('td.excel-td');
        if (tds.length < 7) return false;
        // A4 날짜가 채워졌는지 확인
        return tds[0].textContent.trim().length > 0;
      }, { timeout: 10000 });

      const evaluation = await page.evaluate(() => {
        const card = document.querySelector('.excel-inline-viewer');
        if (!card) return { error: '카드 요소를 찾을 수 없음' };

        const tr3 = card.querySelector('table tbody tr:nth-child(3)');
        const tr4 = card.querySelector('table tbody tr:nth-child(4)');
        if (!tr3 || !tr4) return { error: 'tr 요소를 찾을 수 없음' };

        const tdA3 = tr3.querySelectorAll('td.excel-td')[0];
        const tds4 = tr4.querySelectorAll('td.excel-td');
        const tdA4 = tds4[0];
        const tdE4 = tds4[4];
        const tdG4 = tds4[6];

        const csA3 = window.getComputedStyle(tdA3);
        const csA4 = window.getComputedStyle(tdA4);
        const csE4 = window.getComputedStyle(tdE4);
        const csG4 = window.getComputedStyle(tdG4);

        const footer = card.querySelector('.excel-footer-info');

        return {
          A3_text: tdA3.textContent.trim(),
          A3_bg: csA3.backgroundColor,
          A3_inlineBg: tdA3.style.backgroundColor,
          A4_text: tdA4.textContent.trim(),
          A4_align: csA4.textAlign,
          A4_inlineAlign: tdA4.style.textAlign,
          E4_text: tdE4.textContent.trim(),
          G4_text: tdG4.textContent.trim(),
          footer_text: footer?.textContent?.trim() || '',
          footer_html: footer?.innerHTML || '',
          hasUnstyledBadge: (footer?.innerHTML || '').includes('서식 미적용'),
        };
      });

      console.log('  [실제 측정값]:', JSON.stringify(evaluation, null, 2));

      // 단언 1: A4 날짜 == "2026-11-27"
      assert(
        evaluation.A4_text === '2026-11-27',
        `A4 날짜 불일치 (실제: '${evaluation.A4_text}', 기대: '2026-11-27')`
      );
      console.log(`  ✓ A4 날짜 단언 통과: 실제값 '${evaluation.A4_text}' == 기대값 '2026-11-27'`);

      // 단언 2: A4 가운데 정렬 == "center"
      assert(
        evaluation.A4_align === 'center' || evaluation.A4_inlineAlign === 'center',
        `A4 가운데 정렬 불일치 (실제: computed='${evaluation.A4_align}', inline='${evaluation.A4_inlineAlign}')`
      );
      console.log(`  ✓ A4 정렬 단언 통과: 실제값 '${evaluation.A4_align}' == 기대값 'center'`);

      // 단언 3: E4 금액 == "23,100원"
      assert(
        evaluation.E4_text === '23,100원',
        `E4 금액 불일치 (실제: '${evaluation.E4_text}', 기대: '23,100원')`
      );
      console.log(`  ✓ E4 금액 단언 통과: 실제값 '${evaluation.E4_text}' == 기대값 '23,100원'`);

      // 단언 4: G4 남은 일수 계산값
      assert(
        evaluation.G4_text === expectedG4Text,
        `G4 남은 일수 불일치 (실제: '${evaluation.G4_text}', 기대: '${expectedG4Text}')`
      );
      console.log(`  ✓ G4 남은 일수 단언 통과: 실제값 '${evaluation.G4_text}' == 기대값 '${expectedG4Text}'`);

      // 단언 5: A3 머리글 채우기색 (흰색이나 투명이 아님)
      const isNotWhiteOrTransparent =
        evaluation.A3_bg !== 'rgb(255, 255, 255)' &&
        evaluation.A3_bg !== 'rgba(0, 0, 0, 0)' &&
        evaluation.A3_bg !== 'transparent';
      assert(
        isNotWhiteOrTransparent,
        `A3 채우기색 미적용 (실제: '${evaluation.A3_bg}')`
      );
      console.log(`  ✓ A3 채우기색 단언 통과: 실제값 '${evaluation.A3_bg}' (헤더 채우기 적용됨)`);

      // 단언 6: 서식 미적용 배지 미표시
      assert(
        !evaluation.hasUnstyledBadge,
        `하단에 '서식 미적용' 경고 배지가 표시됨! (footer: ${evaluation.footer_html})`
      );
      console.log(`  ✓ 서식 적용 정상 확인: '(서식 미적용)' 배지 없음`);

      return evaluation;
    }

    // =========================================================================
    // CASE 1: 페이지 로드 직후 (Initial Load)
    // =========================================================================
    console.log('\n[4/6] CASE 1: 페이지 로드 직후 검증');
    await extractAndAssertCells('CASE 1: 페이지 로드 직후');
    const shotCase1 = path.join(outDir, 'real_app_case1_initial_load.png');
    await page.locator('.excel-inline-viewer').screenshot({ path: shotCase1 });
    console.log(`  ✓ CASE 1 스크린샷 저장: ${shotCase1}`);

    // =========================================================================
    // CASE 2: 시트 탭 클릭 후 (Sheet Tab Click)
    // =========================================================================
    console.log('\n[5/6] CASE 2: 시트 탭 클릭 후 전환 검증');
    // '서비스 현황' 시트 탭 클릭
    console.log('  -> [서비스 현황] 시트 탭 클릭...');
    await page.click('.excel-inline-viewer .excel-sheet-pill[data-sheet-name="서비스 현황"]');
    await page.waitForTimeout(500);

    // 다시 '결제 예정' 시트 탭 클릭
    console.log('  -> 다시 [결제 예정] 시트 탭 클릭...');
    await page.click('.excel-inline-viewer .excel-sheet-pill[data-sheet-name="결제 예정"]');
    await page.waitForTimeout(500);

    await extractAndAssertCells('CASE 2: 시트 탭 클릭 후');
    const shotCase2 = path.join(outDir, 'real_app_case2_sheet_tab_click.png');
    await page.locator('.excel-inline-viewer').screenshot({ path: shotCase2 });
    console.log(`  ✓ CASE 2 스크린샷 저장: ${shotCase2}`);

    // =========================================================================
    // CASE 3: 메모 전환 후 다시 돌아왔을 때 (Note Switch Back)
    // =========================================================================
    console.log('\n[6/6] CASE 3: 메모 전환 후 다시 돌아왔을 때 검증');
    console.log('  -> [다른 일반 메모] 클릭하여 메모 전환...');
    await page.click('h3:has-text("다른 일반 메모")');
    await page.waitForTimeout(600);

    // 엑셀 뷰어가 사라졌는지 확인
    const hasExcelInTextNote = await page.evaluate(() => !!document.querySelector('.excel-inline-viewer'));
    assert(!hasExcelInTextNote, '일반 텍스트 메모로 전환 후에도 엑셀 뷰어가 남아 있음');
    console.log('  ✓ 일반 텍스트 메모 전환 성공 (엑셀 뷰어 없음)');

    console.log('  -> 다시 [홈페이지 비용 관리표 메모] 클릭하여 복귀...');
    await page.click('h3:has-text("홈페이지 비용 관리표 메모")');
    await page.waitForTimeout(800);

    await extractAndAssertCells('CASE 3: 메모 전환 후 다시 복귀');
    const shotCase3 = path.join(outDir, 'real_app_case3_note_switch_back.png');
    await page.locator('.excel-inline-viewer').screenshot({ path: shotCase3 });
    console.log(`  ✓ CASE 3 스크린샷 저장: ${shotCase3}`);

    // =========================================================================
    // 추가 검증: 설정 > 정보 커밋 ID 및 빌드 시각 표시 확인 (Requirement 3)
    // =========================================================================
    console.log('\n--- [추가 검증] 설정 > 정보 빌드 커밋 ID 표시 ---');
    await page.click('button[title*="설정"]');
    await page.waitForSelector('.app-commit-hash', { timeout: 5000 });

    const buildInfo = await page.evaluate(() => {
      const commitEl = document.querySelector('.app-commit-hash');
      return {
        commitHash: commitEl?.textContent?.trim() || '',
      };
    });
    assert(buildInfo.commitHash.length > 0, '설정 모달에 커밋 해시가 표시되지 않음');
    console.log(`  ✓ 설정 모달 커밋 해시 확인: '${buildInfo.commitHash}'`);
    const shotCase4 = path.join(outDir, 'real_app_case4_settings_version.png');
    await page.screenshot({ path: shotCase4 });
    console.log(`  ✓ 설정 모달 스크린샷 저장: ${shotCase4}`);

    // =========================================================================
    // 추가 검증: PWA 새 버전 업데이트 알림 배너 확인 (Requirement 4)
    // =========================================================================
    console.log('\n--- [추가 검증] PWA 새 버전 배포 알림 배너 ---');
    // 모달 닫기
    await page.keyboard.press('Escape');
    await page.waitForTimeout(300);

    // PWA 업데이트 이벤트 디스패치
    await page.evaluate(() => {
      window.dispatchEvent(
        new CustomEvent('mymemo_pwa_update_available', {
          detail: {
            update: () => {
              window.__pwa_updated_called = true;
            },
          },
        })
      );
    });

    await page.waitForSelector('text=새 버전이 배포되었습니다', { timeout: 5000 });
    const hasPwaPrompt = await page.evaluate(() => {
      const prompt = document.querySelector('body')?.innerText || '';
      return prompt.includes('새 버전이 배포되었습니다') && prompt.includes('지금 새로고침');
    });
    assert(hasPwaPrompt, 'PWA 새 버전 배포 알림 배너가 정상 표시되지 않음');
    console.log('  ✓ PWA 새 버전 배포 알림 배너 및 [지금 새로고침] 버튼 확인 완료');

    const shotCase5 = path.join(outDir, 'real_app_case5_pwa_update_prompt.png');
    await page.screenshot({ path: shotCase5 });
    console.log(`  ✓ PWA 알림 배너 스크린샷 저장: ${shotCase5}`);

    console.log('\n================================================================');
    console.log('🎉 모든 단언 및 검증이 100% 성공적으로 통과되었습니다!');
    console.log('================================================================');
    exitCode = 0;
  } catch (err) {
    console.error('\n❌ [TEST_FAILED] 오류 발생:', err);
    exitCode = 1;
  } finally {
    console.log('\n[정리 작업] 브라우저 및 서버 종료 중...');
    if (browser) {
      try {
        await browser.close();
        console.log('  ✓ 브라우저 닫기 완료');
      } catch (e) {
        console.error('  ! 브라우저 닫기 오류:', e);
      }
    }
    if (server) {
      try {
        server.kill();
        console.log('  ✓ Vite 서버 종료 완료');
      } catch (e) {
        console.error('  ! Vite 서버 종료 오류:', e);
      }
    }
    clearTimeout(globalTimer);
    console.log(`테스트 종료 (종료 코드: ${exitCode})`);
    process.exit(exitCode);
  }
}

run();
