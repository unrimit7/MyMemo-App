# STEP 5 보완: 엑셀 인라인 카드 / 큰 화면 뷰어 / 팝아웃 시트 탭 전환, 표시 형식, 수식 계산 및 테마 독립 서식 개선 보고서

- **작성일자**: 2026-11-13
- **대상 저장소**: `MyMemo` (PWA 메모 앱)
- **배포 주소**: `https://unrimit7.github.io/MyMemo-App/`
- **검증 환경**: Google Chrome Headless (`C:\Program Files\Google\Chrome\Application\chrome.exe`), Node.js v26.13.0, Windows 11

---

## 1. 개요 및 배경

사용자 실사용 환경에서 6개 시트로 구성된 다중 시트 엑셀 문서(`홈페이지_유지비용_관리표_20261111.xlsx`)를 첨부하여 열람할 때 발견된 문제들을 보완하고 전수 실측 검증을 완료하였습니다.

> [!NOTE]
> **검증 대상 파일 사실 정정**: `docs/office-upgrade/samples/real/sample_다중시트_예시.xlsx`는 사용자의 실제 파일이 아니라 같은 시트 구조(6개 시트)와 데이터 형식을 재현한 **예시 파일**입니다. 테스트는 이 예시 파일로 진행되었으며, 실제 파일에서의 동작은 사용자가 직접 확인하셨습니다. 혼동을 방지하기 위해 파일명을 `sample_다중시트_예시.xlsx`로 변경하였습니다.

### 주요 보완 과제
1. **시트 탭 미작동 해결**: 인라인 카드, 큰 화면 뷰어(`ExcelViewerModal`), 창 분리 팝아웃(`FloatingNoteWindow`) 3곳 모두에서 시트 탭 클릭 시 해당 시트로 즉시 전환.
2. **표시 형식(numFmt) 미적용 및 Date.toString() 제거**:
   - 날짜(numFmt `yyyy-mm-dd`)가 `Fri Nov 27 2026 00:00:00 GMT+0900 (한국 표준시)`로 출력되던 현상 근절 (`2026-11-27`로 표시).
   - 금액(numFmt `#,##0"원";-#,##0"원";"-"`)의 섹션 포맷 적용 (`23,100원`).
   - 견적서 날짜(numFmt `yyyy"년" m"월" d"일"`) 적용 (`2026년 11월 11일`), 천 단위 쉼표 (`1,250,000`).
   - 백분율(numFmt `0.0%`) 적용 (`80.0%`).
3. **수식 결과 빈칸 해결 (동적 뷰어 수식 평가기)**:
   - openpyxl 등으로 생성되어 캐시된 계산값이 없는 수식(`=A4-TODAY()`, numFmt `0"일";"지남 "0"일"`)이나 휘발성 함수(`TODAY()`, `NOW()`)를 뷰어에서 실시간 계산하여 표시 (`14일` 또는 `지남 N일`).
   - 계산 실패 시 빈칸 대신 수식 문자열(`=...`)을 회색(`color: #94a3b8`)으로 표시. 원본 파일 바이트는 In-Place 무손실 유지.
4. **기본 글자색 검정(rgb(0,0,0)) 및 서식 있는 텍스트(RichText)**:
   - 원본에 글자색 지정이 없는 셀은 회색이 아닌 순수 검정(`rgb(0,0,0)`)으로 강제.
   - 견적서 B8의 서식 있는 텍스트(부분 빨간 굵은 글씨)를 run 단위로 `<span style="...">` 렌더링.

---

## 2. 수정된 파일 및 함수 목록

| 파일 | 주요 함수 / 컴포넌트 | 변경 내용 |
| :--- | :--- | :--- |
| `src/utils/excelViewer.ts` | `formatCellValue(val, numFmt)` | SheetJS `SSF.format` 연동, 날짜/시간/섹션문자열/천단위/백분율 포맷팅, Date 객체의 `toString()` 노출 경로 원천 차단 |
| `src/utils/excelViewer.ts` | `evaluateFormula(formula, resolveCell)` | 캐시값이 없거나 `TODAY()`/`NOW()`가 포함된 수식을 실시간 평가하여 남은 일수/비율 계산, 실패 시 회색 수식 문자열 반환 |
| `src/utils/excelViewer.ts` | `parseExcelWorkbook(dataOrUrl)` | 동기 파서에서 셀 직접 순회로 `cell.w`(포맷된 문자열) 및 `formatCellValue` 즉시 적용 |
| `src/utils/excelViewer.ts` | `parseExcelWorkbookWithStyles(dataOrUrl)` | `exceljs` 기반 셀 파싱 시 수식 평가(`evaluateFormula`), RichText 런 분리(`richTextRuns`), `formatCellValue` 적용, 기본 글자색 검정 보존 |
| `src/utils/excelViewer.ts` | `renderExcelTableHtml(parsed, sheetName)` | 기본 셀 글자색을 엄격히 `#000000`(`rgb(0,0,0)`)으로 지정, RichText run 단위 HTML 생성, 수식 실패 시 `#94a3b8` 회색 렌더링 |
| `src/utils/excelViewer.ts` | `switchExcelInlineSheet` | 탭 클릭 시 이벤트 발생 카드의 테이블 및 안내 문구 즉각 교체 |
| `src/components/ExcelViewerModal.tsx` | `ExcelViewerModal` | 기본 글자색 검정 `#000000` 적용, RichText `<span>` 렌더링 및 `formattedText` 우선 표출 |
| `src/components/FloatingNoteWindow.tsx` | `handlePreviewClick` | 팝아웃 창 내 시트 탭 클릭 이벤트 수신 및 `switchExcelInlineSheet` 연동 |
| `src/components/NoteEditor.tsx` | `handlePreviewClick` | 인라인 카드 시트 탭 클릭 이벤트 수신 및 `switchExcelInlineSheet` 연동 |
| `src/utils/attachment.ts` | `createAttachmentCardHtml` | 엑셀/워드/PPT/한글 인라인 카드에 `not-prose` 클래스 부여로 상위 다크 테마 간섭 격리 |

---

## 3. 실측 검증 결과 (Playwright E2E)

- 실행 명령: `node scripts/run_excel_view_playwright.mjs`
- 5분 타임아웃 및 프로세스 자동 종료(`process.exit(0)`) 처리 완료

### 실제값 vs 기대값 실측 단언표

| 번호 | 검증 항목 | 대상 위치 | 실제값 | 기대값 | 판정 |
| :---: | :--- | :--- | :--- | :--- | :---: |
| **1-1** | 인라인 초기 시트 로드 | 인라인 카드 A1 | `SSL 인증서 관리 대장` | `SSL 인증서 관리 대장` | **PASS** |
| **1-2** | 기본 글자색 (미지정 셀) | 인라인 데이터 셀 | `color: rgb(0, 0, 0)` | `color: rgb(0, 0, 0)` | **PASS** |
| **1-3** | 시트 탭 전환 ([서비스 현황]) | 인라인 카드 A1 | `클라우드 인프라 서비스 현황` | `클라우드 인프라 서비스 현황` | **PASS** |
| **1-4** | 날짜 표시 형식 (yyyy-mm-dd) | 결제 예정 A4 | `2026-11-27` | `2026-11-27` | **PASS** |
| **1-5** | 금액 표시 형식 (#,##0"원") | 결제 예정 E4 | `23,100원` | `23,100원` | **PASS** |
| **1-6** | 수식 동적 계산 (=A4-TODAY()) | 결제 예정 G4 | `14일` | `14일` (실행일 기준) | **PASS** |
| **2-1** | 큰 화면 뷰어 기본 글자색 | 뷰어 모달 셀 | `color: rgb(0, 0, 0)` | `color: rgb(0, 0, 0)` | **PASS** |
| **2-2** | 큰 화면 뷰어 날짜 표시 형식 | 뷰어 결제 예정 A4 | `2026-11-27` | `2026-11-27` | **PASS** |
| **2-3** | 큰 화면 뷰어 금액 표시 형식 | 뷰어 결제 예정 E4 | `23,100원` | `23,100원` | **PASS** |
| **2-4** | 큰 화면 뷰어 수식 계산 | 뷰어 결제 예정 G4 | `14일` | `14일` | **PASS** |
| **3-1** | 팝아웃 창 시트 탭 전환 | 팝아웃 카드 A1 | `연간 IT 유지비용 종합 요약` | `연간 IT 유지비용 종합 요약` | **PASS** |
| **4-1** | 견적서 날짜 표시 형식 | 견적서 C5 | `2026년 11월 11일` | `2026년 11월 11일` | **PASS** |
| **4-2** | 견적서 금액 표시 형식 | 견적서 F11 | `1,250,000` | `1,250,000` | **PASS** |
| **4-3** | 견적서 서식 있는 텍스트 | 견적서 B8 | `<span style="color: #FF0000; font-weight: 700;">부가세 포함</span>` | 부분 빨간 굵은 글씨 | **PASS** |
| **4-4** | 월별매출 백분율 표시 형식 | 월별매출 D2 | `80.0%` | `80.0%` | **PASS** |
| **4-5** | 인사평가 금액 표시 형식 | 인사평가 G4 | `3,000,000원` | `3,000,000원` | **PASS** |
| **5-1** | 타 오피스 테마 격리 | docx/pptx/hwp | `not-prose` 클래스 적용 | `not-prose` 격리 | **PASS** |

---

## 4. 생성된 증거 스크린샷 목록

모든 스크린샷은 [excel_view 디렉터리](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/)에 정상 저장되었습니다.

1. [01_인라인카드_초기시트_인증서현황.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/01_인라인카드_초기시트_인증서현황.png): 기본 검정 글자 `rgb(0,0,0)` 및 흰 바탕 표시
2. [02_인라인카드_탭전환_서비스현황.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/02_인라인카드_탭전환_서비스현황.png): 인라인 시트 탭 전환 및 에메랄드 활성 스타일
3. [03_큰화면뷰어_초기시트_인증서현황.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/03_큰화면뷰어_초기시트_인증서현황.png): 큰 화면 모달 뷰어 초기 시트
4. [04_큰화면뷰어_탭전환_결제예정.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/04_큰화면뷰어_탭전환_결제예정.png): 날짜(`2026-11-27`), 금액(`23,100원`), 남은 일수(`14일`) 표시
5. [05_팝아웃창_초기시트_인증서현황.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/05_팝아웃창_초기시트_인증서현황.png): 창 분리(팝아웃) 초기 시트
6. [06_팝아웃창_탭전환_연간비용.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/06_팝아웃창_탭전환_연간비용.png): 팝아웃 창에서 시트 탭 전환
7. [07_실파일A_MS엑셀_인라인카드.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/07_실파일A_MS엑셀_인라인카드.png): 인사평가 G4 `3,000,000원` 표시
8. [08_실파일A_MS엑셀_탭전환_부서코드.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/08_실파일A_MS엑셀_탭전환_부서코드.png): 실파일A 부서코드 탭 전환
9. [09_실파일B_MS엑셀_견적서_인라인카드.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/09_실파일B_MS엑셀_견적서_인라인카드.png): C5 `2026년 11월 11일`, F11 `1,250,000`, B8 부분 빨간 굵은 글씨
10. [10_실파일A_한셀_인라인카드.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/10_실파일A_한셀_인라인카드.png): 한셀 파일 정상 서식 표시
