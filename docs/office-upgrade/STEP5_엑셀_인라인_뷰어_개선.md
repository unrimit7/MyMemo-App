# STEP 5 보완: 엑셀 인라인 카드 / 큰 화면 뷰어 / 팝아웃 시트 탭 전환, 표시 형식, 화이트리스트 수식 계산 및 가로 정렬 개선 보고서

- **작성일자**: 2026-11-13
- **대상 저장소**: `MyMemo` (PWA 메모 앱)
- **배포 주소**: `https://unrimit7.github.io/MyMemo-App/`
- **검증 환경**: Google Chrome Headless (`C:\Program Files\Google\Chrome\Application\chrome.exe`), Node.js v26.13.0, Windows 11

---

## 1. 개요 및 배경

사용자 실사용 환경에서 6개 시트로 구성된 다중 시트 엑셀 문서(`홈페이지_유지비용_관리표_20261111.xlsx`)를 열람할 때 발견된 문제들을 전면 보완하고 실측 검증을 완료하였습니다.

> [!NOTE]
> **검증 대상 파일 사실 정정**: `docs/office-upgrade/samples/real/sample_다중시트_예시.xlsx`는 사용자의 실제 파일이 아니라 같은 시트 구조(6개 시트)와 데이터 형식을 재현한 **예시 파일**입니다. 테스트는 이 예시 파일로 진행되었으며, 실제 파일에서의 동작은 사용자가 직접 확인하셨습니다. 혼동을 방지하기 위해 파일명을 `sample_다중시트_예시.xlsx`로 명확히 변경하였습니다.

### 주요 보완 과제 및 해결 방안

1. **[화이트리스트 기반 안전한 수식 파서 (evaluateFormula)]**:
   - 기존의 함수명·쉼표·`$`·`^`·`%`·`&` 정규식 치환 및 `Function()` 평가 방식을 전면 폐기하였습니다. (치환으로 인해 `=MAX(A1,B1)`→`53`, `=$A$1*2`→`2`, `=A1^2`→`52`, `=A1*10%`→`50`, `=A1&B1`→`53` 등 틀린 숫자가 조용히 반환되던 치명적 결함 원천 해결)
   - **엄격한 화이트리스트 문법 파서(`FormulaTokenizer` & `FormulaParser`)** 구현:
     - 허용: 숫자 리터럴, 동일 시트 단일 셀 참조(`$A$1`, `A1`, `$A1`, `A$1`), 사칙연산(`+`, `-`, `*`, `/`, 단항 `+`, `-`), 괄호, `TODAY()`, `NOW()`, `DATE(y, m, d)`.
     - 미지원 함수(`MAX`, `MIN`, `ROUND`, `SUM`, `IF` 등), 연산자(`^`, `%`, `&` 등), 문자열, 타 시트 참조(`!`), 범위(`:`)가 하나라도 발견되면 즉시 `{ success: false }` 반환.
   - 수식 실패(`success: false`) 시: 파일에 캐시값(`val.result`)이 있으면 캐시값을 표시하고, 캐시값이 없으면 수식 문자열을 회색(`#94a3b8`)으로 표시하여 왜곡된 숫자가 절대 화면에 노출되지 않도록 차단하였습니다.
   - `tests/excelFormula.test.ts` 단위 테스트(18개 항목) 구축 완료: 7가지 왜곡 수식 및 `=A4-TODAY()`, `=IF(...)`, `=MAX(...)` 등 전수 검증.

2. **[숫자·날짜 가로 정렬(Horizontal Alignment) 복원]**:
   - `numFmt` 서식 적용 시 문자열(`string`)로 변환되더라도 원본 값의 타입(`rawType: 'number' | 'date' | 'boolean' | 'error' | 'string'`)을 추적하도록 개선.
   - 원본 셀에 명시적인 가로 정렬 지정이 없는 경우 엑셀 표준 기본 정렬 적용:
     - 숫자 및 날짜: `right` (우측 정렬)
     - 문자열: `left` (좌측 정렬)
     - 논리값(`TRUE`/`FALSE`) 및 오류(`error`): `center` (가운데 정렬)
   - 견적서 번호·수량·단가·금액(F11), 결제 예정 금액(E4)·남은 일수(G4), 인사평가(G4) 등이 모두 정상 우측 정렬됨을 Playwright로 실측 단언.

3. **[캐시 및 파서 수명주기 분리 (`hasDetailedStyles`)]**:
   - 동기 파서(`parseExcelWorkbook`, SheetJS)와 비동기 정밀 스타일 파서(`parseExcelWorkbookWithStyles`, ExcelJS)가 동일 캐시맵을 사용할 때, SheetJS의 불완전 캐시가 반환되어 ExcelJS 스타일 및 수식 평가가 누락되던 현상을 `hasDetailedStyles: boolean` 플래그로 완벽 분리 해결.

4. **[시트 탭 전환, 표시 형식, RichText, 테마 격리]**:
   - 인라인 카드, 큰 화면 뷰어(`ExcelViewerModal`), 창 분리 팝아웃(`FloatingNoteWindow`) 3곳 모두에서 시트 탭 전환 완벽 연동.
   - 날짜 `yyyy-mm-dd` (`2026-11-27`), 금액 `#,##0"원"` (`23,100원`), 잔여일 `0"일"` (`14일`), 천 단위 쉼표, 백분율 서식 적용.
   - 견적서 B8의 서식 있는 텍스트(부분 빨간 굵은 글씨) `<span>` run 단위 렌더링.
   - `not-prose` 클래스 부여로 상위 다크 테마 간섭 차단 및 기본 글자색 검정 `rgb(0,0,0)` 유지.

---

## 2. 수정된 파일 및 함수 목록

| 파일 | 주요 함수 / 인터페이스 | 변경 내용 |
| :--- | :--- | :--- |
| `src/utils/excelViewer.ts` | `FormulaTokenizer` & `FormulaParser` | 화이트리스트 기반 안전한 수식 토크나이저 및 재귀 하향 파서. 정규식 치환 및 `Function()` 완전 삭제 |
| `src/utils/excelViewer.ts` | `evaluateFormula` | 화이트리스트 외 문법/함수 즉시 `success: false` 반환, 안전한 계산 또는 캐시/회색 문자열 폴백 |
| `src/utils/excelViewer.ts` | `formatCellValue` | SheetJS `SSF.format` 연동, 날짜/시간/섹션문자열/천단위/백분율 포맷팅, Date 객체의 `toString()` 노출 차단 |
| `src/utils/excelViewer.ts` | `ParsedWorkbook` | `hasDetailedStyles: boolean` 추가로 SheetJS 캐시와 ExcelJS 상세 스타일 캐시 엄격 분리 |
| `src/utils/excelViewer.ts` | `parseExcelWorkbookWithStyles` | `rawType` 기반 가로 정렬(`right`/`center`/`left`), 수식 실시간 계산, RichText 런 분리, 검정 기본 글자색 보존 |
| `src/utils/excelViewer.ts` | `renderExcelTableHtml` | 셀 인라인 스타일에 `textAlign: cellVisual.align` 적용, 기본 글자색 `#000000`, RichText run 생성 |
| `src/utils/excelViewer.ts` | `switchExcelInlineSheet` | 인라인 시트 탭 클릭 시 ExcelJS 정밀 스타일로 즉시 재렌더링 |
| `src/components/ExcelViewerModal.tsx` | `ExcelViewerModal` | 큰 화면 모달에서 `cellVisual.align` 기반 가로 정렬, 기본 글자색 `#000000`, RichText 렌더링 |
| `scripts/update_sample.mjs` | 시트 재생성 로직 | `while (rowCount > 0)` 무한 루프 버그 수정, UTC 날짜 저장, 4단계 단계별 로깅 및 `process.exit(0)` |
| `scripts/run_excel_view_playwright.mjs` | E2E 검증 스크립트 | 5분 타임아웃 가드, 우측 정렬(`text-align: right`) 단언 8건 추가, `browser.close()` 및 `process.exit(0/1)` 보장 |
| `tests/excelFormula.test.ts` | 단위 테스트 (신규) | 18개 테스트 케이스: 7개 왜곡 수식 차단, 날짜 계산, 화이트리스트 외 문법 차단 검증 |

---

## 3. 단위 테스트 검증 결과 (`tests/excelFormula.test.ts`)

- 실행 명령: `npm test` (`vitest run`)
- 결과: **6개 테스트 파일 전체 통과 (총 62개 테스트 100% 통과)**

```text
 ✓ tests/excelFormula.test.ts (18 tests) 11ms
 ✓ tests/syncOriginalBackup.test.ts (2 tests) 5ms
 ✓ tests/docxFullComparison.test.ts (13 tests) 82ms
 ✓ tests/pptxFullComparison.test.ts (9 tests) 1142ms
 ✓ tests/excelRoundtrip.test.ts (15 tests) 474ms
 ✓ tests/excelFullComparison.test.ts (5 tests) 670ms

 Test Files  6 passed (6)
      Tests  62 passed (62)
```

### `tests/excelFormula.test.ts` 상세 항목
1. `=MAX(A1,B1)`: `success: false`, 왜곡값 `53` 반환 금지 확인 (PASS)
2. `=MIN(A1,B1)`: `success: false`, 왜곡값 `53` 반환 금지 확인 (PASS)
3. `=ROUND(C1,0)`: `success: false`, 임의 평가 금지 확인 (PASS)
4. `=$A$1*2`: `success: true`, 정확한 `10` 반환 (기존 `2` 왜곡 방지) (PASS)
5. `=A1^2`: `success: false`, 왜곡값 `52` 반환 금지 확인 (PASS)
6. `=A1*10%`: `success: false`, 왜곡값 `50` 반환 금지 확인 (PASS)
7. `=A1&B1`: `success: false`, 왜곡값 `53` 반환 금지 확인 (PASS)
8. `=A4-TODAY()`: `success: true`, 잔여 일수 동적 계산 (PASS)
9. `=IF(A4-TODAY()<0,"지남",A4-TODAY())`: `success: false`, 왜곡 숫자 반환 금지 (PASS)
10. `=MAX(0,A4-TODAY())`: `success: false`, 왜곡 숫자 반환 금지 (PASS)
11. `=DATE(2026, 11, 27) - DATE(2026, 11, 20)`: `success: true`, 정확한 `7` 계산 (PASS)
12. 타 시트 참조(`=Sheet1!A1+5`): `success: false` (PASS)
13. 범위 참조(`=SUM(A1:A5)`): `success: false` (PASS)
14. 숫자로 끝나는 함수명(`=LOG10(100)`): 셀 참조 오인 차단, `success: false` (PASS)
15. 0으로 나누기(`=A1/0`): 안전한 실패 처리 (PASS)
16. 복합 사칙연산(`=((A1 + B1) * 2 - 6) / 2`): 정확한 `5` 계산 (PASS)
17. 단항 마이너스 및 `$` 변형(`=-$A1 + A$1`): 정확한 `0` 계산 (PASS)
18. 문자열 숫자 셀(`D1="100"`): 숫자로 변환하여 `=D1 * 2` -> `200` 계산 (PASS)

---

## 4. Playwright E2E 실측 검증 결과

- 실행 명령: `node scripts/run_excel_view_playwright.mjs`
- 프로세스 종료 코드: `0` (정상 종료, 백그라운드 잔류 0개)

### 실제값 vs 기대값 실측 단언표

| 번호 | 검증 항목 | 대상 위치 | 실제값 | 기대값 | 판정 |
| :---: | :--- | :--- | :--- | :--- | :---: |
| **1-1** | 인라인 초기 시트 로드 | 인라인 카드 A1 | `SSL 인증서 관리 대장` | `SSL 인증서 관리 대장` | **PASS** |
| **1-2** | 기본 글자색 (미지정 셀) | 인라인 데이터 셀 | `color: rgb(0, 0, 0)` | `color: rgb(0, 0, 0)` | **PASS** |
| **1-3** | 시트 탭 전환 ([서비스 현황]) | 인라인 카드 A1 | `클라우드 인프라 서비스 현황` | `클라우드 인프라 서비스 현황` | **PASS** |
| **1-4** | 날짜 표시 형식 (yyyy-mm-dd) | 결제 예정 A4 | `2026-11-27` | `2026-11-27` | **PASS** |
| **1-4-1** | **날짜 셀 가로 우측 정렬** | 결제 예정 A4 | `text-align: right` | `text-align: right` | **PASS** |
| **1-5** | 금액 표시 형식 (#,##0"원") | 결제 예정 E4 | `23,100원` | `23,100원` | **PASS** |
| **1-5-1** | **금액 셀 가로 우측 정렬** | 결제 예정 E4 | `text-align: right` | `text-align: right` | **PASS** |
| **1-6** | 수식 동적 계산 (=A4-TODAY()) | 결제 예정 G4 | `14일` | `14일` (실행일 기준) | **PASS** |
| **1-6-1** | **남은 일수 셀 가로 우측 정렬** | 결제 예정 G4 | `text-align: right` | `text-align: right` | **PASS** |
| **2-1** | 큰 화면 뷰어 기본 글자색 | 뷰어 모달 셀 | `color: rgb(0, 0, 0)` | `color: rgb(0, 0, 0)` | **PASS** |
| **2-2** | 큰 화면 뷰어 날짜 표시 형식 | 뷰어 결제 예정 A4 | `2026-11-27` | `2026-11-27` | **PASS** |
| **2-2-1** | **큰 화면 날짜 셀 우측 정렬** | 뷰어 결제 예정 A4 | `text-align: right` | `text-align: right` | **PASS** |
| **2-3** | 큰 화면 뷰어 금액 표시 형식 | 뷰어 결제 예정 E4 | `23,100원` | `23,100원` | **PASS** |
| **2-3-1** | **큰 화면 금액 셀 우측 정렬** | 뷰어 결제 예정 E4 | `text-align: right` | `text-align: right` | **PASS** |
| **2-4** | 큰 화면 뷰어 수식 계산 | 뷰어 결제 예정 G4 | `14일` | `14일` | **PASS** |
| **2-4-1** | **큰 화면 남은 일수 우측 정렬** | 뷰어 결제 예정 G4 | `text-align: right` | `text-align: right` | **PASS** |
| **3-1** | 팝아웃 창 시트 탭 전환 | 팝아웃 카드 A1 | `연간 IT 유지비용 종합 요약` | `연간 IT 유지비용 종합 요약` | **PASS** |
| **4-1** | 견적서 날짜 표시 형식 | 견적서 C5 | `2026년 11월 11일` | `2026년 11월 11일` | **PASS** |
| **4-2** | 견적서 금액 표시 형식 | 견적서 F11 | `1,250,000` | `1,250,000` | **PASS** |
| **4-2-1** | **견적서 금액 셀 우측 정렬** | 견적서 F11 | `text-align: right` | `text-align: right` | **PASS** |
| **4-3** | 견적서 서식 있는 텍스트 | 견적서 B8 | `<span style="color: #FF0000; font-weight: 700;">부가세 포함</span>` | 부분 빨간 굵은 글씨 | **PASS** |
| **4-4** | 월별매출 백분율 표시 형식 | 월별매출 D2 | `80.0%` | `80.0%` | **PASS** |
| **4-5** | 인사평가 금액 표시 형식 | 인사평가 G4 | `3,000,000원` | `3,000,000원` | **PASS** |
| **4-5-1** | **인사평가 금액 셀 우측 정렬** | 인사평가 G4 | `text-align: right` | `text-align: right` | **PASS** |
| **5-1** | 타 오피스 테마 격리 | docx/pptx/hwp | `not-prose` 클래스 적용 | `not-prose` 격리 | **PASS** |

---

## 5. 생성된 증거 스크린샷 목록

모든 스크린샷은 [excel_view 디렉터리](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/)에 정상 저장되었습니다.

1. [01_인라인카드_초기시트_인증서현황.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/01_인라인카드_초기시트_인증서현황.png): 기본 검정 글자 `rgb(0,0,0)` 및 흰 바탕 표시
2. [02_인라인카드_탭전환_서비스현황.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/02_인라인카드_탭전환_서비스현황.png): 인라인 시트 탭 전환 및 에메랄드 활성 스타일
3. [03_큰화면뷰어_초기시트_인증서현황.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/03_큰화면뷰어_초기시트_인증서현황.png): 큰 화면 모달 뷰어 초기 시트
4. [04_큰화면뷰어_탭전환_결제예정.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/04_큰화면뷰어_탭전환_결제예정.png): 날짜(`2026-11-27`), 금액(`23,100원`), 남은 일수(`14일`), 우측 정렬 실측
5. [05_팝아웃창_초기시트_인증서현황.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/05_팝아웃창_초기시트_인증서현황.png): 창 분리(팝아웃) 초기 시트
6. [06_팝아웃창_탭전환_연간비용.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/06_팝아웃창_탭전환_연간비용.png): 팝아웃 창에서 시트 탭 전환
7. [07_실파일A_MS엑셀_인라인카드.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/07_실파일A_MS엑셀_인라인카드.png): 인사평가 G4 `3,000,000원` 및 우측 정렬
8. [08_실파일A_MS엑셀_탭전환_부서코드.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/08_실파일A_MS엑셀_탭전환_부서코드.png): 실파일A 부서코드 탭 전환
9. [09_실파일B_MS엑셀_견적서_인라인카드.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/09_실파일B_MS엑셀_견적서_인라인카드.png): C5 `2026년 11월 11일`, F11 `1,250,000` 우측 정렬, B8 RichText
10. [10_실파일A_한셀_인라인카드.png](file:///D:/Tool/MyMemo/docs/office-upgrade/out/excel_view/10_실파일A_한셀_인라인카드.png): 한셀 파일 정상 서식 표시
