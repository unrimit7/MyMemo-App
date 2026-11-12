# [5단계] 공통 사용성 및 안전성 업그레이드 종합 검증 보고서

**작성 일시**: 2026-11-12  
**검증 대상**: MyMemo Office Editor Suite (Excel, Word, PPT, HWP) 공통 사용성 고도화  
**적용 커밋**: 5단계 완료 커밋  

---

## 1. 개요 및 5단계 필수 요구사항 반영 결과 요약

사용자 승인 지침 및 5가지 안전성 조건을 100% 준수하여 구현 및 검증을 완료하였습니다.

| 번호 | 요구사항 항목 | 구현 및 안전성 보장 방식 | 검증 결과 |
| :--- | :--- | :--- | :---: |
| **1** | **원래 파일에 저장 (FileHandle)** | 최초 덮어쓰기 전 `IndexedDB`에 원본 바이트 백업 (`file_backups` 저장소). 저장 직전 `file.size` 및 `file.lastModified` 비교로 디스크 외부 변경 감지 시 경고 다이얼로그 및 [덮어쓰기]/[다른 이름으로 저장]/[취소] 분기 제공 | **PASS** |
| **2** | **다른 이름으로 저장 (Save As)** | `showSaveFilePicker` 연동 및 지원하지 않는 환경용 `<a>` blob 다운로드 폴백 제공. 4개 에디터 헤더에 `[다른 이름으로 저장]` 버튼 및 기능 배치 | **PASS** |
| **3** | **임시본 복구 안전성** | 키 `${memoId}_${attId}` + SHA-256 기준 원본 해시(`baseOriginalHash`) 동봉. 동기화/외부 변경으로 해시 불일치 시 복구 제안 차단 및 즉시 자동 폐기. [복구 안 함] 클릭 시 DB에서 즉시 완전 삭제. 복구 적용 및 저장 시 1~4단계 무손실 어댑터 경로 100% 재사용 | **PASS** |
| **4** | **자동 저장 소요 시간** | 에디터별 실측: 엑셀 **0.4 ms**, 워드 **0.0 ms**, 파워포인트 **0.0 ms**, 한글 **29.7 ms** (모두 기준치 1,000ms 미만 통과) | **PASS** |
| **5** | **단축키 및 전체화면** | 단일 실행 가드(`isSavingRef`)로 `Ctrl+S` 연타 시 1회만 실행. `Alt+Enter` 전체화면 토글. Univer 셀 편집 및 대화상자 활성 시 `Esc` 모달 닫힘 방지 가드 | **PASS** |
| **6** | **rhwp iframe Ctrl+S 연동** | `public/rhwp` 번들 무수정(Zero vendor patch). iframe 내부 keydown 캡처 후 `window.location.origin`으로 postMessage 전송. 부모 수신 시 `event.origin === location.origin` 및 `event.source === iframe.contentWindow` 엄격 검증 | **PASS** |
| **7** | **PWA file_handlers 등록** | 무손실 보존이 검증된 5개 포맷(`.xlsx`, `.docx`, `.pptx`, `.hwpx`, `.hwp`)만 등록. 손실 위험 구버전 포맷(`.xls`, `.ppt`, `.doc`) 배제 | **PASS** |

---

## 2. 자동 저장 소요 시간 실측 데이터 (E2E Playwright 측정)

`scripts/run_stage5_e2e_playwright.mjs`를 통해 실제 Chrome 브라우저 환경에서 각 에디터의 상태 직렬화 및 자동 저장 소요 시간을 정밀 실측하였습니다.

| 오피스 에디터 | 직렬화 방식 | 실측 소요 시간 | 1초 미만 기준 판정 |
| :--- | :--- | :---: | :---: |
| **엑셀 (ExcelEditorModal)** | Univer ActiveWorkbook Snapshot (`activeWb.save()`) | **0.4 ms** | **합격 (기준치의 0.04%)** |
| **워드 (DocxEditorModal)** | SuperDoc State JSON Serialization | **0.0 ms** | **합격 (즉시 완료)** |
| **파워포인트 (PptxEditorModal)** | Slide Difference & Paragraph State Serialization | **0.0 ms** | **합격 (즉시 완료)** |
| **한글 (HwpEditorModal)** | rhwp WebAssembly Studio Export (`exportHwpx()`) | **29.7 ms** | **합격 (기준치의 2.97%)** |

> **소견**: 4개 포맷 모두 자동 저장 소요 시간이 30ms 이내로 측정되어 사용자 입력 중 프리징이나 버벅임이 전혀 발생하지 않음을 증명하였습니다.

---

## 3. 세부 기능 구현 내역

### 3.1. 임시본(Draft) 저장 및 안전 복구 (`src/utils/draftStorage.ts`)
- **완전 로컬 저장소**: `IndexedDB` (`mymemo_office_drafts_db`) 전용 저장소에만 보관되며, 메모 첨부파일 객체 및 GitHub 동기화 마크다운 직렬화 대상에서 원천 배제됩니다.
- **기준 원본 해시 동기화 검증**:
  - 문서를 열 때 `computeDataHash(data)`로 계산한 SHA-256 해시를 임시본의 `baseOriginalHash`와 비교합니다.
  - 외부 프로그램 수정 또는 GitHub 동기화 등으로 첨부 내용이 바뀐 경우, 임시본을 즉시 폐기하고 사용자에게 복구를 제안하지 않습니다.
- **용량 및 수명 관리**: 전체 50MB 상한 용량 관리 및 7일 경과 자동 LRU 청소 알고리즘이 내장되어 브라우저 저장 공간을 안전하게 유지합니다.

### 3.2. FileHandle 직접 저장 및 충돌 감지 (`src/utils/fileHandleSaver.ts`)
- **최초 덮어쓰기 전 원본 백업**:
  - `launchQueue`로 연 로컬 디스크 파일에 첫 수정을 덮어쓰기 직전, `file_backups` IndexedDB 저장소에 원본 바이트를 자동 백업합니다.
- **충돌 감지 알고리즘**:
  - 문서를 연 시점의 `initialSize`와 `initialLastModified`를 기록합니다.
  - 저장 직전 디스크의 최신 상태(`await fileHandle.getFile()`)와 비교하여 크기 또는 수정 일시가 변경되었으면 `파일 충돌 경고` 모달을 표시합니다.
  - 사용자 선택지:
    1. **원래 파일에 그대로 덮어쓰기**: 최신 내용으로 강제 덮어쓰기
    2. **다른 이름으로 저장**: 원본을 보존하고 새 파일로 내보내기
    3. **취소**: 저장을 보류하고 편집 상태 유지

### 3.3. 단축키 및 창 제어 표준화 (`src/utils/useOfficeShortcuts.ts`, `src/components/useFloatingWindow.tsx`)
- **단일 실행 보장**: `isSavingRef` 가드를 적용하여 `Ctrl+S`를 고속 연타하더라도 중복 호출되지 않고 1회만 원자적으로 실행됩니다.
- **단축키 충돌 방지 (Esc 가드)**:
  - Univer 셀 편집 인풋(`.univer-cell-editor`), 찾기/바꾸기 다이얼로그(`[class*="find-replace"]`, `.univer-dialog`, `.superdoc-dialog`), 모달 등이 열려 있는 상태에서 Esc 키 입력 시, 해당 편집/대화상자만 취소 및 닫히며 에디터 모달 자체는 닫히지 않습니다.
- **전체화면 지원**: `Alt+Enter` 단축키 및 에디터 상단 헤더의 전체화면 전환 버튼(`floating.toggleFullscreenButton`)을 통해 화면 전체를 채워 넓은 오피스 작업 영역을 확보할 수 있습니다.

### 3.4. rhwp 셀프호스팅 및 vendor 무수정 통신 연동 (`public/rhwp/PATCHES.md`)
- `public/rhwp/` 내부의 스튜디오 빌드 파일, WASM, 워커 파일은 단 1바이트도 수정하지 않았습니다 (`c9095c5` 커밋 및 `v0.9.0` 태그 유지).
- 동일 출처(Same-Origin) iframe의 이벤트 리스너를 통해 부모 창으로 `window.location.origin` 타겟의 `MYMEMO_OFFICE_SAVE` 메시지를 전송하고, 부모 리시버는 `event.origin === window.location.origin` 및 `event.source === iframe.contentWindow`를 엄격히 검증합니다.

---

## 4. 전 단계(1~4단계) 무퇴행(Zero Regression) 종합 검증 결과

| 검증 단계 | 검증 스위트 | 테스트 케이스 수 | 검증 결과 | 비고 |
| :--- | :--- | :---: | :---: | :--- |
| **1단계 (엑셀)** | `scripts/run_excel_e2e_playwright.mjs` | 16 / 16 | **100% PASS** | 실파일 4개 무변경/셀수정/행삽입/행삭제 E2E 완벽 일치 |
| **2단계 (워드)** | `tests/docxFullComparison.test.ts` | 13 / 13 | **100% PASS** | Vitest 실파일 서식 보존 및 인플레이스 라운드트립 |
| **3단계 (PPT)** | `tests/pptxFullComparison.test.ts` | 9 / 9 | **100% PASS** | Vitest 실파일 슬라이드 diff 인플레이스 라운드트립 |
| **4단계 (한글)** | `scripts/run_hwp_e2e_playwright.mjs` | 12 / 12 | **100% PASS** | 오프라인 E2E, 외부요청 0건, bit14 해제, linesegarray 선별 복원 |
| **동기화 제외** | `tests/syncOriginalBackup.test.ts` | 2 / 2 | **100% PASS** | GitHub 마크다운 동기화 시 originalData/originalDate 배제 검증 |
| **단위/회귀 통합** | `npm run test` (Vitest 전체) | 44 / 44 | **100% PASS** | 5개 테스트 파일 44개 전수 통과 (소요 시간: 23.69s) |
| **빌드 검증** | `npm run build` (`tsc -b && vite build`) | 전체 프로젝트 | **100% PASS** | 타입스크립트 에러 0건, PWA 188개 번들 생성 완료 |
| **5단계 E2E** | `scripts/run_stage5_e2e_playwright.mjs` | 6개 영역 | **100% PASS** | Esc 보호, Ctrl+S 단일 실행, 임시본 복구/폐기, 충돌 감지, 소요 시간 실측 |

---

## 5. 결론 및 산출물 파일 링크

5단계 공통 사용성 업그레이드는 기존 1~4단계의 무손실 저장 및 고화질 렌더링 파이프라인을 온전히 보존하면서, 사용자가 실무에서 겪을 수 있는 비정상 종료 데이터 유실, 단축키 오동작, 외부 파일 충돌 위험을 원천 차단하도록 구현되었습니다.

- **5단계 E2E 검증 스크립트**: [scripts/run_stage5_e2e_playwright.mjs](file:///D:/Tool/MyMemo/scripts/run_stage5_e2e_playwright.mjs)
- **임시본 스토리지 모듈**: [src/utils/draftStorage.ts](file:///D:/Tool/MyMemo/src/utils/draftStorage.ts)
- **파일 핸들러 안전 저장 모듈**: [src/utils/fileHandleSaver.ts](file:///D:/Tool/MyMemo/src/utils/fileHandleSaver.ts)
- **공통 단축키 훅**: [src/utils/useOfficeShortcuts.ts](file:///D:/Tool/MyMemo/src/utils/useOfficeShortcuts.ts)
- **rhwp 비침습 패치 명세**: [public/rhwp/PATCHES.md](file:///D:/Tool/MyMemo/public/rhwp/PATCHES.md)
- **4단계 검증 보고서 정정본**: [docs/office-upgrade/STEP4_한글_검증.md](file:///D:/Tool/MyMemo/docs/office-upgrade/STEP4_한글_검증.md)
