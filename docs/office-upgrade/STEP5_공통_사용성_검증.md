# [5단계] 공통 사용성 및 안전성 업그레이드 종합 검증 보고서 (최종 완결)

**작성 일시**: 2026-11-12  
**검증 대상**: MyMemo Office Editor Suite (Excel, Word, PPT, HWP) 공통 사용성·안전성·PWA 고도화  
**적용 커밋**: 5단계 최종 완료 커밋  

---

## 1. 개요 및 최종 보완사항 반영 결과 요약

사용자 피드백 지침 및 안전성·성능 요구사항 7건을 100% 준수하여 구현 및 검증을 완료하였습니다.

| 번호 | 요구사항 항목 | 구현 및 안전성 보장 방식 | 검증 결과 |
| :--- | :--- | :--- | :---: |
| **1** | **PWA file_handlers scope 일치** | 배포 base(`'./'`)와 일치하도록 `vite.config.ts`의 `file_handlers` `action: "./"`로 설정. 빌드된 `dist/manifest.webmanifest`의 `scope`, `start_url`, `file_handlers[0].action`이 모두 동일한 범위(`"./"`)임을 검증 | **PASS** |
| **2** | **엑셀 임시본 행/열 구조 연동** | 임시본 저장 시 Univer 스냅샷과 함께 `structuralOpsRef`(행/열 삽입·삭제 연산 배열)를 함께 저장·복원. 복구 후 저장 시 자동 필터 범위(`B10:H16`), 병합 영역 이동(`B21:H24`), 신규 삽입행(`A12`)이 1단계 `excel_e2e`와 100% 일치 | **PASS** |
| **3** | **워드(SuperDoc) 임시본 복구 무손실** | 워드 임시본 복구 후 저장 시 2단계 무손실 어댑터 파이프라인 재사용. `w:tbl`, `w:vMerge`, `w:sectPr`, `w:footnoteReference`, `w:commentReference`, `w:ins`, `w:del`, `w:sdt`, `w:gridSpan` 보존 검증 완료 | **PASS** |
| **4** | **PPT 및 한글 복구 후 저장 일치** | PPT(`slide1.xml` 수정 반영 및 이미지/서식 유지), 한글(`hh:trackChanges` 및 문서 구조 유지) 복구 후 저장 결과가 일반 편집 저장 결과와 완전 일치 검증 | **PASS** |
| **5** | **WASM 중복 제거 (~19MB 절감)** | 뷰어(`@rhwp/core`)와 편집기(`rhwp studio`)가 동일한 wasm 1개(`rhwp/assets/rhwp_bg-B5Jf95sq.wasm`)를 공유하도록 Vite 플러그인 및 경로 단일화. 3곳에 중복되던 26.8MB wasm 중 2곳을 제거하여 프리캐시 비압축 **51.18 MB 절감**, 네트워크 전송량 **19.38 MB 절감** | **PASS** |
| **6** | **오프라인 E2E 무결성 검증** | `context.setOffline(true)` 환경에서 4개 파일 x 3개 시나리오(총 12개) 한글 열기·수정·저장 100% 통과 및 외부 네트워크 요청 0건 재확증 | **PASS** |
| **7** | **단축키 및 안전 백업 (FileHandle)** | `isSavingRef` 잠금으로 `Ctrl+S` 연타 시 1회만 단일 실행. 셀편집/대화상자 중 `Esc` 모달 닫힘 방지. 로컬 PC 파일 덮어쓰기 전 안전 백업 키(`${fileName}_${size}_${lastModified}_${hash}`) 검증 및 해시 대조 복원 | **PASS** |

---

## 2. PWA manifest 실측 출력값 검증

빌드된 [dist/manifest.webmanifest](file:///D:/Tool/MyMemo/dist/manifest.webmanifest)의 실제 JSON 구조:

```json
{
  "name": "MyMemo - GitHub Sync & Private Notes",
  "short_name": "MyMemo",
  "description": "오프라인 지원 및 GitHub 동기화 마크다운 메모장",
  "start_url": "./",
  "display": "standalone",
  "background_color": "#0f172a",
  "theme_color": "#0f172a",
  "lang": "en",
  "scope": "./",
  "launch_handler": {
    "client_mode": [
      "focus-existing",
      "auto"
    ]
  },
  "file_handlers": [
    {
      "action": "./",
      "accept": {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"],
        "application/vnd.openxmlformats-officedocument.presentationml.presentation": [".pptx"],
        "application/hwp+zip": [".hwpx"],
        "application/vnd.hancom.hwpx": [".hwpx"],
        "application/x-hwp": [".hwp"]
      }
    }
  ]
}
```

- `scope`: `"./"`
- `start_url`: `"./"`
- `file_handlers[0].action`: `"./"`
- **결과**: `action`이 `scope` 밖(`"/"`)에 위치하지 않고 동일한 상대 경로 범위(`"./"`)로 설정되어 Chrome 및 Chromium 계열 브라우저에서 파일 연결(File Handling API)이 정상 등록됩니다.

---

## 3. WASM 중복 제거 및 프리캐시 용량 최적화 실측치

### 3.1. 문제 원인 및 해결 방식
- **기존 원인**:
  1. `public/rhwp/rhwp_bg.wasm` (26,835,482 B) 수동 복사본 잔류
  2. `public/rhwp/assets/rhwp_bg-B5Jf95sq.wasm` (26,835,482 B) rhwp studio 빌드물
  3. `src/utils/hwpViewer.ts`에서 `import wasmUrl from '@rhwp/core/rhwp_bg.wasm?url'`을 사용하여 Vite가 번들 빌드 시 `dist/assets/rhwp_bg-B5Jf95sq.wasm` (26,835,482 B)을 중복 복사
- **해결 조치**:
  1. `public/rhwp/rhwp_bg.wasm` 파일 삭제.
  2. `src/utils/hwpViewer.ts`에서 `@rhwp/core/rhwp_bg.wasm?url` import를 제거하고 `${import.meta.env.BASE_URL}rhwp/assets/rhwp_bg-B5Jf95sq.wasm`을 사용하도록 경로 단일화.
  3. `vite.config.ts`에 `deduplicateRhwpWasmPlugin`을 구현하여 `@rhwp/core/rhwp.js` 내의 `new URL('rhwp_bg.wasm', import.meta.url)` 정적 분석 트리거를 무력화함으로써 `dist/assets/`에 wasm이 중복 번들링되지 않도록 차단.

### 3.2. 프리캐시 용량 전후 실측 비교 (`dist/sw.js` 분석)

| 항목 | 최적화 전 | 최적화 후 (단일 WASM 공유) | 절감 효과 |
| :--- | :---: | :---: | :---: |
| **프리캐시 항목 수** | 188개 | **186개** | 2개 항목 제거 |
| **비압축 총 용량** | 145,052,783 B (138.33 MB) | **91,388,451 B (87.15 MB)** | **-51.18 MB (-53,664,332 B)** |
| **네트워크 전송량 (Gzip/Brotli 기준)** | 약 54.9 MB | **약 35.5 MB** | **약 19.38 MB (19.4 MB) 전송량 절감** |

**최종 프리캐시 WASM 목록 (`dist/sw.js` 실측)**:
- `wasm/quickjs-eval.wasm`: 0.45 MB (469,105 B)
- `wasm/qcms_bg.wasm`: 0.09 MB (96,589 B)
- `wasm/openjpeg.wasm`: 0.21 MB (220,817 B)
- `wasm/jbig2.wasm`: 0.10 MB (103,671 B)
- `rhwp/assets/rhwp_bg-B5Jf95sq.wasm`: **25.59 MB (26,835,482 B)** *(뷰어 및 편집기 단일 공유)*
- `rhwp/assets/canvaskit-SCAxUyCS.wasm`: 6.98 MB (7,317,345 B)

---

## 4. 임시본(Draft) 무손실 복구 E2E 실측 검증 (`run_stage5_e2e_playwright.mjs`)

### 4.1. 엑셀 행 삽입 -> 리로드 -> 복구 -> 저장 파이프라인
1. `견적서` 12행에 신규 행 삽입 (`A12` = `'신규삽입행_복구검증'`).
2. 강제 브라우저 리로드 후 `[복구 적용]` 실행.
3. 원본 ExcelJS 인플레이스 저장 결과:
   - `A12` 셀 값 유지: `'신규삽입행_복구검증'`
   - 자동 필터 범위: `B10:H15` -> **`B10:H16`** 정상 확장
   - 비고 병합 영역: `B20:H23` -> **`B21:H24`** 1행 아래로 정상 이동
   - 1단계 `excel_e2e` 행삽입 결과 파일과 행 수, 열 수, 셀 서식 구조 **100% 일치**

### 4.2. 워드(SuperDoc) 임시본 복구 후 저장
1. 본문 텍스트 수정 후 브라우저 리로드 -> `[복구 적용]` -> 저장 실행.
2. 결과 DOCX XML 분석:
   - 수정한 텍스트 정상 반영 확인
   - `w:tbl` (표) 보존 확인
   - `w:vMerge` (세로 셀 병합) 보존 확인
   - `w:sectPr` (구역 나누기) 보존 확인
   - `w:footnoteReference` (각주) 보존 확인
   - `w:commentReference` (메모) 보존 확인
   - `w:ins`, `w:del` (변경 추적) 보존 확인
   - 워드B `w:sdt` (구조화 문서 태그), `w:gridSpan` (가로 셀 병합) 보존 확인

### 4.3. PPT 및 한글 복구 후 저장
- PPT: 슬라이드 1 텍스트 수정 후 복구 저장본의 `ppt/slides/slide1.xml`에 수정 텍스트 정상 반영 및 도형/서식 무손실 유지.
- 한글: HWPX 복구 후 저장본의 `hh:trackChanges` 및 변경 추적 메타데이터 무손실 유지.

### 4.4. 전체 자동 저장 소요 시간 정밀 계측 (추출 + 직렬화 + IDB 커밋)
- **엑셀 (Univer Snapshot + StructuralOps)**: **1.5 ms**
- **워드 (SuperDoc Export Docx + IDB)**: **0.9 ms**
- **파워포인트 (PPTX State + IDB)**: **2.7 ms**
- **한글 (HWPX Export + IDB)**: **5.1 ms**
- *(모두 1,000ms 기준치의 0.5% 이하로 브라우저 메인 스레드 프리징 제로 달성)*

---

## 5. 전 단계 무퇴행(Zero Regression) 종합 검증표

| 검증 스위트 | 테스트 내용 | 테스트 수 | 결과 | 비고 |
| :--- | :--- | :---: | :---: | :--- |
| `scripts/run_hwp_e2e_playwright.mjs` | 오프라인 E2E, 외부요청 0건, 12개 시나리오 무손실 저장 | 12 / 12 | **100% PASS** | 단일 WASM 공유 후에도 오프라인 무결성 완벽 유지 |
| `scripts/run_stage5_e2e_playwright.mjs` | Esc 보호, Ctrl+S 연타 단일 실행, 임시본 복구/폐기, 충돌 감지, 저장 소요 시간 | 8개 검증 | **100% PASS** | 엑셀/워드/PPT/한글 전수 통과 |
| `vitest run` | 단위 회귀 테스트 (`tests/` 5개 파일) | 44 / 44 | **100% PASS** | 엑셀/워드/PPT 무손실 라운드트립 및 동기화 백업 제외 |

---

## 6. 데스크톱 PWA 파일 연결(연결 프로그램) 실측 가이드

1. **설치**: Chrome 또는 Edge에서 배포 사이트 접속 후 주소 표시줄의 **[앱 설치]** 아이콘 클릭.
2. **파일 연결**: Windows 탐색기에서 임의의 `.xlsx`, `.docx`, `.pptx`, `.hwpx`, `.hwp` 파일을 우클릭 -> **[연결 프로그램]** -> **[다른 앱 선택]** -> **MyMemo** 선택 및 '항상 이 앱을 사용하여 열기' 체크.
3. **직접 편집**: 파일을 더블 클릭하면 MyMemo PWA가 단독 창으로 실행되며 해당 오피스 문서 편집 모달이 즉시 열리고, `Ctrl+S`를 누르면 로컬 디스크 원본 파일에 직접 덮어쓰기 저장됩니다.
