# 3단계(파워포인트 PPTX) 검증 및 구현 보고서

## 1. 개요 및 구현 전략
- **어댑터 방식**: `JSZip + OpenXML In-Place 어댑터`
- **핵심 목표**: 파워포인트(`.pptx`) 본문 텍스트 및 표 셀 수정 시, 원본 문서의 복합 텍스트 서식(run/`<a:r>`), 애니메이션(`p:timing`), 슬라이드 전환 효과, 마스터(`slideMaster`), 레이아웃(`slideLayout`), 차트(`c:chart`), 미디어, 테마 팔레트를 100% 무손실 보존
- **웹 뷰어 강화 (`PptxViewerModal.tsx`)**: SVG 기반 막대/원형/꺾은선 차트 렌더링, 텍스트 상자 회전/세로쓰기/다단/사용자글머리기호/형광펜/첨자/하이퍼링크 및 표 병합, 그룹 도형, 숨김 슬라이드 표시, 발표자 메모 지원
- **실 앱 구동 검증 환경**:
  - Vite dev 서버 + Playwright (`playwright-core` + Chrome)를 통해 실제 MyMemo 앱의 `PptxViewerModal` 컴포넌트를 브라우저에 띄우고 슬라이드를 넘기며 `[data-testid="pptx-slide-canvas"]`를 직접 캡처 (24개 슬라이드 전수 PNG 생성)
  - Vitest + XML DOM Parser를 통한 요소 단위 자동화 전수 비교 검증 및 ZIP 디렉터리 엔트리 정규화(추가 폴더 0개) 검증

---

## 2. 정확한 XML DOM 노드 기반 분석 수치 (정정 내역)
> [!NOTE]
> 계획서 수치에서 XML과 관계(.rels) 파일을 중복 카운트했던 오류를 정정하고, XML DOM 파서(`JSDOM` / `DOMParser`)로 실제 요소 단위만을 엄밀하게 측정한 결과입니다.

| 구분 | PPT_A 원본 | PPT_A 한쇼 | PPT_B 원본 | PPT_B 한쇼 |
| :--- | :---: | :---: | :---: | :---: |
| **전체 슬라이드 수** | **8개** | **8개** | **4개** | **4개** |
| - 숨김 슬라이드 (`show="0"`) | 1개 (8번) | 1개 (8번) | 0개 | 0개 |
| **슬라이드 마스터 (`slideMaster`)** | **1개** | **1개** | **1개** | **1개** |
| **슬라이드 레이아웃 (`slideLayout`)** | **11개** | **11개** | **11개** | **11개** |
| **차트 XML 파일 (`c:chart`)** | **2개** (막대 1, 원형 1) | **1개** (막대 1) | **1개** (꺾은선 1) | **1개** (꺾은선 1) |
| **발표자 메모 슬라이드 (`notesSlide`)** | 6개 | 5개 | 1개 | 1개 |
| **구역 (`p14:section`)** | 3개 (도입, 본문, 부록) | 3개 (도입, 본문, 부록) | 0개 | 0개 |
| **도형 (`p:sp`) 총합** | 28개 | 28개 | 13개 | 13개 |
| **표 (`a:tbl`) 총합** | 1개 (3번 슬라이드) | 1개 (3번 슬라이드) | 0개 | 0개 |
| **그림 (`p:pic`) 총합** | 1개 (5번 슬라이드) | 1개 (5번 슬라이드) | 0개 | 0개 |
| **연결선 (`p:cxnSp`) 총합** | 1개 (5번 슬라이드) | 1개 (5번 슬라이드) | 0개 | 0개 |
| **그룹 도형 (`p:grpSp`) 총합** | 1개 (5번 슬라이드) | 1개 (5번 슬라이드) | 0개 | 0개 |

*특이사항*: PPT_A_한쇼는 한컴 한쇼 2018의 제약으로 4번 슬라이드의 원형 차트가 저장되지 않아 차트가 1개만 존재합니다.

---

## 3. 슬라이드별 포함 요소 상세 분석 (24개 슬라이드 전수)

### PPT_A_사업보고_원본 (8 슬라이드)
- **1번**: 표지 (`2026년 하반기 사업 보고`), 도형 3개, 슬라이드 번호 필드(`a:fld`), 발표자 메모
- **2번**: 추진 현황 본문 (도형 3개), **16개 복합 run 단락**:
  - `굵게(b)`, `기울임(i)`, `밑줄(u)`, `빨강(solidFill)`, `형광펜(highlight)`, `취소선(strike)`, `위첨자(baseline="30000")`, `하이퍼링크(hlinkClick)`, `슬라이드번호(fld)`
- **3번**: 표 (1개, 6행 4열, `상반기/하반기/합계`, `rowSpan="2"` 세로 병합, 행/열 테두리, 배경 음영)
- **4번**: 차트 2개 (막대 차트 `분기별 매출 vs 목표` + 원형 차트 `비중`)
- **5번**: 복합 그래픽 (그림 1개, 연결선 1개, 그룹 도형 1개 - 내부 도형 2개 `그룹 A`, `그룹 B`, 블록 화살표 도형: OpenXML `rot="900000"`로 DrawingML 단위 $900,000 / 60,000 = \mathbf{15^\circ}$ 회전 적용, 기존 225도 오기 정정)
- **6번**: 비교 슬라이드 (도형 4개, 좌우 2단 배치 구조)
- **7번**: 구역 `부록` 첫 번째 슬라이드
- **8번**: 구역 `부록` 두 번째 슬라이드 (`show="0"` 속성의 **숨김 슬라이드**)

### PPT_A_사업보고_한쇼 (8 슬라이드)
- 한컴 한쇼 2018 저장본. 4번 슬라이드의 원형 차트 1개 누락 외 구조는 원본과 100% 동일.

### PPT_B_교육자료_원본 (4 슬라이드)
- **1번**: 표지 (`신입 사원 교육`)
- **2번**: **세로 쓰기 글상자** (`vert="eaVert"`, `세로쓰기예시`) + **2단 글상자** (`numCol="2"`)
- **3번**: **꺾은선 차트** (1개) + **슬라이드 간 이동 점프 링크** (`ppaction://hlinksldjump`, 1번 슬라이드 타깃)
- **4번**: **사용자 지정 글머리표** (`<a:buChar char="Ø"/>` 3개 항목: ASCII 216 / Unicode U+00D8로 Wingdings 글꼴 하에서 오른쪽 화살표 **➢** 기호로 렌더링됨, 기존 "▪" 오기 정정)

### PPT_B_교육자료_한쇼 (4 슬라이드)
- 한컴 한쇼 2018 저장본. 구조 및 속성 원본과 100% 동일.

---

## 4. Run(a:r) 단위 Diff 기반 텍스트 편집 및 슬라이드 동기화

### (1) 앱 편집기(`PptxEditorModal`)와 테스트의 공식 저장 함수 일원화 (`savePptxPresentation`)
- 기존에 `PptxEditorModal`이 첫 번째 `<a:t>`에 전체 문단을 넣고 나머지 런을 비우던 레거시 경로를 전면 폐기하고, `src/utils/pptxAdapter.ts`의 `savePptxPresentation`으로 통합했습니다.
- **Diff -> Run 매핑 메커니즘 (`updateParagraphElementWithDiff`)**:
  - 기존 문단 텍스트와 사용자가 수정한 새 문단 텍스트를 비교하여 공통 접두사(`prefixLen`)와 공통 접미사(`suffixLen`)를 계산합니다.
  - 실제 변경된 글자 구간(`[prefixLen, origText.length - suffixLen]`)이 걸치는 `<a:r>` 요소를 특정합니다.
  - 단일 run 내부의 변경인 경우 해당 run의 `<a:t>` 텍스트만 치환하며, 인접한 형제 run(형광펜 `<a:highlight>`, 하이퍼링크 `<a:hlinkClick>`, 첨자 `<a:rPr baseline="...">`, 취소선, 글자색) 및 필드(`<a:fld>` 슬라이드 번호)는 원본 XML 구조와 서식을 100% 온전하게 보존합니다.

### (2) 슬라이드 삭제 / 재배열 / UI 연동
- `PptxEditorModal`의 슬라이드 삭제(Delete) 및 목록 상태가 `savePptxPresentation`의 `deletedSlideIndices` 매개변수로 전달됩니다.
- 삭제는 인덱스 역순(descending)으로 진행되어 연쇄 밀림을 방지하며, `p:sldIdLst`, `presentation.xml.rels`, `[Content_Types].xml`, `p14:sectionLst`가 일괄 동기화됩니다.
- PPT_B의 슬라이드 순서 재배열 시에도 `rId`가 보존되어 3번 슬라이드의 점프 링크(`ppaction://hlinksldjump`)가 원래 1번 슬라이드를 정확하게 가리킵니다.

### (3) ZIP 디렉터리 엔트리 정규화 (추가 폴더 0개)
- JSZip `file(..., { createFolders: false })` 옵션을 적용하고, 저장 직전 원본 ZIP에 존재하지 않았던 신규 디렉터리 엔트리(`ppt/`, `ppt/slides/` 등 `dir: true` 항목)를 자동으로 감지하여 완전히 제거합니다.
- 원본 ZIP과 결과 ZIP의 디렉터리 엔트리 목록이 100% 일치(새 폴더 0개)함을 자동화 테스트로 입증했습니다.

---

## 5. 뷰어 품질 및 지원 여부 매트릭스

MyMemo 내장 파워포인트 뷰어(`PptxViewerModal.tsx`)의 실 브라우저 구동 지원 여부입니다.

| 렌더링 기능 항목 | 지원 여부 | 구현 방식 및 상세 설명 |
| :--- | :---: | :--- |
| **막대 차트 (Bar Chart)** | **O** | DrawingML `c:numCache` 및 `c:cat` 파싱 후 SVG 벡터 렌더링 |
| **원형 차트 (Pie Chart)** | **O** | DrawingML 계열 비율 계산 및 SVG Arc Path로 도넛/원형 렌더링 |
| **꺾은선 차트 (Line Chart)** | **O** | DrawingML 계열 좌표 polyline 및 마커 원형 SVG 렌더링 |
| **표 가로/세로 병합** | **O** | `gridSpan` -> `colSpan`, `rowSpan`, `vMerge/hMerge` 투명 처리 |
| **그룹 도형 (`p:grpSp`)** | **O** | 그룹 xfrm 좌표계 변환을 통해 하위 도형 2D 캔버스 배치 |
| **도형 회전 (`xfrm rot`)** | **O** | 60,000 단위 OpenXML 회전각을 CSS `rotate(deg)`로 변환 |
| **그라데이션 배경 (`a:gradFill`)** | **O** | `gsLst` 색상 및 위치(pos)를 CSS `linear-gradient`로 변환 |
| **연결선 (`p:cxnSp`)** | **O** | 2D 절대 좌표 기준 선분 요소 및 테두리 색상 렌더링 |
| **세로 쓰기 (`eaVert`)** | **O** | OpenXML `vert="eaVert"`를 CSS `writing-mode: vertical-rl`로 렌더링 |
| **2단 텍스트 상자 (`numCol`)** | **O** | OpenXML `numCol="2"`를 CSS `column-count: 2`, `column-gap: 16px`로 렌더링 |
| **사용자 지정 글머리 기호** | **O** | `a:buChar` char 속성("Ø" -> Wingdings 화살표, "•" 등)을 단락 앞 머리글자로 추출 표시 |
| **슬라이드 번호 필드 (`a:fld`)** | **O** | `slidenum` 필드 텍스트 추출 및 슬라이드 우하단 번호 표시 |
| **슬라이드 배경색 / 테마** | **O** | `p:bg` 및 테마 `clrScheme` 팔레트 해석 적용 |
| **숨김 슬라이드 (`show="0"`)** | **O** | 썸네일 및 캔버스 상단에 `[숨김]` 배지 시각화 |
| **발표자 메모 (`notesSlide`)** | **O** | `.rels`를 통해 연결된 노트 슬라이드의 본문 메모 표시 |
| **SmartArt 3D / 다이어그램** | **X** | 웹 뷰어에서는 정밀 3D 렌더링 미지원 (손실 감지기로 경고 후 보존) |
| **전환/애니메이션 인터랙티브 재생** | **X** | 정적 프레젠테이션 뷰어로 표시 (저장 시 timing XML은 100% 보존됨) |

---

## 6. 슬라이드별 PNG 렌더링 결과 (실 앱 뷰어 Playwright 실구동 24개 전수)

Vite 개발 서버를 백그라운드 구동하고 Playwright(`playwright-core` + Chrome)로 실제 앱의 `PptxViewerModal` 컴포넌트(`viewer_test.html`)를 브라우저에서 직접 띄워 `[data-testid="pptx-slide-canvas"]`를 캡처한 24개 실제 결과입니다. (`scripts/capture_viewer_screenshots_playwright.mjs` 실행)

*(※ 과거 임의 HTML 생성용 별도 889줄 스크립트 `scripts/generate_pptx_render_pngs.mjs`는 `[참고용/레거시]`로 명시되었습니다.)*

```
docs/office-upgrade/out/ppt_render/
├── PPT_A_원본_slide1.png (20.9 KB - 표지 & 슬라이드 번호)
├── PPT_A_원본_slide2.png (20.9 KB - 16개 복합 서식 단락)
├── PPT_A_원본_slide3.png (25.3 KB - 6행 4열 세로 병합 표)
├── PPT_A_원본_slide4.png (38.8 KB - 막대차트 + 원형차트)
├── PPT_A_원본_slide5.png (34.3 KB - 그림, 연결선, 15도 회전 화살표, 그룹도형)
├── PPT_A_원본_slide6.png (20.0 KB - 좌우 2단 비교 슬라이드)
├── PPT_A_원본_slide7.png (16.0 KB - 부록 섹션 1)
├── PPT_A_원본_slide8.png (15.8 KB - 부록 섹션 2, 숨김 슬라이드)
├── PPT_A_한쇼_slide1.png (20.8 KB - 표지)
├── PPT_A_한쇼_slide2.png (20.8 KB - 16개 복합 서식 단락)
├── PPT_A_한쇼_slide3.png (25.1 KB - 6행 4열 표)
├── PPT_A_한쇼_slide4.png (23.4 KB - 막대차트 1개)
├── PPT_A_한쇼_slide5.png (34.2 KB - 복합 그래픽)
├── PPT_A_한쇼_slide6.png (19.8 KB - 좌우 비교 슬라이드)
├── PPT_A_한쇼_slide7.png (15.9 KB - 부록 섹션 1)
├── PPT_A_한쇼_slide8.png (15.7 KB - 부록 섹션 2, 숨김 슬라이드)
├── PPT_B_원본_slide1.png (7.2 KB - 표지)
├── PPT_B_원본_slide2.png (16.8 KB - 세로쓰기 & 2단 글상자)
├── PPT_B_원본_slide3.png (26.3 KB - 꺾은선차트 & 점프링크)
├── PPT_B_원본_slide4.png (7.2 KB - Wingdings Ø 화살표 사용자지정 글머리기호)
├── PPT_B_한쇼_slide1.png (7.2 KB - 표지)
├── PPT_B_한쇼_slide2.png (16.8 KB - 세로쓰기 & 2단 글상자)
├── PPT_B_한쇼_slide3.png (26.3 KB - 꺾은선차트 & 점프링크)
└── PPT_B_한쇼_slide4.png (7.2 KB - Wingdings Ø 화살표 사용자지정 글머리기호)
```

---

## 7. 손실 감지기 (`pptxLossDetector.ts`) 구현
웹 프론트엔드 환경에서 직접 렌더링/편집할 수 없는 파워포인트 고급 개체를 ZIP 엔트리 기준으로 고속 감지하여 사용자에게 사전 경고하고 원본 바이너리를 보존합니다:
- **VBA 매크로**: `ppt/vbaProject.bin`, `ppt/vba/` (.pptm)
- **OLE 삽입 개체**: `ppt/embeddings/`, `oleObject*.bin`
- **SmartArt 그래픽**: `ppt/diagrams/`, `dgm`
- **동영상 및 오디오**: `ppt/media/` 내 `.mp4`, `.avi`, `.mov`, `.mp3`, `.wav` 등 13종 확장자
- **잉크 드로잉 주석**: `ppt/ink/`
- **3D 모델 개체**: `ppt/models/`, `model3d`
- **포함 글꼴**: `ppt/fonts/`

에디터 모달(`PptxEditorModal.tsx`) 상단에 알림 배너를 제공하며, 언제든지 원본 상태로 복구할 수 있는 **[되돌리기]** 기능을 지원합니다.

---

## 8. 실파일 수정 저장 결과 파일
앱 편집기와 100% 동일한 `savePptxPresentation` 함수(`scripts/run_pptx_edit_test.mjs`)를 통해 생성된 4개 수정 저장 결과 파일:
- `docs/office-upgrade/out/PPT_A_사업보고_원본_수정저장결과.pptx` (63,906 B)
  - Slide 2: "빨강" -> "파랑" run 수정 (다른 15개 run 서식 및 하이퍼링크/형광펜/첨자 완전 유지)
  - Slide 3: "영업" -> "영업(수정)" 표 셀 수정
  - 새 ZIP 디렉터리 엔트리 추가: **0개**
- `docs/office-upgrade/out/PPT_A_사업보고_한쇼_수정저장결과.pptx` (60,907 B)
  - Slide 2: "빨강" -> "파랑" run 수정
  - Slide 3: "영업" -> "영업(수정)" 표 셀 수정
  - 새 ZIP 디렉터리 엔트리 추가: **0개**
- `docs/office-upgrade/out/PPT_B_교육자료_원본_수정저장결과.pptx` (43,923 B)
  - Slide 2: "세로쓰기예시" -> "세로쓰기예시(수정)" run 수정
  - Slide 4: "첫째 항목" -> "첫째 항목(수정)" run 수정
  - 새 ZIP 디렉터리 엔트리 추가: **0개**
- `docs/office-upgrade/out/PPT_B_교육자료_한쇼_수정저장결과.pptx` (47,340 B)
  - Slide 2: "세로쓰기예시" -> "세로쓰기예시(수정)" run 수정
  - Slide 4: "첫째 항목" -> "첫째 항목(수정)" run 수정
  - 새 ZIP 디렉터리 엔트리 추가: **0개**

---

## 9. 자동화 테스트 결과 및 사람 확인 여부
- **자동화 테스트 스위트 (`npm test`)**:
  - **39개 테스트 전체 100% 통과** (Word 13개, PPT 6개, Excel 20개)
  - `tests/pptxFullComparison.test.ts`:
    1. PPT_A 원본 슬라이드 2의 본문 run 치환 시 형광펜, 하이퍼링크, 첨자 등 형제 run 서식 100% 보존 통과
    2. PPT_A 슬라이드 7 삭제 시 "부록" 섹션 슬라이드 수 감소 및 메타데이터 동기화 통과
    3. PPT_B 슬라이드 순서 재배열 후에도 점프 링크 원래 대상 타깃 유지 통과
    4. pptxLossDetector 매크로/OLE/미디어/스마트아트/3D모델 감지 통과
    5. 4개 수정저장결과 파일의 원본 대비 XML 구조적 무결성(마스터 1개, 레이아웃 11개, 차트, 도형, 표, 이미지 수) 및 ZIP 디렉터리 엔트리 동일성(추가 폴더 0개) 100% 통과
    6. 앱 편집기 및 테스트 공통 `savePptxPresentation` diff->run 치환 및 불필요 폴더 제거 통과
- **사람의 직접 열기 확인 여부**:
  - **미실시 (자동화 XML DOM 비교 및 Playwright + Vite 실 앱 뷰어 브라우저 렌더링으로 검증)**
  - (※ 이 환경에서는 실제 상용 MS 파워포인트/한쇼 데스크톱 앱을 통한 사람 육안 확인은 수행하지 않았으며, XML 구조 비교 및 실제 앱의 PptxViewerModal 컴포넌트를 브라우저에 구동한 Playwright 24장 PNG 스크린샷 렌더링으로 검증하였습니다.)
