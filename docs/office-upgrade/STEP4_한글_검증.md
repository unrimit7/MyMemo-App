# 4단계(한글 HWP/HWPX) 검증 및 구현 보고서

## 1. 개요 및 구현 전략
- **구현 방식**: `rhwp v0.9.0 완전 오프라인 셀프호스팅 + 인플레이스 후처리 어댑터(hwpxAdapter.ts)`
- **버전 및 커밋 고정**:
  - `rhwp`: v0.9.0 (커밋 해시: `c9095c5`)
  - WASM 엔진: `rhwp_bg.wasm` (10,160,971 B)
  - 글꼴: OFL / GUST 라이선스 번들 (26,835,480 B gzip 10,220,890 B)
  - 버전 표기: `public/rhwp/VERSION`에 영구 기록
- **완전 오프라인 무결성**:
  - Playwright `context.setOffline(true)` 환경에서 HWP/HWPX 로드, 편집, 저장 전 과정 수행
  - 외부 도메인 네트워크 요청 **0건** (100% 셀프호스팅 증명)
- **핵심 목표 달성**:
  - 무변경 저장 시 원본 바이트 100% 동일 유지 (재직렬화 원천 차단)
  - HWPX 편집 시 텍스트가 바뀐 문단만 `hp:linesegarray`를 선별 제거하여 텍스트 겹침 방지 및 줄 배치 엔진 자동 재계산 유도
  - 미수정 문단의 변경 추적 마크업(`hp:deleteBegin`, `hp:insertBegin`) 및 `Contents/header.xml`의 `<hh:trackChanges>`, `<hh:trackChangeAuthors>` 완벽 보존
  - HWP 5.0 CFB 내보내기 시 `ViewText` 스트림 부재 시 `FileHeader.Properties`의 bit 14(`0x4000`)를 자동 클리어하여 데스크톱 한글 열기 호환성 100% 확보

---

## 2. 실파일 4개 샘플 구조 분석

| 파일명 | 포맷 | 파일 크기 | 작성 프로그램 | 특징 및 주요 요소 |
| :--- | :---: | :---: | :---: | :--- |
| **한글C_업무계획_원본.hwpx** | HWPX | 14,637 B | python-hwpx | 3페이지, 표 1개(7행 5열), 변경 추적 2건(삭제 1, 삽입 1), linesegarray 없음 |
| **한글C_업무계획_한글.hwpx** | HWPX | 52,043 B | 한글 2018 | 3페이지, 표 1개(7행 5열), 변경 추적 4건(삭제 2, 삽입 2), linesegarray 51개, 덧말/누름틀/체크박스 |
| **한글C_업무계획_원본.hwp** | HWP | 15,872 B | python-hwpx | 3페이지, 표 1개, 변경 추적 포함, FileHeader v5.1.1.0 |
| **한글C_업무계획_한글.hwp** | HWP | 52,224 B | 한글 2018 | 3페이지, 표 1개, 변경 추적 포함, FileHeader v5.1.0.1, ViewText 스트림 포함 |

---

## 3. 주요 기술 해결 내역

### 1) 데스크톱 한글 2018의 HWP 열기 오류 ("%1을(를) 찾지 못했습니다") 원인 규명 및 해결
- **문제**: rhwp에서 수정 저장한 `.hwp` 파일 4개가 한글 2018에서 더블클릭이나 파일 열기 시 `"%1을(를) 찾지 못했습니다"` 오류와 함께 열리지 않음.
- **원인 분석**:
  - 원본 `.hwp`는 `FileHeader.Properties`의 bit 14(`0x4000`, 변경 추적 문서 플래그)가 켜져(`0x4001`) 있고 `ViewText/Section*` 스트림이 존재함.
  - rhwp `exportHwp`는 변경 내용을 본문(`/BodyText/Section*`)에 수락·반영하여 내보내면서 `/ViewText` 스트림을 출력하지 않으나, `FileHeader.Properties`는 원본의 `0x4001`을 그대로 유지함.
  - 데스크톱 한글 2018은 bit 14가 켜져 있으면 필수적으로 `/ViewText` 스토리지를 탐색하며, 스트림이 없으면 열기를 중단함.
- **해결책 (`postProcessHwp`)**:
  - HWP 5.0 CFB 버퍼에서 `ViewText` 스트림(`Buffer.from('ViewText', 'utf16le')`) 존재 여부를 판별.
  - `ViewText` 스트림이 없으면 `FileHeader.Properties`의 bit 14(`0x4000`)를 해제(`props &= ~0x4000`)하여 일반 문서(`0x0001`)로 플래그를 정정.
  - 수정 저장된 4개 `.hwp` 파일 모두 한글 2018에서 오류 없이 정상 열람됨을 실측 확인.

### 2) 데스크톱 한글 2018의 HWPX 변경 추적 표시 누락 원인 규명 및 해결
- **문제**: `.hwpx` 수정 저장본에서 `Contents/section0.xml`의 `hp:deleteBegin`/`hp:insertBegin` 마크업은 보존되었으나, 한글 2018에서 취소선/밑줄이 표시되지 않고 두 문장이 연속된 일반 텍스트로 합쳐져 표시됨.
- **원인 분석**:
  - `Contents/header.xml`의 `<hh:refList>`에서 `<hh:trackChangeAuthors>`(작성자 메타데이터: 테스트작성자 id=1, JUNG id=2)가 누락되었고, `<hh:trackChanges>`가 `<hh:refList>` 바깥에 배치되어 있었음.
  - 한글 2018은 `authorID`가 가리키는 작성자가 `trackChangeAuthors`에 없으면 변경 추적 렌더링 스타일을 비활성화함.
- **해결책 (`postProcessHwpx`)**:
  - `Contents/header.xml`의 `<hh:refList>` 내부 마지막 위치에 원본의 `<hh:trackChanges>`와 `<hh:trackChangeAuthors>`를 원본 순서 및 계층 그대로 복원.
  - `header.xml`의 `<hh:trackchageConfig flags="57"/>` 문서 설정 유지.
  - 한글 2018에서 작성자별 취소선/밑줄 변경 추적이 원본과 동일하게 완벽 렌더링됨을 실측 확인.

### 3) 텍스트 빈도 맵 기반 문단 매칭 및 linesegarray 선별 제거
- 직계 문단 텍스트(`hp:run > hp:t`) 빈도 맵(`Map<string, number>`)을 구성하여 구역/헤더/문단 순서 불일치를 원천 방지.
- 실제 텍스트가 바뀐 문단만 `linesegarray`를 제거하고, 미수정 문단(51개 중 50개)의 `linesegarray`는 100% 원본 바이트로 보존.

---

## 4. E2E Playwright 12개 시나리오 전수 검증 결과

실제 브라우저 환경에서 Vite 개발 서버 기동 후 Playwright를 통해 4개 샘플 파일 x 3개 시나리오를 전수 수행한 실측 결과입니다.

| 순번 | 대상 파일 | 검증 시나리오 | 출력 파일명 | 원본 크기 | 저장 크기 | 결과 |
|:---:|:---|:---|:---|:---:|:---:|:---:|
| 1 | 한글C_업무계획_원본.hwpx | 무변경 저장 | `한글C_업무계획_원본_무변경저장.hwpx` | 14,637 B | 14,637 B | **PASS (바이트 100% 일치)** |
| 2 | 한글C_업무계획_원본.hwpx | 문단 1곳 수정 | `한글C_업무계획_원본_문단수정저장.hwpx` | 14,637 B | 14,358 B | **PASS (TC 2건, TCA 1건 복원)** |
| 3 | 한글C_업무계획_원본.hwpx | 표 셀 1곳 수정 | `한글C_업무계획_원본_표수정저장.hwpx` | 14,637 B | 14,354 B | **PASS (TC 2건, TCA 1건 복원)** |
| 4 | 한글C_업무계획_한글.hwpx | 무변경 저장 | `한글C_업무계획_한글_무변경저장.hwpx` | 52,043 B | 52,043 B | **PASS (바이트 100% 일치)** |
| 5 | 한글C_업무계획_한글.hwpx | 문단 1곳 수정 | `한글C_업무계획_한글_문단수정저장.hwpx` | 52,043 B | 44,728 B | **PASS (TC 4건, TCA 2건 복원)** |
| 6 | 한글C_업무계획_한글.hwpx | 표 셀 1곳 수정 | `한글C_업무계획_한글_표수정저장.hwpx` | 52,043 B | 44,730 B | **PASS (TC 4건, TCA 2건 복원)** |
| 7 | 한글C_업무계획_원본.hwp | 무변경 저장 | `한글C_업무계획_원본_무변경저장.hwp` | 15,872 B | 15,872 B | **PASS (바이트 100% 일치)** |
| 8 | 한글C_업무계획_원본.hwp | 문단 1곳 수정 | `한글C_업무계획_원본_문단수정저장.hwp` | 15,872 B | 16,384 B | **PASS (props: 0x1, bit14 해제)** |
| 9 | 한글C_업무계획_원본.hwp | 표 셀 1곳 수정 | `한글C_업무계획_원본_표수정저장.hwp` | 15,872 B | 16,384 B | **PASS (props: 0x1, bit14 해제)** |
| 10 | 한글C_업무계획_한글.hwp | 무변경 저장 | `한글C_업무계획_한글_무변경저장.hwp` | 52,224 B | 52,224 B | **PASS (바이트 100% 일치)** |
| 11 | 한글C_업무계획_한글.hwp | 문단 1곳 수정 | `한글C_업무계획_한글_문단수정저장.hwp` | 52,224 B | 46,080 B | **PASS (props: 0x1, bit14 해제)** |
| 12 | 한글C_업무계획_한글.hwp | 표 셀 1곳 수정 | `한글C_업무계획_한글_표수정저장.hwp` | 52,224 B | 46,080 B | **PASS (props: 0x1, bit14 해제)** |

- **오프라인 검증**: `context.setOffline(true)` 시 외부 통신 0건 확인.
- **경고 모달 검증**: 변경 추적 문서 수정 시 `변경 추적이 걸린 문단을 직접 수정하면 그 문단의 변경 추적이 일반 텍스트로 합쳐질 수 있습니다` 안내 및 `[취소]`/`[그래도 저장]` 정상 작동 확인.

---

## 5. 한글 2018 데스크톱 프로그램 직접 확인 (사용자 확인 완료, 2026-11-12)

사용자가 실제 Windows 환경의 한글 2018 데스크톱 프로그램에서 생성 결과 파일을 직접 열람하여 최종 검증한 내역입니다:

1. **수정 저장 `.hwp` 4개 파일**:
   - 모두 한글 2018에서 정상적으로 열림 (`"%1을(를) 찾지 못했습니다"` 오류 완벽 해결).
   - `FileHeader.Properties`가 `0x0001`(bit 14 해제)로 일관되게 기록됨을 확인.
2. **`한글C_업무계획_한글_문단수정저장.hwpx`**:
   - "6. 수정 이력" 구역의 변경 추적 내용이 원본과 동일하게 **취소선 및 밑줄(작성자별 색상 매핑)**로 완벽하게 표시됨.
   - `Contents/header.xml`의 `hh:refList` 내에 `<hh:trackChanges>`와 `<hh:trackChangeAuthors>`가 복원되어 작성자 매핑이 성공적으로 동작함.
3. **`한글C_업무계획_한글_표수정저장.hwpx`**:
   - 경고 팝업 없이 정상 열림.
   - 표 셀 내용("영업" → "영업기획") 수정 반영 확인.
   - 표, 셀 병합, 수식, 덧말, 누름틀, 체크박스 등 한글 고급 서식이 모두 정상 유지됨.
4. **`한글C_업무계획_원본_표수정저장.hwp`**:
   - 열기 시 나타나는 "상위 버전에서 작성한 문서" 경고는 원본 샘플(python-hwpx 생성본, FileHeader v5.1.1.0)의 버전 표기를 그대로 유지한 결과로 정상 동작임을 확인.
   - 한글 2018 저장본 계열(`한글C_업무계획_한글.hwp`)은 v5.1.0.1이 유지되어 경고 없이 열림.
5. **날짜 필드 동작**:
   - 날짜 필드는 한글 프로그램이 열람 시 시스템 현재 날짜로 다시 계산하여 변경 추적으로 기록하는 한글 고유의 정상 동작임을 확인.

---

## 6. rhwp 업스트림 이슈 등록 권고안

오픈소스 생태계 기여를 위해 정리한 2건의 기술 이슈 리포트입니다 (등록 여부는 사용자가 결정):

### Issue 1: [Bug] exportHwp retains FileHeader bit 14 (tracked changes) without generating ViewText stream, breaking desktop Hancom Office opening ("%1을(를) 찾지 못했습니다")
- **현상**: 변경 추적이 포함된 HWP 문서를 `exportHwp()`로 내보낼 때 `ViewText` 스트림이 없는데도 `FileHeader.Properties`의 bit 14(`0x4000`)가 켜져 있어 데스크톱 한글에서 열리지 않는 문제.
- **해결책**: `exportHwp` 시 `ViewText` 스트림을 출력하지 않는 경우 `FileHeader.Properties &= ~0x4000` 처리.

### Issue 2: [Bug] exportHwpx omits <hh:trackChanges> and <hh:trackChangeAuthors> from Contents/header.xml
- **현상**: HWPX 문서를 `exportHwpx()`로 내보낼 때 `Contents/header.xml`의 `<hh:refList>`에서 `<hh:trackChanges>`와 `<hh:trackChangeAuthors>`가 누락되어 변경 추적 렌더링이 일반 텍스트로 합쳐지는 문제.
- **해결책**: `exportHwpx` 시 `<hh:refList>` 내 변경 추적 태그 및 작성자 목록 보존.
