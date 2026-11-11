# MyMemo

MyMemo는 오프라인 PWA 지원 및 GitHub 양방향 동기화를 제공하는 개인용 마크다운 메모 및 오피스 문서 편집 앱입니다.

---

## 오피스 문서 편집 엔진 및 오픈소스 라이선스 고지 (Open Source Licenses)

MyMemo는 사용자 문서의 서식 보존 및 로컬 오프라인 편집을 위해 다음 오픈소스 라이브러리를 사용합니다:

1. **SuperDoc** (`superdoc`)
   - **라이선스**: GNU Affero General Public License v3.0 ([AGPL-3.0](https://www.gnu.org/licenses/agpl-3.0.html))
   - **용도**: 워드(.docx) 문서의 인플레이스(In-place) 서식 보존 및 WYSIWYG 편집 엔진
   - **소스 코드 저장소**: [https://github.com/vlad-t/superdoc](https://github.com/vlad-t/superdoc)

2. **Univer** (`@univerjs/*`)
   - **라이선스**: Apache License 2.0 ([Apache-2.0](https://www.apache.org/licenses/LICENSE-2.0))
   - **용도**: 엑셀(.xlsx) 스프레드시트 뷰어 및 사용자 인터랙션 UI

3. **ExcelJS** (`exceljs`)
   - **라이선스**: MIT License ([MIT](https://opensource.org/licenses/MIT))
   - **용도**: 엑셀(.xlsx) 파일의 셀 값, 수식, 서식, 열 너비/행 높이 라운드트립 보존 파서 및 직렬화 엔진

4. **SheetJS CE** (`xlsx`)
   - **라이선스**: Apache License 2.0 ([Apache-2.0](https://www.apache.org/licenses/LICENSE-2.0))
   - **용도**: 경량 스프레드시트 파싱 및 레거시 폴백

5. **rhwp**
   - **라이선스**: MIT License ([MIT](https://opensource.org/licenses/MIT))
   - **용도**: 한컴 한글(.hwp) 오프라인 WebAssembly 뷰어/편집기

---

## PWA & 오프라인 지원
- Service Worker(Workbox)를 통해 외부 서버 통신 없이 100% 로컬 오프라인 환경에서 동작합니다.
- GitHub 동기화는 사용자의 명시적 동기화 시점에만 수행됩니다.
