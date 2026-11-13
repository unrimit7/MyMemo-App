/**
 * 기능 스위치
 *
 * PPTX_EDITING_ENABLED
 *   현재 PPT 편집기(PptxEditorModal)는 슬라이드별 제목/글머리 텍스트만 고치는 양식형 편집기라
 *   PowerPoint 사용자에게 익숙하지 않다(위치·크기·도형·이미지 편집 불가). 2026-11-13 사용자 결정으로
 *   PPT는 미리보기 + 슬라이드 쇼만 제공하고 편집 진입점을 모두 숨긴다.
 *   (메모 카드 '직접 편집' 버튼, 큰 화면 뷰어 '편집' 버튼, 탐색기 파일 연결로 연 .pptx)
 *   PowerPoint와 비슷한 편집기(ONLYOFFICE 브라우저판, BetterOffice 등)를 검증해 도입하면 true로 되돌리거나 교체한다.
 */
export const PPTX_EDITING_ENABLED = false;
