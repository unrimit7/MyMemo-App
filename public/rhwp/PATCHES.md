# rhwp 스튜디오 패치 및 확장 내역 (PATCHES.md)

## 버전 및 커밋 정보
- **rhwp 버전**: v0.9.0
- **커밋 해시**: `c9095c5`
- **WASM 엔진**: `public/rhwp/rhwp_bg.wasm` (10,160,971 B)

## 패치 정책 및 구현 내역
1. **Ctrl+S 단축키 연동 정책**:
   - `public/rhwp`의 빌드 산출물(JS/WASM 번들)을 직접 변조하지 않고, `HwpEditorModal.tsx`에서 동일 출처(same-origin) iframe의 `contentWindow`에 이벤트 리스너를 바인딩하여 `Ctrl+S`를 안전하게 캡처합니다.
   - 키 캡처 시 `targetOrigin`을 `window.location.origin`으로 명시하여 `postMessage({ type: 'RHWP_SAVE_SHORTCUT' }, window.location.origin)`를 상위 창에 전달합니다.
   - 수신 측에서는 `event.origin === window.location.origin` 및 `event.source === iframe.contentWindow`를 엄격히 검증하여 다른 창이나 출처로부터의 메시지 위조를 방지합니다.
   - 따라서 향후 `@rhwp/editor` 및 `public/rhwp` 스튜디오를 상위 버전으로 갱신하더라도 산출물 충돌이나 재패치 작업 없이 원활하게 업그레이드가 가능합니다.
