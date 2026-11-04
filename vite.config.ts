import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';
import { VitePWA } from 'vite-plugin-pwa';

// 모바일 브라우저(Samsung Internet, iOS Safari, 모바일 웹뷰 등) 웹 워커 내 ES2024 호환성 완벽 폴리필 플러그인
function patchPdfWorkerPlugin() {
  const polyfill = `
if (typeof Promise !== 'undefined' && !Promise.withResolvers) {
  Promise.withResolvers = function () {
    let resolve, reject;
    const promise = new Promise((res, rej) => {
      resolve = res;
      reject = rej;
    });
    return { promise, resolve, reject };
  };
}
if (typeof Response !== 'undefined' && !Response.prototype.bytes) {
  Response.prototype.bytes = async function () {
    return new Uint8Array(await this.arrayBuffer());
  };
}
if (typeof Blob !== 'undefined' && !Blob.prototype.bytes) {
  Blob.prototype.bytes = async function () {
    return new Uint8Array(await this.arrayBuffer());
  };
}
if (typeof Array !== 'undefined' && !Array.prototype.toSpliced) {
  Array.prototype.toSpliced = function (start, deleteCount, ...items) {
    const copy = Array.from(this);
    copy.splice(start, deleteCount, ...items);
    return copy;
  };
}
if (typeof Set !== 'undefined' && !Set.prototype.intersection) {
  Set.prototype.intersection = function (other) {
    const res = new Set();
    for (const elem of this) {
      if (other && (typeof other.has === 'function' ? other.has(elem) : false)) {
        res.add(elem);
      }
    }
    return res;
  };
}
`;
  return {
    name: 'patch-pdf-worker',
    generateBundle(_options: any, bundle: Record<string, any>) {
      for (const fileName in bundle) {
        if (fileName.includes('pdf.worker')) {
          const item = bundle[fileName];
          let content = '';
          if (item.type === 'asset') {
            content = typeof item.source === 'string' ? item.source : new TextDecoder().decode(item.source);
          } else if (item.type === 'chunk') {
            content = item.code;
          }
          if (content) {
            let patched = polyfill + content;
            patched = patched.replace(
              /return\s+response\.bytes\(\);/g,
              'return typeof response.bytes === "function" ? response.bytes() : response.arrayBuffer().then((b) => new Uint8Array(b));'
            );
            patched = patched.replace(
              /blob\s*=>\s*blob\.bytes\(\)/g,
              'blob => typeof blob.bytes === "function" ? blob.bytes() : blob.arrayBuffer().then((b) => new Uint8Array(b))'
            );
            if (item.type === 'asset') {
              item.source = typeof item.source === 'string' ? patched : new TextEncoder().encode(patched);
            } else {
              item.code = patched;
            }
            console.log(`[patch-pdf-worker] Successfully patched ${fileName} with ES2024 polyfills`);
          }
        }
      }
    },
  };
}

// https://vite.dev/config/
export default defineConfig({
  plugins: [
    patchPdfWorkerPlugin(),
    react(),
    tailwindcss(),
    VitePWA({
      registerType: 'autoUpdate',
      workbox: {
        maximumFileSizeToCacheInBytes: 25 * 1024 * 1024, // 25MB 이하의 WASM 및 워커 파일 오프라인 캐싱 허용
        globPatterns: ['**/*.{js,css,html,ico,png,svg,wasm,mjs}']
      },
      includeAssets: ['favicon.ico', 'favicon.svg', 'favicon-32x32.png', 'favicon-16x16.png', 'apple-touch-icon.png'],
      manifest: {
        name: 'MyMemo - GitHub Sync & Private Notes',
        short_name: 'MyMemo',
        description: '오프라인 지원 및 GitHub 동기화 마크다운 메모장',
        theme_color: '#0f172a',
        background_color: '#0f172a',
        display: 'standalone',
        // 단일 인스턴스 정책: 이미 앱 창이 열려있다면 새 창을 띄우지 않고 기존 창에 포커스
        launch_handler: {
          client_mode: ['focus-existing', 'auto']
        },
        icons: [
          {
            src: 'favicon-32x32.png',
            sizes: '32x32',
            type: 'image/png'
          },
          {
            src: 'pwa-192x192.png',
            sizes: '192x192',
            type: 'image/png',
            purpose: 'any'
          },
          {
            src: 'pwa-512x512.png',
            sizes: '512x512',
            type: 'image/png',
            purpose: 'any'
          },
          {
            src: 'pwa-maskable-512x512.png',
            sizes: '512x512',
            type: 'image/png',
            purpose: 'maskable'
          }
        ]
      }
    })
  ],
  base: './' // 상대 경로로 빌드하여 GitHub Pages 등 어디서나 서브디렉토리 배포 가능
});
