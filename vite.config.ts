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
if (typeof Map !== 'undefined') {
  if (!Map.prototype.getOrInsertComputed) {
    Map.prototype.getOrInsertComputed = function (key, callback) {
      if (this.has(key)) return this.get(key);
      const val = callback(key);
      this.set(key, val);
      return val;
    };
  }
  if (!Map.prototype.getOrInsert) {
    Map.prototype.getOrInsert = function (key, defaultVal) {
      if (this.has(key)) return this.get(key);
      this.set(key, defaultVal);
      return defaultVal;
    };
  }
}
if (typeof WeakMap !== 'undefined') {
  if (!WeakMap.prototype.getOrInsertComputed) {
    WeakMap.prototype.getOrInsertComputed = function (key, callback) {
      if (this.has(key)) return this.get(key);
      const val = callback(key);
      this.set(key, val);
      return val;
    };
  }
  if (!WeakMap.prototype.getOrInsert) {
    WeakMap.prototype.getOrInsert = function (key, defaultVal) {
      if (this.has(key)) return this.get(key);
      this.set(key, defaultVal);
      return defaultVal;
    };
  }
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
        maximumFileSizeToCacheInBytes: 30 * 1024 * 1024, // rhwp_bg.wasm(26.8MB) 오프라인 프리캐시 허용
        globPatterns: [
          '**/*.{js,css,html,ico,png,svg,wasm,mjs,webmanifest}',
          'rhwp/fonts/RHWPPublicBatang-Regular.woff2',
          'rhwp/fonts/NanumGothic-Regular.woff2',
          'rhwp/fonts/NanumGothic-Bold.woff2',
          'rhwp/fonts/LatinModernMath-Regular.woff2',
          'rhwp/fonts/D2Coding-Regular.woff2',
        ],
        globIgnores: [
          '**/rhwp/samples/**',
          '**/rhwp/fonts/licenses/**',
        ],
        runtimeCaching: [
          {
            urlPattern: /\/rhwp\/fonts\/.*\.(woff2|ttf|otf)$/,
            handler: 'CacheFirst',
            options: {
              cacheName: 'rhwp-fonts-cache',
              expiration: {
                maxEntries: 100,
                maxAgeSeconds: 365 * 24 * 60 * 60, // 1 year
              },
              cacheableResponse: {
                statuses: [0, 200],
              },
            },
          },
        ],
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
        // 무손실 편집이 검증된 5개 오피스 포맷 파일 연결 (구버전 .xls/.ppt/.doc 배제)
        file_handlers: [
          {
            action: '/',
            accept: {
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
              'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
              'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'],
              'application/hwp+zip': ['.hwpx'],
              'application/vnd.hancom.hwpx': ['.hwpx'],
              'application/x-hwp': ['.hwp']
            }
          }
        ],
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
  build: {
    target: ['es2020', 'chrome80', 'safari13', 'edge80'],
  },
  base: './' // 상대 경로로 빌드하여 GitHub Pages 등 어디서나 서브디렉토리 배포 가능
});
