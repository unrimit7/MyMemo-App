// 모바일 브라우저(Samsung Internet, iOS Safari, 구버전 크롬/웹뷰) ES2024/ES2025 호환성 폴리필
if (typeof Promise !== 'undefined' && !(Promise as any).withResolvers) {
  (Promise as any).withResolvers = function () {
    let resolve: any, reject: any;
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
if (typeof Array !== 'undefined' && !(Array.prototype as any).toSpliced) {
  (Array.prototype as any).toSpliced = function (start: number, deleteCount: number, ...items: any[]) {
    const copy = Array.from(this);
    copy.splice(start, deleteCount, ...items);
    return copy;
  };
}
if (typeof Set !== 'undefined' && !(Set.prototype as any).intersection) {
  (Set.prototype as any).intersection = function (other: any) {
    const res = new Set();
    for (const elem of this) {
      if (other && (typeof other.has === 'function' ? other.has(elem) : false)) {
        res.add(elem);
      }
    }
    return res;
  };
}
// TC39 Map/WeakMap getOrInsertComputed & getOrInsert 폴리필 (pdf.js v6 모바일 렌더링 필수)
if (typeof Map !== 'undefined') {
  if (!(Map.prototype as any).getOrInsertComputed) {
    (Map.prototype as any).getOrInsertComputed = function (key: any, callback: (k: any) => any) {
      if (this.has(key)) {
        return this.get(key);
      }
      const value = callback(key);
      this.set(key, value);
      return value;
    };
  }
  if (!(Map.prototype as any).getOrInsert) {
    (Map.prototype as any).getOrInsert = function (key: any, defaultValue: any) {
      if (this.has(key)) {
        return this.get(key);
      }
      this.set(key, defaultValue);
      return defaultValue;
    };
  }
}
if (typeof WeakMap !== 'undefined') {
  if (!(WeakMap.prototype as any).getOrInsertComputed) {
    (WeakMap.prototype as any).getOrInsertComputed = function (key: any, callback: (k: any) => any) {
      if (this.has(key)) {
        return this.get(key);
      }
      const value = callback(key);
      this.set(key, value);
      return value;
    };
  }
  if (!(WeakMap.prototype as any).getOrInsert) {
    (WeakMap.prototype as any).getOrInsert = function (key: any, defaultValue: any) {
      if (this.has(key)) {
        return this.get(key);
      }
      this.set(key, defaultValue);
      return defaultValue;
    };
  }
}

// PWA 단일 인스턴스 (LaunchQueue & BroadcastChannel 포커스 연동 및 파일 연결)
if (typeof window !== 'undefined') {
  if ('launchQueue' in window) {
    (window as any).launchQueue.setConsumer((launchParams: any) => {
      window.focus();
      if (launchParams?.files && launchParams.files.length > 0) {
        const fileHandle = launchParams.files[0];
        window.dispatchEvent(new CustomEvent('mymemo_open_file_handle', { detail: { fileHandle } }));
      }
    });
  }
  if (typeof BroadcastChannel !== 'undefined') {
    const singleInstanceChannel = new BroadcastChannel('mymemo_single_instance');
    singleInstanceChannel.onmessage = (e) => {
      if (e.data === 'REQUEST_FOCUS') {
        window.focus();
      }
    };
    singleInstanceChannel.postMessage('REQUEST_FOCUS');
  }
}

// 새 서비스 워커가 업데이트되어 활성화되면 구버전 캐시 대신 즉시 화면을 갱신
if (typeof navigator !== 'undefined' && 'serviceWorker' in navigator) {
  let refreshing = false;
  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (!refreshing) {
      refreshing = true;
      window.location.reload();
    }
  });
}

import React, { Component, type ErrorInfo, type ReactNode } from 'react';
import { createRoot } from 'react-dom/client';
import { registerSW } from 'virtual:pwa-register';
import './index.css';
import App from './App.tsx';

// PWA 서비스 워커 자동 등록 및 새 배포 버전 감지 시 프롬프트 알림 이벤트 발송
const updateSW = registerSW({
  immediate: true,
  onNeedRefresh() {
    window.dispatchEvent(
      new CustomEvent('mymemo_pwa_update_available', {
        detail: {
          update: () => updateSW(true),
        },
      })
    );
  },
  onOfflineReady() {
    console.log('[PWA] 오프라인 캐시가 준비되었습니다.');
  },
  onRegisteredSW(_swUrl, registration) {
    if (registration) {
      // 1분마다 새 버전 배포 확인
      setInterval(() => {
        registration.update();
      }, 60 * 1000);
    }
  },
});

interface ErrorBoundaryProps {
  children: ReactNode;
}

interface ErrorBoundaryState {
  hasError: boolean;
  error: Error | null;
}

class ErrorBoundary extends Component<ErrorBoundaryProps, ErrorBoundaryState> {
  constructor(props: ErrorBoundaryProps) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { hasError: true, error };
  }

  componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    console.error('MyMemo 런타임 오류:', error, errorInfo);
  }

  render() {
    if (this.state.hasError) {
      return (
        <div className="flex h-screen w-screen flex-col items-center justify-center bg-slate-950 p-6 text-center text-slate-100">
          <div className="max-w-md space-y-4 rounded-2xl border border-rose-500/30 bg-slate-900 p-6 shadow-2xl">
            <div className="text-3xl">⚠️</div>
            <h2 className="text-base font-bold text-rose-400">화면을 표시하는 중 문제가 발생했습니다</h2>
            <p className="text-xs text-slate-400">
              {this.state.error?.message || '일시적인 오류입니다. 새로고침을 진행해 주세요.'}
            </p>
            <div className="flex justify-center gap-2 pt-2">
              <button
                onClick={() => window.location.reload()}
                className="rounded-xl bg-indigo-600 px-4 py-2 text-xs font-semibold text-white hover:bg-indigo-500"
              >
                앱 새로고침
              </button>
              <button
                onClick={() => {
                  localStorage.clear();
                  window.location.reload();
                }}
                className="rounded-xl bg-slate-800 px-4 py-2 text-xs font-semibold text-slate-300 hover:bg-slate-700"
              >
                캐시 초기화 후 재시작
              </button>
            </div>
          </div>
        </div>
      );
    }

    return this.props.children;
  }
}

createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </React.StrictMode>
);
