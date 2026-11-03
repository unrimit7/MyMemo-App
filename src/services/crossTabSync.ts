// 멀티 탭/창 실시간 연동 및 동기화 충돌 방지 서비스

export interface SyncMessage {
  type: 'PING' | 'PONG' | 'DATA_CHANGED';
  tabId: string;
  action?: 'notes' | 'projects' | 'sync_complete';
  timestamp: number;
}

const CHANNEL_NAME = 'mymemo_tab_sync_channel';
export const CURRENT_TAB_ID = `tab_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;

class CrossTabSyncManager {
  private channel: BroadcastChannel | null = null;
  private listeners: ((msg: SyncMessage) => void)[] = [];
  private otherTabsCount: number = 0;
  private onOtherTabsChange?: (count: number) => void;

  constructor() {
    if (typeof window !== 'undefined' && 'BroadcastChannel' in window) {
      try {
        this.channel = new BroadcastChannel(CHANNEL_NAME);
        this.channel.onmessage = (event) => {
          this.handleMessage(event.data);
        };
      } catch (e) {
        console.warn('BroadcastChannel 초기화 실패:', e);
      }
    }

    // 다른 탭 존재 여부 확인 (Ping)
    setTimeout(() => {
      this.send({ type: 'PING', tabId: CURRENT_TAB_ID, timestamp: Date.now() });
    }, 500);

    // 창 닫힐 때 락 해제
    if (typeof window !== 'undefined') {
      window.addEventListener('beforeunload', () => {
        this.releaseSyncLock();
      });
    }
  }

  private handleMessage(msg: SyncMessage) {
    if (!msg || msg.tabId === CURRENT_TAB_ID) return;

    if (msg.type === 'PING') {
      this.otherTabsCount = Math.max(1, this.otherTabsCount + 1);
      this.onOtherTabsChange?.(this.otherTabsCount);
      // PONG으로 응답하여 내 존재를 알림
      this.send({ type: 'PONG', tabId: CURRENT_TAB_ID, timestamp: Date.now() });
    } else if (msg.type === 'PONG') {
      this.otherTabsCount = Math.max(1, this.otherTabsCount + 1);
      this.onOtherTabsChange?.(this.otherTabsCount);
    } else if (msg.type === 'DATA_CHANGED') {
      for (const listener of this.listeners) {
        listener(msg);
      }
    }
  }

  public subscribe(listener: (msg: SyncMessage) => void) {
    this.listeners.push(listener);
    return () => {
      this.listeners = this.listeners.filter((l) => l !== listener);
    };
  }

  public setOtherTabsCallback(cb: (count: number) => void) {
    this.onOtherTabsChange = cb;
  }

  public broadcastChange(action: 'notes' | 'projects' | 'sync_complete') {
    this.send({
      type: 'DATA_CHANGED',
      tabId: CURRENT_TAB_ID,
      action,
      timestamp: Date.now(),
    });
  }

  private send(msg: SyncMessage) {
    if (this.channel) {
      try {
        this.channel.postMessage(msg);
      } catch (e) {
        console.warn('BroadcastChannel 메시지 전송 실패:', e);
      }
    }
  }

  // Cross-Tab 동기화 락 (동일 PC 내 여러 창에서 동시 GitHub 동기화 실행 방지)
  public acquireSyncLock(): boolean {
    if (typeof window === 'undefined') return true;
    const lockKey = 'mymemo_sync_lock';
    try {
      const existing = localStorage.getItem(lockKey);
      if (existing) {
        const parsed = JSON.parse(existing);
        // 락이 걸린 지 45초 이내이고 다른 탭이 보유 중이라면 획득 불가
        if (parsed.tabId !== CURRENT_TAB_ID && Date.now() - parsed.timestamp < 45000) {
          return false;
        }
      }
      localStorage.setItem(
        lockKey,
        JSON.stringify({ tabId: CURRENT_TAB_ID, timestamp: Date.now() })
      );
      return true;
    } catch {
      return true;
    }
  }

  public releaseSyncLock() {
    if (typeof window === 'undefined') return;
    const lockKey = 'mymemo_sync_lock';
    try {
      const existing = localStorage.getItem(lockKey);
      if (existing) {
        const parsed = JSON.parse(existing);
        if (parsed.tabId === CURRENT_TAB_ID) {
          localStorage.removeItem(lockKey);
        }
      }
    } catch {
      // ignore
    }
  }
}

export const crossTabSync = new CrossTabSyncManager();
