/**
 * 오피스 문서 로컬 임시본(Draft) 및 파일 핸들러 원본 백업 스토리지
 *
 * - IndexedDB 기반 완전 로컬 저장 (메모 첨부 및 GitHub 동기화 원천 배제)
 * - 30초 주기 + visibilitychange / blur 시점에만 저장
 * - 기준 원본 해시 일치 검증으로 첨부 변경 시 안전한 폐기
 * - 파일당 1개, 전체 50MB 상한, 7일 경과 자동 청소
 */

export interface OfficeDraft {
  draftKey: string; // `${memoId}_${attId}` 또는 `file_${fileName}`
  memoId?: string;
  attId?: string;
  title: string;
  fileType: 'excel' | 'docx' | 'pptx' | 'hwp' | 'hwpx';
  baseOriginalHash: string; // 문서 열람 시점의 원본 SHA-256 해시
  data: any; // snapshot JSON / diff / dataUrl
  dataType: 'univer_snapshot' | 'superdoc_json' | 'pptx_state' | 'data_url';
  updatedAt: number; // Date.now()
  byteLength: number;
  saveDurationMs: number; // 저장 소요 시간(ms)
}

const DB_NAME = 'mymemo_office_drafts_db';
const DB_VERSION = 2;
const DRAFT_STORE = 'drafts';
const FILE_BACKUP_STORE = 'file_backups';
const MAX_TOTAL_BYTES = 50 * 1024 * 1024; // 50MB
const MAX_AGE_MS = 7 * 24 * 60 * 60 * 1000; // 7일

let dbPromise: Promise<IDBDatabase> | null = null;

function getDb(): Promise<IDBDatabase> {
  if (typeof indexedDB === 'undefined') {
    return Promise.reject(new Error('IndexedDB is not supported in this environment'));
  }

  if (!dbPromise) {
    dbPromise = new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(DRAFT_STORE)) {
          db.createObjectStore(DRAFT_STORE, { keyPath: 'draftKey' });
        }
        if (db.objectStoreNames.contains(FILE_BACKUP_STORE)) {
          db.deleteObjectStore(FILE_BACKUP_STORE);
        }
        db.createObjectStore(FILE_BACKUP_STORE, { keyPath: 'backupKey' });
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => {
        dbPromise = null;
        reject(req.error);
      };
    });
  }
  return dbPromise;
}

/**
 * 데이터의 SHA-256 해시를 계산합니다.
 */
export async function computeDataHash(data: string | Uint8Array): Promise<string> {
  if (typeof crypto === 'undefined' || !crypto.subtle) {
    // Web Crypto 미지원 시 간단한 해시 폴백
    let hash = 0;
    const str = typeof data === 'string' ? data : data.length.toString();
    for (let i = 0; i < str.length; i++) {
      hash = ((hash << 5) - hash) + str.charCodeAt(i);
      hash |= 0;
    }
    return 'fallback_' + Math.abs(hash).toString(16);
  }

  let bytes: Uint8Array;
  if (typeof data === 'string') {
    if (data.startsWith('data:')) {
      const base64 = data.split(',')[1] || '';
      const bin = atob(base64.slice(0, Math.min(base64.length, 1024 * 100))); // 앞부분 100KB 샘플
      bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    } else {
      bytes = new TextEncoder().encode(data);
    }
  } else {
    bytes = data;
  }

  const hashBuf = await crypto.subtle.digest('SHA-256', bytes as any);
  return Array.from(new Uint8Array(hashBuf))
    .map((b) => b.toString(16).padStart(2, '0'))
    .join('');
}

/**
 * 임시본을 IndexedDB에 저장합니다 (전체 용량 및 오래된 임시본 자동 청소 포함).
 * @param draft 임시본 데이터
 * @param startTime 임시본 생성 시작 시각 (performance.now()), 미지정 시 현재 시각 기준
 * @returns 전체 저장 소요 시간 (ms)
 */
export async function saveOfficeDraft(draft: OfficeDraft, startTime?: number): Promise<number> {
  const t0 = startTime ?? performance.now();
  try {
    const db = await getDb();
    const tx = db.transaction([DRAFT_STORE], 'readwrite');
    const store = tx.objectStore(DRAFT_STORE);

    // IDB 쓰기 직전까지의 소요 시간(추출, 직렬화 등)을 포함하여 draft에 기록
    const prepDuration = Math.max(0.1, Number((performance.now() - t0).toFixed(1)));
    const record: OfficeDraft = {
      ...draft,
      saveDurationMs: prepDuration,
    };
    store.put(record);

    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });

    const totalDuration = Math.max(0.1, Number((performance.now() - t0).toFixed(1)));

    // 백그라운드에서 오래된 임시본 정리
    cleanOldOfficeDrafts().catch(() => {});

    return totalDuration;
  } catch (err) {
    console.warn('[DraftStorage] 임시본 저장 실패 (무시됨):', err);
    return Math.max(0.1, Number((performance.now() - t0).toFixed(1)));
  }
}

/**
 * 임시본을 조회합니다.
 */
export async function getOfficeDraft(draftKey: string): Promise<OfficeDraft | null> {
  try {
    const db = await getDb();
    const tx = db.transaction([DRAFT_STORE], 'readonly');
    const store = tx.objectStore(DRAFT_STORE);
    const req = store.get(draftKey);

    return await new Promise<OfficeDraft | null>((resolve, reject) => {
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  } catch (err) {
    console.warn('[DraftStorage] 임시본 조회 실패:', err);
    return null;
  }
}

/**
 * 임시본을 삭제합니다.
 */
export async function deleteOfficeDraft(draftKey: string): Promise<void> {
  try {
    const db = await getDb();
    const tx = db.transaction([DRAFT_STORE], 'readwrite');
    const store = tx.objectStore(DRAFT_STORE);
    store.delete(draftKey);

    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (err) {
    console.warn('[DraftStorage] 임시본 삭제 실패:', err);
  }
}

/**
 * 7일 경과 임시본 정리 및 전체 50MB 초과 시 LRU 정리
 */
export async function cleanOldOfficeDrafts(): Promise<void> {
  try {
    const db = await getDb();
    const tx = db.transaction([DRAFT_STORE], 'readwrite');
    const store = tx.objectStore(DRAFT_STORE);
    const req = store.getAll();

    const items: OfficeDraft[] = await new Promise((resolve, reject) => {
      req.onsuccess = () => resolve(req.result || []);
      req.onerror = () => reject(req.error);
    });

    const now = Date.now();
    let totalBytes = 0;
    const toDelete: string[] = [];

    // 1. 7일 경과 항목 식별
    for (const item of items) {
      if (now - item.updatedAt > MAX_AGE_MS) {
        toDelete.push(item.draftKey);
      } else {
        totalBytes += item.byteLength || 0;
      }
    }

    // 2. 50MB 초과 시 오래된 순(LRU) 정리
    if (totalBytes > MAX_TOTAL_BYTES) {
      const remaining = items
        .filter((it) => !toDelete.includes(it.draftKey))
        .sort((a, b) => a.updatedAt - b.updatedAt);

      for (const item of remaining) {
        if (totalBytes <= MAX_TOTAL_BYTES) break;
        toDelete.push(item.draftKey);
        totalBytes -= item.byteLength || 0;
      }
    }

    for (const key of toDelete) {
      store.delete(key);
    }
  } catch {
    // ignore
  }
}

// ----------------------------------------------------------------------------
// FileSystemFileHandle 연동 파일 백업 (원래 파일 저장 전 원본 바이트 보관)
// ----------------------------------------------------------------------------

export interface FileBackupRecord {
  backupKey: string;
  fileName: string;
  fileSize: number;
  lastModified: number;
  originalHash: string;
  bytes: Uint8Array;
  savedAt: number;
}

export async function saveFileOriginalBackup(
  backupKey: string,
  fileName: string,
  fileSize: number,
  lastModified: number,
  originalHash: string,
  bytes: Uint8Array
): Promise<void> {
  try {
    const db = await getDb();
    const tx = db.transaction([FILE_BACKUP_STORE], 'readwrite');
    const store = tx.objectStore(FILE_BACKUP_STORE);
    store.put({
      backupKey,
      fileName,
      fileSize,
      lastModified,
      originalHash,
      bytes,
      savedAt: Date.now(),
    });

    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (err) {
    console.warn('[DraftStorage] 파일 백업 저장 실패:', err);
  }
}

export async function getFileOriginalBackup(backupKey: string): Promise<FileBackupRecord | null> {
  try {
    const db = await getDb();
    const tx = db.transaction([FILE_BACKUP_STORE], 'readonly');
    const store = tx.objectStore(FILE_BACKUP_STORE);
    const req = store.get(backupKey);

    const res = await new Promise<FileBackupRecord | null>((resolve, reject) => {
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });

    return res;
  } catch {
    return null;
  }
}
