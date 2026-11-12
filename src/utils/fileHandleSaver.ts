import { saveFileOriginalBackup, computeDataHash } from './draftStorage';

export interface FileHandleContext {
  handle: any; // FileSystemFileHandle
  initialSize: number;
  initialLastModified: number;
  fileName: string;
  initialHash?: string;
  backupKey?: string;
  isBackedUp?: boolean;
}

/**
 * 디스크 파일이 외부 프로그램에 의해 수정되었는지 확인합니다.
 */
export async function checkFileHandleConflict(ctx: FileHandleContext): Promise<boolean> {
  try {
    const file = await ctx.handle.getFile();
    return file.size !== ctx.initialSize || file.lastModified !== ctx.initialLastModified;
  } catch {
    return false;
  }
}

/**
 * 최초 덮어쓰기 전 원본 바이트를 IndexedDB에 안전 백업합니다.
 * 백업 키: 파일명 + 크기 + lastModified + 원본 해시 조합
 */
export async function ensureOriginalBackup(
  ctx: FileHandleContext,
  originalBytes: Uint8Array,
  originalHash?: string
): Promise<void> {
  if (ctx.isBackedUp) return;
  try {
    const hash = originalHash || ctx.initialHash || (await computeDataHash(originalBytes));
    ctx.initialHash = hash;
    const backupKey = ctx.backupKey || `${ctx.fileName}_${ctx.initialSize}_${ctx.initialLastModified}_${hash}`;
    ctx.backupKey = backupKey;
    await saveFileOriginalBackup(
      backupKey,
      ctx.fileName,
      ctx.initialSize,
      ctx.initialLastModified,
      hash,
      originalBytes
    );
    ctx.isBackedUp = true;
  } catch (err) {
    console.warn('파일 핸들 원본 백업 실패:', err);
  }
}

/**
 * FileSystemFileHandle을 통해 원래 파일에 무손실 직렬화 바이트를 기록합니다.
 */
export async function saveToFileHandle(ctx: FileHandleContext, bytes: Uint8Array): Promise<void> {
  const writable = await ctx.handle.createWritable();
  await writable.write(bytes as any);
  await writable.close();

  // 저장 완료 후 디스크 파일의 새 크기 및 최종 수정 일시로 갱신
  try {
    const updated = await ctx.handle.getFile();
    ctx.initialSize = updated.size;
    ctx.initialLastModified = updated.lastModified;
  } catch {
    // ignore
  }
}
