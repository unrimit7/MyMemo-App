/**
 * 다른 이름으로 저장 (Save As) 유틸리티
 *
 * - window.showSaveFilePicker 지원 시 데스크톱 네이티브 파일 저장 대화상자 호출
 * - 미지원 환경(또는 취소 시) <a> 다운로드 폴백
 */
export async function saveFileAs(
  bytes: Uint8Array,
  suggestedName: string,
  mimeType: string,
  extension: string
): Promise<{ success: boolean; canceled?: boolean; method: 'picker' | 'download' }> {
  const ext = extension.startsWith('.') ? extension : `.${extension}`;

  if (typeof window !== 'undefined' && 'showSaveFilePicker' in window) {
    try {
      const handle = await (window as any).showSaveFilePicker({
        suggestedName,
        types: [
          {
            description: `${extension.toUpperCase()} 문서`,
            accept: {
              [mimeType]: [ext],
            },
          },
        ],
      });

      const writable = await handle.createWritable();
      await writable.write(bytes as any);
      await writable.close();
      return { success: true, method: 'picker' };
    } catch (err: any) {
      if (err.name === 'AbortError') {
        return { success: false, canceled: true, method: 'picker' };
      }
      console.warn('showSaveFilePicker 실패, 다운로드로 폴백:', err);
    }
  }

  // Fallback to standard Blob download
  try {
    const blob = new Blob([bytes as any], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = suggestedName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    return { success: true, method: 'download' };
  } catch (err) {
    console.error('파일 다운로드 실패:', err);
    return { success: false, method: 'download' };
  }
}
