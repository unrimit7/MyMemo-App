// Data URL 또는 Base64 문자열을 안전하게 Uint8Array로 변환
export function dataUrlToUint8Array(dataOrUrl: string): Uint8Array {
  const commaIdx = dataOrUrl.indexOf(',');
  const base64 = commaIdx !== -1 ? dataOrUrl.slice(commaIdx + 1) : dataOrUrl;
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

// Uint8Array 바이트를 Base64 Data URL로 안전하게 비동기 변환
export function uint8ArrayToDataUrl(bytes: Uint8Array, mimeType: string = 'application/octet-stream'): Promise<string> {
  return new Promise((resolve, reject) => {
    const blob = new Blob([bytes as any], { type: mimeType });
    const reader = new FileReader();
    reader.onloadend = () => {
      resolve(reader.result as string);
    };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

// 16진수 오프셋 주소 포맷팅 (예: 00000000, 00001A30)
export function formatHexOffset(offset: number, padding: number = 8): string {
  return offset.toString(16).toUpperCase().padStart(padding, '0');
}

// 1바이트를 2자리 16진수 문자열로 포맷팅 (예: 4D, 0A)
export function formatByteToHex(byte: number): string {
  return byte.toString(16).toUpperCase().padStart(2, '0');
}

// 1바이트를 표시 가능한 ASCII 문자로 변환 (제어문자는 '.'로 표시)
export function byteToAsciiChar(byte: number): string {
  if (byte >= 32 && byte <= 126) {
    return String.fromCharCode(byte);
  }
  return '.';
}

// 16진수 바이트 배열 패턴 검색 (예: [0x4D, 0x5A])
export function searchHexBytes(
  buffer: Uint8Array,
  pattern: Uint8Array,
  startOffset: number = 0,
  searchForward: boolean = true
): number {
  if (pattern.length === 0 || buffer.length === 0) return -1;
  const pLen = pattern.length;

  if (searchForward) {
    const start = Math.max(0, startOffset);
    const end = buffer.length - pLen;
    for (let i = start; i <= end; i++) {
      let match = true;
      for (let j = 0; j < pLen; j++) {
        if (buffer[i + j] !== pattern[j]) {
          match = false;
          break;
        }
      }
      if (match) return i;
    }
  } else {
    const start = Math.min(buffer.length - pLen, startOffset);
    for (let i = start; i >= 0; i--) {
      let match = true;
      for (let j = 0; j < pLen; j++) {
        if (buffer[i + j] !== pattern[j]) {
          match = false;
          break;
        }
      }
      if (match) return i;
    }
  }

  return -1;
}

// 문자열(Text) 패턴 검색
export function searchStringBytes(
  buffer: Uint8Array,
  query: string,
  startOffset: number = 0,
  caseInsensitive: boolean = true,
  searchForward: boolean = true
): number {
  if (!query || buffer.length === 0) return -1;
  const qUpper = caseInsensitive ? query.toUpperCase() : query;
  const qLen = qUpper.length;

  if (searchForward) {
    const start = Math.max(0, startOffset);
    const end = buffer.length - qLen;
    for (let i = start; i <= end; i++) {
      let match = true;
      for (let j = 0; j < qLen; j++) {
        const char = String.fromCharCode(buffer[i + j]);
        const testChar = caseInsensitive ? char.toUpperCase() : char;
        if (testChar !== qUpper[j]) {
          match = false;
          break;
        }
      }
      if (match) return i;
    }
  } else {
    const start = Math.min(buffer.length - qLen, startOffset);
    for (let i = start; i >= 0; i--) {
      let match = true;
      for (let j = 0; j < qLen; j++) {
        const char = String.fromCharCode(buffer[i + j]);
        const testChar = caseInsensitive ? char.toUpperCase() : char;
        if (testChar !== qUpper[j]) {
          match = false;
          break;
        }
      }
      if (match) return i;
    }
  }

  return -1;
}

export interface InspectorData {
  offset: number;
  uint8: number;
  int8: number;
  uint16LE: number | null;
  uint16BE: number | null;
  int16LE: number | null;
  int16BE: number | null;
  uint32LE: number | null;
  uint32BE: number | null;
  int32LE: number | null;
  int32BE: number | null;
  binary: string;
}

// 특정 오프셋에서의 Data Inspector 값 계산
export function inspectByteValues(buffer: Uint8Array, offset: number): InspectorData | null {
  if (offset < 0 || offset >= buffer.length) return null;

  const uint8 = buffer[offset];
  const int8 = (uint8 << 24) >> 24;
  const binary = uint8.toString(2).padStart(8, '0');

  const view = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength);

  let uint16LE: number | null = null;
  let uint16BE: number | null = null;
  let int16LE: number | null = null;
  let int16BE: number | null = null;

  if (offset + 1 < buffer.length) {
    uint16LE = view.getUint16(offset, true);
    uint16BE = view.getUint16(offset, false);
    int16LE = view.getInt16(offset, true);
    int16BE = view.getInt16(offset, false);
  }

  let uint32LE: number | null = null;
  let uint32BE: number | null = null;
  let int32LE: number | null = null;
  let int32BE: number | null = null;

  if (offset + 3 < buffer.length) {
    uint32LE = view.getUint32(offset, true);
    uint32BE = view.getUint32(offset, false);
    int32LE = view.getInt32(offset, true);
    int32BE = view.getInt32(offset, false);
  }

  return {
    offset,
    uint8,
    int8,
    uint16LE,
    uint16BE,
    int16LE,
    int16BE,
    uint32LE,
    uint32BE,
    int32LE,
    int32BE,
    binary,
  };
}
