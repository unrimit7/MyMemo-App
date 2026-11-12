import JSZip from 'jszip';

function getDomParser(): any {
  if (typeof DOMParser !== 'undefined') {
    return new DOMParser();
  }
  // Node.js fallback
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const { DOMParser: XmldomParser } = require('@xmldom/xmldom');
  return new XmldomParser();
}

function getXmlSerializer(): any {
  if (typeof XMLSerializer !== 'undefined') {
    return new XMLSerializer();
  }
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const { XMLSerializer: XmldomSerializer } = require('@xmldom/xmldom');
  return new XmldomSerializer();
}

function findElements(parent: Document | Element, tagName: string): Element[] {
  const list = parent.getElementsByTagName(tagName);
  if (list && list.length > 0) {
    return Array.from(list);
  }
  const all = parent.getElementsByTagName('*');
  const res: Element[] = [];
  const cleanName = tagName.includes(':') ? tagName.split(':')[1] : tagName;
  for (let i = 0; i < all.length; i++) {
    const el = all[i];
    if (el.localName === cleanName || el.nodeName === tagName) {
      res.push(el);
    }
  }
  return res;
}

/**
 * 문단의 직계 런(hp:run > hp:t) 텍스트만 추출
 * (내포된 표 hp:tbl, 구역 정의 등의 텍스트는 제외하여 순수 문단 텍스트만 비교)
 */
function getDirectParaText(pNode: Element): string {
  let text = '';
  for (let c = pNode.firstChild; c; c = c.nextSibling) {
    if (c.nodeName === 'hp:run' || (c as any).localName === 'run') {
      for (let r = c.firstChild; r; r = r.nextSibling) {
        if (r.nodeName === 'hp:t' || (r as any).localName === 't') {
          text += r.textContent || '';
        }
      }
    }
  }
  return text.replace(/\s+/g, ' ').trim();
}

/**
 * 문단의 직계 hp:linesegarray 자식 요소만 제거 (내포된 표 셀 등의 linesegarray 보존)
 */
function removeDirectLineSegArray(pNode: Element): boolean {
  let removed = false;
  const toRemove: Node[] = [];
  for (let c = pNode.firstChild; c; c = c.nextSibling) {
    if (c.nodeName === 'hp:linesegarray' || (c as any).localName === 'linesegarray') {
      toRemove.push(c);
    }
  }
  toRemove.forEach((n) => {
    pNode.removeChild(n);
    removed = true;
  });
  return removed;
}

/**
 * HWP/HWPX 문서에 변경 추적(Track Changes)이 포함되어 있는지 감지
 */
export async function hasHwpTrackChanges(
  bytes: Uint8Array,
  isHwpx: boolean
): Promise<{ hasTrackChanges: boolean; details: string[] }> {
  const details: string[] = [];
  if (isHwpx) {
    try {
      const zip = await JSZip.loadAsync(bytes);
      for (const name of Object.keys(zip.files)) {
        if (name.endsWith('.xml')) {
          const text = await zip.file(name)!.async('string');
          if (text.includes('trackChange') || text.includes('insertBegin') || text.includes('deleteBegin')) {
            details.push(`${name} 내 변경 추적 태그 포함`);
          }
        }
      }
    } catch {
      // ignore
    }
  } else {
    // HWP 바이너리: OLE2 레코드 및 FileHeader/BodyText 검사
    try {
      const latin = Buffer.from(bytes).toString('latin1');
      if (latin.includes('TrackChange') || latin.includes('TRACKCHANGE')) {
        details.push('HWP 바이너리 내 변경 추적 레코드 포함');
      }
    } catch {
      // ignore
    }
  }
  return {
    hasTrackChanges: details.length > 0,
    details,
  };
}

/**
 * HWPX 문서 저장 후처리 어댑터
 *
 * 1. 원본 문서의 version.xml(버전 정보 1.4/1.5 등) 보존
 * 2. 원본 header.xml의 hh:trackChanges 보존
 * 3. 텍스트가 실제로 수정된 문단(표 셀 문단 포함)만 직계 linesegarray(줄 배치 캐시)를 선별 제거
 * 4. 수정되지 않은 문단 중 변경 추적(deleteBegin/insertBegin)이 포함된 문단은 원본 마크업을 복원하여 보존
 * 5. 수정되지 않은 문단은 원본 linesegarray 및 구조 100% 바이트 보존
 */
export async function postProcessHwpx(
  exportedBytes: Uint8Array,
  originalBytes?: Uint8Array
): Promise<Uint8Array> {
  if (!originalBytes) return exportedBytes;

  try {
    const expZip = await JSZip.loadAsync(exportedBytes);
    const origZip = await JSZip.loadAsync(originalBytes);

    let modified = false;
    const parser = getDomParser();
    const serializer = getXmlSerializer();

    // 1. version.xml 보존 (한컴 한글 호환성 유지)
    const origVer = origZip.file('version.xml');
    if (origVer) {
      const origVerXml = await origVer.async('string');
      expZip.file('version.xml', origVerXml);
      modified = true;
    }

    // 2. Contents/header.xml 내 변경 추적 메타데이터(hh:trackChanges, hh:trackChangeAuthors, 설정) 원본 복원
    const origHdrFile = origZip.file('Contents/header.xml');
    const expHdrFile = expZip.file('Contents/header.xml');
    if (origHdrFile && expHdrFile) {
      const origHdrXml = await origHdrFile.async('string');
      if (origHdrXml.includes('trackChanges') || origHdrXml.includes('trackChangeAuthors')) {
        const expHdrXml = await expHdrFile.async('string');
        const docOrigHdr = parser.parseFromString(origHdrXml, 'application/xml');
        const docExpHdr = parser.parseFromString(expHdrXml, 'application/xml');

        let hdrChanged = false;
        const expRefLists = findElements(docExpHdr, 'hh:refList');
        const expRefList = expRefLists[0];

        if (expRefList) {
          // 기존에 잘못 붙었을 수 있는 docExpHdr 내부 trackChanges/trackChangeAuthors 정리
          const existingExpTC = findElements(docExpHdr, 'hh:trackChanges');
          const existingExpTCA = findElements(docExpHdr, 'hh:trackChangeAuthors');
          for (const el of existingExpTC) {
            if (el.parentNode) el.parentNode.removeChild(el);
          }
          for (const el of existingExpTCA) {
            if (el.parentNode) el.parentNode.removeChild(el);
          }

          // docOrigHdr에서 trackChanges와 trackChangeAuthors 추출하여 expRefList 끝에 원본 순서대로 삽입
          const origTC = findElements(docOrigHdr, 'hh:trackChanges');
          const origTCA = findElements(docOrigHdr, 'hh:trackChangeAuthors');

          if (origTC.length > 0) {
            const imported = docExpHdr.importNode ? docExpHdr.importNode(origTC[0], true) : origTC[0];
            expRefList.appendChild(imported);
            hdrChanged = true;
          }
          if (origTCA.length > 0) {
            const imported = docExpHdr.importNode ? docExpHdr.importNode(origTCA[0], true) : origTCA[0];
            expRefList.appendChild(imported);
            hdrChanged = true;
          }
        }

        // trackchageConfig / trackchangeConfig 복원 (docExpHdr.documentElement 직계)
        const origConfigs = findElements(docOrigHdr, 'hh:trackchageConfig').concat(findElements(docOrigHdr, 'hh:trackchangeConfig'));
        const expConfigs = findElements(docExpHdr, 'hh:trackchageConfig').concat(findElements(docExpHdr, 'hh:trackchangeConfig'));
        if (origConfigs.length > 0 && expConfigs.length === 0) {
          const imported = docExpHdr.importNode ? docExpHdr.importNode(origConfigs[0], true) : origConfigs[0];
          docExpHdr.documentElement.appendChild(imported);
          hdrChanged = true;
        }

        if (hdrChanged) {
          expZip.file('Contents/header.xml', serializer.serializeToString(docExpHdr));
          modified = true;
        }
      }
    }

    // 3. Contents/section*.xml 문단 단위 linesegarray 선별 제거 및 미수정 변경추적 복원
    for (const [fname, expFile] of Object.entries(expZip.files)) {
      if (fname.startsWith('Contents/section') && fname.endsWith('.xml')) {
        const origFile = origZip.file(fname);
        if (!origFile) continue;

        const expXml = await expFile.async('string');
        const origXml = await origFile.async('string');

        if (expXml === origXml) continue;

        const expDoc = parser.parseFromString(expXml, 'application/xml');
        const origDoc = parser.parseFromString(origXml, 'application/xml');

        // Live NodeList 변조 방지를 위해 정적 배열로 변환
        const expParas = Array.from(findElements(expDoc, 'hp:p'));
        const origParas = Array.from(findElements(origDoc, 'hp:p'));

        let sectionChanged = false;

        // 1. 문단 텍스트 빈도 맵 구성 (구역/헤더/문단 순서 불일치 방지)
        const origTextFreq = new Map<string, number>();
        for (const p of origParas) {
          const txt = getDirectParaText(p);
          origTextFreq.set(txt, (origTextFreq.get(txt) || 0) + 1);
        }

        // 2. 텍스트가 실제로 변경된 문단만 직계 linesegarray 선별 제거
        for (const expP of expParas) {
          const txt = getDirectParaText(expP);
          const remaining = origTextFreq.get(txt) || 0;
          if (remaining > 0) {
            origTextFreq.set(txt, remaining - 1);
          } else {
            if (removeDirectLineSegArray(expP)) {
              sectionChanged = true;
            }
          }
        }

        // 3. 미수정 문단 중 변경 추적 마크업(deleteBegin/insertBegin) 복원 (텍스트 기준 매칭)
        for (const origP of origParas) {
          const origPXml = serializer.serializeToString(origP);
          if (origPXml.includes('deleteBegin') || origPXml.includes('insertBegin')) {
            const origText = getDirectParaText(origP);
            const targetExpP = expParas.find((p) => {
              const pXml = serializer.serializeToString(p);
              return getDirectParaText(p) === origText && !pXml.includes('deleteBegin') && !pXml.includes('insertBegin');
            });
            if (targetExpP && targetExpP.parentNode) {
              targetExpP.parentNode.replaceChild(origP, targetExpP);
              sectionChanged = true;
            }
          }
        }

        if (sectionChanged) {
          const newXml = serializer.serializeToString(expDoc);
          expZip.file(fname, newXml);
          modified = true;
        }
      }
    }

    if (modified) {
      return await expZip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    }
    return exportedBytes;
  } catch (err) {
    console.warn('HWPX 후처리 중 오류 (원본 export 유지):', err);
    return exportedBytes;
  }
}

/**
 * HWP 5.0 CFB 내보내기 후처리:
 * ViewText 스트림이 없는데 FileHeader의 bit 14(0x4000, 변경 추적 문서 플래그)가 켜져 있으면
 * 한글 2018 등 데스크톱 프로그램에서 "%1을(를) 찾지 못했습니다" 오류를 발생시킵니다.
 * ViewText가 없으면 bit 14를 해제(props &= ~0x4000)하여 정상 열람되도록 합니다.
 */
export function postProcessHwp(exportedBytes: Uint8Array): Uint8Array {
  try {
    const buf = new Uint8Array(exportedBytes);
    // 1. ViewText 스트림 존재 여부 확인 (UTF-16LE: 'V\0i\0e\0w\0T\0e\0x\0t\0')
    const viewTextBytes = [0x56, 0x00, 0x69, 0x00, 0x65, 0x00, 0x77, 0x00, 0x54, 0x00, 0x65, 0x00, 0x78, 0x00, 0x74, 0x00];
    let hasViewText = false;
    for (let i = 0; i <= buf.length - viewTextBytes.length; i++) {
      let match = true;
      for (let j = 0; j < viewTextBytes.length; j++) {
        if (buf[i + j] !== viewTextBytes[j]) {
          match = false;
          break;
        }
      }
      if (match) {
        hasViewText = true;
        break;
      }
    }

    if (hasViewText) {
      return exportedBytes;
    }

    // 2. FileHeader 시그니처("HWP Document File") 검색
    const sigBytes = [0x48, 0x57, 0x50, 0x20, 0x44, 0x6F, 0x63, 0x75, 0x6D, 0x65, 0x6E, 0x74, 0x20, 0x46, 0x69, 0x6C, 0x65];
    for (let i = 0; i <= buf.length - 40; i++) {
      let match = true;
      for (let j = 0; j < sigBytes.length; j++) {
        if (buf[i + j] !== sigBytes[j]) {
          match = false;
          break;
        }
      }
      if (match) {
        const propsOffset = i + 36;
        const currentProps = buf[propsOffset] | (buf[propsOffset + 1] << 8) | (buf[propsOffset + 2] << 16) | (buf[propsOffset + 3] << 24);
        if ((currentProps & 0x4000) !== 0) {
          const newProps = currentProps & ~0x4000;
          buf[propsOffset] = newProps & 0xff;
          buf[propsOffset + 1] = (newProps >> 8) & 0xff;
          buf[propsOffset + 2] = (newProps >> 16) & 0xff;
          buf[propsOffset + 3] = (newProps >> 24) & 0xff;
        }
      }
    }
    return buf;
  } catch (err) {
    console.warn('HWP 후처리 중 오류:', err);
    return exportedBytes;
  }
}
