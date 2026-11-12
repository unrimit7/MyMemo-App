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
  // First try getElementsByTagName with tag name and prefix
  const list = parent.getElementsByTagName(tagName);
  if (list && list.length > 0) {
    return Array.from(list);
  }
  // Fallback to checking localName
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
 * HWPX 문서 저장 후처리 어댑터
 *
 * 1. 원본 문서의 version.xml(버전 정보 1.4/1.5 등) 보존
 * 2. 텍스트가 실제로 수정된 문단만 linesegarray(줄 배치 캐시)를 제거하여 한컴 한글에서 글자 겹침 현상 원천 방지
 * 3. 수정되지 않은 문단은 원본 linesegarray 및 구조 100% 바이트 보존
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

    // 1. version.xml 보존 (한컴 한글 호환성 유지)
    const origVer = origZip.file('version.xml');
    if (origVer) {
      const origVerXml = await origVer.async('string');
      expZip.file('version.xml', origVerXml);
      modified = true;
    }

    const parser = getDomParser();
    const serializer = getXmlSerializer();

    // 2. Contents/section*.xml 문단 단위 검사
    for (const [fname, expFile] of Object.entries(expZip.files)) {
      if (fname.startsWith('Contents/section') && fname.endsWith('.xml')) {
        const origFile = origZip.file(fname);
        if (!origFile) continue;

        const expXml = await expFile.async('string');
        const origXml = await origFile.async('string');

        if (expXml === origXml) continue;

        const expDoc = parser.parseFromString(expXml, 'application/xml');
        const origDoc = parser.parseFromString(origXml, 'application/xml');

        const expParas = findElements(expDoc, 'hp:p');
        const origParas = findElements(origDoc, 'hp:p');

        let sectionChanged = false;

        // 원본 문단의 텍스트를 map으로도 기록 (id 기준 및 index 기준)
        const origParaMap = new Map<string, string>();
        origParas.forEach((p, idx) => {
          const id = p.getAttribute('id') || String(idx);
          const text = (p.textContent || '').replace(/\s+/g, ' ').trim();
          origParaMap.set(id, text);
        });

        expParas.forEach((expP, idx) => {
          const id = expP.getAttribute('id') || String(idx);
          const expText = (expP.textContent || '').replace(/\s+/g, ' ').trim();
          const origText = origParaMap.has(id)
            ? origParaMap.get(id)!
            : (origParas[idx] ? (origParas[idx].textContent || '').replace(/\s+/g, ' ').trim() : '');

          // 텍스트가 변경된 문단만 linesegarray 제거
          if (expText !== origText) {
            const lsList = findElements(expP, 'hp:linesegarray');
            lsList.forEach((ls) => {
              if (ls.parentNode) {
                ls.parentNode.removeChild(ls);
                sectionChanged = true;
              }
            });
          }
        });

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
