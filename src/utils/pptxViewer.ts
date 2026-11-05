import JSZip from 'jszip';

export interface PptxTextRun {
  text: string;
  isBold?: boolean;
  isItalic?: boolean;
  size?: number;
  color?: string;
}

export interface PptxParagraph {
  text: string;
  runs: PptxTextRun[];
  level: number;
  isTitle?: boolean;
}

export interface PptxTable {
  headers: string[];
  rows: string[][];
}

export interface PptxImage {
  id: string;
  dataUrl: string;
  contentType: string;
}

export interface PptxSlide {
  slideNumber: number;
  title: string;
  paragraphs: PptxParagraph[];
  tables: PptxTable[];
  images: PptxImage[];
  notes?: string;
  slidePath: string;
}

export interface PptxPresentation {
  title: string;
  slideCount: number;
  slides: PptxSlide[];
  rawZip?: JSZip;
}

// Data URL에서 Base64 바이너리 추출
function dataUrlToBinary(dataUrl: string): string {
  const commaIdx = dataUrl.indexOf(',');
  const base64 = commaIdx !== -1 ? dataUrl.slice(commaIdx + 1) : dataUrl;
  return atob(base64);
}

// 바이너리 문자열을 Uint8Array로 변환
function binaryToUint8Array(binary: string): Uint8Array {
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

// 캐시 맵
const presentationCache = new Map<string, { presentation: PptxPresentation; timestamp: number }>();

export function invalidatePptxCache(attId?: string) {
  if (attId) {
    presentationCache.delete(attId);
  } else {
    presentationCache.clear();
  }
}

// XML 문자열을 DOM으로 파싱
function parseXml(xmlText: string): Document {
  const parser = new DOMParser();
  return parser.parseFromString(xmlText, 'application/xml');
}

// 노드에서 로컬 태그 이름으로 요소들 찾기 (네임스페이스 무관)
function getElementsByLocalName(parent: Element | Document, localName: string): Element[] {
  const results: Element[] = [];
  const elements = parent.getElementsByTagName('*');
  for (let i = 0; i < elements.length; i++) {
    const el = elements[i];
    if (el.localName === localName) {
      results.push(el);
    }
  }
  return results;
}

// 특정 노드의 직속 자식 중 로컬 태그 이름이 일치하는 요소들 찾기
function getDirectChildrenByLocalName(parent: Element, localName: string): Element[] {
  const results: Element[] = [];
  for (let i = 0; i < parent.children.length; i++) {
    const child = parent.children[i];
    if (child.localName === localName) {
      results.push(child);
    }
  }
  return results;
}

// PPTX 파일 파싱
export async function parsePptxPresentation(dataUrl: string): Promise<PptxPresentation> {
  const binary = dataUrlToBinary(dataUrl);
  const bytes = binaryToUint8Array(binary);
  const zip = await JSZip.loadAsync(bytes);

  // 1. 슬라이드 순서 및 파일 목록 확인
  let slidePaths: string[] = [];

  // presentation.xml 및 _rels/presentation.xml.rels 확인
  const presRelsFile = zip.file('ppt/_rels/presentation.xml.rels');
  const presFile = zip.file('ppt/presentation.xml');

  if (presRelsFile && presFile) {
    try {
      const relsText = await presRelsFile.async('text');
      const presText = await presFile.async('text');

      const relsDoc = parseXml(relsText);
      const presDoc = parseXml(presText);

      // rId -> Target 매핑
      const relMap = new Map<string, string>();
      const relElements = getElementsByLocalName(relsDoc, 'Relationship');
      for (const rel of relElements) {
        const id = rel.getAttribute('Id');
        const target = rel.getAttribute('Target');
        if (id && target) {
          // 상대 경로 처리: 'slides/slide1.xml' -> 'ppt/slides/slide1.xml'
          const normalized = target.startsWith('ppt/') ? target : `ppt/${target.replace(/^\//, '')}`;
          relMap.set(id, normalized);
        }
      }

      // presentation.xml 내 sldId 목록 순서대로 추출
      const sldIdElements = getElementsByLocalName(presDoc, 'sldId');
      for (const sld of sldIdElements) {
        const rId = sld.getAttribute('r:id') || sld.getAttribute('id');
        if (rId && relMap.has(rId)) {
          slidePaths.push(relMap.get(rId)!);
        }
      }
    } catch (err) {
      console.warn('presentation.xml.rels 파싱 실패, 파일명 검색 대체:', err);
    }
  }

  // 만약 rels를 통해 못 찾았다면 zip 파일 목록에서 직접 탐색
  if (slidePaths.length === 0) {
    const allFiles = Object.keys(zip.files);
    const regex = /^ppt\/slides\/slide(\d+)\.xml$/;
    const slideEntries: { path: string; num: number }[] = [];

    for (const f of allFiles) {
      const match = f.match(regex);
      if (match) {
        slideEntries.push({ path: f, num: parseInt(match[1], 10) });
      }
    }

    slideEntries.sort((a, b) => a.num - b.num);
    slidePaths = slideEntries.map((e) => e.path);
  }

  // 2. 각 슬라이드 파싱
  const slides: PptxSlide[] = [];
  let presentationTitle = '';

  for (let idx = 0; idx < slidePaths.length; idx++) {
    const sPath = slidePaths[idx];
    const sFile = zip.file(sPath);
    if (!sFile) continue;

    const sXmlText = await sFile.async('text');
    const sDoc = parseXml(sXmlText);

    // 슬라이드 관계(이미지 등) 파싱
    // 예: ppt/slides/slide1.xml -> ppt/slides/_rels/slide1.xml.rels
    const pathParts = sPath.split('/');
    const fileName = pathParts.pop() || '';
    const relsPath = `${pathParts.join('/')}/_rels/${fileName}.rels`;
    const relsFile = zip.file(relsPath);

    const slideImageMap = new Map<string, string>(); // rId -> target media path
    if (relsFile) {
      try {
        const sRelsText = await relsFile.async('text');
        const sRelsDoc = parseXml(sRelsText);
        const rels = getElementsByLocalName(sRelsDoc, 'Relationship');
        for (const r of rels) {
          const type = r.getAttribute('Type') || '';
          if (type.includes('image')) {
            const id = r.getAttribute('Id');
            let target = r.getAttribute('Target') || '';
            if (id && target) {
              // '../media/image1.png' -> 'ppt/media/image1.png'
              if (target.startsWith('../')) {
                target = `ppt/${target.substring(3)}`;
              } else if (!target.startsWith('ppt/')) {
                target = `ppt/slides/${target}`;
              }
              slideImageMap.set(id, target);
            }
          }
        }
      } catch (e) {
        console.warn(`슬라이드 rels(${relsPath}) 파싱 오류:`, e);
      }
    }

    // 슬라이드 내 이미지 추출
    const images: PptxImage[] = [];
    const blipElements = getElementsByLocalName(sDoc, 'blip');
    for (const blip of blipElements) {
      const embedId = blip.getAttribute('r:embed');
      if (embedId && slideImageMap.has(embedId)) {
        const mediaPath = slideImageMap.get(embedId)!;
        const mediaFile = zip.file(mediaPath);
        if (mediaFile) {
          try {
            const ext = mediaPath.split('.').pop()?.toLowerCase() || 'png';
            const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : ext === 'gif' ? 'image/gif' : ext === 'svg' ? 'image/svg+xml' : 'image/png';
            const base64 = await mediaFile.async('base64');
            images.push({
              id: embedId,
              dataUrl: `data:${mime};base64,${base64}`,
              contentType: mime,
            });
          } catch (imgErr) {
            console.warn(`이미지(${mediaPath}) 읽기 오류:`, imgErr);
          }
        }
      }
    }

    // 슬라이드 내 텍스트 및 제목 추출
    let slideTitle = '';
    const paragraphs: PptxParagraph[] = [];

    // 도형(Shape)들 탐색
    const spElements = getElementsByLocalName(sDoc, 'sp');
    for (const sp of spElements) {
      // 제목 여부 판별 (nvSpPr -> nvPr -> ph type="title" | "ctrTitle")
      let isShapeTitle = false;
      const phElements = getElementsByLocalName(sp, 'ph');
      for (const ph of phElements) {
        const phType = ph.getAttribute('type');
        if (phType === 'title' || phType === 'ctrTitle' || phType === 'subTitle') {
          isShapeTitle = true;
          break;
        }
      }

      // 텍스트 바디(txBody)
      const txBodyList = getElementsByLocalName(sp, 'txBody');
      for (const txBody of txBodyList) {
        const pElements = getDirectChildrenByLocalName(txBody, 'p');
        for (const p of pElements) {
          // 들여쓰기 레벨
          let level = 0;
          const pPrList = getDirectChildrenByLocalName(p, 'pPr');
          if (pPrList.length > 0) {
            const lvlAttr = pPrList[0].getAttribute('lvl');
            if (lvlAttr) level = parseInt(lvlAttr, 10);
          }

          // 텍스트 런(run)들 수집
          const runs: PptxTextRun[] = [];
          const rElements = getDirectChildrenByLocalName(p, 'r');
          let paragraphFullText = '';

          for (const r of rElements) {
            const tElements = getDirectChildrenByLocalName(r, 't');
            for (const t of tElements) {
              const textContent = t.textContent || '';
              if (!textContent) continue;

              let isBold = false;
              let isItalic = false;
              const rPrList = getDirectChildrenByLocalName(r, 'rPr');
              if (rPrList.length > 0) {
                isBold = rPrList[0].getAttribute('b') === '1' || rPrList[0].getAttribute('b') === 'true';
                isItalic = rPrList[0].getAttribute('i') === '1' || rPrList[0].getAttribute('i') === 'true';
              }

              runs.push({
                text: textContent,
                isBold,
                isItalic,
              });
              paragraphFullText += textContent;
            }
          }

          // 직속 t 요소가 있는 경우 (예: 필드 fld 내의 t)
          if (runs.length === 0) {
            const allT = getElementsByLocalName(p, 't');
            for (const t of allT) {
              const textContent = t.textContent || '';
              if (textContent) {
                runs.push({ text: textContent });
                paragraphFullText += textContent;
              }
            }
          }

          const trimmed = paragraphFullText.trim();
          if (trimmed) {
            if (isShapeTitle && !slideTitle) {
              slideTitle = trimmed;
              if (idx === 0 && !presentationTitle) {
                presentationTitle = trimmed;
              }
            }

            paragraphs.push({
              text: trimmed,
              runs,
              level,
              isTitle: isShapeTitle,
            });
          }
        }
      }
    }

    // 표(Table) 추출
    const tables: PptxTable[] = [];
    const tblElements = getElementsByLocalName(sDoc, 'tbl');
    for (const tbl of tblElements) {
      const rows: string[][] = [];
      const trElements = getElementsByLocalName(tbl, 'tr');
      for (const tr of trElements) {
        const rowCells: string[] = [];
        const tcElements = getElementsByLocalName(tr, 'tc');
        for (const tc of tcElements) {
          const tElements = getElementsByLocalName(tc, 't');
          const cellText = tElements.map((t) => t.textContent || '').join(' ').trim();
          rowCells.push(cellText);
        }
        if (rowCells.some((c) => c.length > 0)) {
          rows.push(rowCells);
        }
      }

      if (rows.length > 0) {
        tables.push({
          headers: rows[0] || [],
          rows: rows.slice(1),
        });
      }
    }

    // 슬라이드 제목 기본값 설정
    if (!slideTitle) {
      if (paragraphs.length > 0) {
        slideTitle = paragraphs[0].text;
      } else {
        slideTitle = `슬라이드 ${idx + 1}`;
      }
    }

    slides.push({
      slideNumber: idx + 1,
      title: slideTitle,
      paragraphs,
      tables,
      images,
      slidePath: sPath,
    });
  }

  const result: PptxPresentation = {
    title: presentationTitle || (slides[0]?.title ?? '파워포인트 프레젠테이션'),
    slideCount: slides.length,
    slides,
    rawZip: zip,
  };

  return result;
}

// 캐시 기반 파싱
export async function getCachedPptxPresentation(attId: string, dataUrl: string): Promise<PptxPresentation> {
  const cached = presentationCache.get(attId);
  if (cached && Date.now() - cached.timestamp < 1000 * 60 * 30) {
    return cached.presentation;
  }
  const presentation = await parsePptxPresentation(dataUrl);
  presentationCache.set(attId, { presentation, timestamp: Date.now() });
  return presentation;
}

// 마크다운 및 일반 텍스트 추출 (메모로 가져오기용)
export async function extractPptxTextAndMarkdown(dataUrl: string): Promise<{ title: string; markdown: string; text: string }> {
  const pres = await parsePptxPresentation(dataUrl);

  const mdLines: string[] = [];
  const textLines: string[] = [];

  mdLines.push(`# 📊 ${pres.title || '파워포인트 발표 자료'}\n`);
  textLines.push(`${pres.title || '파워포인트 발표 자료'}\n`);

  for (const slide of pres.slides) {
    mdLines.push(`\n## 📄 슬라이드 ${slide.slideNumber}: ${slide.title}\n`);
    textLines.push(`\n[슬라이드 ${slide.slideNumber}] ${slide.title}\n`);

    for (const p of slide.paragraphs) {
      if (p.isTitle && p.text === slide.title) continue; // 중복 제목 생략
      const indent = '  '.repeat(Math.max(0, p.level));
      mdLines.push(`${indent}- ${p.text}`);
      textLines.push(`${indent}• ${p.text}`);
    }

    for (const tbl of slide.tables) {
      if (tbl.headers.length > 0) {
        mdLines.push('\n| ' + tbl.headers.join(' | ') + ' |');
        mdLines.push('| ' + tbl.headers.map(() => '---').join(' | ') + ' |');
        for (const row of tbl.rows) {
          mdLines.push('| ' + row.join(' | ') + ' |');
        }
        mdLines.push('');
      }
    }
  }

  return {
    title: pres.title,
    markdown: mdLines.join('\n'),
    text: textLines.join('\n'),
  };
}

// 슬라이드 텍스트 수정 후 새로운 PPTX Data URL 생성
export async function updatePptxSlideTexts(
  originalDataUrl: string,
  updatedSlides: { slideNumber: number; title: string; bodyText: string }[]
): Promise<string> {
  const binary = dataUrlToBinary(originalDataUrl);
  const bytes = binaryToUint8Array(binary);
  const zip = await JSZip.loadAsync(bytes);

  for (const upd of updatedSlides) {
    const slidePath = `ppt/slides/slide${upd.slideNumber}.xml`;
    const sFile = zip.file(slidePath);
    if (!sFile) continue;

    const xmlText = await sFile.async('text');
    const doc = parseXml(xmlText);

    // 슬라이드 내 모든 텍스트 단락(p) 찾기
    const paragraphs = getElementsByLocalName(doc, 'p');
    if (paragraphs.length > 0) {
      // 1. 첫 단락 또는 제목 단락에 제목 적용
      const firstP = paragraphs[0];
      const tElements = getElementsByLocalName(firstP, 't');
      if (tElements.length > 0) {
        tElements[0].textContent = upd.title;
        for (let i = 1; i < tElements.length; i++) {
          tElements[i].textContent = '';
        }
      }

      // 2. 본문 텍스트 적용 (여러 줄을 단락들에 분배하거나 새 텍스트로 치환)
      const bodyLines = upd.bodyText
        .split('\n')
        .map((l) => l.trim().replace(/^[-*•]\s*/, ''))
        .filter(Boolean);

      if (bodyLines.length > 0) {
        // 기존 2번째 단락부터 수정
        for (let i = 0; i < bodyLines.length; i++) {
          const targetPIndex = i + 1;
          if (targetPIndex < paragraphs.length) {
            const p = paragraphs[targetPIndex];
            const pTs = getElementsByLocalName(p, 't');
            if (pTs.length > 0) {
              pTs[0].textContent = bodyLines[i];
              for (let k = 1; k < pTs.length; k++) pTs[k].textContent = '';
            }
          } else {
            // 단락이 부족하면 마지막 단락을 복제하여 추가
            const lastP = paragraphs[paragraphs.length - 1];
            const cloneP = lastP.cloneNode(true) as Element;
            const cloneTs = getElementsByLocalName(cloneP, 't');
            if (cloneTs.length > 0) {
              cloneTs[0].textContent = bodyLines[i];
              for (let k = 1; k < cloneTs.length; k++) cloneTs[k].textContent = '';
              lastP.parentNode?.appendChild(cloneP);
            }
          }
        }
      }
    }

    // 직렬화하여 zip에 다시 쓰기
    const serializer = new XMLSerializer();
    const updatedXml = serializer.serializeToString(doc);
    zip.file(slidePath, updatedXml);
  }

  // base64 data url 생성
  const updatedBase64 = await zip.generateAsync({
    type: 'base64',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });

  return `data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,${updatedBase64}`;
}
