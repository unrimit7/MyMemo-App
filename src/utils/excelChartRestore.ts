// ExcelJS 저장본에 원본 xlsx의 차트·그리기 부분(xl/drawings, xl/charts와 rels·Content_Types)을 다시 넣어 보존하는 모듈
import JSZip from 'jszip';

const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_DRAWING = `${NS_REL}/drawing`;

export interface ChartRestoreResult {
  bytes: Uint8Array;
  /** 다시 넣은 차트 수 (0이면 넣을 것이 없었거나 넣지 못함) */
  restoredCharts: number;
  /** 안전하게 넣지 못한 이유. null이면 문제 없음 */
  blockedReason: string | null;
}

export interface ChartRestoreContext {
  /** 편집기에서 행·열을 삽입·삭제한 기록 (차트 참조 범위가 어긋날 수 있음) */
  structuralOpCount?: number;
}

interface SheetInfo {
  name: string;
  path: string; // xl/worksheets/sheetN.xml
}

const decodeXml = (s: string) =>
  s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');

/** a/b/../c 같은 경로를 정리한다 */
function resolvePath(baseFile: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = baseFile.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.' && seg !== '') parts.push(seg);
  }
  return parts.join('/');
}

function relsPathOf(partPath: string): string {
  const idx = partPath.lastIndexOf('/');
  return `${partPath.slice(0, idx)}/_rels/${partPath.slice(idx + 1)}.rels`;
}

interface Rel {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

function parseRels(xml: string): Rel[] {
  const out: Rel[] = [];
  for (const m of xml.matchAll(/<(?:\w+:)?Relationship\b([^>]*?)\/?>/g)) {
    const a = m[1];
    const get = (k: string) => (a.match(new RegExp(`\\s${k}\\s*=\\s*"([^"]*)"`)) || [])[1] || '';
    out.push({ id: get('Id'), type: get('Type'), target: decodeXml(get('Target')), external: get('TargetMode') === 'External' });
  }
  return out;
}

async function readText(zip: JSZip, path: string): Promise<string | null> {
  const f = zip.file(path);
  return f ? f.async('string') : null;
}

async function listSheets(zip: JSZip): Promise<SheetInfo[]> {
  const wb = (await readText(zip, 'xl/workbook.xml')) || '';
  const rels = parseRels((await readText(zip, 'xl/_rels/workbook.xml.rels')) || '');
  const out: SheetInfo[] = [];
  for (const m of wb.matchAll(/<(?:\w+:)?sheet\b([^>]*?)\/?>/g)) {
    const a = m[1];
    const name = decodeXml((a.match(/\sname\s*=\s*"([^"]*)"/) || [])[1] || '');
    const rid = (a.match(/\sr:id\s*=\s*"([^"]*)"/) || [])[1] || '';
    const rel = rels.find((r) => r.id === rid);
    if (rel) out.push({ name, path: resolvePath('xl/workbook.xml', rel.target) });
  }
  return out;
}

/** 차트 수식(<c:f>)이 참조하는 시트 이름들 */
function chartReferencedSheets(chartXml: string): string[] {
  const names = new Set<string>();
  for (const m of chartXml.matchAll(/<(?:\w+:)?f>([\s\S]*?)<\/(?:\w+:)?f>/g)) {
    const f = decodeXml(m[1]);
    for (const r of f.matchAll(/(?:'((?:[^']|'')+)'|([^\s,()'!]+))!/g)) {
      names.add((r[1] ? r[1].replace(/''/g, "'") : r[2]).trim());
    }
  }
  return [...names];
}

// CT_Worksheet에서 <drawing>보다 뒤에 와야 하는 요소들
const AFTER_DRAWING = ['legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

function insertDrawingElement(sheetXml: string, rid: string): string {
  const tag = `<drawing r:id="${rid}"/>`;
  let pos = -1;
  for (const name of AFTER_DRAWING) {
    const m = sheetXml.match(new RegExp(`<${name}[\\s>/]`));
    if (m && m.index !== undefined && (pos === -1 || m.index < pos)) pos = m.index;
  }
  if (pos === -1) pos = sheetXml.lastIndexOf('</worksheet>');
  let out = sheetXml.slice(0, pos) + tag + sheetXml.slice(pos);
  if (!/<worksheet\b[^>]*\sxmlns:r=/.test(out)) {
    out = out.replace(/<worksheet\b/, `<worksheet xmlns:r="${NS_REL}"`);
  }
  return out;
}

function nextRelId(rels: Rel[]): string {
  let n = 1;
  const ids = new Set(rels.map((r) => r.id));
  while (ids.has(`rId${n}`)) n++;
  return `rId${n}`;
}

const contentTypeOverride = (ctXml: string, part: string) =>
  (ctXml.match(new RegExp(`<Override\\b[^>]*PartName\\s*=\\s*"/${part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`)) || [])[0] || null;

/**
 * 원본에만 있고 저장본에서 빠진 시트 그리기(차트 포함)를 저장본에 다시 넣는다.
 * 안전하게 넣을 수 없으면 저장본을 그대로 돌려주고 blockedReason을 채운다.
 */
export async function restoreChartsFromOriginal(
  savedBytes: Uint8Array,
  origBytes: Uint8Array,
  ctx: ChartRestoreContext = {},
): Promise<ChartRestoreResult> {
  const unchanged = (reason: string | null): ChartRestoreResult => ({ bytes: savedBytes, restoredCharts: 0, blockedReason: reason });
  const orig = await JSZip.loadAsync(origBytes);
  const chartFiles = Object.keys(orig.files).filter((n) => /^xl\/charts\/chart\d+\.xml$/i.test(n));
  if (chartFiles.length === 0) return unchanged(null);

  const saved = await JSZip.loadAsync(savedBytes);
  const origSheets = await listSheets(orig);
  const savedSheets = await listSheets(saved);
  const savedNames = new Set(savedSheets.map((s) => s.name));

  // 1. 차트가 참조하는 시트가 모두 그대로 있는지, 행·열 구조가 바뀌지 않았는지
  if (ctx.structuralOpCount && ctx.structuralOpCount > 0) {
    return unchanged('행·열을 삽입하거나 삭제해서 차트가 가리키는 범위가 어긋날 수 있습니다.');
  }
  for (const cf of chartFiles) {
    const xml = (await readText(orig, cf)) || '';
    for (const sn of chartReferencedSheets(xml)) {
      if (!savedNames.has(sn)) return unchanged(`차트가 참조하는 시트 '${sn}'의 이름이 바뀌었거나 삭제되었습니다.`);
    }
  }

  // 2. 원본 시트별 그리기 → 저장본 같은 이름 시트에 다시 연결
  const ctOrig = (await readText(orig, '[Content_Types].xml')) || '';
  let ctSaved = (await readText(saved, '[Content_Types].xml')) || '';
  const toCopy = new Set<string>();
  const sheetEdits: { savedSheet: string; drawingPath: string }[] = [];

  for (const os of origSheets) {
    const sheetXml = (await readText(orig, os.path)) || '';
    const dm = sheetXml.match(/<(?:\w+:)?drawing\b[^>]*\sr:id\s*=\s*"([^"]+)"/);
    if (!dm) continue;
    const rel = parseRels((await readText(orig, relsPathOf(os.path))) || '').find((r) => r.id === dm[1]);
    if (!rel || rel.type !== REL_DRAWING) continue;
    const drawingPath = resolvePath(os.path, rel.target);
    // 그리기에서 rels로 닿는 모든 부분(차트, 차트 스타일·색, 그림 등)을 모은다
    const parts = new Set<string>();
    const queue = [drawingPath];
    while (queue.length) {
      const part = queue.shift()!;
      if (parts.has(part)) continue;
      if (!orig.file(part)) return unchanged(`원본의 '${part}'를 찾을 수 없습니다.`);
      parts.add(part);
      const relsPath = relsPathOf(part);
      const relsXml = await readText(orig, relsPath);
      if (relsXml) {
        parts.add(relsPath);
        for (const r of parseRels(relsXml)) if (!r.external) queue.push(resolvePath(part, r.target));
      }
    }
    // 차트가 없는 그리기(그림만)는 저장 엔진이 이미 보존하므로 건드리지 않는다
    if (![...parts].some((x) => /^xl\/charts\/chart\d+\.xml$/i.test(x))) continue;

    const ss = savedSheets.find((s) => s.name === os.name);
    if (!ss) return unchanged(`차트가 있던 시트 '${os.name}'가 저장본에 없습니다.`);
    const savedSheetXml = (await readText(saved, ss.path)) || '';
    if (/<(?:\w+:)?drawing\b/.test(savedSheetXml)) {
      return unchanged(`시트 '${os.name}'에 그림과 차트가 함께 있어 원본 차트를 다시 넣을 수 없습니다.`);
    }
    sheetEdits.push({ savedSheet: ss.path, drawingPath });
    parts.forEach((x) => toCopy.add(x));
  }
  if (sheetEdits.length === 0) return unchanged(null);

  // 3. 이름이 겹치는 부분이 있으면 (내용이 다르면) 넣지 않는다
  for (const part of toCopy) {
    const existing = saved.file(part);
    if (existing) {
      const [a, b] = await Promise.all([existing.async('uint8array'), orig.file(part)!.async('uint8array')]);
      if (a.length !== b.length || a.some((v, i) => v !== b[i])) {
        return unchanged(`저장본에 같은 이름의 다른 파일(${part})이 있어 차트를 넣을 수 없습니다.`);
      }
    }
  }

  // 4. 복사 + Content_Types
  for (const part of toCopy) {
    saved.file(part, await orig.file(part)!.async('uint8array'));
    if (part.endsWith('.rels')) continue;
    if (!contentTypeOverride(ctSaved, part)) {
      const ov = contentTypeOverride(ctOrig, part);
      if (ov) {
        ctSaved = ctSaved.replace('</Types>', `${ov}</Types>`);
      } else {
        const ext = part.split('.').pop()!.toLowerCase();
        if (!new RegExp(`<Default\\b[^>]*Extension\\s*=\\s*"${ext}"`, 'i').test(ctSaved)) {
          const def = (ctOrig.match(new RegExp(`<Default\\b[^>]*Extension\\s*=\\s*"${ext}"[^>]*/>`, 'i')) || [])[0];
          if (!def) return unchanged(`'${part}'의 형식 정보를 원본에서 찾을 수 없습니다.`);
          ctSaved = ctSaved.replace(/<Types\b[^>]*>/, (t) => t + def);
        }
      }
    }
  }
  saved.file('[Content_Types].xml', ctSaved);

  // 5. 시트에 <drawing> 연결과 rels 추가
  for (const { savedSheet, drawingPath } of sheetEdits) {
    const relsPath = relsPathOf(savedSheet);
    let relsXml = (await readText(saved, relsPath)) ||
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    const rid = nextRelId(parseRels(relsXml));
    const target = `../${drawingPath.replace(/^xl\//, '')}`;
    relsXml = relsXml.replace('</Relationships>', `<Relationship Id="${rid}" Type="${REL_DRAWING}" Target="${target}"/></Relationships>`);
    saved.file(relsPath, relsXml);
    saved.file(savedSheet, insertDrawingElement((await readText(saved, savedSheet)) || '', rid));
  }

  const bytes = await saved.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  return { bytes, restoredCharts: [...toCopy].filter((p) => /^xl\/charts\/chart\d+\.xml$/i.test(p)).length, blockedReason: null };
}

/** 편집기가 저장 전에 "차트가 사라질 수 있음" 경고를 띄워야 하는지 판단 (null이면 보존 가능) */
export async function chartRestoreBlockReason(
  origBytes: Uint8Array,
  currentSheetNames: string[],
  structuralOpCount: number,
): Promise<string | null> {
  const orig = await JSZip.loadAsync(origBytes);
  const chartFiles = Object.keys(orig.files).filter((n) => /^xl\/charts\/chart\d+\.xml$/i.test(n));
  if (chartFiles.length === 0) return null;
  if (structuralOpCount > 0) return '행·열을 삽입하거나 삭제해서 차트가 가리키는 범위가 어긋날 수 있습니다.';
  const names = new Set(currentSheetNames);
  for (const s of await listSheets(orig)) {
    const xml = (await readText(orig, s.path)) || '';
    if (/<(?:\w+:)?drawing\b/.test(xml) && !names.has(s.name)) return `그리기가 있던 시트 '${s.name}'의 이름이 바뀌었거나 삭제되었습니다.`;
  }
  for (const cf of chartFiles) {
    for (const sn of chartReferencedSheets((await readText(orig, cf)) || '')) {
      if (!names.has(sn)) return `차트가 참조하는 시트 '${sn}'의 이름이 바뀌었거나 삭제되었습니다.`;
    }
  }
  return null;
}
