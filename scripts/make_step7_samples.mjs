// STEP7 큰 문서 점검용 가상 시험 파일(워드·PDF 60쪽+ 그림 多, 한글 hwpx 본문 반복)을 만드는 스크립트
// 모든 내용은 가상 데이터다. 결과는 용량이 커서 저장소에 넣지 않는다(.gitignore). 실행: node scripts/make_step7_samples.mjs
import fs from 'fs';
import path from 'path';
import zlib from 'zlib';
import { fileURLToPath } from 'url';
import JSZip from 'jszip';
import { Document, Packer, Paragraph, TextRun, ImageRun, PageBreak, HeadingLevel, Table, TableRow, TableCell, WidthType } from 'docx';
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '..');
const outDir = path.join(root, 'docs', 'office-upgrade', 'samples', 'step7');
fs.mkdirSync(outDir, { recursive: true });

// --- 간단한 PNG 인코더 (잡음 섞인 그라데이션: 압축이 잘 안 돼 실제 사진처럼 용량이 큼) ---
const CRC_TABLE = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
const crc32 = (buf) => { let c = 0xffffffff; for (const b of buf) c = CRC_TABLE[(c ^ b) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
function chunk(type, data) {
  const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type), data]);
  const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(td));
  return Buffer.concat([len, td, crc]);
}
function makePng(w, h, seed) {
  let s = seed * 9301 + 49297;
  const rand = () => ((s = (s * 9301 + 49297) % 233280) / 233280);
  const raw = Buffer.alloc((w * 3 + 1) * h);
  for (let y = 0; y < h; y++) {
    raw[y * (w * 3 + 1)] = 0;
    for (let x = 0; x < w; x++) {
      const o = y * (w * 3 + 1) + 1 + x * 3;
      raw[o] = (x * 255 / w + seed * 40 + rand() * 60) & 255;
      raw[o + 1] = (y * 255 / h + rand() * 60) & 255;
      raw[o + 2] = ((x + y) * 128 / (w + h) + seed * 20 + rand() * 60) & 255;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

const PAGES = 60;
const IMAGES = 30;
const imgs = Array.from({ length: IMAGES }, (_, i) => makePng(480, 320, i + 1));
const para = (i, j) => `가상 문단 ${i}-${j}: 이 문서는 큰 파일 점검을 위한 시험용 가상 문서입니다. 실제 업무 내용은 없습니다. 숫자 ${i * 100 + j}.`;

async function makeDocx() {
  const children = [];
  for (let p = 1; p <= PAGES; p++) {
    children.push(new Paragraph({ text: `${p}쪽. 큰 문서 점검 (가상)`, heading: HeadingLevel.HEADING_1 }));
    for (let j = 1; j <= 6; j++) children.push(new Paragraph({ children: [new TextRun(para(p, j))] }));
    if (p % 2 === 0) {
      const img = imgs[(p / 2 - 1) % IMAGES];
      children.push(new Paragraph({ children: [new ImageRun({ type: 'png', data: img, transformation: { width: 480, height: 320 } })] }));
    } else {
      children.push(new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: Array.from({ length: 4 }, (_, r) => new TableRow({ children: Array.from({ length: 4 }, (_, c) => new TableCell({ children: [new Paragraph(`${p}-${r + 1}-${c + 1}`)] })) })),
      }));
    }
    if (p < PAGES) children.push(new Paragraph({ children: [new PageBreak()] }));
  }
  const doc = new Document({ sections: [{ children }] });
  fs.writeFileSync(path.join(outDir, 'step7_가상_큰워드_60쪽_그림30.docx'), await Packer.toBuffer(doc));
}

async function makePdf() {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const embedded = [];
  for (const im of imgs) embedded.push(await pdf.embedPng(im));
  for (let p = 1; p <= PAGES; p++) {
    const page = pdf.addPage([595, 842]);
    page.drawText(`Big PDF test page ${p} (virtual data)`, { x: 50, y: 790, size: 18, font, color: rgb(0.1, 0.1, 0.4) });
    for (let j = 0; j < 8; j++) page.drawText(`Line ${j + 1}: virtual test data ${p * 100 + j}`, { x: 50, y: 750 - j * 18, size: 11, font });
    page.drawImage(embedded[(p - 1) % IMAGES], { x: 57, y: 160, width: 480, height: 320 });
  }
  fs.writeFileSync(path.join(outDir, 'step7_가상_큰PDF_60쪽_그림30.pdf'), await pdf.save());
}

/** 최상위 <hp:p> 문단들의 [시작, 끝] 위치 (표·글상자 안 문단은 제외) */
function topLevelParagraphs(xml, bodyStart) {
  const out = [];
  const re = /<hp:p\b|<\/hp:p>/g;
  re.lastIndex = bodyStart;
  let depth = 0; let start = -1; let m;
  while ((m = re.exec(xml))) {
    if (m[0] === '</hp:p>') {
      depth--;
      if (depth === 0) out.push([start, m.index + 7]);
    } else {
      if (depth === 0) start = m.index;
      depth++;
    }
  }
  return out;
}

async function makeHwpx() {
  const src = fs.readFileSync(path.join(root, 'docs', 'office-upgrade', 'samples', 'real', '한글C_업무계획_원본.hwpx'));
  const zip = await JSZip.loadAsync(src);
  const xml = await zip.file('Contents/section0.xml').async('string');
  const bodyStart = xml.indexOf('>', xml.indexOf('<hs:sec')) + 1;
  const paras = topLevelParagraphs(xml, bodyStart);
  // 첫 문단(구역 설정·머리말·꼬리말)은 한 번만, 변경 추적·각주·필드가 든 문단은 반복하지 않는다
  const body = paras.slice(1).map(([a, b]) => xml.slice(a, b)).filter((p) => !/TcId|footNote|fieldBegin|endNote/.test(p));
  // 반복마다 다른 그림(BIN0002~)을 쓰도록 바꾸고 BinData·header.xml·content.hpf에 등록한 뒤 50회 반복
  const REPEAT = 50;
  let rep = '';
  let header = await zip.file('Contents/header.xml').async('string');
  let hpf = await zip.file('Contents/content.hpf').async('string');
  const binItems = [];
  for (let i = 0; i < REPEAT; i++) {
    let chunkXml = body.join('').replace(/<hp:t>([^<]*)<\/hp:t>/, `<hp:t>[반복 ${i + 1}] $1</hp:t>`);
    if (i < IMAGES) {
      const id = `BIN${String(i + 2).padStart(4, '0')}`;
      chunkXml = chunkXml.split('binaryItemIDRef="BIN0001"').join(`binaryItemIDRef="${id}"`);
      zip.file(`BinData/${id}.png`, imgs[i]);
      binItems.push(id);
    }
    rep += chunkXml;
  }
  const total = 1 + binItems.length;
  header = header.replace(/<hh:binDataList itemCnt="\d+">/, `<hh:binDataList itemCnt="${total}">`)
    .replace('</hh:binDataList>', binItems.map((id, k) => `<hh:binItem id="${k + 1}" Type="Embedding" BinData="${id}.png" Format="png"/>`).join('') + '</hh:binDataList>');
  hpf = hpf.replace(/(<opf:item id="BIN0001"[^>]*\/>)/, (m) => m + binItems.map((id) => `<opf:item id="${id}" href="BinData/${id}.png" media-type="image/png" isEmbeded="1" />`).join(''));
  zip.file('Contents/header.xml', header);
  zip.file('Contents/content.hpf', hpf);
  const last = paras[paras.length - 1][1];
  const out = xml.slice(0, last) + rep + xml.slice(last);
  zip.file('Contents/section0.xml', out);
  // 미리보기 글자도 시험용으로
  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', mimeType: 'application/hwp+zip' });
  fs.writeFileSync(path.join(outDir, 'step7_가상_큰한글_본문50회반복_그림30.hwpx'), buf);
  return { paras: paras.length, repeated: body.length * REPEAT, pics: (out.match(/<hp:pic\b/g) || []).length, tbls: (out.match(/<hp:tbl\b/g) || []).length };
}

await makeDocx();
await makePdf();
const h = await makeHwpx();
console.log('hwpx', JSON.stringify(h));
for (const f of fs.readdirSync(outDir)) console.log(f, fs.statSync(path.join(outDir, f)).size);
