/**
 * [참고용/레거시]
 * 본 스크립트는 실 앱 뷰어 연동 전 독립 XML->HTML 변환 렌더링 검증용으로 작성된 레거시 스크립트입니다.
 * 실제 공식 스크린샷은 MyMemo 앱 내장 PptxViewerModal을 Vite dev 서버 + Playwright로
 * 실 브라우저 구동하여 캡처하는 `scripts/capture_viewer_screenshots_playwright.mjs`에서 생성됩니다.
 */
import fs from 'fs';
import path from 'path';
import { execSync } from 'child_process';
import JSZip from 'jszip';
import { JSDOM } from 'jsdom';

const CHROME_PATH = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const OUT_DIR = path.resolve('docs/office-upgrade/out/ppt_render');
const TEMP_HTML_DIR = path.resolve('docs/office-upgrade/out/temp_html');

if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });
if (!fs.existsSync(TEMP_HTML_DIR)) fs.mkdirSync(TEMP_HTML_DIR, { recursive: true });

function getElementsByLocalName(parent, localName) {
  const elements = parent.getElementsByTagName('*');
  const results = [];
  for (let i = 0; i < elements.length; i++) {
    if (elements[i].localName === localName) {
      results.push(elements[i]);
    }
  }
  return results;
}

function getDirectChildrenByLocalName(parent, localName) {
  const results = [];
  for (let i = 0; i < parent.children.length; i++) {
    const child = parent.children[i];
    if (child.localName === localName) results.push(child);
  }
  return results;
}

function parseDrawingColor(parent, themeMap = {}) {
  const srgb = getElementsByLocalName(parent, 'srgbClr');
  if (srgb.length > 0) {
    const val = srgb[0].getAttribute('val');
    if (val) return `#${val}`;
  }
  const scheme = getElementsByLocalName(parent, 'schemeClr');
  if (scheme.length > 0) {
    const val = scheme[0].getAttribute('val');
    if (val) {
      if (themeMap[val]) return themeMap[val];
      const defaultSchemeMap = {
        accent1: '#4F81BD', accent2: '#C0504D', accent3: '#5CB85C',
        accent4: '#8064A2', accent5: '#4BACC6', accent6: '#F79646',
        dk1: '#111827', lt1: '#FFFFFF', dk2: '#1F497D', lt2: '#F3F4F6',
        bg1: '#FFFFFF', tx1: '#111827', bg2: '#F3F4F6', tx2: '#1F497D',
        hlink: '#2563EB', folHlink: '#7C3AED',
      };
      if (defaultSchemeMap[val]) return defaultSchemeMap[val];
    }
  }
  const prst = getElementsByLocalName(parent, 'prstClr');
  if (prst.length > 0) {
    const val = prst[0].getAttribute('val');
    if (val) return val;
  }
  const gradFill = getElementsByLocalName(parent, 'gradFill');
  if (gradFill.length > 0) {
    const gsList = getElementsByLocalName(gradFill[0], 'gs');
    if (gsList.length >= 2) {
      const stops = [];
      for (let i = 0; i < gsList.length; i++) {
        const gs = gsList[i];
        const pos = Math.round((parseInt(gs.getAttribute('pos') || '0', 10) / 100000) * 100);
        const c = parseDrawingColor(gs, themeMap) || '#4F81BD';
        stops.push(`${c} ${pos}%`);
      }
      return `linear-gradient(135deg, ${stops.join(', ')})`;
    }
  }
  return undefined;
}

function parseElementCoordinates(parent, slideWidth, slideHeight) {
  const xfrmList = getElementsByLocalName(parent, 'xfrm');
  if (xfrmList.length === 0) return null;
  const xfrm = xfrmList[0];
  const offList = getElementsByLocalName(xfrm, 'off');
  const extList = getElementsByLocalName(xfrm, 'ext');
  if (offList.length === 0 || extList.length === 0) return null;

  const rawX = parseInt(offList[0].getAttribute('x') || '0', 10);
  const rawY = parseInt(offList[0].getAttribute('y') || '0', 10);
  const rawW = parseInt(extList[0].getAttribute('cx') || '0', 10);
  const rawH = parseInt(extList[0].getAttribute('cy') || '0', 10);

  const left = Math.max(0, Math.min(100, (rawX / slideWidth) * 100));
  const top = Math.max(0, Math.min(100, (rawY / slideHeight) * 100));
  const width = Math.max(0, Math.min(100, (rawW / slideWidth) * 100));
  const height = Math.max(0, Math.min(100, (rawH / slideHeight) * 100));

  const rotAttr = xfrm.getAttribute('rot');
  const rotation = rotAttr ? Math.round(parseInt(rotAttr, 10) / 60000) : undefined;

  return { left, top, width, height, rotation };
}

function parseParagraph(p, themeMap = {}) {
  let level = 0;
  let align = 'left';
  let bulletChar = undefined;

  const pPrList = getDirectChildrenByLocalName(p, 'pPr');
  if (pPrList.length > 0) {
    const lvlAttr = pPrList[0].getAttribute('lvl');
    if (lvlAttr) level = parseInt(lvlAttr, 10);
    const algnAttr = pPrList[0].getAttribute('algn');
    if (algnAttr === 'ctr') align = 'center';
    else if (algnAttr === 'r') align = 'right';
    else if (algnAttr === 'just') align = 'justify';
    else if (algnAttr === 'l') align = 'left';

    const buChar = getElementsByLocalName(pPrList[0], 'buChar');
    if (buChar.length > 0) {
      bulletChar = buChar[0].getAttribute('char') || '•';
    }
  }

  const runs = [];
  let fullText = '';

  for (let i = 0; i < p.children.length; i++) {
    const child = p.children[i];
    const localName = child.localName;

    if (localName === 'br') {
      runs.push({ text: '\n' });
      fullText += '\n';
      continue;
    }
    if (localName === 'fld' || localName === 'r') {
      const tEls = getDirectChildrenByLocalName(child, 't');
      for (const t of tEls) {
        const textContent = t.textContent || '';
        if (!textContent) continue;

        let isBold = false;
        let isItalic = false;
        let isUnderline = false;
        let isStrike = false;
        let baseline = undefined;
        let highlightColor = undefined;
        let hyperlink = undefined;
        let fontSize = undefined;
        let color = undefined;

        const rPrList = getDirectChildrenByLocalName(child, 'rPr');
        if (rPrList.length > 0) {
          const rPr = rPrList[0];
          isBold = rPr.getAttribute('b') === '1' || rPr.getAttribute('b') === 'true';
          isItalic = rPr.getAttribute('i') === '1' || rPr.getAttribute('i') === 'true';
          isUnderline = rPr.getAttribute('u') === 'sng';
          isStrike = rPr.getAttribute('strike') === 'sngStrike';
          const baseAttr = rPr.getAttribute('baseline');
          if (baseAttr) baseline = parseInt(baseAttr, 10);
          const szAttr = rPr.getAttribute('sz');
          if (szAttr) fontSize = Math.round(parseInt(szAttr, 10) / 100);

          color = parseDrawingColor(rPr, themeMap);
          const hlList = getElementsByLocalName(rPr, 'highlight');
          if (hlList.length > 0) highlightColor = parseDrawingColor(hlList[0], themeMap) || '#FEF08A';
          const hlinkList = getElementsByLocalName(rPr, 'hlinkClick');
          if (hlinkList.length > 0) hyperlink = hlinkList[0].getAttribute('action') || hlinkList[0].getAttribute('r:id') || '#';
        }

        runs.push({
          text: textContent,
          isBold,
          isItalic,
          isUnderline,
          isStrike,
          baseline,
          highlightColor,
          hyperlink,
          fontSize,
          color,
        });
        fullText += textContent;
      }
    }
  }

  return { text: fullText, runs, level, align, bulletChar };
}

function parseChartXml(chartDoc) {
  let chartType = 'bar';
  if (getElementsByLocalName(chartDoc, 'pieChart').length > 0) chartType = 'pie';
  else if (getElementsByLocalName(chartDoc, 'lineChart').length > 0) chartType = 'line';

  let title = '';
  const titleList = getElementsByLocalName(chartDoc, 'title');
  if (titleList.length > 0) {
    const tEls = getElementsByLocalName(titleList[0], 't');
    title = tEls.map(t => t.textContent || '').join(' ').trim();
  }

  const categories = [];
  const catList = getElementsByLocalName(chartDoc, 'cat');
  if (catList.length > 0) {
    const ptEls = getElementsByLocalName(catList[0], 'pt');
    for (const pt of ptEls) {
      const v = getElementsByLocalName(pt, 'v')[0]?.textContent || '';
      categories.push(v);
    }
  }

  const series = [];
  const serList = getElementsByLocalName(chartDoc, 'ser');
  for (const ser of serList) {
    let name = '';
    const txList = getElementsByLocalName(ser, 'tx');
    if (txList.length > 0) {
      name = getElementsByLocalName(txList[0], 'v')[0]?.textContent || '';
    }
    const values = [];
    const valList = getElementsByLocalName(ser, 'val');
    if (valList.length > 0) {
      const ptEls = getElementsByLocalName(valList[0], 'pt');
      for (const pt of ptEls) {
        const valStr = getElementsByLocalName(pt, 'v')[0]?.textContent || '0';
        values.push(parseFloat(valStr) || 0);
      }
    }
    series.push({ name, values });
  }

  return { chartType, title, categories, series };
}

async function renderSlideToHtml(slide, slideIndex, totalSlides, presentationTitle) {
  const elementsHtml = slide.elements.map(el => {
    if (el.type === 'shape') {
      const isLine = el.shapeType === 'line';
      if (isLine) {
        return `
          <div style="position:absolute; left:${el.left}%; top:${el.top}%; width:${Math.max(0.2, el.width)}%; height:${Math.max(1, el.borderWidth || 1)}px; background-color:${el.borderColor || '#334155'}; z-index:${el.zIndex};"></div>
        `;
      }

      const rotStyle = el.rotation ? `transform: rotate(${el.rotation}deg); transform-origin: center center;` : '';
      const vertStyle = el.isVerticalText ? 'writing-mode: vertical-rl;' : '';
      const colStyle = el.columnCount && el.columnCount > 1 ? `column-count: ${el.columnCount}; column-gap: 16px;` : '';
      const borderStyle = el.borderColor ? `border: ${el.borderWidth || 1}px solid ${el.borderColor};` : '';
      const bgStyle = el.fillColor ? (el.fillColor.startsWith('linear-gradient') ? `background: ${el.fillColor};` : `background-color: ${el.fillColor};`) : 'background-color: transparent;';
      const radiusStyle = el.isRounded ? 'border-radius: 16px;' : el.shapeType === 'ellipse' ? 'border-radius: 9999px;' : 'border-radius: 3px;';

      const paragraphsHtml = el.paragraphs.map(p => {
        const alignStyle = `text-align: ${p.align || 'left'};`;
        const indentStyle = p.level > 0 ? `padding-left: ${p.level * 16}px;` : '';
        const bulletHtml = p.bulletChar ? `<span style="margin-right: 4px; color: #64748B;">${p.bulletChar}</span>` : '';

        const runsHtml = p.runs.map(r => {
          let textStyle = '';
          if (r.color) textStyle += `color: ${r.color};`;
          if (r.highlightColor) textStyle += `background-color: ${r.highlightColor};`;
          if (r.fontSize) textStyle += `font-size: ${Math.max(11, Math.min(32, r.fontSize * 0.8))}px;`;
          else if (el.isTitle) textStyle += 'font-size: 24px;';
          else textStyle += 'font-size: 13px;';

          if (r.isBold || el.isTitle) textStyle += 'font-weight: bold;';
          if (r.isItalic) textStyle += 'font-style: italic;';
          if (r.isStrike) textStyle += 'text-decoration: line-through;';
          else if (r.isUnderline) textStyle += 'text-decoration: underline;';
          if (r.baseline && r.baseline > 0) textStyle += 'vertical-align: super; font-size: 0.75em;';
          else if (r.baseline && r.baseline < 0) textStyle += 'vertical-align: sub; font-size: 0.75em;';

          const cls = r.hyperlink ? 'class="link"' : '';
          return `<span style="${textStyle}" ${cls}>${escapeHtml(r.text)}</span>`;
        }).join('');

        return `<div style="${alignStyle} ${indentStyle} margin-bottom: 3px; white-space: pre-wrap; line-height: 1.35;">${bulletHtml}${runsHtml}</div>`;
      }).join('');

      return `
        <div style="position:absolute; left:${el.left}%; top:${el.top}%; width:${el.width}%; height:${el.height}%; ${bgStyle} ${borderStyle} ${radiusStyle} ${rotStyle} ${vertStyle} ${colStyle} z-index:${el.zIndex}; overflow:hidden; display:flex; flex-direction:column; justify-content:${el.verticalAlign === 'middle' ? 'center' : el.verticalAlign === 'bottom' ? 'flex-end' : 'flex-start'}; padding: 6px 10px; box-sizing: border-box;">
          ${paragraphsHtml}
        </div>
      `;
    }

    if (el.type === 'table') {
      const rowsHtml = el.rows.map((row, rIdx) => {
        const cellsHtml = row.cells.map((cell, cIdx) => {
          const bg = cell.fillColor || (rIdx === 0 ? '#F1F5F9' : '#FFFFFF');
          const pContent = cell.paragraphs.map(p => {
            const rText = p.runs.map(r => `<span style="font-weight:${r.isBold?'bold':'normal'}; color:${r.color || '#0F172A'}; font-size:${r.fontSize ? `${r.fontSize * 0.75}px` : '12px'};">${escapeHtml(r.text)}</span>`).join('');
            return `<div style="text-align:${p.align || 'left'};">${rText}</div>`;
          }).join('');

          return `<td colspan="${cell.colSpan || 1}" rowspan="${cell.rowSpan || 1}" style="background-color:${bg}; border: 1px solid #CBD5E1; padding: 5px 8px; vertical-align:${cell.verticalAlign || 'middle'};">${pContent}</td>`;
        }).join('');
        return `<tr>${cellsHtml}</tr>`;
      }).join('');

      return `
        <div style="position:absolute; left:${el.left}%; top:${el.top}%; width:${el.width}%; height:${el.height}%; z-index:${el.zIndex}; overflow:hidden; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
          <table style="width:100%; height:100%; border-collapse:collapse; font-size:12px; table-layout:fixed;">
            <tbody>${rowsHtml}</tbody>
          </table>
        </div>
      `;
    }

    if (el.type === 'image') {
      return `
        <div style="position:absolute; left:${el.left}%; top:${el.top}%; width:${el.width}%; height:${el.height}%; z-index:${el.zIndex}; display:flex; align-items:center; justify-content:center;">
          <img src="${el.dataUrl}" style="width:100%; height:100%; object-fit:contain;" />
        </div>
      `;
    }

    if (el.type === 'chart') {
      const palette = ['#3B82F6', '#F59E0B', '#10B981', '#EC4899', '#8B5CF6'];
      let maxVal = 1;
      for (const s of el.series) {
        for (const v of s.values) {
          if (v > maxVal) maxVal = v;
        }
      }

      let chartSvg = '';
      if (el.chartType === 'bar') {
        const catCount = Math.max(1, el.categories.length);
        const groupW = 260 / catCount;
        let barsSvg = '';
        el.categories.forEach((cat, cIdx) => {
          const gx = 30 + cIdx * groupW;
          const barW = Math.max(5, (groupW * 0.7) / Math.max(1, el.series.length));
          el.series.forEach((s, sIdx) => {
            const val = s.values[cIdx] || 0;
            const barH = (val / maxVal) * 110;
            const bx = gx + 4 + sIdx * barW;
            const by = 135 - barH;
            barsSvg += `<rect x="${bx}" y="${by}" width="${barW - 1}" height="${Math.max(1, barH)}" fill="${palette[sIdx % palette.length]}" rx="1" />`;
          });
          barsSvg += `<text x="${gx + groupW / 2}" y="150" text-anchor="middle" font-size="9" fill="#64748B">${escapeHtml(cat)}</text>`;
        });
        chartSvg = `
          <svg viewBox="0 0 300 170" style="width:100%; height:100%;">
            <line x1="30" y1="25" x2="290" y2="25" stroke="#E2E8F0" stroke-dasharray="2,2" stroke-width="0.8"/>
            <line x1="30" y1="80" x2="290" y2="80" stroke="#E2E8F0" stroke-dasharray="2,2" stroke-width="0.8"/>
            <line x1="30" y1="135" x2="290" y2="135" stroke="#94A3B8" stroke-width="1.2"/>
            ${barsSvg}
          </svg>
        `;
      } else if (el.chartType === 'line') {
        const catCount = Math.max(1, el.categories.length);
        let linesSvg = '';
        el.series.forEach((s, sIdx) => {
          const color = palette[sIdx % palette.length];
          const pts = s.values.map((v, vIdx) => {
            const step = catCount > 1 ? 250 / (catCount - 1) : 125;
            const x = 35 + vIdx * step;
            const y = 135 - (v / maxVal) * 110;
            return `${x},${y}`;
          }).join(' ');
          linesSvg += `<polyline fill="none" stroke="${color}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" points="${pts}" />`;
          s.values.forEach((v, vIdx) => {
            const step = catCount > 1 ? 250 / (catCount - 1) : 125;
            const x = 35 + vIdx * step;
            const y = 135 - (v / maxVal) * 110;
            linesSvg += `<circle cx="${x}" cy="${y}" r="3.5" fill="${color}" />`;
          });
        });
        el.categories.forEach((cat, cIdx) => {
          const step = catCount > 1 ? 250 / (catCount - 1) : 125;
          const x = 35 + cIdx * step;
          linesSvg += `<text x="${x}" y="150" text-anchor="middle" font-size="9" fill="#64748B">${escapeHtml(cat)}</text>`;
        });
        chartSvg = `
          <svg viewBox="0 0 300 170" style="width:100%; height:100%;">
            <line x1="30" y1="25" x2="290" y2="25" stroke="#E2E8F0" stroke-dasharray="2,2" stroke-width="0.8"/>
            <line x1="30" y1="80" x2="290" y2="80" stroke="#E2E8F0" stroke-dasharray="2,2" stroke-width="0.8"/>
            <line x1="30" y1="135" x2="290" y2="135" stroke="#94A3B8" stroke-width="1.2"/>
            ${linesSvg}
          </svg>
        `;
      } else if (el.chartType === 'pie') {
        const s = el.series[0];
        let slicesSvg = '';
        if (s && s.values.length > 0) {
          const total = s.values.reduce((a, b) => a + b, 0) || 1;
          let cum = 0;
          const cx = 95, cy = 80, r = 58;
          s.values.forEach((val, idx) => {
            const angle = (val / total) * 360;
            const sA = cum;
            const eA = cum + angle;
            cum = eA;
            const sRad = (sA - 90) * (Math.PI / 180);
            const eRad = (eA - 90) * (Math.PI / 180);
            const x1 = cx + r * Math.cos(sRad);
            const y1 = cy + r * Math.sin(sRad);
            const x2 = cx + r * Math.cos(eRad);
            const y2 = cy + r * Math.sin(eRad);
            const largeArc = angle > 180 ? 1 : 0;
            const d = `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${largeArc} 1 ${x2} ${y2} Z`;
            slicesSvg += `<path d="${d}" fill="${palette[idx % palette.length]}" stroke="#FFFFFF" stroke-width="1" />`;
          });
        }
        chartSvg = `<svg viewBox="0 0 190 160" style="width:100%; height:100%;">${slicesSvg}</svg>`;
      }

      const legendHtml = (el.chartType === 'pie' ? el.categories : el.series.map(s => s.name)).map((label, idx) => `
        <div style="display:flex; align-items:center; gap:3px; margin: 0 4px;">
          <span style="display:inline-block; width:8px; height:8px; background-color:${palette[idx % palette.length]}; border-radius:1px;"></span>
          <span>${escapeHtml(label)}</span>
        </div>
      `).join('');

      return `
        <div style="position:absolute; left:${el.left}%; top:${el.top}%; width:${el.width}%; height:${el.height}%; z-index:${el.zIndex}; background:#FFFFFF; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 1px 3px rgba(0,0,0,0.06); padding:8px; display:flex; flex-direction:column; box-sizing:border-box;">
          ${el.title ? `<div style="text-align:center; font-size:12px; font-weight:bold; color:#1E293B; margin-bottom:4px;">${escapeHtml(el.title)}</div>` : ''}
          <div style="flex:1; width:100%; height:100%; min-height:0;">${chartSvg}</div>
          <div style="display:flex; justify-content:center; flex-wrap:wrap; font-size:10px; color:#475569; margin-top:4px;">${legendHtml}</div>
        </div>
      `;
    }

    return '';
  }).join('\n');

  const hiddenBadge = slide.isHidden ? `
    <div style="position:absolute; top:16px; right:20px; z-index:40; background:#E11D48; color:#FFFFFF; font-size:12px; font-weight:bold; padding:4px 10px; border-radius:9999px; box-shadow:0 2px 4px rgba(0,0,0,0.2);">
      숨김 슬라이드
    </div>
  ` : '';

  const notesHtml = slide.notes ? `
    <div style="position:absolute; bottom:12px; left:20px; max-width:600px; background:rgba(15,23,42,0.85); color:#F1F5F9; font-size:11px; padding:6px 12px; border-radius:6px; z-index:30;">
      <span style="color:#F59E0B; font-weight:bold;">📝 메모:</span> ${escapeHtml(slide.notes)}
    </div>
  ` : '';

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>${escapeHtml(slide.title)}</title>
  <style>
    * { box-sizing: border-box; }
    body {
      margin: 0;
      padding: 0;
      width: 1280px;
      height: 720px;
      overflow: hidden;
      font-family: -apple-system, BlinkMacSystemFont, "Malgun Gothic", "맑은 고딕", Dotum, sans-serif;
      background: #F1F5F9;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    .slide-canvas {
      position: relative;
      width: 1280px;
      height: 720px;
      background: ${slide.backgroundColor || '#FFFFFF'};
      overflow: hidden;
    }
    .accent-bar {
      position: absolute;
      top: 0; left: 0; right: 0;
      height: 6px;
      background: linear-gradient(90deg, #F97316, #EAB308, #6366F1);
      z-index: 50;
    }
    .slide-number {
      position: absolute;
      bottom: 12px;
      right: 20px;
      font-size: 13px;
      color: #94A3B8;
      font-family: monospace;
      z-index: 30;
    }
    .link {
      color: #2563EB;
      text-decoration: underline;
    }
  </style>
</head>
<body>
  <div class="slide-canvas">
    <div class="accent-bar"></div>
    ${hiddenBadge}
    ${elementsHtml}
    ${notesHtml}
    <div class="slide-number">${slideIndex + 1}</div>
  </div>
</body>
</html>`;
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

async function parsePptxFile(filePath) {
  const buf = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(buf);

  // Parse presentation.xml & rels
  const presXml = await zip.file('ppt/presentation.xml').async('text');
  const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('text');
  const pDom = new JSDOM(presXml, { contentType: 'text/xml' });
  const prDom = new JSDOM(presRelsXml, { contentType: 'text/xml' });

  let slideWidth = 12192000;
  let slideHeight = 6858000;
  const sldSz = getElementsByLocalName(pDom.window.document, 'sldSz')[0];
  if (sldSz) {
    slideWidth = parseInt(sldSz.getAttribute('cx') || '12192000', 10);
    slideHeight = parseInt(sldSz.getAttribute('cy') || '6858000', 10);
  }

  const relMap = new Map();
  for (const r of prDom.window.document.getElementsByTagName('Relationship')) {
    relMap.set(r.getAttribute('Id'), r.getAttribute('Target'));
  }

  const sldIds = pDom.window.document.getElementsByTagName('p:sldId');
  const slidePaths = [];
  for (let i = 0; i < sldIds.length; i++) {
    const rId = sldIds[i].getAttribute('r:id');
    const target = relMap.get(rId);
    if (target) {
      slidePaths.push(target.startsWith('ppt/') ? target : `ppt/${target}`);
    }
  }

  // Theme map
  const themeMap = {};
  const themeFiles = Object.keys(zip.files).filter(f => f.startsWith('ppt/theme/theme') && f.endsWith('.xml'));
  if (themeFiles.length > 0) {
    try {
      const tText = await zip.file(themeFiles[0]).async('text');
      const tDom = new JSDOM(tText, { contentType: 'text/xml' });
      const clrScheme = getElementsByLocalName(tDom.window.document, 'clrScheme')[0];
      if (clrScheme) {
        for (let i = 0; i < clrScheme.children.length; i++) {
          const child = clrScheme.children[i];
          const tag = child.localName;
          const srgb = getElementsByLocalName(child, 'srgbClr')[0];
          if (srgb) themeMap[tag] = `#${srgb.getAttribute('val')}`;
        }
      }
    } catch (e) {}
  }

  const parsedSlides = [];

  for (let idx = 0; idx < slidePaths.length; idx++) {
    const sPath = slidePaths[idx];
    const sXml = await zip.file(sPath).async('text');
    const sDom = new JSDOM(sXml, { contentType: 'text/xml' });
    const sDoc = sDom.window.document;

    // Rels
    const parts = sPath.split('/');
    const fname = parts.pop();
    const relsPath = `${parts.join('/')}/_rels/${fname}.rels`;
    const relsFile = zip.file(relsPath);
    const imageMap = new Map();
    const chartMap = new Map();
    let notesPath = undefined;

    if (relsFile) {
      const rDom = new JSDOM(await relsFile.async('text'), { contentType: 'text/xml' });
      for (const r of rDom.window.document.getElementsByTagName('Relationship')) {
        const type = r.getAttribute('Type') || '';
        const id = r.getAttribute('Id');
        let target = r.getAttribute('Target') || '';
        if (target.startsWith('../')) target = `ppt/${target.substring(3)}`;
        else if (!target.startsWith('ppt/')) target = `ppt/slides/${target}`;

        if (type.includes('image') && id) imageMap.set(id, target);
        else if (type.includes('chart') && id) chartMap.set(id, target);
        else if (type.includes('notesSlide')) notesPath = target;
      }
    }

    let slideNotes = undefined;
    if (notesPath && zip.file(notesPath)) {
      try {
        const nDom = new JSDOM(await zip.file(notesPath).async('text'), { contentType: 'text/xml' });
        const ts = Array.from(nDom.window.document.getElementsByTagName('a:t')).map(t => t.textContent).join(' ').trim();
        if (ts) slideNotes = ts;
      } catch (e) {}
    }

    const isHidden = sDoc.documentElement.getAttribute('show') === '0';
    let backgroundColor = undefined;
    const bgList = getElementsByLocalName(sDoc, 'bg');
    if (bgList.length > 0) {
      backgroundColor = parseDrawingColor(bgList[0], themeMap);
    }

    const elements = [];
    let zOrder = 1;

    // spTree
    const spTree = getElementsByLocalName(sDoc, 'spTree')[0];
    const shapeContainers = spTree ? Array.from(spTree.children) : [];

    for (const item of shapeContainers) {
      const localName = item.localName;

      if (localName === 'sp') {
        const coords = parseElementCoordinates(item, slideWidth, slideHeight);
        if (!coords) continue;

        let fillColor = undefined;
        let borderColor = undefined;
        let borderWidth = undefined;

        const spPrList = getElementsByLocalName(item, 'spPr');
        if (spPrList.length > 0) {
          fillColor = parseDrawingColor(spPrList[0], themeMap);
          const lnList = getElementsByLocalName(spPrList[0], 'ln');
          if (lnList.length > 0) {
            borderColor = parseDrawingColor(lnList[0], themeMap) || '#CBD5E1';
            const w = lnList[0].getAttribute('w');
            borderWidth = w ? Math.max(1, Math.round(parseInt(w, 10) / 12700)) : 1;
          }
        }

        let verticalAlign = 'top';
        let isVerticalText = false;
        let columnCount = 1;
        const txBodyList = getElementsByLocalName(item, 'txBody');
        if (txBodyList.length > 0) {
          const bodyPr = getElementsByLocalName(txBodyList[0], 'bodyPr')[0];
          if (bodyPr) {
            const anchor = bodyPr.getAttribute('anchor');
            if (anchor === 'ctr' || anchor === 'mid') verticalAlign = 'middle';
            else if (anchor === 'b') verticalAlign = 'bottom';

            const vert = bodyPr.getAttribute('vert');
            if (vert === 'eaVert' || vert === 'vert' || vert === 'vert270') isVerticalText = true;

            const numCol = bodyPr.getAttribute('numCol');
            if (numCol) columnCount = parseInt(numCol, 10);
          }
        }

        const phList = getElementsByLocalName(item, 'ph');
        const isTitle = phList.length > 0 && (phList[0].getAttribute('type') === 'title' || phList[0].getAttribute('type') === 'ctrTitle');

        const paragraphs = [];
        const pElements = getElementsByLocalName(item, 'p');
        for (const p of pElements) {
          paragraphs.push(parseParagraph(p, themeMap));
        }

        elements.push({
          id: `sp-${zOrder}`,
          type: 'shape',
          shapeType: 'rect',
          left: coords.left,
          top: coords.top,
          width: coords.width,
          height: coords.height,
          fillColor,
          borderColor,
          borderWidth,
          paragraphs,
          isTitle,
          verticalAlign,
          rotation: coords.rotation,
          isVerticalText,
          columnCount,
          zIndex: zOrder++,
        });
      } else if (localName === 'graphicFrame') {
        const coords = parseElementCoordinates(item, slideWidth, slideHeight);
        const chartList = getElementsByLocalName(item, 'chart');
        const tblList = getElementsByLocalName(item, 'tbl');

        if (coords && chartList.length > 0) {
          const rId = chartList[0].getAttribute('r:id');
          if (rId && chartMap.has(rId)) {
            const cPath = chartMap.get(rId);
            const cFile = zip.file(cPath);
            if (cFile) {
              const cXml = await cFile.async('text');
              const cDom = new JSDOM(cXml, { contentType: 'text/xml' });
              const chartData = parseChartXml(cDom.window.document);
              elements.push({
                id: `chart-${zOrder}`,
                type: 'chart',
                chartType: chartData.chartType,
                title: chartData.title,
                categories: chartData.categories,
                series: chartData.series,
                left: coords.left,
                top: coords.top,
                width: coords.width,
                height: coords.height,
                zIndex: zOrder++,
              });
            }
          }
        } else if (coords && tblList.length > 0) {
          const tbl = tblList[0];
          const trElements = getElementsByLocalName(tbl, 'tr');
          const rows = [];
          for (const tr of trElements) {
            const tcElements = getElementsByLocalName(tr, 'tc');
            const cells = [];
            for (const tc of tcElements) {
              if (tc.getAttribute('hMerge') === '1' || tc.getAttribute('vMerge') === '1') continue;
              const colSpan = parseInt(tc.getAttribute('gridSpan') || '1', 10);
              const rowSpan = parseInt(tc.getAttribute('rowSpan') || '1', 10);
              let fillColor = undefined;
              const tcPr = getElementsByLocalName(tc, 'tcPr')[0];
              if (tcPr) {
                const sFill = getDirectChildrenByLocalName(tcPr, 'solidFill')[0];
                if (sFill) fillColor = parseDrawingColor(sFill, themeMap);
              }
              const pElements = getElementsByLocalName(tc, 'p');
              const paragraphs = pElements.map(p => parseParagraph(p, themeMap));
              cells.push({ colSpan, rowSpan, fillColor, paragraphs });
            }
            if (cells.length > 0) rows.push({ cells });
          }
          elements.push({
            id: `tbl-${zOrder}`,
            type: 'table',
            left: coords.left,
            top: coords.top,
            width: coords.width,
            height: coords.height,
            rows,
            zIndex: zOrder++,
          });
        }
      } else if (localName === 'pic') {
        const coords = parseElementCoordinates(item, slideWidth, slideHeight);
        const blip = getElementsByLocalName(item, 'blip')[0];
        if (coords && blip) {
          const embedId = blip.getAttribute('r:embed');
          if (embedId && imageMap.has(embedId)) {
            const mPath = imageMap.get(embedId);
            const mFile = zip.file(mPath);
            if (mFile) {
              const mBuf = await mFile.async('nodebuffer');
              const ext = path.extname(mPath).toLowerCase();
              const mime = ext === '.png' ? 'image/png' : ext === '.jpg' || ext === '.jpeg' ? 'image/jpeg' : 'image/png';
              const dataUrl = `data:${mime};base64,${mBuf.toString('base64')}`;
              elements.push({
                id: `img-${zOrder}`,
                type: 'image',
                left: coords.left,
                top: coords.top,
                width: coords.width,
                height: coords.height,
                dataUrl,
                zIndex: zOrder++,
              });
            }
          }
        }
      } else if (localName === 'cxnSp') {
        const coords = parseElementCoordinates(item, slideWidth, slideHeight);
        if (coords) {
          let lineColor = '#334155';
          let lineWidth = 1;
          const ln = getElementsByLocalName(item, 'ln')[0];
          if (ln) {
            lineColor = parseDrawingColor(ln, themeMap) || '#334155';
            const w = ln.getAttribute('w');
            lineWidth = w ? Math.max(1, Math.round(parseInt(w, 10) / 12700)) : 1;
          }
          elements.push({
            id: `cxn-${zOrder}`,
            type: 'shape',
            shapeType: 'line',
            left: coords.left,
            top: coords.top,
            width: coords.width,
            height: coords.height,
            borderColor: lineColor,
            borderWidth: lineWidth,
            paragraphs: [],
            zIndex: zOrder++,
          });
        }
      } else if (localName === 'grpSp') {
        const childSps = getElementsByLocalName(item, 'sp');
        for (const childSp of childSps) {
          const spCoords = parseElementCoordinates(childSp, slideWidth, slideHeight);
          if (!spCoords) continue;
          let fillColor = '#E2E8F0';
          let borderColor = undefined;
          let borderWidth = 1;
          const spPr = getElementsByLocalName(childSp, 'spPr')[0];
          if (spPr) {
            fillColor = parseDrawingColor(spPr, themeMap) || '#E2E8F0';
            const ln = getElementsByLocalName(spPr, 'ln')[0];
            if (ln) {
              borderColor = parseDrawingColor(ln, themeMap) || '#CBD5E1';
              const w = ln.getAttribute('w');
              borderWidth = w ? Math.max(1, Math.round(parseInt(w, 10) / 12700)) : 1;
            }
          }
          const pElements = getElementsByLocalName(childSp, 'p');
          const paragraphs = pElements.map(p => parseParagraph(p, themeMap));
          elements.push({
            id: `grp-sp-${zOrder}`,
            type: 'shape',
            shapeType: 'rect',
            left: spCoords.left,
            top: spCoords.top,
            width: spCoords.width,
            height: spCoords.height,
            fillColor,
            borderColor,
            borderWidth,
            paragraphs,
            verticalAlign: 'middle',
            zIndex: zOrder++,
          });
        }
      }
    }

    const firstText = elements.find(el => el.paragraphs && el.paragraphs.length > 0)?.paragraphs[0]?.text || `슬라이드 ${idx + 1}`;
    parsedSlides.push({
      title: firstText,
      isHidden,
      backgroundColor,
      notes: slideNotes,
      elements,
    });
  }

  return { title: path.basename(filePath), slides: parsedSlides };
}

async function renderPresentation(filePath, prefix) {
  console.log(`\n렌더링 시작: ${filePath} (${prefix})...`);
  const pres = await parsePptxFile(filePath);
  console.log(`슬라이드 총 ${pres.slides.length}개 발견`);

  for (let idx = 0; idx < pres.slides.length; idx++) {
    const slide = pres.slides[idx];
    const htmlContent = await renderSlideToHtml(slide, idx, pres.slides.length, pres.title);
    const htmlPath = path.join(TEMP_HTML_DIR, `${prefix}_slide${idx + 1}.html`);
    const pngPath = path.join(OUT_DIR, `${prefix}_slide${idx + 1}.png`);

    fs.writeFileSync(htmlPath, htmlContent, 'utf-8');

    // Launch headless Chrome
    const cmd = `"${CHROME_PATH}" --headless=new --screenshot="${pngPath}" --window-size=1280,720 --hide-scrollbars "file:///${htmlPath.replace(/\\/g, '/')}"`;
    execSync(cmd, { stdio: 'pipe' });

    const size = fs.existsSync(pngPath) ? fs.statSync(pngPath).size : 0;
    console.log(`  ✓ [슬라이드 ${idx + 1}/${pres.slides.length}] ${path.basename(pngPath)} (${(size / 1024).toFixed(1)} KB)`);
  }
}

async function main() {
  console.log('4개 샘플 파일 24개 슬라이드 전체 PNG 렌더링 시작...');
  const t0 = Date.now();

  await renderPresentation('docs/office-upgrade/samples/real/PPT_A_사업보고_원본.pptx', 'PPT_A_원본');
  await renderPresentation('docs/office-upgrade/samples/real/PPT_A_사업보고_한쇼.pptx', 'PPT_A_한쇼');
  await renderPresentation('docs/office-upgrade/samples/real/PPT_B_교육자료_원본.pptx', 'PPT_B_원본');
  await renderPresentation('docs/office-upgrade/samples/real/PPT_B_교육자료_한쇼.pptx', 'PPT_B_한쇼');

  const totalPngs = fs.readdirSync(OUT_DIR).filter(f => f.endsWith('.png'));
  console.log(`\n총 ${totalPngs.length}개 PNG 생성 완료 (${((Date.now() - t0) / 1000).toFixed(1)}초 소요).`);
}

main().catch(err => {
  console.error('렌더링 중 오류:', err);
  process.exit(1);
});
