// PPT 카드 첫 슬라이드 미리보기(K8)가 큰 화면 뷰어와 같은 렌더러로 실제 슬라이드 내용을 그리는지 확인하는 회귀 테스트
// @vitest-environment jsdom
import { describe, test, expect } from 'vitest';
import fs from 'fs';
import path from 'path';
import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { parsePptxPresentation } from '../src/utils/pptxViewer';
import { PptxSlideThumbnail, PPTX_DESIGN_WIDTH } from '../src/components/PptxSlideRenderer';

const file = path.join(__dirname, '..', 'docs', 'office-upgrade', 'samples', 'real', 'PPT_A_사업보고_원본.pptx');
const dataUrl = `data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,${fs.readFileSync(file).toString('base64')}`;

describe('PPT 카드 첫 슬라이드 (K8)', () => {
  test('첫 슬라이드의 제목·본문이 카드 크기로 축소되어 그려진다', async () => {
    const pres = await parsePptxPresentation(dataUrl);
    expect(pres.slides.length).toBeGreaterThan(1);
    const html = renderToStaticMarkup(
      React.createElement(PptxSlideThumbnail, { slide: pres.slides[0], aspect: pres.aspectRatio, width: 480 }),
    );
    expect(html).toContain('data-testid="pptx-card-slide"');
    expect(html).toContain('2026년 하반기 사업 보고');
    expect(html).not.toContain('PRESENTATION SLIDE');
    expect(html).toContain(`scale(${480 / PPTX_DESIGN_WIDTH})`);
  });
});
