import PptxGenJS from 'pptxgenjs';
import { logos } from './assetManifest.js';
import path from 'path';

/**
 * Generates a PPTX file from Claude's structured slide data.
 *
 * @param {object} slideData - Structured slide data from Claude
 * @returns {Promise<Buffer>} PPTX file as a Node.js Buffer
 */
export async function generatePptx(slideData) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_16x9';

  const slide = pres.addSlide();

  // Background color
  const bg = stripHash(slideData.backgroundColor || '#FFFFFF');
  slide.background = { color: bg };

  // Left accent bar
  const accent = stripHash(slideData.accentColor || '#000000');
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 0.08, h: 5.625,
    fill: { color: accent },
    line: { color: accent },
  });

  const titleColor = stripHash(slideData.titleColor || '#000000');
  const bodyColor = stripHash(slideData.bodyColor || '#333333');
  const headingFont = slideData.headingFont || 'Calibri';
  const bodyFont = slideData.bodyFont || 'Calibri';

  // Title
  slide.addText(slideData.title || '', {
    x: 0.3, y: 0.4, w: 8.8, h: 1.0,
    fontSize: 32,
    bold: true,
    color: titleColor,
    fontFace: headingFont,
    wrap: true,
  });

  let bodyY = 1.6;

  // Optional subtitle
  if (slideData.subtitle) {
    slide.addText(slideData.subtitle, {
      x: 0.3, y: 1.5, w: 8.0, h: 0.6,
      fontSize: 18,
      color: bodyColor,
      fontFace: bodyFont,
      italic: true,
      wrap: true,
    });
    bodyY = 2.3;
  }

  // Bullet body
  if (Array.isArray(slideData.body) && slideData.body.length > 0) {
    const bullets = slideData.body.map((text) => ({
      text,
      options: { bullet: true, fontSize: 16, color: bodyColor, fontFace: bodyFont },
    }));
    slide.addText(bullets, {
      x: 0.3, y: bodyY, w: 8.8, h: 5.625 - bodyY - 0.3,
      wrap: true,
    });
  }

  // Logo (top-right corner)
  if (slideData.logoId) {
    const logoEntry = logos.find((l) => l.id === slideData.logoId);
    if (logoEntry) {
      const logoPath = path.join(process.cwd(), 'public', 'assets', 'logos', logoEntry.file);
      slide.addImage({
        path: logoPath,
        x: 8.5, y: 0.15, w: 1.2, h: 0.6,
        sizing: { type: 'contain', w: 1.2, h: 0.6 },
      });
    }
  }

  return await pres.write({ outputType: 'nodebuffer' });
}

function stripHash(hex) {
  return hex.startsWith('#') ? hex.slice(1) : hex;
}
