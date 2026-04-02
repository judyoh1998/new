import PptxGenJS from 'pptxgenjs';
import { logos } from './assetManifest.js';
import path from 'path';

/**
 * Generates a PPTX file from Claude's structured slide data,
 * an optional extracted background image, and extracted template shapes.
 *
 * @param {object} slideData      - Structured slide data from Claude
 * @param {object|null} backgroundImage - { base64, mimeType } or null
 * @param {Array}  templateShapes - Array of { x, y, w, h, fillColor, rotation, transparency }
 * @returns {Promise<Buffer>}
 */
export async function generatePptx(slideData, backgroundImage = null, templateShapes = []) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_16x9';

  const slide = pres.addSlide();

  // --- Background ---
  if (backgroundImage?.base64) {
    slide.background = { data: `data:${backgroundImage.mimeType};base64,${backgroundImage.base64}` };
  } else {
    slide.background = { color: stripHash(slideData.backgroundColor || '#FFFFFF') };
  }

  const titleColor  = stripHash(slideData.titleColor  || '#000000');
  const bodyColor   = stripHash(slideData.bodyColor   || '#333333');
  const accentColor = stripHash(slideData.accentColor || '#000000');
  const headingFont = slideData.headingFont || 'Verdana';
  const bodyFont    = slideData.bodyFont    || 'Verdana';
  const hasShapes   = templateShapes.length > 0;

  // --- Template overlay shapes (rendered before text so text sits on top) ---
  for (const shape of templateShapes) {
    slide.addShape(pres.ShapeType[shape.type], {
      x: shape.x,
      y: shape.y,
      w: shape.w,
      h: shape.h,
      fill: {
        color: stripHash(shape.fillColor),
        transparency: shape.transparency || 0,
      },
      line: { type: 'none' },
      rotate: shape.rotation || 0,
    });
  }

  // --- Accent bar — only when no template shapes ---
  if (!hasShapes) {
    slide.addShape(pres.ShapeType.rect, {
      x: 0, y: 0, w: 0.08, h: 5.625,
      fill: { color: accentColor },
      line: { color: accentColor },
    });
  }

  // --- Title ---
  slide.addText(slideData.title || '', {
    x: 0.3, y: 0.4, w: 8.8, h: 1.0,
    fontSize: 32,
    bold: slideData.titleBold ?? true,
    color: slideData.titleColor || titleColor,
    fontFace: slideData.titleFont || headingFont,
    wrap: true,
  });

  let bodyY = 1.6;

  // --- Subtitle ---
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

  // --- Bullet body ---
  if (Array.isArray(slideData.body) && slideData.body.length > 0) {
    const bullets = slideData.body.map(text => ({
      text: item.text ?? item,
      options: { bullet: item.bullet ?? false, fontSize: 14, color: bodyColor, fontFace: bodyFont },
    }));
    slide.addText(bullets, {
      x: 0.3, y: bodyY, w: 8.8, h: 5.625 - bodyY - 0.3,
      wrap: true,
    });
  }

  // --- Logo (top-right) ---
  if (slideData.logoId) {
    const logoEntry = logos.find(l => l.id === slideData.logoId);
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
