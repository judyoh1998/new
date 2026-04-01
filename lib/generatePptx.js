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
  const headingFont = slideData.headingFont || 'Calibri';
  const bodyFont    = slideData.bodyFont    || 'Calibri';
  const hasShapes   = templateShapes.length > 0;

  // --- Template overlay shapes (rendered before text so text sits on top) ---
  for (const shape of templateShapes) {
    slide.addShape(pres.ShapeType.rect, {
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

  // --- Label (eyebrow text above title) ---
  if (slideData.label) {
    slide.addText(slideData.label, {
      x: 0.3, y: 0.1, w: 8.8, h: 0.25,
      fontSize: 11, color: accentColor, fontFace: headingFont, charSpacing: 1.5,
    });
  }

  // --- Title ---
  const titleY = slideData.label ? 0.42 : 0.4;
  slide.addText(slideData.title || '', {
    x: 0.3, y: titleY, w: 8.8, h: 1.0,
    fontSize: 32,
    bold: true,
    color: titleColor,
    fontFace: headingFont,
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

  // --- Sections (named groups with sub-headers + bullets) ---
  if (Array.isArray(slideData.sections) && slideData.sections.length > 0) {
    let sY = bodyY;
    const statsPresent = Array.isArray(slideData.stats) && slideData.stats.length > 0;
    const maxY = statsPresent ? 4.1 : 5.3;
    for (const section of slideData.sections) {
      if (sY >= maxY) break;
      slide.addText(section.header || '', {
        x: 0.3, y: sY, w: 8.8, h: 0.38,
        fontSize: 15, bold: true, color: accentColor, fontFace: headingFont,
      });
      sY += 0.38;
      const bullets = (section.bullets || []).map(text => ({
        text,
        options: { bullet: true, fontSize: 13, color: bodyColor, fontFace: bodyFont },
      }));
      if (bullets.length > 0) {
        const bH = Math.min(bullets.length * 0.32 + 0.1, maxY - sY);
        slide.addText(bullets, { x: 0.3, y: sY, w: 8.8, h: bH, wrap: true });
        sY += bH + 0.15;
      }
    }
  } else if (Array.isArray(slideData.body) && slideData.body.length > 0) {
    // --- Bullet body ---
    const statsPresent = Array.isArray(slideData.stats) && slideData.stats.length > 0;
    const bodyH = statsPresent ? 4.1 - bodyY : 5.625 - bodyY - 0.3;
    const bullets = slideData.body.map(text => ({
      text,
      options: { bullet: true, fontSize: 16, color: bodyColor, fontFace: bodyFont },
    }));
    slide.addText(bullets, {
      x: 0.3, y: bodyY, w: 8.8, h: bodyH,
      wrap: true,
    });
  }

  // --- Stats (big numbers at bottom) ---
  if (Array.isArray(slideData.stats) && slideData.stats.length > 0) {
    const statsY = 4.2;
    const colW = 9.0 / slideData.stats.length;
    slideData.stats.forEach((stat, i) => {
      const xPos = 0.3 + i * colW;
      slide.addText(stat.value || '', {
        x: xPos, y: statsY, w: colW - 0.1, h: 0.7,
        fontSize: 32, bold: true, color: accentColor, fontFace: headingFont,
      });
      slide.addText(stat.label || '', {
        x: xPos, y: statsY + 0.72, w: colW - 0.1, h: 0.45,
        fontSize: 11, color: bodyColor, fontFace: bodyFont, wrap: true,
      });
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
