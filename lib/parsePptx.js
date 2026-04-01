import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';

const EMU = 914400; // EMU per inch
const ROT = 60000;  // units per degree

/**
 * Extracts brand colors, fonts, background image, and template shapes
 * from a PPTX file. Runs entirely in the browser (no Node.js APIs used).
 *
 * @param {ArrayBuffer|Buffer} buffer
 * @returns {{ colors, fonts, backgroundImage, templateShapes }}
 */
export async function parsePptx(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', isArray: () => false });

  // --- Colors & fonts from theme ---
  const { colors, fonts } = await extractTheme(zip, parser);

  // --- Background image: try slides 1–3 ---
  let backgroundImage = null;

  for (let n = 1; n <= 3; n++) {
    const slideFile = zip.file(`ppt/slides/slide${n}.xml`);
    if (!slideFile) break;

    const slideXml = await slideFile.async('string');
    const slideDoc = parser.parse(slideXml);

    const relsFile = zip.file(`ppt/slides/_rels/slide${n}.xml.rels`);
    const relsXml = relsFile ? await relsFile.async('string') : null;
    const relsDoc = relsXml ? parser.parse(relsXml) : null;

    backgroundImage = await extractBackground(zip, slideDoc, relsDoc);
    if (backgroundImage) break;
  }

  // --- Template shapes: prefer slide master, then slide layouts, then slide 1 ---
  let templateShapes = [];

  const masterFile = zip.file('ppt/slideMasters/slideMaster1.xml');
  if (masterFile) {
    const masterXml = await masterFile.async('string');
    const masterDoc = parser.parse(masterXml);
    templateShapes = extractTemplateShapes(masterDoc, 'p:sldMaster', colors);
  }

  if (templateShapes.length === 0) {
    for (let n = 1; n <= 3; n++) {
      const layoutFile = zip.file(`ppt/slideLayouts/slideLayout${n}.xml`);
      if (!layoutFile) break;
      const layoutXml = await layoutFile.async('string');
      const layoutDoc = parser.parse(layoutXml);
      templateShapes = extractTemplateShapes(layoutDoc, 'p:sldLayout', colors);
      if (templateShapes.length > 0) break;
    }
  }

  if (templateShapes.length === 0) {
    const slideFile = zip.file('ppt/slides/slide1.xml');
    if (slideFile) {
      const slideXml = await slideFile.async('string');
      const slideDoc = parser.parse(slideXml);
      templateShapes = extractTemplateShapes(slideDoc, 'p:sld', colors);
    }
  }

  return { colors, fonts, backgroundImage, templateShapes };
}

// ---------------------------------------------------------------------------

async function extractTheme(zip, parser) {
  const colors = [];
  const fonts = [];

  const themeFile = zip.file('ppt/theme/theme1.xml');
  if (!themeFile) return { colors, fonts };

  const xml = await themeFile.async('string');
  const doc = parser.parse(xml);

  try {
    const elements = doc?.['a:theme']?.['a:themeElements'];
    if (!elements) return { colors, fonts };

    const clrScheme = elements['a:clrScheme'];
    if (clrScheme) {
      const slots = ['a:dk1','a:lt1','a:dk2','a:lt2',
                     'a:accent1','a:accent2','a:accent3',
                     'a:accent4','a:accent5','a:accent6'];
      for (const slot of slots) {
        const hex = extractHex(clrScheme[slot]);
        if (hex) colors.push(hex);
      }
    }

    const fontScheme = elements['a:fontScheme'];
    if (fontScheme) {
      const major = fontScheme['a:majorFont']?.['a:latin']?.['@_typeface'];
      const minor = fontScheme['a:minorFont']?.['a:latin']?.['@_typeface'];
      if (major && major !== '+mj-lt') fonts.push(major);
      if (minor && minor !== '+mn-lt') fonts.push(minor);
    }
  } catch { /* partial results ok */ }

  return { colors, fonts };
}

// ---------------------------------------------------------------------------

async function extractBackground(zip, slideDoc, relsDoc) {
  try {
    const bg = slideDoc?.['p:sld']?.['p:cSld']?.['p:bg']?.['p:bgPr'];
    const blip = bg?.['a:blipFill']?.['a:blip'];
    const rId = blip?.['@_r:embed'] || blip?.['@_r:link'];
    if (!rId || !relsDoc) return null;

    // Resolve rId → media path
    let rels = relsDoc?.['Relationships']?.['Relationship'];
    if (!rels) return null;
    if (!Array.isArray(rels)) rels = [rels];
    const rel = rels.find(r => r['@_Id'] === rId);
    if (!rel) return null;

    // Target is like "../media/image1.jpeg" — normalise to zip path
    const target = rel['@_Target'].replace(/^\.\.\//, 'ppt/');
    const imgFile = zip.file(target);
    if (!imgFile) return null;

    const bytes = await imgFile.async('arraybuffer');
    const ext = target.split('.').pop().toLowerCase();
    const mimeType = ext === 'png' ? 'image/png' : ext === 'gif' ? 'image/gif' : 'image/jpeg';

    return await compressImage(bytes, mimeType);
  } catch {
    return null;
  }
}

// Compress image to 1280×720 JPEG @80% quality using browser Canvas API
function compressImage(arrayBuffer, mimeType) {
  return new Promise((resolve) => {
    try {
      const blob = new Blob([arrayBuffer], { type: mimeType });
      const url = URL.createObjectURL(blob);
      const img = new Image();
      img.onload = () => {
        try {
          const canvas = document.createElement('canvas');
          canvas.width = 1280;
          canvas.height = 720;
          const ctx = canvas.getContext('2d');
          ctx.drawImage(img, 0, 0, 1280, 720);
          const dataUrl = canvas.toDataURL('image/jpeg', 0.8);
          URL.revokeObjectURL(url);
          resolve({ base64: dataUrl.split(',')[1], mimeType: 'image/jpeg' });
        } catch {
          URL.revokeObjectURL(url);
          resolve(null);
        }
      };
      img.onerror = () => { URL.revokeObjectURL(url); resolve(null); };
      img.src = url;
    } catch {
      resolve(null);
    }
  });
}

// ---------------------------------------------------------------------------

function extractTemplateShapes(doc, rootTag, themeColors) {
  const shapes = [];
  try {
    const spTree = doc?.[rootTag]?.['p:cSld']?.['p:spTree'];
    if (!spTree) return shapes;
    collectShapes(spTree, themeColors, shapes);
  } catch { /* return whatever collected */ }
  return shapes;
}

function collectShapes(spTree, themeColors, shapes) {
  // Direct shapes
  let spList = spTree['p:sp'];
  if (spList) {
    if (!Array.isArray(spList)) spList = [spList];
    for (const sp of spList) {
      try { extractShape(sp, themeColors, shapes); } catch { /* skip */ }
    }
  }

  // Recurse into group shapes
  let grpList = spTree['p:grpSp'];
  if (grpList) {
    if (!Array.isArray(grpList)) grpList = [grpList];
    for (const grp of grpList) {
      try { collectShapes(grp, themeColors, shapes); } catch { /* skip */ }
    }
  }
}

function extractShape(sp, themeColors, shapes) {
  // Skip OOXML content placeholders (title frame, body frame, content area, etc.)
  if (sp['p:nvSpPr']?.['p:nvPr']?.['p:ph'] !== undefined) return;

  const spPr = sp['p:spPr'];
  if (!spPr) return;

  const xfrm = spPr['a:xfrm'];
  if (!xfrm) return;

  const x = (Number(xfrm['a:off']?.['@_x']) || 0) / EMU;
  const y = (Number(xfrm['a:off']?.['@_y']) || 0) / EMU;
  const w = (Number(xfrm['a:ext']?.['@_cx']) || 0) / EMU;
  const h = (Number(xfrm['a:ext']?.['@_cy']) || 0) / EMU;

  // Skip zero-size shapes only
  if (w < 0.05 && h < 0.05) return;

  const rotation = (Number(xfrm['@_rot']) || 0) / ROT;

  // Fill color — check spPr first, then style fallback
  const solidFill = spPr['a:solidFill'] || sp['p:style']?.['a:fillRef'];
  let fillColor = null;
  let transparency = 0;

  if (solidFill) {
    const srgb = solidFill['a:srgbClr'];
    if (srgb) {
      fillColor = `#${srgb['@_val']}`;
      const alphaVal = Number(srgb['a:alpha']?.['@_val']);
      if (alphaVal) transparency = Math.round(100 - alphaVal / 1000);
    }
    const schemeClr = solidFill['a:schemeClr'];
    if (!fillColor && schemeClr) {
      fillColor = resolveSchemeColor(schemeClr['@_val'], themeColors);
      const alphaVal = Number(schemeClr['a:alpha']?.['@_val']);
      if (alphaVal) transparency = Math.round(100 - alphaVal / 1000);
    }
  }

  if (!fillColor) return;

  // Skip content placeholders (shapes with actual text)
  const txBody = sp['p:txBody'];
  if (txBody) {
    const paras = txBody['a:p'];
    const paraList = Array.isArray(paras) ? paras : paras ? [paras] : [];
    const hasText = paraList.some(p => {
      const runs = p['a:r'];
      const runList = Array.isArray(runs) ? runs : runs ? [runs] : [];
      return runList.some(r => r['a:t']?.trim());
    });
    if (hasText) return;
  }

  shapes.push({ x, y, w, h, fillColor, rotation, transparency });
}

// Map OOXML scheme color names to theme colors array indices
const SCHEME_MAP = {
  dk1: 0, lt1: 1, dk2: 2, lt2: 3,
  accent1: 4, accent2: 5, accent3: 6,
  accent4: 7, accent5: 8, accent6: 9,
};

function resolveSchemeColor(name, themeColors) {
  const idx = SCHEME_MAP[name];
  return (idx !== undefined && themeColors[idx]) ? themeColors[idx] : null;
}

function extractHex(entry) {
  const srgb = entry?.['a:srgbClr']?.['@_val'];
  if (srgb) return `#${srgb}`;
  const sys = entry?.['a:sysClr']?.['@_lastClr'];
  if (sys) return `#${sys}`;
  return null;
}
