import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';

/**
 * Extracts brand colors and fonts from a PPTX file buffer.
 * Reads ppt/theme/theme1.xml inside the PPTX (which is a ZIP archive).
 *
 * @param {Buffer} buffer - Raw PPTX file buffer
 * @returns {{ colors: string[], fonts: string[] }}
 *   colors: hex strings ordered by theme prominence (dark1, light1, dark2, light2, accent1–6)
 *   fonts: font family names [headingFont, bodyFont, ...]
 */
export async function parsePptx(buffer) {
  const zip = await JSZip.loadAsync(buffer);

  const themeFile = zip.file('ppt/theme/theme1.xml');
  if (!themeFile) {
    return { colors: [], fonts: [] };
  }

  const xml = await themeFile.async('string');
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });
  const doc = parser.parse(xml);

  const colors = [];
  const fonts = [];

  try {
    const fmtScheme = doc?.['a:theme']?.['a:themeElements'];
    if (!fmtScheme) return { colors: [], fonts: [] };

    // Extract colors from color scheme
    const clrScheme = fmtScheme?.['a:clrScheme'];
    if (clrScheme) {
      const slots = ['a:dk1', 'a:lt1', 'a:dk2', 'a:lt2',
                     'a:accent1', 'a:accent2', 'a:accent3',
                     'a:accent4', 'a:accent5', 'a:accent6'];
      for (const slot of slots) {
        const entry = clrScheme[slot];
        if (!entry) continue;
        const hex = extractHex(entry);
        if (hex) colors.push(hex);
      }
    }

    // Extract fonts from font scheme
    const fontScheme = fmtScheme?.['a:fontScheme'];
    if (fontScheme) {
      const majorFont = fontScheme?.['a:majorFont']?.['a:latin']?.['@_typeface'];
      const minorFont = fontScheme?.['a:minorFont']?.['a:latin']?.['@_typeface'];
      if (majorFont && majorFont !== '+mj-lt') fonts.push(majorFont);
      if (minorFont && minorFont !== '+mn-lt') fonts.push(minorFont);
    }
  } catch {
    // Partial extraction is fine — return whatever was collected
  }

  return { colors, fonts };
}

function extractHex(entry) {
  // Colors can be srgbClr (hex) or sysClr (system color with lastClr fallback)
  const srgb = entry?.['a:srgbClr']?.['@_val'];
  if (srgb) return `#${srgb}`;

  const sysLastClr = entry?.['a:sysClr']?.['@_lastClr'];
  if (sysLastClr) return `#${sysLastClr}`;

  return null;
}
