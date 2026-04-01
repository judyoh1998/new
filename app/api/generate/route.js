import Anthropic from '@anthropic-ai/sdk';
import { generatePptx } from '../../../lib/generatePptx.js';
import { logos } from '../../../lib/assetManifest.js';

export const maxDuration = 60;

export async function POST(request) {
  if (!process.env.ANTHROPIC_API_KEY) {
    return Response.json({ error: 'ANTHROPIC_API_KEY is not configured in environment variables.' }, { status: 500 });
  }

  const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

  try {
    const {
      title, description, content,
      colors = [], fonts = [],
      backgroundImage = null,
      templateShapes = [],
    } = await request.json();

    if (!title?.trim() || !description?.trim() || !content?.trim()) {
      return Response.json({ error: 'Title, description and content are required.' }, { status: 400 });
    }

    const colorList = colors.length > 0 ? colors.join(', ') : 'no brand colors detected';
    const headingFont = fonts[0] || 'Calibri';
    const bodyFont = fonts[1] || fonts[0] || 'Calibri';

    const logoList = logos.length > 0
      ? logos.map((l) => `- ${l.id}: ${l.description}`).join('\n')
      : '(none available)';

    // Describe extracted template shapes so Claude positions content appropriately
    const templateDescription = templateShapes.length > 0
      ? `\nThe slide template has these overlay shapes extracted from the uploaded brand file:\n${
          templateShapes.map(s =>
            `- x:${s.x.toFixed(1)}", y:${s.y.toFixed(1)}", w:${s.w.toFixed(1)}", h:${s.h.toFixed(1)}", color:${s.fillColor}`
          ).join('\n')
        }\nPosition text title and bullets to sit within the largest overlay shape's area.`
      : '';

    const systemPrompt = `You are a professional slide designer. Output only a single JSON object — no explanation, no markdown, no code fences.

Brand colors (in order of prominence): ${colorList}
Brand fonts — heading: ${headingFont}, body: ${bodyFont}
${templateDescription}

Available logos:
${logoList}

Return exactly this JSON schema:
{
  "title": "string — refined slide headline",
  "subtitle": "string or null — optional subtitle, max 15 words",
  "body": ["string", "..."],
  "layout": "title-only" | "title-body" | "title-subtitle-body",
  "backgroundColor": "hex color from brand palette",
  "titleColor": "hex color from brand palette",
  "bodyColor": "hex color from brand palette",
  "headingFont": "heading font name",
  "bodyFont": "body font name",
  "accentColor": "hex color from brand palette for decorative elements",
  "logoId": "one of the available logo ids, or null"
}

Rules:
- body must have 2–5 items, each max 12 words
- All color values must be valid hex strings (e.g. #1A3A5C)
- If no brand colors were detected, use professional neutral defaults
- Choose logoId only if it genuinely fits the slide content; otherwise null`;

    const userMessage = `Title: ${title}\nDescription: ${description}\nContent: ${content}`;

    const message = await client.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 800,
      system: systemPrompt,
      messages: [{ role: 'user', content: userMessage }],
    });

    let rawText = message.content[0]?.text || '';
    rawText = rawText.replace(/^```json\s*/i, '').replace(/\s*```$/, '').trim();
    const slideData = JSON.parse(rawText);

    const pptxBuffer = await generatePptx(slideData, backgroundImage, templateShapes);
    const pptxBase64 = pptxBuffer.toString('base64');
    const previewHtml = buildPreviewHtml(slideData, logos, backgroundImage, templateShapes);

    return Response.json({ previewHtml, pptxBase64 });
  } catch (err) {
    console.error('Generate error:', err);
    return Response.json({ error: `Failed to generate slide: ${err?.message || 'Unknown error'}` }, { status: 500 });
  }
}

// ---------------------------------------------------------------------------

function esc(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function hexToRgb(hex) {
  const clean = hex.replace('#', '');
  const r = parseInt(clean.slice(0, 2), 16);
  const g = parseInt(clean.slice(2, 4), 16);
  const b = parseInt(clean.slice(4, 6), 16);
  return `${r},${g},${b}`;
}

function buildPreviewHtml(s, logoManifest, backgroundImage, templateShapes) {
  const bg = esc(s.backgroundColor || '#FFFFFF');
  const titleColor = esc(s.titleColor || '#000000');
  const bodyColor = esc(s.bodyColor || '#333333');
  const accentColor = esc(s.accentColor || '#000000');
  const headingFont = esc(s.headingFont || 'Calibri');
  const bodyFont = esc(s.bodyFont || 'Calibri');
  const hasShapes = templateShapes && templateShapes.length > 0;

  // Background — photo if extracted, else solid color
  const bgStyle = backgroundImage
    ? `background-image:url(data:${backgroundImage.mimeType};base64,${backgroundImage.base64});background-size:cover;background-position:center;`
    : `background:${bg};`;

  // Template shapes as absolutely-positioned divs
  const shapeDivs = hasShapes
    ? templateShapes.map(shape => {
        const left   = (shape.x / 10 * 100).toFixed(1);
        const top    = (shape.y / 5.625 * 100).toFixed(1);
        const width  = (shape.w / 10 * 100).toFixed(1);
        const height = (shape.h / 5.625 * 100).toFixed(1);
        const alpha  = ((100 - (shape.transparency || 0)) / 100).toFixed(2);
        const rgb    = hexToRgb(shape.fillColor);
        const rot    = shape.rotation ? `transform:rotate(${shape.rotation}deg);transform-origin:center;` : '';
        return `<div style="position:absolute;left:${left}%;top:${top}%;width:${width}%;height:${height}%;background:rgba(${rgb},${alpha});${rot}overflow:hidden;"></div>`;
      }).join('')
    : '';

  // Accent bar — only shown when no template shapes
  const accentBar = hasShapes
    ? ''
    : `<div style="position:absolute;top:0;left:0;width:8px;height:100%;background:${accentColor};"></div>`;

  // Logo
  let logoHtml = '';
  if (s.logoId) {
    const logoEntry = logoManifest.find((l) => l.id === s.logoId);
    if (logoEntry) {
      logoHtml = `<img src="${esc(`/assets/logos/${logoEntry.file}`)}" alt="logo" style="position:absolute;top:16px;right:20px;height:44px;object-fit:contain;z-index:2;">`;
    }
  }

  // Subtitle
  const subtitleHtml = s.subtitle
    ? `<div style="font-family:'${headingFont}',sans-serif;font-size:16px;font-style:italic;color:${bodyColor};margin-bottom:16px;position:relative;z-index:2;">${esc(s.subtitle)}</div>`
    : '';

  // Bullets
  const bulletsHtml = Array.isArray(s.body) && s.body.length > 0
    ? s.body.map(b =>
        `<div style="display:flex;align-items:flex-start;gap:8px;margin-bottom:8px;position:relative;z-index:2;">
           <span style="width:7px;height:7px;border-radius:50%;background:${accentColor};flex-shrink:0;margin-top:5px;"></span>
           <span style="font-family:'${bodyFont}',sans-serif;font-size:14px;color:${bodyColor};">${esc(b)}</span>
         </div>`
      ).join('')
    : '';

  return `<div style="width:800px;height:450px;${bgStyle}position:relative;border-radius:8px;overflow:hidden;padding:40px 52px 40px 60px;box-sizing:border-box;">
  ${shapeDivs}
  ${accentBar}
  ${logoHtml}
  <div style="font-family:'${headingFont}',sans-serif;font-size:30px;font-weight:700;color:${titleColor};margin-bottom:10px;line-height:1.2;position:relative;z-index:2;">${esc(s.title)}</div>
  ${subtitleHtml}
  ${bulletsHtml}
</div>`;
}
