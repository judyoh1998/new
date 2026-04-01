import Anthropic from '@anthropic-ai/sdk';
import { generatePptx } from '../../../lib/generatePptx.js';
import { logos } from '../../../lib/assetManifest.js';

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

export async function POST(request) {
  try {
    const { title: rawTitle, description: rawDescription, content: rawContent, colors = [], fonts = [] } = await request.json();
    const title = rawTitle?.trim();
    const description = rawDescription?.trim();
    const content = rawContent?.trim();

    if (!title || !description || !content) {
      return Response.json({ error: 'All fields are required.' }, { status: 400 });
    }

    // Build Claude prompt
    const colorList = colors.length > 0 ? colors.join(', ') : 'no brand colors detected';
    const headingFont = fonts[0] || 'Calibri';
    const bodyFont = fonts[1] || fonts[0] || 'Calibri';

    const logoList = logos.length > 0
      ? logos.map((l) => `- ${l.id}: ${l.description}`).join('\n')
      : '(none available)';

    const systemPrompt = `You are a professional slide designer.
Return only a single valid JSON object matching this exact schema. No explanation, no markdown, no code fences.

{
  "layout": "title-body",
  "title": "Slide title here",
  "subtitle": null,
  "body": ["Bullet one", "Bullet two", "Bullet three"],
  "backgroundColor": "#FFFFFF",
  "titleColor": "#1A3A5C",
  "bodyColor": "#333333",
  "accentColor": "#2E86AB",
  "headingFont": "${headingFont}",
  "bodyFont": "${bodyFont}",
  "logoId": null
}

Brand colors: ${colorList}
Brand fonts: heading ${headingFont}, body ${bodyFont}

Available logos:
${logoList}

Rules:
- layout must be one of: "title-only", "title-body", "title-subtitle-body"
- body must have 2 to 5 items; each item max 12 words
- subtitle is a string when layout is "title-subtitle-body", otherwise null
- all colors must be valid hex strings like #1A3A5C
- use brand colors for backgroundColor, titleColor, accentColor; if none detected, use professional neutral defaults
- choose logoId only if it genuinely fits the content; otherwise null`;

    const userMessage = `Title: ${title}\nDescription: ${description}\nContent: ${content}`;

    const message = await client.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 800,
      system: systemPrompt,
      messages: [{ role: 'user', content: userMessage }],
    });

    // Parse Claude's JSON response
    let rawText = message.content[0]?.text || '';
    rawText = rawText.replace(/^```json\s*/i, '').replace(/\s*```$/, '').trim();
    
    let slideData;
    try {
      slideData = JSON.parse(rawText);
    } catch (parseErr) {
      console.error('JSON parse error:', parseErr, 'Raw text:', rawText);
      return Response.json({ error: 'Invalid response format from AI.' }, { status: 500 });
    }

    // Generate PPTX
    const pptxBuffer = await generatePptx(slideData);
    const pptxBase64 = pptxBuffer.toString('base64');

    // Build HTML preview
    const previewHtml = buildPreviewHtml(slideData, logos);

    return Response.json({ previewHtml, pptxBase64 });
  } catch (err) {
    console.error('Generate error:', err);
    return Response.json({ error: 'Failed to generate slide. Please try again.' }, { status: 500 });
  }
}

function esc(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function buildPreviewHtml(s, logoManifest) {
  const bg = esc(s.backgroundColor || '#FFFFFF');
  const titleColor = esc(s.titleColor || '#000000');
  const bodyColor = esc(s.bodyColor || '#333333');
  const accentColor = esc(s.accentColor || '#000000');
  const headingFont = esc(s.headingFont || 'Calibri');
  const bodyFont = esc(s.bodyFont || 'Calibri');

  // Logo
  let logoHtml = '';
  if (s.logoId) {
    const logoEntry = logoManifest.find((l) => l.id === s.logoId);
    if (logoEntry) {
      const logoSrc = esc(`/assets/logos/${logoEntry.file}`);
      logoHtml = `<img src="${logoSrc}" alt="logo" style="position:absolute;top:16px;right:20px;height:44px;object-fit:contain;">`;
    }
  }

  // Subtitle
  const subtitleHtml = s.subtitle
    ? `<div style="font-family:'${headingFont}',sans-serif;font-size:16px;font-style:italic;color:${bodyColor};margin-bottom:16px;">${esc(s.subtitle)}</div>`
    : '';

  // Bullets
  const bulletsHtml = Array.isArray(s.body) && s.body.length > 0
    ? s.body.map((b) =>
        `<div style="display:flex;align-items:flex-start;gap:8px;margin-bottom:8px;">
           <span style="width:7px;height:7px;border-radius:50%;background:${accentColor};flex-shrink:0;margin-top:5px;"></span>
           <span style="font-family:'${bodyFont}',sans-serif;font-size:14px;color:${bodyColor};">${esc(b)}</span>
         </div>`
      ).join('')
    : '';

  return `<div style="width:800px;height:450px;background:${bg};position:relative;border-radius:8px;overflow:hidden;padding:40px 52px 40px 60px;box-sizing:border-box;">
  <div style="position:absolute;top:0;left:0;width:8px;height:100%;background:${accentColor};"></div>
  ${logoHtml}
  <div style="font-family:'${headingFont}',sans-serif;font-size:30px;font-weight:700;color:${titleColor};margin-bottom:10px;line-height:1.2;">${esc(s.title)}</div>
  ${subtitleHtml}
  ${bulletsHtml}
</div>`;
}
