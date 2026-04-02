import Anthropic from '@anthropic-ai/sdk';
import { generatePptx } from '../../../lib/generatePptx.js';
import { logos } from '../../../lib/assetManifest.js';

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

export async function POST(request) {
  try {
    const { title: rawTitle, description: rawDescription, content: rawContent, colors = [], fonts = [], backgroundImage = null, templateShapes = [] } = await request.json();
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

    const systemPrompt = `You are a presentation strategist and visual designer. Your job is two things:
1. STORYTELLING — Sharpen the content so the slide lands a clear, memorable message. Rewrite bullets to be punchy, parallel, and outcome-focused. Add a subtitle that frames the "so what" for the audience.
2. DESIGN — Choose layout, colors, and an icon that reinforce the message and create strong visual hierarchy.

Return only a single valid JSON object matching this exact schema. No explanation, no markdown, no code fences.

{
  "layout": "title-subtitle-body",
  "label": null,
  "title": "Sharp, compelling slide title",
  "subtitle": "One sentence framing why this matters to the audience",
  "body": ["Punchy insight one", "Punchy insight two", "Punchy insight three"],
  "sections": null,
  "stats": null,
  "backgroundColor": "#1A3A5C",
  "titleColor": "#F0A500",
  "bodyColor": "#FFFFFF",
  "accentColor": "#F0A500",
  "headingFont": "${headingFont}",
  "bodyFont": "${bodyFont}",
  "logoId": "data-science"
}

Brand colors: ${colorList}
Brand fonts: heading ${headingFont}, body ${bodyFont}

Available icons — pick the one that best reinforces the message:
${logoList}

Layout guide:
- "title-only": bold single statement, section break, or striking statistic
- "title-body": facts or data points that stand on their own without extra framing
- "title-subtitle-body": subtitle reframes the audience's perspective or adds the "so what"
- "title-sections": content has 2–3 named groups (Challenge/Solution, Before/After, Problem/Approach/Outcome) — use sections array, set body to null
- "title-stats": the story IS the numbers — use stats array, set body to null
- "title-subtitle-stats": subtitle + big numbers for context-then-proof flow

Content field guide:
- label: optional ALL-CAPS eyebrow text (e.g. "CASE STUDY #1", "KEY FINDING", "PHASE 2") — use when content has a clear category or series label; otherwise null
- sections: array of { header, bullets } when content splits into 2–3 named groups; set body to null when using sections
- stats: array of { value, label } for 2–4 big numbers that prove the point; can appear alongside sections

Color rules:
- backgroundColor: use the darkest brand color for impact; light neutral only if content is data-heavy
- titleColor: must contrast sharply against backgroundColor
- accentColor: use the most vibrant brand color — applied to section headers, stats, bullets, and accent bar
- bodyColor: white on dark backgrounds, dark on light backgrounds
- all colors must be valid hex strings like #1A3A5C; if no brand colors detected, use professional neutral defaults

Word limits — these are hard limits, count the words:
- title: 5–7 words maximum
- subtitle: 10–14 words maximum
- each bullet: 7–9 words maximum, NO conjunctions joining two ideas
- section headers: 1–3 words maximum
- stat labels: 4–7 words maximum

Content rules:
- Rewrite and SHORTEN the provided content to meet the word limits above
- Every bullet must be one standalone insight, active voice
- The subtitle answers "why should the audience care right now?" in one tight phrase
- Always pick a logoId that reinforces the core message; only use null if truly nothing fits`;

    const userMessage = `Title: ${title}\nDescription: ${description}\nContent: ${content}`;

    const message = await client.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 1024,
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
    const pptxBuffer = await generatePptx(slideData, backgroundImage, templateShapes);
    const pptxBase64 = pptxBuffer.toString('base64');

    // Build HTML preview
    const previewHtml = buildPreviewHtml(slideData, logos, templateShapes, backgroundImage);

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

// Preview scale: PPTX LAYOUT_16x9 is 10" × 5.625"; preview canvas is 800px × 450px → 80 px/inch
const PX_PER_INCH = 80;

function buildPreviewHtml(s, logoManifest, templateShapes = [], backgroundImage = null) {
  const bg = esc(s.backgroundColor || '#FFFFFF');
  const titleColor = esc(s.titleColor || '#000000');
  const bodyColor = esc(s.bodyColor || '#333333');
  const accentColor = esc(s.accentColor || '#000000');
  const headingFont = esc(s.headingFont || 'Calibri');
  const bodyFont = esc(s.bodyFont || 'Calibri');

  // Background — photo takes priority over solid color
  const bgStyle = backgroundImage?.base64
    ? `background-image:url('data:${esc(backgroundImage.mimeType)};base64,${backgroundImage.base64}');background-size:cover;background-position:center;`
    : `background:${bg};`;

  // Template shapes — filter large shapes that sit in the main content zone
  // (they obscure text and are usually placeholders that slipped through XML parsing)
  const renderableShapes = templateShapes.filter((shape) => {
    if (shape.w > 3 && shape.h > 1.5) {
      const cX1 = 0.3, cY1 = 0.8, cX2 = 9.7, cY2 = 4.5;
      const overlapW = Math.max(0, Math.min(shape.x + shape.w, cX2) - Math.max(shape.x, cX1));
      const overlapH = Math.max(0, Math.min(shape.y + shape.h, cY2) - Math.max(shape.y, cY1));
      if (overlapW * overlapH > 2.0) return false; // sq inches
    }
    return true;
  });

  const shapesHtml = renderableShapes.map((shape) => {
    const left = Math.round(shape.x * PX_PER_INCH);
    const top = Math.round(shape.y * PX_PER_INCH);
    const width = Math.round(shape.w * PX_PER_INCH);
    const height = Math.round(shape.h * PX_PER_INCH);
    const opacity = shape.transparency > 0 ? (1 - shape.transparency / 100).toFixed(2) : '1';
    const rotate = shape.rotation ? `transform:rotate(${shape.rotation}deg);` : '';
    const fillColor = esc(shape.fillColor || '#000000');
    return `<div style="position:absolute;left:${left}px;top:${top}px;width:${width}px;height:${height}px;background:${fillColor};opacity:${opacity};${rotate}"></div>`;
  }).join('');

  // Logo / icon (top-right)
  let logoHtml = '';
  if (s.logoId) {
    const logoEntry = logoManifest.find((l) => l.id === s.logoId);
    if (logoEntry) {
      const logoSrc = esc(`/assets/logos/${logoEntry.file}`);
      logoHtml = `<img src="${logoSrc}" alt="logo" style="position:absolute;top:16px;right:20px;height:44px;object-fit:contain;">`;
    }
  }

  // Label (eyebrow text)
  const labelHtml = s.label
    ? `<div style="font-family:'${headingFont}',sans-serif;font-size:10px;font-weight:700;color:${accentColor};letter-spacing:2.5px;text-transform:uppercase;margin-bottom:4px;">${esc(s.label)}</div>`
    : '';

  // Subtitle — clamped to 2 lines max
  const subtitleHtml = s.subtitle
    ? `<div style="font-family:'${headingFont}',sans-serif;font-size:13px;font-style:italic;color:${bodyColor};margin-bottom:12px;line-height:1.45;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;">${esc(s.subtitle)}</div>`
    : '';

  // Flat bullets
  const bulletsHtml = Array.isArray(s.body) && s.body.length > 0
    ? s.body.map((b) =>
        `<div style="display:flex;align-items:flex-start;gap:8px;margin-bottom:6px;">
           <span style="width:6px;height:6px;border-radius:50%;background:${accentColor};flex-shrink:0;margin-top:4px;"></span>
           <span style="font-family:'${bodyFont}',sans-serif;font-size:13px;line-height:1.4;color:${bodyColor};">${esc(b)}</span>
         </div>`
      ).join('')
    : '';

  // Sections (named groups with sub-headers)
  const sectionsHtml = Array.isArray(s.sections) && s.sections.length > 0
    ? s.sections.map((section) =>
        `<div style="margin-bottom:8px;">
           <div style="font-family:'${headingFont}',sans-serif;font-size:13px;font-weight:700;color:${accentColor};margin-bottom:4px;text-transform:uppercase;letter-spacing:0.5px;">${esc(section.header)}</div>
           ${(section.bullets || []).map((b) =>
             `<div style="display:flex;align-items:flex-start;gap:7px;margin-bottom:3px;">
                <span style="width:4px;height:4px;border-radius:50%;background:${accentColor};flex-shrink:0;margin-top:5px;"></span>
                <span style="font-family:'${bodyFont}',sans-serif;font-size:12px;line-height:1.35;color:${bodyColor};">${esc(b)}</span>
              </div>`).join('')}
         </div>`).join('')
    : '';

  const contentHtml = sectionsHtml || bulletsHtml;

  // Stats (big numbers anchored to bottom)
  const statsHtml = Array.isArray(s.stats) && s.stats.length > 0
    ? `<div style="display:flex;gap:32px;position:absolute;bottom:24px;left:60px;right:52px;">
         ${s.stats.map((stat) =>
           `<div>
              <div style="font-family:'${headingFont}',sans-serif;font-size:28px;font-weight:700;color:${accentColor};line-height:1;">${esc(stat.value)}</div>
              <div style="font-family:'${bodyFont}',sans-serif;font-size:10px;color:${bodyColor};margin-top:3px;max-width:160px;line-height:1.3;">${esc(stat.label)}</div>
            </div>`).join('')}
       </div>`
    : '';

  const hasShapes = renderableShapes.length > 0;
  const accentBar = hasShapes ? '' : `<div style="position:absolute;top:0;left:0;width:6px;height:100%;background:${accentColor};"></div>`;

  return `<div style="width:800px;height:450px;${bgStyle}position:relative;border-radius:8px;overflow:hidden;padding:36px 48px 36px 56px;box-sizing:border-box;">
  ${shapesHtml}
  ${accentBar}
  ${logoHtml}
  <div style="position:relative;max-height:370px;overflow:hidden;">
    ${labelHtml}
    <div style="font-family:'${headingFont}',sans-serif;font-size:26px;font-weight:700;color:${titleColor};margin-bottom:8px;line-height:1.2;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;">${esc(s.title)}</div>
    ${subtitleHtml}
    ${contentHtml}
  </div>
  ${statsHtml}
</div>`;
}
