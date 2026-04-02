import Anthropic from '@anthropic-ai/sdk';
import { logos } from '../../../lib/assetManifest.js';

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

function buildClaudeMessages(userMessage, backgroundImage) {
  if (!backgroundImage?.base64) {
    return [{ role: 'user', content: userMessage }];
  }
  return [{
    role: 'user',
    content: [
      {
        type: 'image',
        source: { type: 'base64', media_type: 'image/jpeg', data: backgroundImage.base64 },
      },
      { type: 'text', text: userMessage },
    ],
  }];
}

export async function POST(request) {
  try {
    const { title, description, content, colors = [], fonts = [], backgroundImage = null } = await request.json();

    if (!title?.trim() || !description?.trim() || !content?.trim()) {
      return Response.json({ error: 'All fields are required.' }, { status: 400 });
    }

    const colorList = colors.length > 0 ? colors.join(', ') : 'no brand colors detected';
    const headingFont = fonts[0] || 'Calibri';
    const bodyFont = fonts[1] || fonts[0] || 'Calibri';

    const logoList = logos.length > 0
      ? logos.map((l) => `- ${l.id}: ${l.description}`).join('\n')
      : '(none available)';

    const hasImage = !!backgroundImage?.base64;

    const systemPrompt = `You are a professional slide designer.${hasImage ? ' The user has provided a reference slide image — study it carefully to understand the visual layout before choosing a format.' : ''}
Return only a single valid JSON object. No explanation, no markdown, no code fences.

Brand colors: ${colorList}
Brand fonts: heading ${headingFont}, body ${bodyFont}

Available logos:
${logoList}

Choose "format" based on the content structure:
- "bullets"      — simple single-column bullet list
- "two-column"   — two items being compared or contrasted
- "four-column"  — four parallel concepts, phases, or pillars
- "timeline"     — sequential steps or phases over time
- "comparison"   — two options with a clear dividing line

Rules:
- all colors must be valid hex strings like #1A3A5C
- if no brand colors were detected, use professional neutral defaults
- choose logoId only if it genuinely fits the content; otherwise null
- sections length must match the chosen format (1 for bullets, 2 for two-column/comparison, 4 for four-column, 3-6 for timeline)

JSON schema (use exactly these field names):
{
  "format": "bullets" | "two-column" | "four-column" | "timeline" | "comparison",
  "title": string,
  "subtitle": string | null,
  "titleColor": hex string,
  "bodyColor": hex string,
  "accentColor": hex string,
  "headingFont": string,
  "bodyFont": string,
  "titleBold": boolean,
  "logoId": string | null,
  "sections": [{ "heading": string | null, "bullets": string[] }]
}`;

    const userMessage = hasImage
      ? `Looking at the reference slide above, generate content for:\nTitle: ${title}\nDescription: ${description}\nContent: ${content}`
      : `Title: ${title}\nDescription: ${description}\nContent: ${content}`;

    const message = await client.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 1200,
      system: systemPrompt,
      messages: buildClaudeMessages(userMessage, backgroundImage),
    });

    let rawText = message.content[0]?.text || '';
    rawText = rawText
      .replace(/^[\s\S]*?```(?:json)?\s*/i, '')
      .replace(/\s*```[\s\S]*$/, '')
      .trim();

    let slideData;
    try {
      slideData = JSON.parse(rawText);
    } catch (parseErr) {
      console.error('JSON parse error:', parseErr, 'Raw text:', rawText);
      return Response.json({ error: 'Invalid response format from AI.' }, { status: 500 });
    }

    const previewHtml = buildPreviewHtml(slideData, logos, backgroundImage);

    return Response.json({ previewHtml });
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

function buildPreviewHtml(s, logoManifest, backgroundImage) {
  const hasBg = !!backgroundImage?.base64;
  const titleColor = esc(s.titleColor || '#000000');
  const bodyColor = esc(s.bodyColor || '#333333');
  const accentColor = esc(s.accentColor || '#1A3A5C');
  const headingFont = esc(s.headingFont || 'Calibri');
  const bodyFont = esc(s.bodyFont || 'Calibri');

  const containerBg = hasBg ? `background:#000;` : `background:${esc(s.accentColor || '#1A3A5C')};`;
  // Use an <img> tag (not CSS background-image) so html2canvas can capture it reliably
  const bgImgHtml = hasBg
    ? `<img src="data:image/jpeg;base64,${backgroundImage.base64}" alt="" style="position:absolute;top:0;left:0;width:100%;height:100%;object-fit:cover;z-index:0;">`
    : '';

  // Overlay helper for text legibility over background images
  const overlay = (content, extra = '') =>
    hasBg
      ? `<div style="background:rgba(0,0,0,0.45);border-radius:4px;padding:4px 10px;display:inline-block;${extra}">${content}</div>`
      : `<div style="${extra}">${content}</div>`;

  // Logo
  let logoHtml = '';
  if (s.logoId) {
    const logoEntry = logoManifest.find((l) => l.id === s.logoId);
    if (logoEntry) {
      const logoSrc = esc(`/assets/logos/${logoEntry.file}`);
      logoHtml = `<img src="${logoSrc}" alt="logo" style="position:absolute;top:16px;right:20px;height:44px;object-fit:contain;">`;
    }
  }

  // Accent bar — only when no background image
  const accentBar = hasBg ? '' : `<div style="position:absolute;top:0;left:0;width:8px;height:100%;background:${accentColor};"></div>`;

  // Title + subtitle
  const titleWeight = s.titleBold !== false ? '700' : '400';
  const titleHtml = `<div style="font-family:'${headingFont}',sans-serif;font-size:30px;font-weight:${titleWeight};color:${titleColor};margin-bottom:8px;line-height:1.2;">${esc(s.title)}</div>`;
  const subtitleHtml = s.subtitle
    ? `<div style="font-family:'${headingFont}',sans-serif;font-size:16px;font-style:italic;color:${bodyColor};margin-bottom:14px;">${esc(s.subtitle)}</div>`
    : '';

  const sections = Array.isArray(s.sections) && s.sections.length > 0 ? s.sections : [];

  function renderBullets(bullets) {
    if (!Array.isArray(bullets) || bullets.length === 0) return '';
    return bullets.map((b) =>
      `<div style="display:flex;align-items:flex-start;gap:8px;margin-bottom:6px;">
        <span style="width:6px;height:6px;border-radius:50%;background:${accentColor};flex-shrink:0;margin-top:5px;"></span>
        <span style="font-family:'${bodyFont}',sans-serif;font-size:13px;color:${bodyColor};">${esc(b)}</span>
      </div>`
    ).join('');
  }

  let contentHtml = '';
  const format = s.format || 'bullets';

  if (format === 'four-column') {
    const cols = sections.slice(0, 4);
    contentHtml = `<div style="display:grid;grid-template-columns:repeat(${cols.length},1fr);gap:12px;margin-top:4px;">
      ${cols.map((sec) => `
        <div style="background:${hasBg ? 'rgba(0,0,0,0.4)' : 'rgba(255,255,255,0.08)'};border-radius:6px;padding:10px;">
          ${sec.heading ? `<div style="font-family:'${headingFont}',sans-serif;font-size:13px;font-weight:700;color:${accentColor};margin-bottom:6px;border-bottom:2px solid ${accentColor};padding-bottom:4px;">${esc(sec.heading)}</div>` : ''}
          ${renderBullets(sec.bullets)}
        </div>`).join('')}
    </div>`;
  } else if (format === 'two-column') {
    const cols = sections.slice(0, 2);
    contentHtml = `<div style="display:grid;grid-template-columns:repeat(2,1fr);gap:20px;margin-top:4px;">
      ${cols.map((sec) => `
        <div>
          ${sec.heading ? `<div style="font-family:'${headingFont}',sans-serif;font-size:14px;font-weight:700;color:${accentColor};margin-bottom:8px;">${esc(sec.heading)}</div>` : ''}
          ${renderBullets(sec.bullets)}
        </div>`).join('')}
    </div>`;
  } else if (format === 'timeline') {
    contentHtml = `<div style="display:flex;align-items:flex-start;gap:0;margin-top:8px;overflow:hidden;">
      ${sections.map((sec, i) => `
        <div style="flex:1;position:relative;padding:0 8px;">
          <div style="display:flex;align-items:center;margin-bottom:8px;">
            <div style="width:28px;height:28px;border-radius:50%;background:${accentColor};color:#fff;font-family:'${headingFont}',sans-serif;font-size:13px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0;">${i + 1}</div>
            ${i < sections.length - 1 ? `<div style="flex:1;height:2px;background:${accentColor};opacity:0.4;"></div>` : ''}
          </div>
          ${sec.heading ? `<div style="font-family:'${headingFont}',sans-serif;font-size:12px;font-weight:700;color:${hasBg ? '#fff' : accentColor};margin-bottom:4px;">${esc(sec.heading)}</div>` : ''}
          <div style="${hasBg ? 'background:rgba(0,0,0,0.4);border-radius:4px;padding:6px;' : ''}">
            ${renderBullets(sec.bullets)}
          </div>
        </div>`).join('')}
    </div>`;
  } else if (format === 'comparison') {
    const left = sections[0] || { heading: null, bullets: [] };
    const right = sections[1] || { heading: null, bullets: [] };
    contentHtml = `<div style="display:grid;grid-template-columns:1fr 4px 1fr;gap:16px;margin-top:4px;align-items:start;">
      <div>
        ${left.heading ? `<div style="font-family:'${headingFont}',sans-serif;font-size:15px;font-weight:700;color:${hasBg ? '#fff' : accentColor};margin-bottom:8px;">${esc(left.heading)}</div>` : ''}
        ${renderBullets(left.bullets)}
      </div>
      <div style="background:${accentColor};width:4px;align-self:stretch;border-radius:2px;"></div>
      <div>
        ${right.heading ? `<div style="font-family:'${headingFont}',sans-serif;font-size:15px;font-weight:700;color:${hasBg ? '#fff' : accentColor};margin-bottom:8px;">${esc(right.heading)}</div>` : ''}
        ${renderBullets(right.bullets)}
      </div>
    </div>`;
  } else {
    // bullets (default)
    const sec = sections[0] || { heading: null, bullets: [] };
    contentHtml = `<div style="${hasBg ? 'background:rgba(0,0,0,0.4);border-radius:6px;padding:10px;' : ''}">
      ${sec.heading ? `<div style="font-family:'${headingFont}',sans-serif;font-size:14px;font-weight:700;color:${hasBg ? '#fff' : accentColor};margin-bottom:8px;">${esc(sec.heading)}</div>` : ''}
      ${renderBullets(sec.bullets)}
    </div>`;
  }

  const titleBlock = hasBg
    ? `<div style="background:rgba(0,0,0,0.5);border-radius:6px;padding:8px 12px;margin-bottom:14px;display:inline-block;max-width:100%;">
        ${titleHtml}${subtitleHtml}
       </div>`
    : `${titleHtml}${subtitleHtml}`;

  return `<div style="width:800px;height:450px;${containerBg}position:relative;border-radius:8px;overflow:hidden;box-sizing:border-box;">
  ${bgImgHtml}
  <div style="position:relative;z-index:1;padding:36px 48px 36px 56px;height:100%;box-sizing:border-box;">
  ${accentBar}
  ${logoHtml}
  ${titleBlock}
  ${contentHtml}
  </div>
</div>`;
}
