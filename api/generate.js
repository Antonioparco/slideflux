const PptxGenJS = require("pptxgenjs");

const themes = {
  professional: { titleBg: "1E2761", titleText: "FFFFFF", slideBg: "FFFFFF", slideText: "1A2320", accent: "4472C4", subText: "666666" },
  teal:         { titleBg: "1A9E8F", titleText: "FFFFFF", slideBg: "FFFFFF", slideText: "1A2320", accent: "D95B2A", subText: "4A5553" },
  warm:         { titleBg: "D95B2A", titleText: "FFFFFF", slideBg: "FAFAF8", slideText: "1A2320", accent: "B04520", subText: "6B4C3B" },
  minimal:      { titleBg: "1A2320", titleText: "FFFFFF", slideBg: "FFFFFF", slideText: "1A2320", accent: "1A2320", subText: "888888" },
  berry:        { titleBg: "6D2E46", titleText: "FFFFFF", slideBg: "FAF7F4", slideText: "1A2320", accent: "A26769", subText: "6D4C55" },
  forest:       { titleBg: "2C5F2D", titleText: "FFFFFF", slideBg: "FAFAF8", slideText: "1A2320", accent: "97BC62", subText: "4A6B3A" }
};

async function fetchImageAsBase64(url) {
  try {
    const res = await fetch(url);
    const buffer = await res.arrayBuffer();
    const base64 = Buffer.from(buffer).toString('base64');
    const contentType = res.headers.get('content-type') || 'image/jpeg';
    return `${contentType};base64,${base64}`;
  } catch (e) { return null; }
}

async function prepareImage(imageData) {
  if (!imageData) return null;
  if (imageData.startsWith('data:')) return imageData.replace('data:', '');
  const b64 = await fetchImageAsBase64(imageData);
  return b64;
}

async function buildPptx(outline, style, slideImages) {
  const theme = themes[style] || themes.professional;
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  for (let index = 0; index < outline.length; index++) {
    const slide = outline[index];
    const s = pres.addSlide();
    const isDark = slide.type === "title" || slide.type === "conclusion" || slide.type === "cta" || index === 0;
    const imageData = await prepareImage(slideImages?.[index]);

    s.background = { color: isDark ? theme.titleBg : theme.slideBg };

    if (isDark) {
      if (imageData) { try { s.addImage({ data: imageData, x: 0, y: 0, w: 10, h: 5.625, transparency: 70 }); } catch(e) {} }
      s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.08, h: 5.625, fill: { color: theme.accent }, line: { color: theme.accent } });
      s.addText(slide.title, { x: 0.6, y: 1.8, w: 8.8, h: 1.4, fontSize: 38, fontFace: "Calibri", bold: true, color: theme.titleText, align: "left" });
      if (slide.bullets?.[0]) s.addText(slide.bullets[0], { x: 0.6, y: 3.4, w: 7.5, h: 0.6, fontSize: 16, fontFace: "Calibri", color: theme.titleText, align: "left" });
    } else {
      s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent }, line: { color: theme.accent } });
      if (imageData) {
        s.addText(slide.title, { x: 0.5, y: 0.25, w: 5.8, h: 0.75, fontSize: 22, fontFace: "Calibri", bold: true, color: theme.slideText, align: "left", margin: 0 });
        s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.05, w: 1.2, h: 0.04, fill: { color: theme.accent }, line: { color: theme.accent } });
        if (slide.bullets?.length) {
          s.addText(slide.bullets.map((b,i)=>({text:b,options:{bullet:true,fontSize:13,fontFace:"Calibri",color:theme.slideText,paraSpaceAfter:6,breakLine:i<slide.bullets.length-1}})), { x: 0.5, y: 1.3, w: 5.8, h: 3.8, valign: "top" });
        }
        try { s.addImage({ data: imageData, x: 6.4, y: 0.06, w: 3.6, h: 5.565, sizing: { type: 'cover', w: 3.6, h: 5.565 } }); } catch(e) {}
      } else {
        s.addText(slide.title, { x: 0.5, y: 0.25, w: 8.5, h: 0.75, fontSize: 26, fontFace: "Calibri", bold: true, color: theme.slideText, align: "left", margin: 0 });
        s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.05, w: 1.2, h: 0.04, fill: { color: theme.accent }, line: { color: theme.accent } });
        if (slide.bullets?.length) {
          s.addText(slide.bullets.map((b,i)=>({text:b,options:{bullet:true,fontSize:15,fontFace:"Calibri",color:theme.slideText,paraSpaceAfter:8,breakLine:i<slide.bullets.length-1}})), { x: 0.5, y: 1.3, w: 8.8, h: 3.8, valign: "top" });
        }
      }
      if (slide.speakerNote) s.addNotes(slide.speakerNote);
    }
  }
  return await pres.write({ outputType: "nodebuffer" });
}

function hexToRgb(hex) {
  const r = parseInt(hex.slice(0,2),16), g = parseInt(hex.slice(2,4),16), b = parseInt(hex.slice(4,6),16);
  return `rgb(${r},${g},${b})`;
}

async function buildPdfBuffer(outline, style, slideImages, orientation) {
  const theme = themes[style] || themes.professional;
  const isLandscape = orientation !== 'portrait';

  let slidesHtml = '';
  for (let i = 0; i < outline.length; i++) {
    const slide = outline[i];
    const isDark = slide.type === 'title' || slide.type === 'conclusion' || slide.type === 'cta' || i === 0;
    let imageData = slideImages?.[i] || null;
    if (imageData && !imageData.startsWith('data:')) {
      const b64 = await fetchImageAsBase64(imageData);
      if (b64) imageData = 'data:' + b64;
    }

    const bg = isDark ? hexToRgb(theme.titleBg) : hexToRgb(theme.slideBg);
    const accentColor = hexToRgb(theme.accent);
    const titleColor = isDark ? '#fff' : hexToRgb(theme.slideText);
    const bulletColor = hexToRgb(theme.slideText);
    const fs = isLandscape ? { title: isDark?'32px':'20px', bullet: '13px', sub: '14px' } : { title: isDark?'24px':'16px', bullet: '11px', sub: '11px' };

    const imgStyle = isDark
      ? 'position:absolute;inset:0;width:100%;height:100%;object-fit:cover;opacity:0.2'
      : 'position:absolute;right:0;top:0;bottom:0;width:38%;object-fit:cover;height:100%';
    const contentRight = imageData && !isDark ? '42%' : '4%';

    const bulletsHtml = (slide.bullets||[]).map(b =>
      `<li style="margin-bottom:5px;font-size:${fs.bullet};color:${bulletColor};line-height:1.5">${b}</li>`
    ).join('');

    const darkContent = `
      ${imageData ? `<img src="${imageData}" style="${imgStyle}" />` : ''}
      <div style="position:absolute;left:0;top:0;bottom:0;width:5px;background:${accentColor}"></div>
      <div style="position:absolute;left:4%;top:35%;right:4%">
        <div style="font-size:${fs.title};font-weight:700;color:#fff;line-height:1.25;margin-bottom:12px">${slide.title}</div>
        <div style="font-size:${fs.sub};color:rgba(255,255,255,0.75)">${(slide.bullets||[])[0]||''}</div>
      </div>
      <div style="position:absolute;top:6px;right:10px;font-size:9px;color:rgba(255,255,255,0.35)">${i+1}</div>`;

    const lightContent = `
      ${imageData ? `<img src="${imageData}" style="${imgStyle}" />` : ''}
      <div style="position:absolute;top:0;left:0;right:0;height:4px;background:${accentColor}"></div>
      <div style="position:absolute;top:12px;left:4%;right:${contentRight};font-size:${fs.title};font-weight:700;color:${titleColor};line-height:1.25">${slide.title}</div>
      <div style="position:absolute;top:42px;left:4%;width:50px;height:3px;background:${accentColor};border-radius:2px"></div>
      <div style="position:absolute;top:54px;left:4%;right:${contentRight};bottom:8px;overflow:hidden">
        <ul style="margin:0;padding-left:18px;list-style-type:disc">${bulletsHtml}</ul>
      </div>
      <div style="position:absolute;top:6px;right:10px;font-size:8px;color:#bbb">${i+1} / ${outline.length}</div>`;

    const slideStyle = isLandscape
      ? 'width:257mm;height:144mm;margin:15mm auto;page-break-after:always;'
      : 'width:190mm;height:107mm;margin:10mm auto;page-break-after:always;';

    slidesHtml += `<div style="${slideStyle}position:relative;background:${bg};border-radius:4px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,0.15)">${isDark ? darkContent : lightContent}</div>`;
  }

  const pageSize = isLandscape ? 'A4 landscape' : 'A4 portrait';
  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
  <style>
    @page { size: ${pageSize}; margin: 0; }
    * { box-sizing: border-box; }
    body { margin: 0; padding: 0; background: #d0d0d0; font-family: Arial, Helvetica, sans-serif; }
    @media print { body { background: white; } div { box-shadow: none !important; } }
  </style>
  </head><body>${slidesHtml}</body></html>`;

  // Use puppeteer if available, otherwise fall back to html
  try {
    const puppeteer = require('puppeteer-core');
    const chromium = require('@sparticuz/chromium');
    const browser = await puppeteer.launch({
      args: chromium.args,
      defaultViewport: chromium.defaultViewport,
      executablePath: await chromium.executablePath(),
      headless: chromium.headless,
    });
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'networkidle0' });
    const pdfBuffer = await page.pdf({
      format: 'A4',
      landscape: isLandscape,
      printBackground: true,
      margin: { top: '0', right: '0', bottom: '0', left: '0' }
    });
    await browser.close();
    return { buffer: Buffer.from(pdfBuffer), isPdf: true };
  } catch (e) {
    // Puppeteer not available — return HTML that user can print to PDF
    console.log('Puppeteer not available, returning printable HTML:', e.message);
    return { buffer: Buffer.from(html, 'utf8'), isPdf: false };
  }
}

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try {
    const { action, input, slideCount, style, title, outline, format, orientation, slideImages } = req.body;
    const apiKey = process.env.ANTHROPIC_API_KEY;
    const safeName = (title || "presentation").replace(/[^a-z0-9]/gi, "_");

    // ── OUTLINE ──
    if (action === "outline") {
      const prompt = `You are a presentation design expert. Based on the input below, create a structured ${slideCount}-slide outline.\n\n${input}\n\nReturn ONLY a raw JSON array with exactly ${slideCount} objects. Each object:\n- "title": short slide title (3-7 words)\n- "type": one of "title","agenda","content","data","quote","cta","conclusion"\n- "bullets": 2-4 concise bullets (5-10 words each)\n- "speakerNote": one sentence of guidance\n\nNo markdown, no explanation, raw JSON array only.`;
      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json", "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
        body: JSON.stringify({ model: "claude-sonnet-4-20250514", max_tokens: 1000, messages: [{ role: "user", content: prompt }] })
      });
      const data = await response.json();
      if (!response.ok) return res.status(response.status).json({ error: data.error?.message || "API error" });
      const text = data.content.map(b => b.text || "").join("");
      const clean = text.replace(/```json|```/g, "").trim();
      return res.status(200).json({ outline: JSON.parse(clean) });
    }

    // ── PPTX ──
    if (action === "pptx" && format === "pptx") {
      const buffer = await buildPptx(outline, style, slideImages);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
      res.setHeader("Content-Disposition", `attachment; filename="${safeName}.pptx"`);
      return res.status(200).send(buffer);
    }

    // ── PDF ──
    if (action === "pptx" && format === "pdf") {
      const { buffer, isPdf } = await buildPdfBuffer(outline, style, slideImages, orientation);
      if (isPdf) {
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", `attachment; filename="${safeName}.pdf"`);
      } else {
        res.setHeader("Content-Type", "text/html; charset=utf-8");
        res.setHeader("Content-Disposition", `attachment; filename="${safeName}_print.html"`);
      }
      return res.status(200).send(buffer);
    }

    // ── BOTH ──
    if (action === "pptx" && format === "both") {
      const buffer = await buildPptx(outline, style, slideImages);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
      res.setHeader("Content-Disposition", `attachment; filename="${safeName}.pptx"`);
      return res.status(200).send(buffer);
    }

    return res.status(400).json({ error: "Invalid action" });

  } catch (err) {
    console.error("Error:", err);
    return res.status(500).json({ error: err.message });
  }
};
