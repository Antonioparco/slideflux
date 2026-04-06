const PptxGenJS = require("pptxgenjs");

const THEMES = {
  professional: { dark: "1E2761", light: "FFFFFF", text: "1A2320", accent: "4472C4" },
  teal:         { dark: "1A9E8F", light: "FFFFFF", text: "1A2320", accent: "D95B2A" },
  warm:         { dark: "D95B2A", light: "FAFAF8", text: "1A2320", accent: "B04520" },
  minimal:      { dark: "1A2320", light: "FFFFFF", text: "1A2320", accent: "555555" },
  berry:        { dark: "6D2E46", light: "FAF7F4", text: "1A2320", accent: "A26769" },
  forest:       { dark: "2C5F2D", light: "FAFAF8", text: "1A2320", accent: "97BC62" }
};

function toHex(c) {
  if (!c) return "FFFFFF";
  c = String(c).trim();
  if (/^#?[0-9A-Fa-f]{6}$/.test(c.replace("#", ""))) return c.replace("#", "").toUpperCase();
  const m = c.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
  if (m) return [m[1], m[2], m[3]].map(v => ("0" + parseInt(v, 10).toString(16)).slice(-2)).join("").toUpperCase();
  return "FFFFFF";
}

function getTheme(style, brandOn, brandColors) {
  if (brandOn && brandColors) {
    return {
      dark: toHex(brandColors.primary || "#1E2761"),
      light: "FFFFFF",
      text: toHex(brandColors.text || "#1A2320"),
      accent: toHex(brandColors.accent || "#4472C4")
    };
  }
  return THEMES[style] || THEMES.professional;
}

async function fetchB64(url) {
  try {
    const r = await fetch(url);
    if (!r.ok) return null;
    const buf = await r.arrayBuffer();
    const ct = r.headers.get("content-type") || "image/jpeg";
    return ct + ";base64," + Buffer.from(buf).toString("base64");
  } catch (e) {
    return null;
  }
}

async function prepImg(src) {
  if (!src) return null;
  if (src.startsWith("data:")) return src.replace("data:", "");
  return await fetchB64(src);
}

async function getImgDims(d) {
  try {
    const b64 = d.includes(";base64,") ? d.split(";base64,")[1] : d;
    const buf = Buffer.from(b64, "base64");
    if (buf.length > 24 && buf[1] === 0x50 && buf[2] === 0x4E && buf[3] === 0x47) {
      return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
    }
    let i = 2;
    while (i < buf.length - 8) {
      if (buf[i] === 0xFF && (buf[i + 1] === 0xC0 || buf[i + 1] === 0xC2)) {
        return { w: buf.readUInt16BE(i + 7), h: buf.readUInt16BE(i + 5) };
      }
      if (i + 2 >= buf.length) break;
      i += buf.readUInt16BE(i + 2) + 2;
    }
  } catch (e) {}
  return null;
}

async function addLogo(s, pres, data, pos, wb, isCover) {
  if (!data) return;
  const isBot = pos === "bottom-left";
  const maxW = isCover ? 1.6 : 0.72;
  const maxH = isCover ? 0.7 : 0.3;
  let fw = maxW, fh = maxH;
  const dims = await getImgDims(data);
  if (dims && dims.w > 0 && dims.h > 0) {
    const r = dims.w / dims.h;
    if (r > maxW / maxH) { fw = maxW; fh = maxW / r; }
    else { fh = maxH; fw = maxH * r; }
  }
  const x = 0.22, y = isBot ? (5.625 - fh - 0.15) : 0.13;
  if (wb) {
    const bh = Math.max(fh + 0.15, isCover ? 0.9 : 0.48);
    s.addShape(pres.ShapeType.rect, {
      x: 0, y: isBot ? 5.625 - bh : 0, w: 10, h: bh,
      fill: { color: "FFFFFF" }, line: { color: "FFFFFF" }
    });
  }
  try { s.addImage({ data, x, y, w: fw, h: fh }); } catch (e) {}
}

function fitText(text, maxCharsPerLine, maxLines) {
  const words = String(text || "").trim().split(/\s+/).filter(Boolean);
  if (!words.length) return "";
  const lines = [];
  let cur = "";
  for (const word of words) {
    const next = cur ? cur + " " + word : word;
    if (next.length <= maxCharsPerLine) cur = next;
    else {
      if (cur) lines.push(cur);
      cur = word;
      if (lines.length >= maxLines) break;
    }
  }
  if (cur && lines.length < maxLines) lines.push(cur);
  return lines.slice(0, maxLines).join("\n");
}

function cleanOutline(outline) {
  return (Array.isArray(outline) ? outline : []).map((sl, i, arr) => {
    const isFirst = i === 0;
    const isLast = i === arr.length - 1;
    const layout = sl.layout || (isFirst ? "cover-center" : isLast ? "closing" : "default");
    return {
      title: fitText(sl.title || `Slide ${i + 1}`, 24, 2),
      type: sl.type || (isFirst ? "title" : isLast ? "cta" : "content"),
      dark: typeof sl.dark === "boolean" ? sl.dark : (isFirst || isLast),
      layout,
      heading: fitText(sl.heading || sl.title || `Slide ${i + 1}`, layout.startsWith("cover") ? 28 : 36, layout.startsWith("cover") ? 3 : 2),
      subheading: fitText(sl.subheading || "", 42, 2),
      bullets: (Array.isArray(sl.bullets) ? sl.bullets : [])
        .slice(0, 4)
        .map(b => fitText(b, 42, 2)),
      paragraph: fitText(sl.paragraph || "", 52, 3),
      stat: fitText(sl.stat || "", 14, 2),
      quote: fitText(sl.quote || "", 46, 4),
      author: fitText(sl.author || "", 28, 2),
      imageKeyword: fitText(sl.imageKeyword || "", 24, 1),
      speakerNote: fitText(sl.speakerNote || "", 120, 2)
    };
  });
}

function shapeLineOpts(el, fallbackColor) {
  return {
    color: toHex(el.stroke || fallbackColor || "#000000"),
    pt: Math.max(1, Math.round((el.strokeWidth || 2) * 0.75))
  };
}

async function buildPptx(slides, theme, pres, logoImg, logoPos, logoWb) {
  const CW = 960, CH = 540, PW = 10, PH = 5.625;
  const ix = px => Math.round(px / CW * PW * 1000) / 1000;
  const iy = px => Math.round(px / CH * PH * 1000) / 1000;
  const clamp = (v, mn, mx) => Math.max(mn, Math.min(mx, v));

  for (let si = 0; si < slides.length; si++) {
    const sd = slides[si];
    const s = pres.addSlide();
    s.background = { color: toHex(sd.background || "#FFFFFF") };

    const preparedLogo = await prepImg(logoImg);
    await addLogo(s, pres, preparedLogo, logoPos || "top-left", logoWb, si === 0);

    for (const el of (sd.elements || [])) {
      const x = ix(clamp(el.left || 0, -10, 960));
      const y = iy(clamp(el.top || 0, -10, 540));
      const w = ix(clamp(el.width || 100, 2, 960));
      const h = iy(clamp(el.height || 20, 1, 540));

      if (el.type === "text") {
        const pt = clamp(Math.round((el.fontSize || 18) * 0.75), 6, 60);
        const col = toHex(el.color || "#1A2320");
        const lh = el.lineHeight || 1.35;
        const raw = String(el.text || "");
        const baseOpts = {
          x, y, w, h,
          fontFace: el.fontFamily || "Calibri",
          bold: !!el.bold,
          italic: !!el.italic,
          fontSize: pt,
          color: col,
          align: el.align || "left",
          valign: "top",
          wrap: true,
          margin: [2, 4, 2, 4]
        };
        if (raw.includes("\n")) {
          const lines = raw.split("\n");
          s.addText(
            lines.map((ln, j) => ({
              text: ln,
              options: {
                fontSize: pt,
                fontFace: el.fontFamily || "Calibri",
                bold: !!el.bold,
                italic: !!el.italic,
                color: col,
                paraSpaceAfter: Math.round((lh - 1) * pt * 0.5),
                breakLine: j < lines.length - 1
              }
            })),
            { x, y, w, h, valign: "top", wrap: true, margin: [2, 4, 2, 4] }
          );
        } else {
          s.addText(raw, baseOpts);
        }
      } else if (el.type === "image") {
        const d = await prepImg(el.src);
        if (d) {
          try { s.addImage({ data: d, x, y, w, h, sizing: { type: "cover", w, h } }); } catch (e) {}
        }
      } else if (el.type === "circle") {
        if (w < 0.01 || h < 0.01) continue;
        const fill = el.fill && el.fill !== "transparent" ? { color: toHex(el.fill) } : { transparency: 100 };
        const line = el.stroke && el.stroke !== "transparent" ? shapeLineOpts(el, el.fill || "#000000") : { transparency: 100 };
        s.addShape(pres.ShapeType.ellipse, { x, y, w, h, fill, line });
      } else if (el.type === "line" || (el.type === "shape" && el.shapeKind === "line")) {
        const line = shapeLineOpts(el, "#000000");
        s.addShape(pres.ShapeType.line, {
          x, y,
          w: Math.max(w, 0.02),
          h: Math.max(h, 0.02),
          line
        });
      } else if (el.type === "star" || (el.type === "shape" && el.shapeKind === "star")) {
        const fill = el.fill && el.fill !== "transparent" ? { color: toHex(el.fill) } : { transparency: 100 };
        const line = el.stroke && el.stroke !== "transparent" ? shapeLineOpts(el, el.fill || "#000000") : { transparency: 100 };
        s.addShape(pres.ShapeType.star5, { x, y, w, h, fill, line });
      } else if (el.type === "shape") {
        if (w < 0.01 || h < 0.01) continue;
        const m = (el.fill || "").match(/rgba\([\d,\s]+,([\d.]+)\)/);
        if (m && parseFloat(m[1]) < 0.06) continue;
        const fillHex = toHex(el.fill || "#cccccc");
        const line = el.stroke && el.stroke !== "transparent" ? shapeLineOpts(el, fillHex) : { color: fillHex };
        const rounded = el.shapeKind === "rounded-rect" || el.rx || el.ry;
        s.addShape(rounded ? pres.ShapeType.roundRect : pres.ShapeType.rect, {
          x, y, w, h,
          fill: { color: fillHex },
          line
        });
      }
    }
  }
}

function extractJSONArray(text) {
  const clean = String(text || "").replace(/```json|```/g, "").trim();
  const s = clean.indexOf("[");
  const e = clean.lastIndexOf("]");
  if (s === -1 || e === -1 || e <= s) return null;
  return clean.slice(s, e + 1);
}

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try {
    const body = req.body || {};
    const { action, input, slideCount, style, title, slides, logoData, logoPos, logoWhiteBg, brandOn, brandColors } = body;
    const theme = getTheme(style, brandOn, brandColors);

    if (action === "outline") {
      const apiKey = process.env.ANTHROPIC_API_KEY;
      if (!apiKey) return res.status(500).json({ error: "Missing ANTHROPIC_API_KEY" });

      const prompt = `You are an expert presentation consultant. Create a ${slideCount}-slide presentation.

USER INPUT:
${input}

STRICT STRUCTURE RULES:
- Return ONLY raw JSON array. No markdown. No code fences.
- First slide must be layout "cover-center" and dark true.
- Last slide must be layout "closing" and dark true.
- Use varied layouts for the remaining slides.
- Keep content visually safe:
  - heading: max 10 words
  - subheading: max 16 words
  - bullets: EXACTLY 4 bullets, each max 12 words
  - paragraph: max 28 words
  - quote: max 28 words
  - stat: max 8 words
  - author: max 6 words
  - imageKeyword: max 4 words
  - speakerNote: one short sentence

Each object must contain:
{"title":"3-6 words","type":"title|content|data|quote|cta","dark":true,"layout":"cover-center|cover-split|cover-circle|cover-dark|default|two-col|three-col|title-body|quote|big-statement|agenda|closing|stat|three-stats|timeline|four-icons|two-icons|comparison|process|pyramid|img-right|img-left|img-full|img-top|two-images|three-images|img-mosaic|img-hero","heading":"specific headline","subheading":"","bullets":["One","Two","Three","Four"],"paragraph":"brief paragraph","stat":"","quote":"","author":"","imageKeyword":"keyword","speakerNote":"one sentence"}`;

      const r = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": apiKey,
          "anthropic-version": "2023-06-01"
        },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 4000,
          messages: [{ role: "user", content: prompt }]
        })
      });

      const d = await r.json().catch(() => ({}));
      if (!r.ok) return res.status(r.status).json({ error: d?.error?.message || "API error" });

      const text = (Array.isArray(d.content) ? d.content : []).map(b => b.text || "").join("");
      const jsonText = extractJSONArray(text);
      if (!jsonText) return res.status(500).json({ error: "Model did not return a valid slide array" });

      let outline;
      try {
        outline = JSON.parse(jsonText);
      } catch (e) {
        return res.status(500).json({ error: "Failed to parse model response" });
      }

      return res.status(200).json({ outline: cleanOutline(outline) });
    }

    if (action === "pptx") {
      if (!Array.isArray(slides) || !slides.length) {
        return res.status(400).json({ error: "No slides to export" });
      }

      const pres = new PptxGenJS();
      pres.layout = "LAYOUT_WIDE";
      pres.author = "Slideflux";
      pres.company = "Slideflux";
      pres.subject = String(title || "Presentation");
      pres.title = String(title || "Presentation");
      pres.lang = "en-US";
      pres.theme = { headFontFace: "Arial", bodyFontFace: "Arial", lang: "en-US" };

      await buildPptx(slides, theme, pres, logoData, logoPos, !!logoWhiteBg);
      const buf = await pres.write({ outputType: "nodebuffer" });
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
      res.setHeader("Content-Disposition", `attachment; filename="${String(title || "presentation").replace(/[^a-z0-9_\-]+/gi, "_")}.pptx"`);
      return res.status(200).send(buf);
    }

    return res.status(400).json({ error: "Invalid action" });
  } catch (e) {
    return res.status(500).json({ error: e.message || "Server error" });
  }
};
