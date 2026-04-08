const PptxGenJS = require("pptxgenjs");

const THEMES = {
  professional: { dark: "1E2761", light: "FFFFFF", text: "1A2320", accent: "4472C4" },
  teal:         { dark: "1A9E8F", light: "FFFFFF", text: "1A2320", accent: "D95B2A" },
  warm:         { dark: "D95B2A", light: "FAFAF8", text: "1A2320", accent: "B04520" },
  minimal:      { dark: "1A2320", light: "FFFFFF", text: "1A2320", accent: "555555" },
  berry:        { dark: "6D2E46", light: "FAF7F4", text: "1A2320", accent: "A26769" },
  forest:       { dark: "2C5F2D", light: "FAFAF8", text: "1A2320", accent: "97BC62" },
};

function toHex(c) {
  if (!c || c === "transparent") return null;
  c = String(c).trim();
  const h = c.replace("#", "");
  if (/^[0-9A-Fa-f]{6}$/.test(h)) return h.toUpperCase();
  const m = c.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?\s*\)/i);
  if (m) {
    if (m[4] !== undefined && parseFloat(m[4]) < 0.06) return null;
    return [m[1], m[2], m[3]].map(v => ("0" + parseInt(v).toString(16)).slice(-2)).join("").toUpperCase();
  }
  return null;
}

function getTheme(style, brandOn, brandColors) {
  if (brandOn && brandColors) {
    return {
      dark:   brandColors.primary.replace("#", ""),
      light:  "FFFFFF",
      text:   brandColors.text.replace("#", ""),
      accent: brandColors.accent.replace("#", ""),
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
  } catch (e) { return null; }
}

async function prepImg(src) {
  if (!src) return null;
  if (src.startsWith("data:")) return src.slice(5);
  if (src.includes(";base64,")) return src;
  return await fetchB64(src);
}

async function getImgDims(d) {
  try {
    const b64 = d.includes(";base64,") ? d.split(";base64,")[1] : d;
    const buf = Buffer.from(b64, "base64");
    if (buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47)
      return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
    let i = 2;
    while (i < buf.length - 8) {
      if (buf[i] === 0xff && (buf[i + 1] === 0xc0 || buf[i + 1] === 0xc2))
        return { w: buf.readUInt16BE(i + 7), h: buf.readUInt16BE(i + 5) };
      const seg = buf.readUInt16BE(i + 2);
      if (seg < 2) break;
      i += seg + 2;
    }
  } catch (e) {}
  return null;
}

async function addLogo(s, pres, data, pos, withBg, isCover) {
  if (!data) return;
  const isBot = pos === "bottom-left";
  const maxW = isCover ? 1.6 : 0.72;
  const maxH = isCover ? 0.7  : 0.3;
  let fw = maxW, fh = maxH;
  const dims = await getImgDims(data);
  if (dims && dims.w > 0) {
    const r = dims.w / dims.h;
    if (r > maxW / maxH) { fw = maxW; fh = maxW / r; }
    else { fh = maxH; fw = maxH * r; }
  }
  const x = 0.22;
  const y = isBot ? (5.625 - fh - 0.15) : 0.13;
  if (withBg) {
    const bh = Math.max(fh + 0.15, isCover ? 0.9 : 0.48);
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: isBot ? 5.625 - bh : 0, w: 10, h: bh,
      fill: { color: "FFFFFF" }, line: { color: "FFFFFF" },
    });
  }
  const imgData = data.startsWith("data:") ? data.slice(5) : data;
  try { s.addImage({ data: imgData, x, y, w: fw, h: fh }); }
  catch (e) { console.error("addLogo:", e.message); }
}

async function buildPptx(slides, theme, pres, logoImg, logoPos, logoWb) {
  // ── Design system (matches reference decks exactly) ─────────────────────
  // Slide: 10" x 5.625" (16:9)
  const W = 10, H = 5.625;
  const ML = 0.55, MR = 0.55, CW = W - ML - MR; // margins & content width
  // Y positions
  const TY = 0.36;   // title top
  const SY = 0.90;   // subtitle / eyebrow
  const CY = 1.15;   // content start
  const BH = 0.62;   // bullet row height
  // Font sizes (pt)
  const F_COVER  = 44;
  const F_TITLE  = 26;
  const F_HEAD   = 17;
  const F_BODY   = 13;
  const F_CAP    = 10;
  const F_STAT   = 56;
  const F_QUOTE  = 24;
  // Card widths for 3-col (matches deck1: 2.98" per card, 0.12" gap)
  const CW3 = 2.98, CG3 = 0.12;
  const CX3 = [ML, ML + CW3 + CG3, ML + (CW3 + CG3) * 2];

  const clamp = (v, mn, mx) => Math.max(mn, Math.min(mx, v));

  // shadow factory — never reuse objects (PptxGenJS mutates them)
  const mk = () => ({ type:"outer", color:"000000", blur:8, offset:2, angle:135, opacity:0.12 });

  for (let si = 0; si < slides.length; si++) {
    const sd = slides[si];
    const lay = sd.layout || 'default';
    const s = pres.addSlide();
    const bgHex = toHex(sd.background || "#FFFFFF") || "FFFFFF";
    s.background = { color: bgHex };
    await addLogo(s, pres, logoImg, logoPos || "top-left", logoWb, si === 0);

    // helpers scoped to this slide
    const txt = (text, x, y, w, h, opts={}) => {
      if (!text) return;
      const o = {
        x, y, w, h,
        fontSize: opts.sz || F_BODY,
        fontFace: opts.ff || "Calibri",
        bold: !!opts.bold, italic: !!opts.it,
        color: toHex(opts.c || "#1A2320") || "1A2320",
        align: opts.al || "left", valign: opts.va || "top",
        wrap: true, margin: [0,0,0,0],
        lineSpacingMultiple: opts.lh || 1.3,
      };
      if (opts.charSpacing) o.charSpacing = opts.charSpacing;
      s.addText(text, o);
    };
    const rect = (x, y, w, h, color, transparency=0) => {
      s.addShape(pres.shapes.RECTANGLE, {
        x, y, w, h,
        fill: { color: toHex(color) || "CCCCCC", transparency },
        line: { type: "none" },
      });
    };
    const card = (x, y, w, h, headerH, accentCol, bodyAlpha=90) => {
      s.addShape(pres.shapes.RECTANGLE, { x, y, w, h,
        fill: { color: toHex(accentCol) || "CCCCCC", transparency: bodyAlpha },
        line: { color: "DDDDDD", width: 0.5 }, shadow: mk() });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: headerH,
        fill: { color: toHex(accentCol) || "CCCCCC", transparency: 0 },
        line: { type:"none" } });
    };
    const imgZone = async (x, y, w, h, src) => {
      if (!src) return;
      const d = await prepImg(src);
      if (d) {
        try { s.addImage({ data:d, x, y, w, h, sizing:{ type:"cover", w, h } }); }
        catch(e) { console.error("addImage:", e.message); }
      }
    };

    // Extract elements
    const els = sd.elements || [];
    // Separate image elements from text/shape elements
    const imgEls = els.filter(e => e.type === 'image');
    const textEls = els.filter(e => e.type !== 'image');

    // ── Render elements ────────────────────────────────────────────────────
    // Process in order — images first (background), then shapes, then text
    for (const el of els) {
      const ix = px => Math.round(px / 960 * W * 1000) / 1000;
      const iy = px => Math.round(px / 540 * H * 1000) / 1000;
      const ex = ix(clamp(el.left   || 0, -10, 960));
      const ey = iy(clamp(el.top    || 0, -10, 540));
      const ew = ix(clamp(el.width  || 100, 2, 960));
      const eh = iy(clamp(el.height || 20,  1, 540));

      if (el.type === "text") {
        const pt     = clamp(Math.round((el.fontSize || 18) * 0.75), 6, 60);
        const colHex = toHex(el.color || "#1A2320") || "1A2320";
        const lh     = el.lineHeight || 1.35;
        const raw    = el.text || "";
        if (!raw.trim()) continue;
        const baseOpts = {
          x: ex, y: ey, w: ew, h: eh,
          fontFace: el.fontFamily || "Calibri",
          bold: !!el.bold, italic: !!el.italic,
          fontSize: pt, color: colHex,
          align: el.align || "left", valign: "top",
          wrap: true, margin: [0,0,0,0],
          lineSpacingMultiple: lh,
        };
        if (raw.includes("\n")) {
          const lines = raw.split("\n");
          s.addText(
            lines.map((ln, j) => ({
              text: ln,
              options: {
                fontSize: pt, fontFace: el.fontFamily || "Calibri",
                bold: !!el.bold, italic: !!el.italic, color: colHex,
                paraSpaceAfter: Math.round((lh - 1) * pt * 0.5),
                breakLine: j < lines.length - 1,
              },
            })),
            { x: ex, y: ey, w: ew, h: eh, valign: "top", wrap: true, margin: [0,0,0,0] }
          );
        } else {
          s.addText(raw, baseOpts);
        }

      } else if (el.type === "image") {
        const d = await prepImg(el.src);
        if (d) {
          try { s.addImage({ data:d, x:ex, y:ey, w:ew, h:eh, sizing:{ type:"cover", w:ew, h:eh } }); }
          catch(e) { console.error("addImage:", e.message); }
        }

      } else if (el.type === "shape") {
        if (ew < 0.01 || eh < 0.01) continue;
        const fill = toHex(el.fill || "#cccccc");
        if (!fill) continue;
        s.addShape(pres.shapes.RECTANGLE, { x:ex, y:ey, w:ew, h:eh,
          fill:{ color:fill }, line:{ color:fill } });

      } else if (el.type === "circle") {
        if (ew < 0.01 || eh < 0.01) continue;
        const cfHex = el.fill && el.fill !== "transparent" ? toHex(el.fill) : null;
        const csHex = el.stroke ? toHex(el.stroke) : null;
        if (!cfHex && !csHex) continue;
        const cf = cfHex ? { color: cfHex } : { type: "none" };
        const cs = csHex
          ? { color: csHex, pt: Math.max(1, Math.round((el.strokeWidth || 2) * 0.75)) }
          : { type: "none" };
        s.addShape(pres.shapes.ELLIPSE, { x:ex, y:ey, w:ew, h:eh, fill:cf, line:cs });
      }
    }
  }
}

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });
  if (!req.body || typeof req.body !== "object")
    return res.status(400).json({ error: "Request body missing or not JSON" });

  try {
    const {
      action, input, slideCount, style, title,
      slides, logoData, logoPos, logoWhiteBg,
      brandOn, brandColors,
    } = req.body;

    const apiKey = process.env.ANTHROPIC_API_KEY;
    const model  = process.env.CLAUDE_MODEL || "claude-sonnet-4-20250514";

    if (action === "outline") {
      const count = Math.max(2, Math.min(20, parseInt(slideCount) || 8));
      const prompt = `You are a world-class presentation designer. Create a ${count}-slide deck.
USER INPUT:\n${input}

CONTENT BUDGET — these limits are HARD, never exceed them:
- "heading": MAX 8 words. Sharp, specific, data-driven. e.g. "Solar Costs Fell 89% Since 2010" not "Cost Trends"
- "bullets": MAX 3 bullets. Each MAX 8 words. No full sentences. Use fragments like "↑42% YoY revenue growth"
- "paragraph": MAX 20 words. Only include on text-heavy layouts (default, title-body, two-col). Leave EMPTY "" on all image, stat, quote, cover, closing layouts.
- "subheading": MAX 10 words. Only on cover and closing slides.
- "stat": single striking number with unit e.g. "$4.2B" or "89%" or "3.2×"
- "quote": MAX 15 words. Punchy, attributed.
- "author": "First Last, Title" — short

LAYOUT RULES:
- Slide 1: layout "cover-center", dark:true — cinematic full-image cover
- Last slide: layout "closing", dark:true — full-image close
- Use image layouts (img-full, img-right, img-left, img-hero) for at least 30% of slides
- For image layouts paragraph MUST be ""
- Mix variety: never use the same layout twice in a row
- Use stat layout when there is a striking number
- Use quote layout when there is a strong quote or claim

IMAGEKEY: always provide a vivid 3-word Pexels search term matching the slide topic e.g. "solar panels desert"

Return ONLY a raw JSON array of exactly ${count} objects. No markdown, no fences.
Each object MUST have ALL these exact fields — no extras, no missing:
{"title":"max 5 words","type":"title|content|data|quote|cta","dark":false,"layout":"cover-center|cover-split|cover-circle|cover-dark|default|two-col|three-col|title-body|quote|big-statement|agenda|closing|stat|three-stats|timeline|four-icons|two-icons|comparison|process|pyramid|img-right|img-left|img-full|img-top|two-images|three-images|img-mosaic|img-hero","heading":"max 8 words","subheading":"","bullets":["max 8 words","max 8 words","max 8 words"],"paragraph":"","stat":"","quote":"","author":"","imageKeyword":"3 vivid words","speakerNote":"one sentence"}`;

      const r = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": apiKey,
          "anthropic-version": "2023-06-01",
        },
        body: JSON.stringify({ model, max_tokens: 8000, messages: [{ role: "user", content: prompt }] }),
      });
      const d = await r.json();
      if (!r.ok) return res.status(r.status).json({ error: d.error?.message || "API error" });
      const text  = d.content.map(b => b.text || "").join("");
      const clean = text.replace(/```json|```/g, "").trim();
      const s     = clean.indexOf("["), e = clean.lastIndexOf("]");
      if (s === -1 || e === -1) return res.status(500).json({ error: "No JSON array in response" });
      let outline;
      try { outline = JSON.parse(clean.slice(s, e + 1)); }
      catch (pe) { return res.status(500).json({ error: "Invalid JSON from model: " + pe.message }); }
      return res.status(200).json({ outline });
    }

    if (action === "pptx") {
      const theme = getTheme(style, brandOn, brandColors);
      const pres  = new PptxGenJS();
      pres.layout = "LAYOUT_16x9";
      pres.title  = title || "Presentation";
      let logoImg = null;
      if (logoData) {
        logoImg = logoData.startsWith("data:") || logoData.includes(";base64,")
          ? logoData
          : await fetchB64(logoData);
      }
      try { await buildPptx(slides, theme, pres, logoImg, logoPos, logoWhiteBg); }
      catch (buildErr) {
        console.error("buildPptx:", buildErr);
        return res.status(500).json({ error: "Build failed: " + buildErr.message });
      }
      const buf  = await pres.write({ outputType: "nodebuffer" });
      const safe = ((title || "presentation").trim() || "presentation").replace(/[^a-z0-9]/gi, "_");
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
      res.setHeader("Content-Disposition", `attachment; filename="${safe}.pptx"`);
      res.setHeader("Content-Length", buf.length);
      return res.status(200).send(buf);
    }

    return res.status(400).json({ error: "Invalid action" });

  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message });
  }
};
