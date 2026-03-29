const PptxGenJS = require("pptxgenjs");

const themes = {
  professional: { titleBg: "1E2761", titleText: "FFFFFF", slideBg: "FFFFFF", slideText: "1A2320", accent: "4472C4", subText: "666666" },
  teal: { titleBg: "1A9E8F", titleText: "FFFFFF", slideBg: "FFFFFF", slideText: "1A2320", accent: "D95B2A", subText: "4A5553" },
  warm: { titleBg: "D95B2A", titleText: "FFFFFF", slideBg: "FAFAF8", slideText: "1A2320", accent: "B04520", subText: "6B4C3B" },
  minimal: { titleBg: "1A2320", titleText: "FFFFFF", slideBg: "FFFFFF", slideText: "1A2320", accent: "1A2320", subText: "888888" },
  berry: { titleBg: "6D2E46", titleText: "FFFFFF", slideBg: "FAF7F4", slideText: "1A2320", accent: "A26769", subText: "6D4C55" },
  forest: { titleBg: "2C5F2D", titleText: "FFFFFF", slideBg: "FAFAF8", slideText: "1A2320", accent: "97BC62", subText: "4A6B3A" }
};

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try {
    const { action, input, slideCount, style, title, outline } = req.body;
    const apiKey = process.env.ANTHROPIC_API_KEY;

    // ── GENERATE OUTLINE ──
    if (action === "outline") {
      const prompt = `You are a presentation design expert. Based on the input below, create a structured ${slideCount}-slide outline.\n\n${input}\n\nReturn ONLY a raw JSON array with exactly ${slideCount} objects. Each object:\n- "title": short slide title (3-7 words)\n- "type": one of "title","agenda","content","data","quote","cta","conclusion"\n- "bullets": 2-4 concise bullets (5-10 words each)\n- "speakerNote": one sentence of guidance\n\nNo markdown, no explanation, raw JSON array only.`;

      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": apiKey,
          "anthropic-version": "2023-06-01"
        },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 1000,
          messages: [{ role: "user", content: prompt }]
        })
      });

      const data = await response.json();
      if (!response.ok) return res.status(response.status).json({ error: data.error?.message || "API error" });

      const text = data.content.map(b => b.text || "").join("");
      const clean = text.replace(/```json|```/g, "").trim();
      const outlineData = JSON.parse(clean);
      return res.status(200).json({ outline: outlineData });
    }

    // ── GENERATE PPTX ──
    if (action === "pptx") {
      const theme = themes[style] || themes.professional;
      const slideImages = req.body.slideImages || {};
      const pres = new PptxGenJS();
      pres.layout = "LAYOUT_16x9";
      pres.title = title || "Presentation";

      outline.forEach((slide, index) => {
        const s = pres.addSlide();
        const isDark = slide.type === "title" || slide.type === "conclusion" || slide.type === "cta" || index === 0;
        const imageData = slideImages[index];
        s.background = { color: isDark ? theme.titleBg : theme.slideBg };

        if (isDark) {
          // Add background image with overlay if exists
          if (imageData) {
            try {
              // If it's a URL (picsum), use path; if base64, use data
              if (imageData.startsWith('data:')) {
                s.addImage({ data: imageData, x: 0, y: 0, w: 10, h: 5.625, transparency: 70 });
              } else {
                s.addImage({ path: imageData, x: 0, y: 0, w: 10, h: 5.625, transparency: 70 });
              }
            } catch(e) { console.log('Image error:', e.message); }
          }
          s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.08, h: 5.625, fill: { color: theme.accent }, line: { color: theme.accent } });
          s.addText(slide.title, { x: 0.6, y: 1.8, w: 8.8, h: 1.4, fontSize: 38, fontFace: "Calibri", bold: true, color: theme.titleText, align: "left" });
          if (slide.bullets && slide.bullets.length > 0) {
            s.addText(slide.bullets[0], { x: 0.6, y: 3.4, w: 7.5, h: 0.6, fontSize: 16, fontFace: "Calibri", color: theme.titleText, align: "left" });
          }
        } else {
          s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent }, line: { color: theme.accent } });

          if (imageData) {
            // Image on right half, text on left
            const textWidth = 5.8;
            s.addText(slide.title, { x: 0.5, y: 0.25, w: textWidth, h: 0.75, fontSize: 22, fontFace: "Calibri", bold: true, color: theme.slideText, align: "left", margin: 0 });
            s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.05, w: 1.2, h: 0.04, fill: { color: theme.accent }, line: { color: theme.accent } });
            if (slide.bullets && slide.bullets.length > 0) {
              const bulletItems = slide.bullets.map((b, i) => ({
                text: b,
                options: { bullet: true, fontSize: 13, fontFace: "Calibri", color: theme.slideText, paraSpaceAfter: 6, breakLine: i < slide.bullets.length - 1 }
              }));
              s.addText(bulletItems, { x: 0.5, y: 1.3, w: textWidth, h: 3.8, valign: "top" });
            }
            try {
              if (imageData.startsWith('data:')) {
                s.addImage({ data: imageData, x: 6.4, y: 0.06, w: 3.6, h: 5.565, sizing: { type: 'cover', w: 3.6, h: 5.565 } });
              } else {
                s.addImage({ path: imageData, x: 6.4, y: 0.06, w: 3.6, h: 5.565, sizing: { type: 'cover', w: 3.6, h: 5.565 } });
              }
            } catch(e) { console.log('Image error:', e.message); }
          } else {
            // No image — full width layout
            s.addText(slide.title, { x: 0.5, y: 0.25, w: 8.5, h: 0.75, fontSize: 26, fontFace: "Calibri", bold: true, color: theme.slideText, align: "left", margin: 0 });
            s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.05, w: 1.2, h: 0.04, fill: { color: theme.accent }, line: { color: theme.accent } });
            if (slide.bullets && slide.bullets.length > 0) {
              const bulletItems = slide.bullets.map((b, i) => ({
                text: b,
                options: { bullet: true, fontSize: 15, fontFace: "Calibri", color: theme.slideText, paraSpaceAfter: 8, breakLine: i < slide.bullets.length - 1 }
              }));
              s.addText(bulletItems, { x: 0.5, y: 1.3, w: 8.8, h: 3.8, valign: "top" });
            }
          }
          if (slide.speakerNote) s.addNotes(slide.speakerNote);
        }
      });

      const buffer = await pres.write({ outputType: "nodebuffer" });
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
      res.setHeader("Content-Disposition", `attachment; filename="${(title || "presentation").replace(/[^a-z0-9]/gi, "_")}.pptx"`);
      res.setHeader("Content-Length", buffer.length);
      return res.status(200).send(buffer);
    }

    return res.status(400).json({ error: "Invalid action" });

  } catch (err) {
    console.error("Error:", err);
    return res.status(500).json({ error: err.message });
  }
};
