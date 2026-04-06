const PptxGenJS = require("pptxgenjs");

const THEMES = {
  professional: { dark: "1E2761", light: "FFFFFF", text: "1A2320", accent: "4472C4" },
  teal:         { dark: "1A9E8F", light: "FFFFFF", text: "1A2320", accent: "D95B2A" },
  warm:         { dark: "D95B2A", light: "FAFAF8", text: "1A2320", accent: "B04520" },
  minimal:      { dark: "1A2320", light: "FFFFFF", text: "1A2320", accent: "666666" },
  berry:        { dark: "6D2E46", light: "FAF7F4", text: "1A2320", accent: "A26769" },
  forest:       { dark: "2C5F2D", light: "FAFAF8", text: "1A2320", accent: "97BC62" }
};

function safeHex(input, fallback = "FFFFFF") {
  if (!input) return fallback;
  const value = String(input).trim();
  if (/^#?[0-9a-fA-F]{6}$/.test(value)) return value.replace("#", "").toUpperCase();
  const rgb = value.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
  if (rgb) {
    return [rgb[1], rgb[2], rgb[3]]
      .map((v) => Math.max(0, Math.min(255, Number(v))))
      .map((v) => v.toString(16).padStart(2, "0"))
      .join("")
      .toUpperCase();
  }
  return fallback;
}

function getTheme(style, brandOn, brandColors) {
  if (brandOn && brandColors) {
    return {
      dark: safeHex(brandColors.primary, "1E2761"),
      light: "FFFFFF",
      text: safeHex(brandColors.text, "1A2320"),
      accent: safeHex(brandColors.accent, "4472C4")
    };
  }
  return THEMES[style] || THEMES.professional;
}

function cleanJsonText(text) {
  const raw = String(text || "").trim().replace(/```json|```/gi, "");
  const firstBracket = raw.indexOf("[");
  const lastBracket = raw.lastIndexOf("]");
  if (firstBracket === -1 || lastBracket === -1) return "[]";
  return raw.slice(firstBracket, lastBracket + 1);
}

function normalizeOutline(slides, theme) {
  return (Array.isArray(slides) ? slides : []).map((slide, index, arr) => {
    const isCover = index === 0;
    const isClosing = index === arr.length - 1;
    const defaultLayout = isCover ? "cover" : isClosing ? "closing" : "content";
    return {
      title: String(slide.title || slide.heading || `Slide ${index + 1}`),
      heading: String(slide.heading || slide.title || `Slide ${index + 1}`),
      subheading: String(slide.subheading || ""),
      bullets: Array.isArray(slide.bullets) ? slide.bullets.slice(0, 5).map(String) : [],
      paragraph: String(slide.paragraph || ""),
      speakerNote: String(slide.speakerNote || ""),
      imageKeyword: String(slide.imageKeyword || ""),
      layout: String(slide.layout || defaultLayout),
      dark: typeof slide.dark === "boolean" ? slide.dark : isCover || isClosing,
      background: (slide.dark || isCover || isClosing) ? `#${theme.dark}` : `#${theme.light}`
    };
  });
}

async function fetchBase64(url) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`Image fetch failed: ${response.status}`);
  const contentType = response.headers.get("content-type") || "image/png";
  const buffer = Buffer.from(await response.arrayBuffer());
  return `data:${contentType};base64,${buffer.toString("base64")}`;
}

async function toPptImageData(src) {
  if (!src) return null;
  if (src.startsWith("data:")) return src.replace(/^data:/, "");
  const fetched = await fetchBase64(src);
  return fetched.replace(/^data:/, "");
}

async function addLogo(slide, logoData) {
  const imageData = await toPptImageData(logoData);
  if (!imageData) return;
  slide.addImage({ data: imageData, x: 0.25, y: 0.12, w: 0.9, h: 0.35 });
}

function pxToInX(px) {
  return (Math.max(0, Math.min(960, Number(px) || 0)) / 960) * 10;
}

function pxToInY(px) {
  return (Math.max(0, Math.min(540, Number(px) || 0)) / 540) * 5.625;
}

async function buildPptx(pres, slides, logoData) {
  for (const slideData of slides) {
    const slide = pres.addSlide();
    slide.background = { color: safeHex(slideData.background, "FFFFFF") };

    if (logoData) {
      try {
        await addLogo(slide, logoData);
      } catch (_) {}
    }

    for (const element of slideData.elements || []) {
      const x = pxToInX(element.left);
      const y = pxToInY(element.top);
      const w = pxToInX((Number(element.left) || 0) + (Number(element.width) || 120)) - x;
      const h = pxToInY((Number(element.top) || 0) + (Number(element.height) || 40)) - y;

      if (element.type === "text") {
        const text = String(element.text || "");
        if (!text.trim()) continue;
        slide.addText(text, {
          x,
          y,
          w: Math.max(0.15, w),
          h: Math.max(0.15, h),
          fontFace: element.fontFamily || "Aptos",
          fontSize: Math.max(8, Math.min(34, Math.round((Number(element.fontSize) || 18) * 0.75))),
          color: safeHex(element.color, "1A2320"),
          bold: !!element.bold,
          italic: !!element.italic,
          underline: !!element.underline,
          align: element.align || "left",
          valign: "top",
          margin: 0.06,
          breakLine: false
        });
      } else if (element.type === "shape") {
        slide.addShape(pres.ShapeType.rect, {
          x,
          y,
          w: Math.max(0.08, w),
          h: Math.max(0.08, h),
          line: { color: safeHex(element.stroke || element.fill || "CCCCCC") },
          fill: { color: safeHex(element.fill || "CCCCCC") }
        });
      } else if (element.type === "circle") {
        slide.addShape(pres.ShapeType.ellipse, {
          x,
          y,
          w: Math.max(0.08, w),
          h: Math.max(0.08, h),
          line: element.stroke ? { color: safeHex(element.stroke), pt: Math.max(1, Number(element.strokeWidth) || 1) } : { color: safeHex(element.fill || "CCCCCC") },
          fill: element.fill && element.fill !== "transparent" ? { color: safeHex(element.fill) } : { transparency: 100 }
        });
      } else if (element.type === "image" && element.src) {
        try {
          const imageData = await toPptImageData(element.src);
          if (imageData) {
            slide.addImage({
              data: imageData,
              x,
              y,
              w: Math.max(0.2, w),
              h: Math.max(0.2, h)
            });
          }
        } catch (_) {}
      }
    }

    if (slideData.note) {
      slide.addNotes(String(slideData.note));
    }
  }
}

function getOutlinePrompt({ input, slideCount }) {
  return `You are an elite presentation strategist.

Create a ${slideCount}-slide business presentation based on the brief below.

BRIEF:
${input}

OUTPUT RULES:
- Return ONLY raw JSON.
- Return EXACTLY ${slideCount} array items.
- Every slide object must contain:
  "title", "heading", "subheading", "bullets", "paragraph", "speakerNote", "imageKeyword", "layout", "dark"
- "bullets" must contain 3 to 5 concise bullets.
- Make the first slide a strong cover slide and the last slide a closing slide.
- Use a mix of layouts across the deck.
- Allowed layouts: "cover", "content", "two-column", "big-number", "quote", "image-left", "image-right", "closing"
- Keep writing specific and polished. No fluff.

SCHEMA:
[
  {
    "title": "Short slide title",
    "heading": "Insight-driven headline",
    "subheading": "Short supporting line",
    "bullets": ["Bullet 1", "Bullet 2", "Bullet 3"],
    "paragraph": "A short presenter paragraph.",
    "speakerNote": "One sentence presenter note.",
    "imageKeyword": "2-4 word image idea",
    "layout": "content",
    "dark": false
  }
]`;
}

async function callAnthropic(apiKey, prompt) {
  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 5000,
      messages: [{ role: "user", content: prompt }]
    })
  });

  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(data?.error?.message || `Anthropic API error (${response.status})`);
  }

  const text = Array.isArray(data.content)
    ? data.content.map((block) => block?.text || "").join("")
    : "";

  return cleanJsonText(text);
}

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try {
    const {
      action,
      input = "",
      slideCount = 8,
      style = "professional",
      slides = [],
      logoData = "",
      brandOn = false,
      brandColors = null
    } = req.body || {};

    const theme = getTheme(style, brandOn, brandColors);

    if (action === "outline") {
      const apiKey = process.env.ANTHROPIC_API_KEY;
      if (!apiKey) {
        return res.status(500).json({ error: "Missing ANTHROPIC_API_KEY on Vercel." });
      }

      const prompt = getOutlinePrompt({
        input: String(input || "").trim(),
        slideCount: Math.max(3, Math.min(20, Number(slideCount) || 8))
      });

      const rawJson = await callAnthropic(apiKey, prompt);
      const parsed = JSON.parse(rawJson);
      const outline = normalizeOutline(parsed, theme);

      return res.status(200).json({ outline, theme });
    }

    if (action === "export") {
      if (!Array.isArray(slides) || slides.length === 0) {
        return res.status(400).json({ error: "No slides to export." });
      }

      const pres = new PptxGenJS();
      pres.layout = "LAYOUT_WIDE";
      pres.author = "OpenAI";
      pres.company = "Slideflux";
      pres.subject = "Generated presentation";
      pres.title = "Slideflux deck";
      pres.lang = "en-GB";
      pres.theme = {
        headFontFace: "Aptos Display",
        bodyFontFace: "Aptos",
        lang: "en-GB"
      };

      await buildPptx(pres, slides, logoData);
      const buffer = await pres.write({ outputType: "nodebuffer" });

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
      );
      res.setHeader("Content-Disposition", 'attachment; filename="slideflux-deck.pptx"');
      return res.status(200).send(buffer);
    }

    return res.status(400).json({ error: "Invalid action." });
  } catch (error) {
    return res.status(500).json({ error: error.message || "Unexpected server error." });
  }
};
