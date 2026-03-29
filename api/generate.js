const PptxGenJS = require("pptxgenjs");

// ── STYLE THEMES ──
const themes = {
  professional: {
    titleBg: "1E2761", titleText: "FFFFFF",
    slideBg: "FFFFFF", slideText: "1A2320",
    accent: "4472C4", subText: "666666"
  },
  teal: {
    titleBg: "1A9E8F", titleText: "FFFFFF",
    slideBg: "FFFFFF", slideText: "1A2320",
    accent: "D95B2A", subText: "4A5553"
  },
  warm: {
    titleBg: "D95B2A", titleText: "FFFFFF",
    slideBg: "FAFAF8", slideText: "1A2320",
    accent: "B04520", subText: "6B4C3B"
  },
  minimal: {
    titleBg: "1A2320", titleText: "FFFFFF",
    slideBg: "FFFFFF", slideText: "1A2320",
    accent: "1A2320", subText: "888888"
  },
  berry: {
    titleBg: "6D2E46", titleText: "FFFFFF",
    slideBg: "FAF7F4", slideText: "1A2320",
    accent: "A26769", subText: "6D4C55"
  },
  forest: {
    titleBg: "2C5F2D", titleText: "FFFFFF",
    slideBg: "FAFAF8", slideText: "1A2320",
    accent: "97BC62", subText: "4A6B3A"
  }
};

module.exports = async function handler(req, res) {

  // Allow requests from any origin (CORS)
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  // Handle preflight
  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const { title, outline, style, format } = req.body;

    // Get theme colours
    const theme = themes[style] || themes.professional;

    // ── BUILD POWERPOINT ──
    const pres = new PptxGenJS();
    pres.layout = "LAYOUT_16x9";
    pres.title = title || "Presentation";

    outline.forEach((slide, index) => {
      const s = pres.addSlide();

      const isTitleSlide = slide.type === "title" || index === 0;
      const isConclusion = slide.type === "conclusion" || slide.type === "cta";
      const isDark = isTitleSlide || isConclusion;

      // Background
      s.background = { color: isDark ? theme.titleBg : theme.slideBg };

      if (isDark) {
        // ── DARK SLIDE (title / conclusion) ──

        // Accent bar on left
        s.addShape(pres.shapes.RECTANGLE, {
          x: 0, y: 0, w: 0.08, h: 5.625,
          fill: { color: theme.accent },
          line: { color: theme.accent }
        });

        // Slide number (top right)
        s.addText(`${index + 1}`, {
          x: 9, y: 0.2, w: 0.8, h: 0.3,
          fontSize: 9, color: "FFFFFF", opacity: 0.4,
          align: "right"
        });

        // Main title
        s.addText(slide.title, {
          x: 0.6, y: 1.8, w: 8.8, h: 1.4,
          fontSize: 38, fontFace: "Calibri",
          bold: true, color: theme.titleText,
          align: "left"
        });

        // Bullets as subtitle
        if (slide.bullets && slide.bullets.length > 0) {
          s.addText(slide.bullets[0], {
            x: 0.6, y: 3.4, w: 7.5, h: 0.6,
            fontSize: 16, fontFace: "Calibri",
            color: theme.titleText, opacity: 0.75,
            align: "left"
          });
        }

      } else {
        // ── LIGHT SLIDE (content) ──

        // Top accent bar
        s.addShape(pres.shapes.RECTANGLE, {
          x: 0, y: 0, w: 10, h: 0.06,
          fill: { color: theme.accent },
          line: { color: theme.accent }
        });

        // Slide number
        s.addText(`${index + 1} / ${outline.length}`, {
          x: 8.5, y: 0.15, w: 1.3, h: 0.25,
          fontSize: 8, color: theme.subText,
          align: "right"
        });

        // Slide title
        s.addText(slide.title, {
          x: 0.5, y: 0.25, w: 8.5, h: 0.75,
          fontSize: 26, fontFace: "Calibri",
          bold: true, color: theme.slideText,
          align: "left", margin: 0
        });

        // Divider line
        s.addShape(pres.shapes.RECTANGLE, {
          x: 0.5, y: 1.05, w: 1.2, h: 0.04,
          fill: { color: theme.accent },
          line: { color: theme.accent }
        });

        // Bullet points
        if (slide.bullets && slide.bullets.length > 0) {
          const bulletItems = slide.bullets.map((b, i) => ({
            text: b,
            options: {
              bullet: true,
              fontSize: 15,
              fontFace: "Calibri",
              color: slide.type === "data" && i === 0 ? theme.accent : theme.slideText,
              bold: slide.type === "data" && i === 0,
              paraSpaceAfter: 8,
              breakLine: i < slide.bullets.length - 1
            }
          }));

          s.addText(bulletItems, {
            x: 0.5, y: 1.3, w: 8.8, h: 3.8,
            valign: "top"
          });
        }

        // Speaker note
        if (slide.speakerNote) {
          s.addNotes(slide.speakerNote);
        }
      }
    });

    // ── WRITE FILE TO BUFFER ──
    const buffer = await pres.write({ outputType: "nodebuffer" });

    // Send the file back
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", `attachment; filename="${(title || "presentation").replace(/[^a-z0-9]/gi, "_")}.pptx"`);
    res.setHeader("Content-Length", buffer.length);

    return res.status(200).send(buffer);

  } catch (err) {
    console.error("Generation error:", err);
    return res.status(500).json({ error: "Failed to generate file", detail: err.message });
  }
};
