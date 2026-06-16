const PptxGenJS = require("pptxgenjs");

/* ───────────────────────────────────────────
   THEME DEFINITIONS (3 beautiful themes)
   ─────────────────────────────────────────── */

const THEMES = {
  ocean: {
    name: "海洋藍",
    primary: "065A82",
    secondary: "1C7293",
    accent: "F0A500",
    dark: "0F2B3D",
    light: "EBF4F6",
    white: "FFFFFF",
    textDark: "1E293B",
    textMuted: "64748B",
    bgContent: "F8FAFB",
  },
  forest: {
    name: "森林綠",
    primary: "2C5F2D",
    secondary: "558B2F",
    accent: "D4A574",
    dark: "1B3319",
    light: "F0F5EE",
    white: "FFFFFF",
    textDark: "1E293B",
    textMuted: "64748B",
    bgContent: "FAFCF9",
  },
  coral: {
    name: "暖橙紅",
    primary: "B85042",
    secondary: "D4796A",
    accent: "2F3C7E",
    dark: "3D1E19",
    light: "FEF5F3",
    white: "FFFFFF",
    textDark: "1E293B",
    textMuted: "64748B",
    bgContent: "FEFCFB",
  },
};

/* ───────────────────────────────────────────
   MARKDOWN PARSER — Slide-by-Slide
   ─────────────────────────────────────────── */

/**
 * Parses structured markdown into slide-level objects.
 * 
 * Rules:
 *  - `# ...` line  → presentation title (used for filename + title slide)
 *  - `## ...` line → new slide boundary (slide title)
 *  - Everything between `##` lines → that slide's content
 *  - First `##` slide → rendered as Title Slide (dark bg, big type)
 *  - Every subsequent `##` slide → Content Slide (light bg)
 *  - Content slides auto-detect lists (`1. ` / `- `) and render with
 *    number circles or bullet dots; plain text is rendered as prose.
 */
function parseLessonPlan(text) {
  const lines = text.split("\n");

  // Extract presentation-level title from first `# ` line
  let title = "";
  for (const line of lines) {
    if (/^#\s+/.test(line)) {
      title = line.replace(/^#\s+/, "").trim();
      break;
    }
  }

  // Extract slides: each `## ` opens a new slide
  const slides = [];
  let currentSlide = null;

  for (const line of lines) {
    const h2Match = line.match(/^##\s+(.+)/);
    if (h2Match) {
      // Save previous slide
      if (currentSlide) {
        currentSlide.content = currentSlide.content.trim();
        slides.push(currentSlide);
      }
      currentSlide = { title: h2Match[1].trim(), content: "" };
    } else if (currentSlide) {
      currentSlide.content += line + "\n";
    }
  }

  // Don't forget the last slide
  if (currentSlide) {
    currentSlide.content = currentSlide.content.trim();
    slides.push(currentSlide);
  }

  // Fallback: if no `# ` title found, use first slide's title
  if (!title && slides.length > 0) {
    title = slides[0].title;
  }

  return { title, slides };
}

/* ───────────────────────────────────────────
   CONTENT PARSING HELPERS
   ─────────────────────────────────────────── */

/** Parse `- **key**：value` pairs into { key: value } */
function parseKeyValuePairs(text) {
  const kv = {};
  const re = /[-*]\s*\*{0,2}([^*：:]+)\*{0,2}[：:]\s*(.+)/;
  for (const line of text.split("\n")) {
    const m = line.trim().match(re);
    if (m) kv[m[1].trim()] = m[2].trim();
  }
  return kv;
}

/**
 * Detect list items in slide content.
 * Returns array of { type: 'number'|'bullet', text: string }.
 * If fewer than 2 list lines found, returns empty (fall back to plain text).
 */
function detectListItems(content) {
  const lines = content.split("\n");
  const items = [];
  let plainLines = 0;

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;

    const numMatch = trimmed.match(/^\d+[.)]\s+(.+)/);
    const bulletMatch = trimmed.match(/^[-*]\s+(.+)/);

    if (numMatch) {
      items.push({ type: "number", text: numMatch[1] });
    } else if (bulletMatch) {
      items.push({ type: "bullet", text: bulletMatch[1] });
    } else {
      plainLines++;
      // If there's a preceding item, append as continuation
      if (items.length > 0 && !trimmed.match(/^[-*\d]/)) {
        items[items.length - 1].text += " " + trimmed;
        plainLines--; // don't count continuation as separate plain line
      }
    }
  }

  // Only use list mode if the majority of non-empty lines are list items
  const totalLines = items.length + plainLines;
  if (items.length < 2 || items.length < totalLines * 0.5) {
    return [];
  }

  return items;
}

/* ───────────────────────────────────────────
   PPTX GENERATOR
   ─────────────────────────────────────────── */

// Widescreen 16:9 dimensions in inches
const SLIDE_W = 13.33;
const SLIDE_H = 7.5;
const MARGIN = 0.7;

function generatePptx(parsed, themeKey) {
  const t = THEMES[themeKey] || THEMES.ocean;
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "WIDE", width: SLIDE_W, height: SLIDE_H });
  pptx.layout = "WIDE";
  pptx.author = "Lesson Plan to PPTX";
  pptx.title = parsed.title || "教案";

  let slideNum = 0;

  if (parsed.slides.length === 0) {
    // Empty input — just title + end
    createTitleSlide(pptx, parsed.title, null, t, ++slideNum);
  } else {
    parsed.slides.forEach((slideData, idx) => {
      if (idx === 0) {
        // First slide = Title Slide (dark background)
        createTitleSlide(pptx, parsed.title, slideData, t, ++slideNum);
      } else {
        // Content slide (auto-detect list vs prose)
        createContentSlide(pptx, slideData, t, ++slideNum);
      }
    });
  }

  // Always append End Slide
  createEndSlide(pptx, t, ++slideNum);

  return pptx;
}

/* ── TITLE SLIDE ── */
function createTitleSlide(pptx, presentationTitle, slideData, t, slideNum) {
  const slide = pptx.addSlide();
  slide.background = { fill: t.primary };

  // Large decorative circle (semi-transparent)
  slide.addShape(pptx.shapes.OVAL, {
    x: SLIDE_W - 3.5,
    y: -1.2,
    w: 5.5,
    h: 5.5,
    fill: { color: t.secondary, transparency: 55 },
  });

  // Small decorative circle
  slide.addShape(pptx.shapes.OVAL, {
    x: SLIDE_W - 1.2,
    y: 5.2,
    w: 2.2,
    h: 2.2,
    fill: { color: t.secondary, transparency: 70 },
  });

  // Display title: use slide title if available, else presentation title
  const displayTitle = (slideData && slideData.title) || presentationTitle || "教案";
  slide.addText(displayTitle, {
    x: MARGIN,
    y: 1.6,
    w: 9.5,
    h: 1.8,
    fontSize: 42,
    fontFace: "Microsoft JhengHei",
    color: t.white,
    bold: true,
    align: "left",
    valign: "middle",
  });

  // Accent line
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: MARGIN,
    y: 3.5,
    w: 1.4,
    h: 0.06,
    fill: { color: t.accent },
  });

  // Meta info from first slide's key-value pairs
  if (slideData && slideData.content) {
    const meta = parseKeyValuePairs(slideData.content);
    const metaParts = [];
    if (meta["科目"] || meta["學科"]) metaParts.push(meta["科目"] || meta["學科"]);
    if (meta["年級"] || meta["班級"]) metaParts.push(meta["年級"] || meta["班級"]);
    if (meta["時間"] || meta["課時"]) metaParts.push(meta["時間"] || meta["課時"]);

    if (metaParts.length > 0) {
      slide.addText(metaParts.join("  ·  "), {
        x: MARGIN,
        y: 3.85,
        w: 9.5,
        h: 0.6,
        fontSize: 16,
        fontFace: "Microsoft JhengHei",
        color: t.light,
        align: "left",
      });
    }
  }

  addSlideNumber(slide, slideNum, t);
}

/* ── CONTENT SLIDE (unified) ── */
function createContentSlide(pptx, slideData, t, slideNum) {
  const slide = pptx.addSlide();
  slide.background = { fill: t.bgContent };

  // Top accent bar
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0,
    y: 0,
    w: SLIDE_W,
    h: 0.08,
    fill: { color: t.primary },
  });

  // Slide title
  slide.addText(slideData.title, {
    x: MARGIN,
    y: 0.45,
    w: SLIDE_W - 2 * MARGIN,
    h: 0.7,
    fontSize: 30,
    fontFace: "Microsoft JhengHei",
    color: t.primary,
    bold: true,
  });

  // Detect content type
  const items = detectListItems(slideData.content);

  if (items.length > 0) {
    // Render as list with visual indicators
    renderList(slide, items, t, pptx);
  } else if (slideData.content) {
    // Render as plain text
    slide.addText(slideData.content, {
      x: MARGIN,
      y: 1.5,
      w: SLIDE_W - 2 * MARGIN,
      h: 5.5,
      fontSize: 16,
      fontFace: "Microsoft JhengHei",
      color: t.textDark,
      lineSpacingMultiple: 1.6,
      valign: "top",
    });
  }

  addSlideNumber(slide, slideNum, t);
}

/* ── LIST RENDERER (number circles + bullet dots) ── */
function renderList(slide, items, t, pptx) {
  const startY = 1.4;
  const itemH = 0.9;
  const maxItems = Math.min(items.length, 6);

  items.slice(0, maxItems).forEach((item, idx) => {
    const y = startY + idx * itemH;

    if (item.type === "number") {
      // Numbered circle
      slide.addShape(pptx.shapes.OVAL, {
        x: MARGIN,
        y: y + 0.08,
        w: 0.52,
        h: 0.52,
        fill: { color: t.primary },
      });
      slide.addText(String(idx + 1), {
        x: MARGIN,
        y: y + 0.08,
        w: 0.52,
        h: 0.52,
        fontSize: 16,
        fontFace: "Calibri",
        color: t.white,
        bold: true,
        align: "center",
        valign: "middle",
      });
    } else {
      // Bullet dot
      slide.addShape(pptx.shapes.OVAL, {
        x: MARGIN + 0.16,
        y: y + 0.23,
        w: 0.2,
        h: 0.2,
        fill: { color: t.accent },
      });
    }

    // Item text
    slide.addText(item.text, {
      x: MARGIN + 0.75,
      y: y,
      w: SLIDE_W - 2 * MARGIN - 0.75,
      h: itemH,
      fontSize: 17,
      fontFace: "Microsoft JhengHei",
      color: t.textDark,
      valign: "middle",
    });
  });

  if (items.length > maxItems) {
    slide.addText(`還有 ${items.length - maxItems} 項…`, {
      x: MARGIN + 0.75,
      y: startY + maxItems * itemH,
      w: 5,
      h: 0.5,
      fontSize: 14,
      fontFace: "Microsoft JhengHei",
      color: t.textMuted,
      italic: true,
    });
  }
}

/* ── END SLIDE ── */
function createEndSlide(pptx, t, slideNum) {
  const slide = pptx.addSlide();
  slide.background = { fill: t.primary };

  // Decorative circle
  slide.addShape(pptx.shapes.OVAL, {
    x: -1.5,
    y: -1.5,
    w: 5.5,
    h: 5.5,
    fill: { color: t.secondary, transparency: 55 },
  });

  slide.addText("謝謝", {
    x: 0,
    y: 1.5,
    w: SLIDE_W,
    h: 2,
    fontSize: 48,
    fontFace: "Microsoft JhengHei",
    color: t.white,
    bold: true,
    align: "center",
    valign: "middle",
  });

  slide.addText("Thank You", {
    x: 0,
    y: 3.6,
    w: SLIDE_W,
    h: 1,
    fontSize: 20,
    fontFace: "Calibri",
    color: t.light,
    align: "center",
  });

  addSlideNumber(slide, slideNum, t);
}

/* ── HELPER: Slide number ── */
function addSlideNumber(slide, num, t) {
  slide.addText(String(num), {
    x: SLIDE_W - 0.8,
    y: SLIDE_H - 0.5,
    w: 0.5,
    h: 0.35,
    fontSize: 10,
    fontFace: "Calibri",
    color: t.textMuted,
    align: "right",
  });
}

/* ───────────────────────────────────────────
   VERCEL SERVERLESS HANDLER
   ─────────────────────────────────────────── */

module.exports = async (req, res) => {
  // CORS preflight
  if (req.method === "OPTIONS") {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type");
    return res.status(204).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed — use POST" });
  }

  try {
    const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
    const { lessonPlan, theme = "ocean" } = body || {};

    if (!lessonPlan || lessonPlan.trim().length < 10) {
      return res.status(400).json({
        error: "請貼上教案內容（至少 10 個字）",
        detail: "Lesson plan text is required and must be at least 10 characters",
      });
    }

    if (!THEMES[theme]) {
      return res.status(400).json({
        error: `無效的主題：${theme}。請選擇：${Object.keys(THEMES).join(", ")}`,
      });
    }

    // Parse
    const parsed = parseLessonPlan(lessonPlan);

    if (parsed.slides.length === 0) {
      return res.status(400).json({
        error: "未偵測到任何 Slide。請確保教案包含至少一個 `## Slide 標題`。",
        detail: "No slides detected. Add `## Slide Title` sections.",
      });
    }

    // Generate PPTX
    const pptx = generatePptx(parsed, theme);
    const buffer = await pptx.write({ outputType: "nodebuffer" });

    // Sanitize filename
    const safeTitle = (parsed.title || "lesson-plan")
      .replace(/[<>:"/\\|?*]/g, "")
      .replace(/\s+/g, "-")
      .slice(0, 60);
    const encodedFilename = encodeURIComponent(safeTitle + ".pptx");

    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename*=UTF-8''${encodedFilename}`
    );
    res.setHeader("Content-Length", buffer.length);

    return res.status(200).send(buffer);
  } catch (err) {
    console.error("PPTX generation error:", err);
    return res.status(500).json({
      error: "生成 PPTX 時發生錯誤，請檢查教案格式或稍後再試",
      detail: err.message,
    });
  }
};
