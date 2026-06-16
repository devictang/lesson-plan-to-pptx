const PptxGenJS = require("pptxgenjs");

/* ═══════════════════════════════════════════════════════════════
   8 THEMES — curated from real design systems
   Model: palette(5) + fonts(3) + scale(4) + radius
   ═══════════════════════════════════════════════════════════════ */

const THEMES = {
  /* ── ① Notion 文青 ────────────────────────── */
  notion: {
    name: "Notion 文青",
    description: "溫暖紙白 · 人文氣息",
    palette: {
      bg: "FBFAF8",      // warm off-white page
      surface: "FFFFFF",  // card white
      text: "1A1A1A",    // near-black warm ink
      muted: "9B9A97",   // warm gray captions
      accent: "E16259",  // warm coral-red highlight
    },
    fonts: {
      display: "Georgia",
      body: "Calibri",
      cjk: "Microsoft JhengHei",
    },
    scale: { hero: 48, title: 30, body: 17, meta: 14 },
    radius: 10,
  },

  /* ── ② Stripe 藍紫 ────────────────────────── */
  stripe: {
    name: "Stripe 藍紫",
    description: "專業科技 · 精準自信",
    palette: {
      bg: "FFFFFF",
      surface: "F8F9FC",
      text: "061B31",    // deep navy (not pure black)
      muted: "64748D",   // slate descriptions
      accent: "533AFD",  // signature purple
    },
    fonts: {
      display: "Calibri Light",
      body: "Calibri",
      cjk: "Microsoft JhengHei",
    },
    scale: { hero: 44, title: 28, body: 16, meta: 13 },
    radius: 6,
  },

  /* ── ③ Linear 極黑 ────────────────────────── */
  linear: {
    name: "Linear 極黑",
    description: "極簡暗黑 · 精緻專注",
    palette: {
      bg: "0D0D0D",
      surface: "1A1A1A",
      text: "ECEDEE",    // near-white
      muted: "6C6E75",   // dark gray
      accent: "5E6AD2",  // muted purple
    },
    fonts: {
      display: "Calibri Light",
      body: "Calibri",
      cjk: "Microsoft JhengHei",
    },
    scale: { hero: 46, title: 28, body: 16, meta: 13 },
    radius: 4,
  },

  /* ── ④ Apple 純白 ─────────────────────────── */
  apple: {
    name: "Apple 純白",
    description: "簡約純白 · 極致留白",
    palette: {
      bg: "FFFFFF",
      surface: "F5F5F7",
      text: "1D1D1F",    // Apple dark gray
      muted: "86868B",   // Apple gray
      accent: "0071E3",  // Apple blue
    },
    fonts: {
      display: "Calibri Light",
      body: "Calibri",
      cjk: "Microsoft JhengHei",
    },
    scale: { hero: 52, title: 30, body: 17, meta: 14 },
    radius: 8,
  },

  /* ── ⑤ Airbnb 暖橙 ───────────────────────── */
  airbnb: {
    name: "Airbnb 暖橙",
    description: "溫暖活力 · 圓潤親切",
    palette: {
      bg: "FFFFFF",
      surface: "F7F7F7",
      text: "222222",
      muted: "717171",
      accent: "FF385C",  // Airbnb coral
    },
    fonts: {
      display: "Calibri",
      body: "Calibri",
      cjk: "Microsoft JhengHei",
    },
    scale: { hero: 46, title: 28, body: 16, meta: 13 },
    radius: 14,
  },

  /* ── ⑥ 海洋學術 ───────────────────────────── */
  ocean: {
    name: "海洋學術",
    description: "專業沉穩 · 學術風範",
    palette: {
      bg: "F0F5F9",
      surface: "FFFFFF",
      text: "1E293B",
      muted: "64748B",
      accent: "065A82",  // deep ocean blue
    },
    fonts: {
      display: "Calibri",
      body: "Calibri",
      cjk: "Microsoft JhengHei",
    },
    scale: { hero: 44, title: 28, body: 16, meta: 13 },
    radius: 10,
  },

  /* ── ⑦ 森林自然 ───────────────────────────── */
  forest: {
    name: "森林自然",
    description: "清新自然 · 平靜柔和",
    palette: {
      bg: "F2F7F0",
      surface: "FFFFFF",
      text: "1E293B",
      muted: "64748B",
      accent: "2C5F2D",  // deep forest green
    },
    fonts: {
      display: "Calibri",
      body: "Calibri",
      cjk: "Microsoft JhengHei",
    },
    scale: { hero: 44, title: 28, body: 16, meta: 13 },
    radius: 10,
  },

  /* ── ⑧ 日落暖暮 ───────────────────────────── */
  sunset: {
    name: "日落暖暮",
    description: "溫暖濃郁 · 人文溫度",
    palette: {
      bg: "FEF9F6",
      surface: "FFFFFF",
      text: "2D1E1B",
      muted: "9B8E8A",
      accent: "B85042",  // warm terracotta
    },
    fonts: {
      display: "Georgia",
      body: "Calibri",
      cjk: "Microsoft JhengHei",
    },
    scale: { hero: 46, title: 28, body: 16, meta: 13 },
    radius: 12,
  },
};

/* ═══════════════════════════════════════════════════════════════
   MARKDOWN PARSER — Slide-by-Slide
   ═══════════════════════════════════════════════════════════════ */

function parseLessonPlan(text) {
  const lines = text.split("\n");
  let title = "";
  for (const l of lines) {
    if (/^#\s+/.test(l)) { title = l.replace(/^#\s+/, "").trim(); break; }
  }
  const slides = [];
  let cur = null;
  for (const l of lines) {
    const m = l.match(/^##\s+(.+)/);
    if (m) {
      if (cur) { cur.content = cur.content.trim(); slides.push(cur); }
      cur = { title: m[1].trim(), content: "" };
    } else if (cur) {
      cur.content += l + "\n";
    }
  }
  if (cur) { cur.content = cur.content.trim(); slides.push(cur); }
  if (!title && slides.length) title = slides[0].title;
  return { title, slides };
}

/* ── Helpers ── */

function parseKeyValuePairs(text) {
  const kv = {};
  const re = /[-*]\s*\*{0,2}([^*：:]+)\*{0,2}[：:]\s*(.+)/;
  for (const line of text.split("\n")) {
    const m = line.trim().match(re);
    if (m) kv[m[1].trim()] = m[2].trim();
  }
  return kv;
}

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
      if (items.length > 0 && !trimmed.match(/^[-*\d]/)) {
        items[items.length - 1].text += " " + trimmed;
        plainLines--;
      }
    }
  }
  const totalLines = items.length + plainLines;
  if (items.length < 2 || items.length < totalLines * 0.5) return [];
  return items;
}

/* ═══════════════════════════════════════════════════════════════
   PPTX GENERATOR
   ═══════════════════════════════════════════════════════════════ */

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
    createTitleSlide(pptx, parsed.title, null, t, ++slideNum);
  } else {
    parsed.slides.forEach((sd, idx) => {
      if (idx === 0) createTitleSlide(pptx, parsed.title, sd, t, ++slideNum);
      else createContentSlide(pptx, sd, t, ++slideNum);
    });
  }

  createEndSlide(pptx, t, ++slideNum);
  return pptx;
}

/* ── TITLE SLIDE ── */
function createTitleSlide(pptx, presentationTitle, slideData, t, slideNum) {
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.accent };

  // Single subtle decorative shape (NOT two loud circles)
  slide.addShape(pptx.shapes.OVAL, {
    x: SLIDE_W - 3.8,
    y: -1.5,
    w: 6,
    h: 6,
    fill: { color: "FFFFFF", transparency: 92 },
  });

  // Title — use slide title if available, else presentation title
  const displayTitle = (slideData && slideData.title) || presentationTitle || "教案";
  slide.addText(displayTitle, {
    x: MARGIN,
    y: 1.4,
    w: 9.5,
    h: 2.0,
    fontSize: t.scale.hero,
    fontFace: t.fonts.display,
    color: "FFFFFF",
    bold: true,
    align: "left",
    valign: "middle",
  });

  // Meta info (NO accent line — anti-slop rule)
  if (slideData && slideData.content) {
    const meta = parseKeyValuePairs(slideData.content);
    const metaParts = [];
    if (meta["科目"] || meta["學科"]) metaParts.push(meta["科目"] || meta["學科"]);
    if (meta["年級"] || meta["班級"]) metaParts.push(meta["年級"] || meta["班級"]);
    if (meta["時間"] || meta["課時"]) metaParts.push(meta["時間"] || meta["課時"]);
    if (meta["課題"]) metaParts.push(meta["課題"]);

    if (metaParts.length > 0) {
      slide.addText(metaParts.join("  ·  "), {
        x: MARGIN,
        y: 3.65,
        w: 9.5,
        h: 0.6,
        fontSize: t.scale.meta,
        fontFace: t.fonts.body,
        color: "FFFFFF",
        transparency: 25,  // 75% opacity white
      });
    }
  }

  addSlideNumber(slide, slideNum, t);
}

/* ── CONTENT SLIDE ── */
function createContentSlide(pptx, slideData, t, slideNum) {
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.bg };

  // Thin top accent bar (subtle, not loud)
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0,
    y: 0,
    w: SLIDE_W,
    h: 0.06,
    fill: { color: t.palette.accent, transparency: 30 },
  });

  // Slide title
  slide.addText(slideData.title, {
    x: MARGIN,
    y: 0.45,
    w: SLIDE_W - 2 * MARGIN,
    h: 0.7,
    fontSize: t.scale.title,
    fontFace: t.fonts.display,
    color: t.palette.text,
    bold: true,
  });

  // Detect content type
  const items = detectListItems(slideData.content);

  if (items.length > 0) {
    renderList(slide, items, t, pptx);
  } else if (slideData.content) {
    slide.addText(slideData.content, {
      x: MARGIN,
      y: 1.5,
      w: SLIDE_W - 2 * MARGIN,
      h: 5.3,
      fontSize: t.scale.body,
      fontFace: t.fonts.body,
      color: t.palette.text,
      lineSpacingMultiple: 1.65,
      valign: "top",
    });
  }

  addSlideNumber(slide, slideNum, t);
}

/* ── LIST RENDERER ── */
function renderList(slide, items, t, pptx) {
  const startY = 1.4;
  const itemH = 0.9;
  const maxItems = Math.min(items.length, 6);
  const circleSize = 0.48;

  items.slice(0, maxItems).forEach((item, idx) => {
    const y = startY + idx * itemH;

    if (item.type === "number") {
      slide.addShape(pptx.shapes.OVAL, {
        x: MARGIN,
        y: y + 0.10,
        w: circleSize,
        h: circleSize,
        fill: { color: t.palette.accent },
      });
      slide.addText(String(idx + 1), {
        x: MARGIN,
        y: y + 0.10,
        w: circleSize,
        h: circleSize,
        fontSize: 15,
        fontFace: "Calibri",
        color: "FFFFFF",
        bold: true,
        align: "center",
        valign: "middle",
      });
    } else {
      // Bullet dot — uses accent color
      slide.addShape(pptx.shapes.OVAL, {
        x: MARGIN + 0.16,
        y: y + 0.25,
        w: 0.18,
        h: 0.18,
        fill: { color: t.palette.accent },
      });
    }

    slide.addText(item.text, {
      x: MARGIN + 0.72,
      y: y,
      w: SLIDE_W - 2 * MARGIN - 0.72,
      h: itemH,
      fontSize: t.scale.body,
      fontFace: t.fonts.body,
      color: t.palette.text,
      valign: "middle",
    });
  });

  if (items.length > maxItems) {
    slide.addText(`還有 ${items.length - maxItems} 項…`, {
      x: MARGIN + 0.72,
      y: startY + maxItems * itemH,
      w: 5,
      h: 0.5,
      fontSize: t.scale.meta,
      fontFace: t.fonts.body,
      color: t.palette.muted,
      italic: true,
    });
  }
}

/* ── END SLIDE ── */
function createEndSlide(pptx, t, slideNum) {
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.accent };

  // Single subtle shape
  slide.addShape(pptx.shapes.OVAL, {
    x: -2,
    y: -2,
    w: 6,
    h: 6,
    fill: { color: "FFFFFF", transparency: 92 },
  });

  slide.addText("謝謝", {
    x: 0,
    y: 1.5,
    w: SLIDE_W,
    h: 2,
    fontSize: t.scale.hero,
    fontFace: t.fonts.display,
    color: "FFFFFF",
    bold: true,
    align: "center",
    valign: "middle",
  });

  slide.addText("Thank You", {
    x: 0,
    y: 3.8,
    w: SLIDE_W,
    h: 0.8,
    fontSize: 18,
    fontFace: t.fonts.body,
    color: "FFFFFF",
    transparency: 30,
    align: "center",
  });

  addSlideNumber(slide, slideNum, t);
}

/* ── HELPER ── */
function addSlideNumber(slide, num, t) {
  slide.addText(String(num), {
    x: SLIDE_W - 0.8,
    y: SLIDE_H - 0.5,
    w: 0.5,
    h: 0.35,
    fontSize: 10,
    fontFace: "Calibri",
    color: t.palette.muted,
    align: "right",
  });
}

/* ═══════════════════════════════════════════════════════════════
   VERCEL SERVERLESS HANDLER
   ═══════════════════════════════════════════════════════════════ */

module.exports = async (req, res) => {
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
      });
    }

    if (!THEMES[theme]) {
      return res.status(400).json({
        error: `無效的主題：${theme}。請選擇：${Object.keys(THEMES).join(", ")}`,
      });
    }

    const parsed = parseLessonPlan(lessonPlan);

    if (parsed.slides.length === 0) {
      return res.status(400).json({
        error: "未偵測到任何 Slide。請確保教案包含至少一個 `## Slide 標題`。",
      });
    }

    const pptx = generatePptx(parsed, theme);
    const buffer = await pptx.write({ outputType: "nodebuffer" });

    const safeTitle = (parsed.title || "lesson-plan")
      .replace(/[<>:"/\\|?*]/g, "")
      .replace(/\s+/g, "-")
      .slice(0, 60);
    const encodedFilename = encodeURIComponent(safeTitle + ".pptx");

    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodedFilename}`);
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
