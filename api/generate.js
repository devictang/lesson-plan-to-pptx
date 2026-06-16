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
   MARKDOWN PARSER
   ─────────────────────────────────────────── */

function parseLessonPlan(text) {
  const result = {
    title: "",
    meta: {},
    objectives: [],
    materials: [],
    procedures: [],
    assessment: [],
    reflection: "",
  };

  const lines = text.split("\n");

  // Extract title from the first `# ` line
  const titleLine = lines.find((l) => /^#\s+/.test(l));
  if (titleLine) {
    result.title = titleLine.replace(/^#\s+/, "").trim();
    // Strip common prefix labels like "教學主題：" or "主題："
    result.title = result.title.replace(
      /^(教學主題|主題|課題|單元)[：:]\s*/,
      ""
    );
  }

  // Split text into sections by `## ` headers
  const sectionRE = /^##\s+(.+)$/;
  let sections = [];
  let currentHeader = null;
  let currentLines = [];

  for (const line of lines) {
    const m = line.match(sectionRE);
    if (m) {
      if (currentHeader) {
        sections.push({ header: currentHeader, content: currentLines.join("\n") });
      }
      currentHeader = m[1].trim();
      currentLines = [];
    } else if (currentHeader) {
      currentLines.push(line);
    }
  }
  if (currentHeader) {
    sections.push({ header: currentHeader, content: currentLines.join("\n") });
  }

  // Parse each section
  for (const sec of sections) {
    const h = sec.header;
    const c = sec.content.trim();

    if (/基本[資资]訊/.test(h)) {
      result.meta = parseKeyValuePairs(c);
    } else if (/教學目標|學習目標|課程目標/.test(h)) {
      result.objectives = parseNumberedList(c);
    } else if (/教學資源|教材|教具|準備/.test(h)) {
      result.materials = parseBulletList(c);
    } else if (/教學流程|教學過程|教學活動|教學步驟/.test(h)) {
      result.procedures = parseProcedures(c);
    } else if (/評估|評量/.test(h)) {
      result.assessment = parseBulletList(c);
    } else if (/反思|檢討|備註/.test(h)) {
      result.reflection = c;
    }
  }

  return result;
}

/** Parse lines like `- **科目**：中國歷史` into { "科目": "中國歷史" } */
function parseKeyValuePairs(text) {
  const kv = {};
  const re = /[-*]\s*\*{0,2}([^*：:]+)\*{0,2}[：:]\s*(.+)/;
  for (const line of text.split("\n")) {
    const m = line.trim().match(re);
    if (m) kv[m[1].trim()] = m[2].trim();
  }
  return kv;
}

/** Parse `1. item\n2. item` or `- item` into string array */
function parseNumberedList(text) {
  const items = [];
  for (let line of text.split("\n")) {
    line = line.trim();
    // Match "1. ...", "2. ...", or "- ...", or "* ..."
    const m = line.match(/^(?:\d+[.)]\s*|[-*]\s+)(.+)/);
    if (m) items.push(m[1].trim());
  }
  return items;
}

/** Parse bullet list into string array */
function parseBulletList(text) {
  return parseNumberedList(text);
}

/** Parse `### Step Title\ncontent` into structured steps */
function parseProcedures(text) {
  const steps = [];
  const lines = text.split("\n");
  let current = null;

  for (const line of lines) {
    const subMatch = line.match(/^###\s+(.+)/);
    if (subMatch) {
      if (current) steps.push(current);
      current = { title: subMatch[1].trim(), content: "" };
    } else if (current) {
      const trimmed = line.trim();
      if (trimmed) {
        current.content += (current.content ? "\n" : "") + trimmed;
      }
    }
  }
  if (current) steps.push(current);

  // If no ### subheaders found, split by double newline as fallback
  if (steps.length === 0 && text.trim()) {
    const chunks = text.split(/\n\n+/).filter(Boolean);
    chunks.forEach((chunk, i) => {
      const lines = chunk.trim().split("\n");
      const title = lines[0].replace(/^[-*\d.]+\s*/, "");
      const content = lines.slice(1).join("\n").trim();
      steps.push({ title, content: content || chunk.trim() });
    });
  }

  return steps;
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

  // ─── SLIDE 1: TITLE ─────────────────────────
  createTitleSlide(pptx, parsed, t, ++slideNum);

  // ─── SLIDE 2: OBJECTIVES ────────────────────
  if (parsed.objectives.length > 0) {
    createIconListSlide(pptx, {
      title: "教學目標",
      items: parsed.objectives,
      iconType: "number",
      theme: t,
      slideNum: ++slideNum,
    });
  }

  // ─── SLIDE 3: MATERIALS ─────────────────────
  if (parsed.materials.length > 0) {
    createIconListSlide(pptx, {
      title: "教學資源",
      items: parsed.materials,
      iconType: "bullet",
      theme: t,
      slideNum: ++slideNum,
    });
  }

  // ─── SLIDES 4+: PROCEDURES ──────────────────
  parsed.procedures.forEach((step, idx) => {
    createProcedureSlide(pptx, step, t, ++slideNum, idx + 1, parsed.procedures.length);
  });

  // ─── ASSESSMENT ────────────────────────────
  if (parsed.assessment.length > 0) {
    createIconListSlide(pptx, {
      title: "評估方法",
      items: parsed.assessment,
      iconType: "bullet",
      theme: t,
      slideNum: ++slideNum,
    });
  }

  // ─── END SLIDE ─────────────────────────────
  createEndSlide(pptx, t, ++slideNum);

  return pptx;
}

/* ── TITLE SLIDE ── */
function createTitleSlide(pptx, parsed, t, slideNum) {
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

  // Title
  slide.addText(parsed.title || "教案", {
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

  // Subtitle accent line
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: MARGIN,
    y: 3.5,
    w: 1.4,
    h: 0.06,
    fill: { color: t.accent },
  });

  // Meta info
  const metaParts = [];
  if (parsed.meta["科目"] || parsed.meta["學科"]) metaParts.push(parsed.meta["科目"] || parsed.meta["學科"]);
  if (parsed.meta["年級"] || parsed.meta["班級"]) metaParts.push(parsed.meta["年級"] || parsed.meta["班級"]);
  if (parsed.meta["時間"] || parsed.meta["課時"]) metaParts.push(parsed.meta["時間"] || parsed.meta["課時"]);

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

  // Slide number
  addSlideNumber(slide, slideNum, t);
}

/* ── ICON LIST SLIDE (objectives / materials / assessment) ── */
function createIconListSlide(pptx, opts) {
  const { title, items, iconType, theme: t } = opts;
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

  // Title
  slide.addText(title, {
    x: MARGIN,
    y: 0.45,
    w: SLIDE_W - 2 * MARGIN,
    h: 0.7,
    fontSize: 30,
    fontFace: "Microsoft JhengHei",
    color: t.primary,
    bold: true,
  });

  // Items grid
  const startY = 1.4;
  const itemH = 0.85;
  const maxItems = Math.min(items.length, 6);

  items.slice(0, maxItems).forEach((item, idx) => {
    const y = startY + idx * itemH;

    if (iconType === "number") {
      // Number circle
      const circleX = MARGIN;
      const circleSize = 0.52;
      slide.addShape(pptx.shapes.OVAL, {
        x: circleX,
        y: y + 0.08,
        w: circleSize,
        h: circleSize,
        fill: { color: t.primary },
      });
      slide.addText(String(idx + 1), {
        x: circleX,
        y: y + 0.08,
        w: circleSize,
        h: circleSize,
        fontSize: 16,
        fontFace: "Calibri",
        color: t.white,
        bold: true,
        align: "center",
        valign: "middle",
      });
    } else {
      // Bullet dot
      const dotX = MARGIN;
      slide.addShape(pptx.shapes.OVAL, {
        x: dotX,
        y: y + 0.23,
        w: 0.2,
        h: 0.2,
        fill: { color: t.accent },
      });
    }

    // Item text
    slide.addText(item, {
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

  // If more items than fit, add "還有 X 項"
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

  addSlideNumber(slide, opts.slideNum, t);
}

/* ── PROCEDURE SLIDE ── */
function createProcedureSlide(pptx, step, t, slideNum, stepIdx, totalSteps) {
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

  // Section label
  slide.addText("教學流程", {
    x: MARGIN,
    y: 0.3,
    w: 4,
    h: 0.4,
    fontSize: 13,
    fontFace: "Microsoft JhengHei",
    color: t.textMuted,
    bold: false,
  });

  // Large step number
  slide.addText(`0${stepIdx}`, {
    x: MARGIN,
    y: 0.8,
    w: 1.2,
    h: 1.2,
    fontSize: 54,
    fontFace: "Calibri",
    color: t.primary,
    bold: true,
    align: "center",
    valign: "middle",
  });

  // Vertical accent line
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: MARGIN + 1.45,
    y: 0.8,
    w: 0.04,
    h: 5.0,
    fill: { color: t.secondary, transparency: 40 },
  });

  // Step title
  slide.addText(step.title, {
    x: MARGIN + 1.8,
    y: 0.8,
    w: SLIDE_W - MARGIN - 1.8 - MARGIN,
    h: 0.75,
    fontSize: 24,
    fontFace: "Microsoft JhengHei",
    color: t.primary,
    bold: true,
  });

  // Step content
  if (step.content) {
    slide.addText(step.content, {
      x: MARGIN + 1.8,
      y: 1.7,
      w: SLIDE_W - MARGIN - 1.8 - MARGIN,
      h: 3.8,
      fontSize: 15,
      fontFace: "Microsoft JhengHei",
      color: t.textDark,
      lineSpacingMultiple: 1.5,
      valign: "top",
    });
  }

  // Progress bar at bottom
  const barY = 6.85;
  const barW = SLIDE_W - 2 * MARGIN;
  const barH = 0.06;
  const progress = stepIdx / totalSteps;

  // Background bar
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: MARGIN,
    y: barY,
    w: barW,
    h: barH,
    fill: { color: "E2E8F0" },
  });
  // Filled portion
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: MARGIN,
    y: barY,
    w: barW * progress,
    h: barH,
    fill: { color: t.primary },
  });

  addSlideNumber(slide, slideNum, t);
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
