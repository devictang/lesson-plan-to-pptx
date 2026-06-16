const PptxGenJS = require("pptxgenjs");

/* ═══════════════════════════════════════════════════════════════
   8 THEMES — curated from real design systems
   Model: palette(5) + fonts(3) + scale(4) + radius
   ═══════════════════════════════════════════════════════════════ */

const THEMES = {
  notion: { name:"Notion 文青", desc:"溫暖紙白 · 人文氣息", palette:{ bg:"FBFAF8", surface:"FFFFFF", text:"1A1A1A", muted:"9B9A97", accent:"E16259" }, fonts:{ display:"Georgia", body:"Calibri", cjk:"Microsoft JhengHei" }, scale:{ hero:48, title:30, body:17, meta:14 }, radius:10 },
  stripe: { name:"Stripe 藍紫", desc:"專業科技 · 精準自信", palette:{ bg:"FFFFFF", surface:"F8F9FC", text:"061B31", muted:"64748D", accent:"533AFD" }, fonts:{ display:"Calibri Light", body:"Calibri", cjk:"Microsoft JhengHei" }, scale:{ hero:44, title:28, body:16, meta:13 }, radius:6 },
  linear: { name:"Linear 極黑", desc:"極簡暗黑 · 精緻專注", palette:{ bg:"0D0D0D", surface:"1A1A1A", text:"ECEDEE", muted:"6C6E75", accent:"5E6AD2" }, fonts:{ display:"Calibri Light", body:"Calibri", cjk:"Microsoft JhengHei" }, scale:{ hero:46, title:28, body:16, meta:13 }, radius:4 },
  apple:  { name:"Apple 純白",  desc:"簡約純白 · 極致留白", palette:{ bg:"FFFFFF", surface:"F5F5F7", text:"1D1D1F", muted:"86868B", accent:"0071E3" }, fonts:{ display:"Calibri Light", body:"Calibri", cjk:"Microsoft JhengHei" }, scale:{ hero:52, title:30, body:17, meta:14 }, radius:8 },
  airbnb: { name:"Airbnb 暖橙",desc:"溫暖活力 · 圓潤親切", palette:{ bg:"FFFFFF", surface:"F7F7F7", text:"222222", muted:"717171", accent:"FF385C" }, fonts:{ display:"Calibri", body:"Calibri", cjk:"Microsoft JhengHei" }, scale:{ hero:46, title:28, body:16, meta:13 }, radius:14 },
  ocean:  { name:"海洋學術",   desc:"專業沉穩 · 學術風範", palette:{ bg:"F0F5F9", surface:"FFFFFF", text:"1E293B", muted:"64748B", accent:"065A82" }, fonts:{ display:"Calibri", body:"Calibri", cjk:"Microsoft JhengHei" }, scale:{ hero:44, title:28, body:16, meta:13 }, radius:10 },
  forest: { name:"森林自然",   desc:"清新自然 · 平靜柔和", palette:{ bg:"F2F7F0", surface:"FFFFFF", text:"1E293B", muted:"64748B", accent:"2C5F2D" }, fonts:{ display:"Calibri", body:"Calibri", cjk:"Microsoft JhengHei" }, scale:{ hero:44, title:28, body:16, meta:13 }, radius:10 },
  sunset: { name:"日落暖暮",   desc:"溫暖濃郁 · 人文溫度", palette:{ bg:"FEF9F6", surface:"FFFFFF", text:"2D1E1B", muted:"9B8E8A", accent:"B85042" }, fonts:{ display:"Georgia", body:"Calibri", cjk:"Microsoft JhengHei" }, scale:{ hero:46, title:28, body:16, meta:13 }, radius:12 },
};

/* ═══════════════════════════════════════════════════════════════
   PARSER
   ═══════════════════════════════════════════════════════════════ */

function parseLessonPlan(text) {
  const lines = text.split("\n");
  let title = "";
  for (const l of lines) { if (/^#\s+/.test(l)) { title = l.replace(/^#\s+/, "").trim(); break; } }
  const slides = [];
  let cur = null;
  for (const l of lines) {
    const m = l.match(/^##\s+(.+)/);
    if (m) {
      if (cur) { cur.content = cur.content.trim(); slides.push(cur); }
      cur = { title: m[1].trim(), content: "" };
    } else if (cur) { cur.content += l + "\n"; }
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
      items.push({ type:"number", text:numMatch[1] });
    } else if (bulletMatch) {
      items.push({ type:"bullet", text:bulletMatch[1] });
    } else {
      plainLines++;
      if (items.length>0 && !trimmed.match(/^[-*\d]/)) { items[items.length-1].text += " "+trimmed; plainLines--; }
    }
  }
  const totalLines = items.length + plainLines;
  if (items.length<2 || items.length<totalLines*0.5) return [];
  return items;
}

/** Detect slide type for routing */
function detectSlideType(slideData) {
  const content = slideData.content;
  const title = slideData.title;

  // Divider: title is "---" or starts with "--- "
  if (title === "---" || /^---\s/.test(title)) {
    const label = title.replace(/^---\s*/, "").trim() || content.replace(/^[-*]\s*/, "").trim() || "";
    return { type:"divider", label };
  }

  // Quote: content starts with "> "
  if (/^>\s/.test(content)) return { type:"quote" };

  // Table: content has pipe-table format
  if (/\|.+\|/.test(content)) {
    const tbl = parseTable(content);
    if (tbl) return { type:"table", table:tbl };
  }

  // Stats: content is ONLY key-value pairs (1–4) with no other text
  const lines = content.split("\n").map(l=>l.trim()).filter(Boolean);
  const kvRe = /^[-*]\s*\*{0,2}([^*：:]+)\*{0,2}[：:]\s*.+/;
  const kvLines = lines.filter(l => kvRe.test(l));
  const nonKvLines = lines.filter(l => !kvRe.test(l) && l.length>0);
  if (kvLines.length>=1 && kvLines.length<=4 && nonKvLines.length===0) {
    const stats = [];
    for (const l of kvLines) {
      const m = l.match(kvRe);
      if (m) stats.push({ label:m[1].trim(), value:m[2] ? m[2].trim() : l.split(/[：:]/)[1]?.trim() || "" });
    }
    // Extract value from the full match more reliably
    const stats2 = kvLines.map(l => {
      const sep = l.includes("：") ? "：" : ":";
      const parts = l.replace(/^[-*]\s*\*{0,2}/,"").replace(/\*{0,2}$/,"").split(sep);
      return { label:(parts[0]||"").trim(), value:(parts[1]||"").trim() };
    });
    return { type:"stats", stats:stats2 };
  }

  // Default: content slide
  return { type:"content" };
}

/** Parse pipe-table markdown into { headers, rows } */
function parseTable(content) {
  const lines = content.split("\n");
  const pipeLines = [];
  for (const l of lines) {
    const t = l.trim();
    if (t.startsWith("|") && t.endsWith("|")) pipeLines.push(t);
  }
  if (pipeLines.length<2) return null;

  let headerIdx = 0;
  let dataStart = 1;
  // Check if row 1 is a separator like |---|---|
  if (pipeLines[1] && /^\|[-: |]+\|$/.test(pipeLines[1])) {
    dataStart = 2;
  }

  const headers = pipeLines[headerIdx].split("|").slice(1,-1).map(h=>h.trim());
  const rows = [];
  for (let i=dataStart; i<pipeLines.length; i++) {
    const cells = pipeLines[i].split("|").slice(1,-1).map(c=>c.trim());
    if (cells.length>0) rows.push(cells);
  }

  if (headers.length===0 || rows.length===0) return null;
  return { headers, rows };
}

/* ═══════════════════════════════════════════════════════════════
   PPTX GENERATOR
   ═══════════════════════════════════════════════════════════════ */

const SLIDE_W = 13.33;
const SLIDE_H = 7.5;
const MARGIN = 0.85;   // increased from 0.7 for more whitespace

function generatePptx(parsed, themeKey) {
  const t = THEMES[themeKey] || THEMES.ocean;
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name:"WIDE", width:SLIDE_W, height:SLIDE_H });
  pptx.layout = "WIDE";
  pptx.author = "Lesson Plan to PPTX";
  pptx.title = parsed.title || "教案";

  let slideNum = 0;

  if (parsed.slides.length===0) {
    createTitleSlide(pptx, parsed.title, null, t, ++slideNum);
  } else {
    parsed.slides.forEach((sd, idx) => {
      if (idx===0) {
        // First slide = always Title
        createTitleSlide(pptx, parsed.title, sd, t, ++slideNum);
      } else {
        const st = detectSlideType(sd);
        switch(st.type) {
          case "divider": createDividerSlide(pptx, st.label || sd.title, t, ++slideNum); break;
          case "quote":   createQuoteSlide(pptx, sd, t, ++slideNum); break;
          case "table":   createTableSlide(pptx, sd, st.table, t, ++slideNum); break;
          case "stats":   createStatsSlide(pptx, sd, st.stats, t, ++slideNum); break;
          default:        createContentSlide(pptx, sd, t, ++slideNum);
        }
      }
    });
  }

  createEndSlide(pptx, t, ++slideNum);
  return pptx;
}

/* ── TITLE SLIDE ── */
function createTitleSlide(pptx, presentationTitle, slideData, t, slideNum) {
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.accent };

  slide.addShape(pptx.shapes.OVAL, {
    x:SLIDE_W-3.8, y:-1.5, w:6, h:6,
    fill:{ color:"FFFFFF", transparency:92 },
  });

  const displayTitle = (slideData && slideData.title) || presentationTitle || "教案";
  slide.addText(displayTitle, {
    x:MARGIN, y:1.4, w:9.5, h:2.0,
    fontSize:t.scale.hero, fontFace:t.fonts.display,
    color:"FFFFFF", bold:true, align:"left", valign:"middle",
  });

  if (slideData && slideData.content) {
    const meta = parseKeyValuePairs(slideData.content);
    const metaParts = [];
    if (meta["科目"]||meta["學科"]) metaParts.push(meta["科目"]||meta["學科"]);
    if (meta["年級"]||meta["班級"]) metaParts.push(meta["年級"]||meta["班級"]);
    if (meta["時間"]||meta["課時"]) metaParts.push(meta["時間"]||meta["課時"]);
    if (meta["課題"]) metaParts.push(meta["課題"]);
    if (metaParts.length>0) {
      slide.addText(metaParts.join("  ·  "), {
        x:MARGIN, y:3.65, w:9.5, h:0.6,
        fontSize:t.scale.meta, fontFace:t.fonts.body,
        color:"FFFFFF", transparency:25,
      });
    }
  }

  addSlideNumber(slide, slideNum, t);
}

/* ── DIVIDER SLIDE ── */
function createDividerSlide(pptx, label, t, slideNum) {
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.accent };

  // Single subtle circle
  slide.addShape(pptx.shapes.OVAL, {
    x:SLIDE_W/2-2.5, y:SLIDE_H/2-2.5, w:5, h:5,
    fill:{ color:"FFFFFF", transparency:93 },
  });

  slide.addText(label, {
    x:MARGIN, y:0, w:SLIDE_W-2*MARGIN, h:SLIDE_H,
    fontSize:t.scale.hero+4, fontFace:t.fonts.display,
    color:"FFFFFF", bold:true, align:"center", valign:"middle",
  });

  // NO slide number on divider
}

/* ── QUOTE SLIDE ── */
function createQuoteSlide(pptx, slideData, t, slideNum) {
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.bg };

  // Top accent bar
  slide.addShape(pptx.shapes.RECTANGLE, {
    x:0, y:0, w:SLIDE_W, h:0.06,
    fill:{ color:t.palette.accent, transparency:30 },
  });

  // Parse quote lines
  const lines = slideData.content.split("\n");
  const quoteLines = [];
  let attribution = "";
  for (const l of lines) {
    const tL = l.trim();
    if (/^>\s*[—–-]\s/.test(tL)) {
      attribution = tL.replace(/^>\s*[—–-]\s*/, "").trim();
    } else if (/^>\s/.test(tL)) {
      quoteLines.push(tL.replace(/^>\s*/, "").trim());
    }
  }
  const quoteText = quoteLines.join("\n") || slideData.content.replace(/^>\s*/gm, "").trim();

  // Oversized opening quote mark
  slide.addText("\u201C", {
    x:MARGIN+0.3, y:0.9, w:1.8, h:2.0,
    fontSize:t.scale.hero*1.4, fontFace:t.fonts.display,
    color:t.palette.accent, bold:true, align:"left", valign:"top",
  });

  // Quote body
  slide.addText(quoteText, {
    x:MARGIN+1.5, y:1.6, w:SLIDE_W-2*MARGIN-1.5, h:3.2,
    fontSize:t.scale.title, fontFace:t.fonts.display,
    color:t.palette.text, italic:true, lineSpacingMultiple:1.5,
    valign:"middle",
  });

  // Attribution
  if (attribution) {
    slide.addText(attribution, {
      x:MARGIN+1.5, y:5.0, w:SLIDE_W-2*MARGIN-1.5, h:0.5,
      fontSize:t.scale.meta, fontFace:t.fonts.body,
      color:t.palette.muted, align:"right",
    });
  }

  // Slide title as small section label
  slide.addText(slideData.title, {
    x:MARGIN, y:SLIDE_H-0.65, w:SLIDE_W-2*MARGIN, h:0.35,
    fontSize:11, fontFace:t.fonts.body,
    color:t.palette.muted, align:"right",
  });

  addSlideNumber(slide, slideNum, t);
}

/* ── STATS SLIDE ── */
function createStatsSlide(pptx, slideData, stats, t, slideNum) {
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.bg };

  slide.addShape(pptx.shapes.RECTANGLE, {
    x:0, y:0, w:SLIDE_W, h:0.06,
    fill:{ color:t.palette.accent, transparency:30 },
  });

  // Title
  slide.addText(slideData.title, {
    x:MARGIN, y:0.45, w:SLIDE_W-2*MARGIN, h:0.7,
    fontSize:t.scale.title, fontFace:t.fonts.display,
    color:t.palette.text, bold:true,
  });

  const count = stats.length;
  const statY = 2.0;
  const statH = 3.2;

  if (count===1) {
    // Centered big number
    slide.addText(stats[0].value, {
      x:MARGIN, y:statY, w:SLIDE_W-2*MARGIN, h:statH-0.6,
      fontSize:t.scale.hero*1.4, fontFace:t.fonts.display,
      color:t.palette.accent, bold:true, align:"center", valign:"middle",
    });
    slide.addText(stats[0].label, {
      x:MARGIN, y:statY+statH-0.8, w:SLIDE_W-2*MARGIN, h:0.5,
      fontSize:t.scale.meta, fontFace:t.fonts.body,
      color:t.palette.muted, align:"center",
    });
  } else {
    // Multiple stats: side-by-side columns
    const colW = (SLIDE_W-2*MARGIN)/count;
    stats.forEach((s, i) => {
      const cx = MARGIN + i*colW;
      slide.addText(s.value, {
        x:cx, y:statY, w:colW, h:statH-0.6,
        fontSize:t.scale.hero*0.9, fontFace:t.fonts.display,
        color:t.palette.accent, bold:true, align:"center", valign:"middle",
      });
      slide.addText(s.label, {
        x:cx, y:statY+statH-0.8, w:colW, h:0.5,
        fontSize:t.scale.meta, fontFace:t.fonts.body,
        color:t.palette.muted, align:"center",
      });
      // Vertical separator (except last)
      if (i<count-1) {
        slide.addShape(pptx.shapes.RECTANGLE, {
          x:cx+colW-0.01, y:statY+0.3, w:0.02, h:statH-1.0,
          fill:{ color:t.palette.muted, transparency:50 },
        });
      }
    });
  }

  addSlideNumber(slide, slideNum, t);
}

/* ── TABLE SLIDE ── */
function createTableSlide(pptx, slideData, table, t, slideNum) {
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.bg };

  slide.addShape(pptx.shapes.RECTANGLE, {
    x:0, y:0, w:SLIDE_W, h:0.06,
    fill:{ color:t.palette.accent, transparency:30 },
  });

  // Title
  slide.addText(slideData.title, {
    x:MARGIN, y:0.45, w:SLIDE_W-2*MARGIN, h:0.7,
    fontSize:t.scale.title, fontFace:t.fonts.display,
    color:t.palette.text, bold:true,
  });

  // Build table data
  const colCount = table.headers.length;
  const colW = (SLIDE_W-2*MARGIN)/colCount;
  const rowH = 0.55;

  const tableData = [
    // Header row
    table.headers.map(h => ({
      text:h,
      options:{ bold:true, color:"FFFFFF", fill:{ color:t.palette.accent }, fontSize:t.scale.meta+1, fontFace:t.fonts.body, align:"center", valign:"middle" }
    })),
    // Data rows with alternating fill
    ...table.rows.map((row, ri) =>
      row.map(cell => ({
        text:cell,
        options:{
          fontSize:t.scale.body-2, fontFace:t.fonts.body, color:t.palette.text,
          fill:{ color: ri%2===0 ? "FFFFFF" : t.palette.surface },
          valign:"middle", align:"center",
        }
      }))
    ),
  ];

  slide.addTable(tableData, {
    x:MARGIN, y:1.6,
    w:SLIDE_W-2*MARGIN,
    colW: Array(colCount).fill(colW),
    rowH: rowH,
    border:{ type:"solid", pt:0.5, color:t.palette.muted },
    autoPage:false,
  });

  addSlideNumber(slide, slideNum, t);
}

/* ── CONTENT SLIDE ── */
function createContentSlide(pptx, slideData, t, slideNum) {
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.bg };

  slide.addShape(pptx.shapes.RECTANGLE, {
    x:0, y:0, w:SLIDE_W, h:0.06,
    fill:{ color:t.palette.accent, transparency:30 },
  });

  // Title — tighter to top
  slide.addText(slideData.title, {
    x:MARGIN, y:0.45, w:SLIDE_W-2*MARGIN, h:0.7,
    fontSize:t.scale.title, fontFace:t.fonts.display,
    color:t.palette.text, bold:true,
  });

  const items = detectListItems(slideData.content);

  if (items.length>0) {
    renderList(slide, items, t, pptx);
  } else if (slideData.content) {
    // More whitespace, better line-height
    slide.addText(slideData.content, {
      x:MARGIN, y:1.6, w:SLIDE_W-2*MARGIN, h:5.2,
      fontSize:t.scale.body, fontFace:t.fonts.body,
      color:t.palette.text, lineSpacingMultiple:1.75,
      valign:"top",
    });
  }

  addSlideNumber(slide, slideNum, t);
}

/* ── LIST RENDERER ── */
function renderList(slide, items, t, pptx) {
  const startY = 1.55;
  const itemH = 0.95;
  const maxItems = Math.min(items.length, 6);
  const circleSize = 0.48;

  items.slice(0, maxItems).forEach((item, idx) => {
    const y = startY + idx*itemH;

    if (item.type==="number") {
      slide.addShape(pptx.shapes.OVAL, {
        x:MARGIN, y:y+0.10, w:circleSize, h:circleSize,
        fill:{ color:t.palette.accent },
      });
      slide.addText(String(idx+1), {
        x:MARGIN, y:y+0.10, w:circleSize, h:circleSize,
        fontSize:15, fontFace:"Calibri", color:"FFFFFF", bold:true,
        align:"center", valign:"middle",
      });
    } else {
      slide.addShape(pptx.shapes.OVAL, {
        x:MARGIN+0.16, y:y+0.28, w:0.18, h:0.18,
        fill:{ color:t.palette.accent },
      });
    }

    slide.addText(item.text, {
      x:MARGIN+0.72, y:y, w:SLIDE_W-2*MARGIN-0.72, h:itemH,
      fontSize:t.scale.body, fontFace:t.fonts.body,
      color:t.palette.text, valign:"middle",
    });
  });

  if (items.length>maxItems) {
    slide.addText(`還有 ${items.length-maxItems} 項…`, {
      x:MARGIN+0.72, y:startY+maxItems*itemH, w:5, h:0.5,
      fontSize:t.scale.meta, fontFace:t.fonts.body,
      color:t.palette.muted, italic:true,
    });
  }
}

/* ── END SLIDE ── */
function createEndSlide(pptx, t, slideNum) {
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.accent };

  slide.addShape(pptx.shapes.OVAL, {
    x:-2, y:-2, w:6, h:6,
    fill:{ color:"FFFFFF", transparency:92 },
  });

  slide.addText("謝謝", {
    x:0, y:1.5, w:SLIDE_W, h:2,
    fontSize:t.scale.hero, fontFace:t.fonts.display,
    color:"FFFFFF", bold:true, align:"center", valign:"middle",
  });

  slide.addText("Thank You", {
    x:0, y:3.8, w:SLIDE_W, h:0.8,
    fontSize:18, fontFace:t.fonts.body,
    color:"FFFFFF", transparency:30, align:"center",
  });

  addSlideNumber(slide, slideNum, t);
}

/* ── HELPER ── */
function addSlideNumber(slide, num, t) {
  slide.addText(String(num), {
    x:SLIDE_W-0.8, y:SLIDE_H-0.5, w:0.5, h:0.35,
    fontSize:10, fontFace:"Calibri", color:t.palette.muted, align:"right",
  });
}

/* ═══════════════════════════════════════════════════════════════
   VERCEL SERVERLESS HANDLER
   ═══════════════════════════════════════════════════════════════ */

module.exports = async (req, res) => {
  if (req.method==="OPTIONS") {
    res.setHeader("Access-Control-Allow-Origin","*");
    res.setHeader("Access-Control-Allow-Methods","POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers","Content-Type");
    return res.status(204).end();
  }
  if (req.method!=="POST") return res.status(405).json({ error:"Method not allowed — use POST" });

  try {
    const body = typeof req.body==="string" ? JSON.parse(req.body) : req.body;
    const { lessonPlan, theme="ocean" } = body || {};
    if (!lessonPlan || lessonPlan.trim().length<10) return res.status(400).json({ error:"請貼上教案內容（至少 10 個字）" });
    if (!THEMES[theme]) return res.status(400).json({ error:`無效的主題：${theme}。請選擇：${Object.keys(THEMES).join(", ")}` });

    const parsed = parseLessonPlan(lessonPlan);
    if (parsed.slides.length===0) return res.status(400).json({ error:"未偵測到任何 Slide。請確保教案包含至少一個 `## Slide 標題`。" });

    const pptx = generatePptx(parsed, theme);
    const buffer = await pptx.write({ outputType:"nodebuffer" });

    const safeTitle = (parsed.title||"lesson-plan").replace(/[<>:"/\\|?*]/g,"").replace(/\s+/g,"-").slice(0,60);
    res.setHeader("Access-Control-Allow-Origin","*");
    res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition",`attachment; filename*=UTF-8''${encodeURIComponent(safeTitle+".pptx")}`);
    res.setHeader("Content-Length", buffer.length);
    return res.status(200).send(buffer);
  } catch(err) {
    console.error("PPTX generation error:", err);
    return res.status(500).json({ error:"生成 PPTX 時發生錯誤，請檢查教案格式或稍後再試", detail:err.message });
  }
};
