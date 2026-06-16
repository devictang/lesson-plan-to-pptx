const PptxGenJS = require("pptxgenjs");

/* ═══════════════════════════════════════════════════════════════
   8 THEMES
   ═══════════════════════════════════════════════════════════════ */

const THEMES = {
  notion: { name:"Notion 文青", palette:{ bg:"FBFAF8", surface:"FFFFFF", text:"1A1A1A", muted:"9B9A97", accent:"E16259" }, fonts:{ display:"Georgia", body:"Calibri", cjk:"Microsoft JhengHei" }, scale:{ hero:48, title:30, body:17, meta:14 }, radius:10 },
  stripe: { name:"Stripe 藍紫", palette:{ bg:"FFFFFF", surface:"F8F9FC", text:"061B31", muted:"64748D", accent:"533AFD" }, fonts:{ display:"Calibri Light", body:"Calibri", cjk:"Microsoft JhengHei" }, scale:{ hero:44, title:28, body:16, meta:13 }, radius:6 },
  linear: { name:"Linear 極黑", palette:{ bg:"0D0D0D", surface:"1A1A1A", text:"ECEDEE", muted:"6C6E75", accent:"5E6AD2" }, fonts:{ display:"Calibri Light", body:"Calibri", cjk:"Microsoft JhengHei" }, scale:{ hero:46, title:28, body:16, meta:13 }, radius:4 },
  apple:  { name:"Apple 純白",  palette:{ bg:"FFFFFF", surface:"F5F5F7", text:"1D1D1F", muted:"86868B", accent:"0071E3" }, fonts:{ display:"Calibri Light", body:"Calibri", cjk:"Microsoft JhengHei" }, scale:{ hero:52, title:30, body:17, meta:14 }, radius:8 },
  airbnb: { name:"Airbnb 暖橙",palette:{ bg:"FFFFFF", surface:"F7F7F7", text:"222222", muted:"717171", accent:"FF385C" }, fonts:{ display:"Calibri", body:"Calibri", cjk:"Microsoft JhengHei" }, scale:{ hero:46, title:28, body:16, meta:13 }, radius:14 },
  ocean:  { name:"海洋學術",   palette:{ bg:"F0F5F9", surface:"FFFFFF", text:"1E293B", muted:"64748B", accent:"065A82" }, fonts:{ display:"Calibri", body:"Calibri", cjk:"Microsoft JhengHei" }, scale:{ hero:44, title:28, body:16, meta:13 }, radius:10 },
  forest: { name:"森林自然",   palette:{ bg:"F2F7F0", surface:"FFFFFF", text:"1E293B", muted:"64748B", accent:"2C5F2D" }, fonts:{ display:"Calibri", body:"Calibri", cjk:"Microsoft JhengHei" }, scale:{ hero:44, title:28, body:16, meta:13 }, radius:10 },
  sunset: { name:"日落暖暮",   palette:{ bg:"FEF9F6", surface:"FFFFFF", text:"2D1E1B", muted:"9B8E8A", accent:"B85042" }, fonts:{ display:"Georgia", body:"Calibri", cjk:"Microsoft JhengHei" }, scale:{ hero:46, title:28, body:16, meta:13 }, radius:12 },
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
    if (numMatch) { items.push({ type:"number", text:numMatch[1] }); }
    else if (bulletMatch) { items.push({ type:"bullet", text:bulletMatch[1] }); }
    else { plainLines++; if (items.length>0 && !trimmed.match(/^[-*\d]/)) { items[items.length-1].text += " "+trimmed; plainLines--; } }
  }
  const totalLines = items.length + plainLines;
  if (items.length<2 || items.length<totalLines*0.5) return [];
  return items;
}

function parseTable(content) {
  const lines = content.split("\n");
  const pipeLines = [];
  for (const l of lines) { const t = l.trim(); if (t.startsWith("|") && t.endsWith("|")) pipeLines.push(t); }
  if (pipeLines.length<2) return null;
  let dataStart = 1;
  if (pipeLines[1] && /^\|[-: |]+\|$/.test(pipeLines[1])) dataStart = 2;
  const headers = pipeLines[0].split("|").slice(1,-1).map(h=>h.trim());
  const rows = [];
  for (let i=dataStart; i<pipeLines.length; i++) {
    const cells = pipeLines[i].split("|").slice(1,-1).map(c=>c.trim());
    if (cells.length>0) rows.push(cells);
  }
  if (headers.length===0 || rows.length===0) return null;
  return { headers, rows };
}

function detectSlideType(slideData) {
  const content = slideData.content;
  const title = slideData.title;

  // Divider: title is "---" or starts with "--- "
  if (title==="---" || /^---\s/.test(title)) {
    const label = title.replace(/^---\s*/, "").trim() || content.replace(/^[-*]\s*/, "").trim() || "";
    return { type:"divider", label };
  }

  // Split / Comparison: title contains "||"
  if (title.includes("||")) {
    const [leftTitle, rightTitle] = title.split("||").map(s=>s.trim());
    const parts = content.split(/\n---\n/);
    return {
      type:"split",
      leftTitle: leftTitle || "",
      rightTitle: rightTitle || "",
      leftContent: (parts[0]||"").trim(),
      rightContent: (parts[1]||"").trim(),
    };
  }

  // Quote: content starts with "> "
  if (/^>\s/.test(content)) return { type:"quote" };

  // Table: pipe-table
  if (/\|.+\|/.test(content)) {
    const tbl = parseTable(content);
    if (tbl) return { type:"table", table:tbl };
  }

  // Stats: ONLY key-value pairs (1–4)
  const lines = content.split("\n").map(l=>l.trim()).filter(Boolean);
  const kvRe = /^[-*]\s*\*{0,2}([^*：:]+)\*{0,2}[：:]\s*.+/;
  const kvLines = lines.filter(l=>kvRe.test(l));
  const nonKv = lines.filter(l=>!kvRe.test(l) && l.length>0);
  if (kvLines.length>=1 && kvLines.length<=4 && nonKv.length===0) {
    const stats = kvLines.map(l => {
      const sep = l.includes("：") ? "：" : ":";
      const parts = l.replace(/^[-*]\s*\*{0,2}/,"").replace(/\*{0,2}$/,"").split(sep);
      return { label:(parts[0]||"").trim(), value:(parts.slice(1).join(sep)||"").trim() };
    });
    return { type:"stats", stats };
  }

  return { type:"content" };
}

/* ═══════════════════════════════════════════════════════════════
   GENERATOR
   ═══════════════════════════════════════════════════════════════ */

const SLIDE_W = 13.33;
const SLIDE_H = 7.5;
const MARGIN = 0.85;

function generatePptx(parsed, themeKey) {
  const t = THEMES[themeKey] || THEMES.ocean;
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name:"WIDE", width:SLIDE_W, height:SLIDE_H });
  pptx.layout = "WIDE";
  pptx.author = "Lesson Plan to PPTX";
  pptx.title = parsed.title || "教案";

  const total = parsed.slides.length + 1; // +1 for end slide
  let n = 0;
  const ctx = { pptx, t, total, slideNum:()=>++n };

  if (parsed.slides.length===0) {
    createTitleSlide(ctx, parsed.title, null);
  } else {
    parsed.slides.forEach((sd, idx) => {
      if (idx===0) {
        createTitleSlide(ctx, parsed.title, sd);
      } else {
        const st = detectSlideType(sd);
        ctx.st = st; ctx.sd = sd;
        switch(st.type) {
          case "divider": createDividerSlide(ctx, st.label||sd.title); break;
          case "quote":   createQuoteSlide(ctx, sd); break;
          case "table":   createTableSlide(ctx, sd, st.table); break;
          case "stats":   createStatsSlide(ctx, sd, st.stats); break;
          case "split":   createSplitSlide(ctx, sd, st); break;
          default:        createContentSlide(ctx, sd);
        }
      }
    });
  }
  createEndSlide(ctx);
  return pptx;
}

/* ═══════════ SHARED POLISH ═══════════ */

function addProgressBar(slide, ctx) {
  const { t, total, slideNum } = ctx;
  const current = slideNum();
  const barY = SLIDE_H - 0.22;
  const barW = SLIDE_W - 2 * MARGIN;
  const barH = 0.035;
  slide.addShape(ctx.pptx.shapes.RECTANGLE, {
    x:MARGIN, y:barY, w:barW, h:barH,
    fill:{ color:t.palette.muted, transparency:75 },
  });
  slide.addShape(ctx.pptx.shapes.RECTANGLE, {
    x:MARGIN, y:barY, w:barW*(current/total), h:barH,
    fill:{ color:t.palette.accent, transparency:25 },
  });
}

function addWatermark(slide, ctx) {
  const { t } = ctx;
  const num = String(ctx.slideNum());
  slide.addText(num, {
    x:SLIDE_W-2.8, y:SLIDE_H-2.6, w:2.3, h:2.3,
    fontSize:130, fontFace:t.fonts.display,
    color:t.palette.muted, transparency:90, bold:true,
    align:"right", valign:"bottom",
  });
}

function addSlideNumber(slide, ctx) {
  slide.addText(String(ctx.slideNum()), {
    x:SLIDE_W-0.8, y:SLIDE_H-0.5, w:0.5, h:0.35,
    fontSize:10, fontFace:"Calibri", color:ctx.t.palette.muted, align:"right",
  });
}

/* ═══════════ SLIDE TYPES ═══════════ */

function createTitleSlide(ctx, presTitle, slideData) {
  const { pptx, t } = ctx;
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.accent };
  slide.addShape(pptx.shapes.OVAL, {
    x:SLIDE_W-3.8, y:-1.5, w:6, h:6,
    fill:{ color:"FFFFFF", transparency:92 },
  });
  const displayTitle = (slideData && slideData.title) || presTitle || "教案";
  slide.addText(displayTitle, {
    x:MARGIN, y:1.4, w:9.5, h:2.0,
    fontSize:t.scale.hero, fontFace:t.fonts.display,
    color:"FFFFFF", bold:true, align:"left", valign:"middle",
  });
  if (slideData && slideData.content) {
    const meta = parseKeyValuePairs(slideData.content);
    const mp = [];
    if (meta["科目"]||meta["學科"]) mp.push(meta["科目"]||meta["學科"]);
    if (meta["年級"]||meta["班級"]) mp.push(meta["年級"]||meta["班級"]);
    if (meta["時間"]||meta["課時"]) mp.push(meta["時間"]||meta["課時"]);
    if (meta["課題"]) mp.push(meta["課題"]);
    if (mp.length>0) {
      slide.addText(mp.join("  ·  "), {
        x:MARGIN, y:3.65, w:9.5, h:0.6,
        fontSize:t.scale.meta, fontFace:t.fonts.body,
        color:"FFFFFF", transparency:25,
      });
    }
  }
  addWatermark(slide, ctx);
  addSlideNumber(slide, ctx);
}

/* ── DIVIDER ── */
function createDividerSlide(ctx, label) {
  const { pptx, t } = ctx;
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.accent };
  slide.addShape(pptx.shapes.OVAL, {
    x:SLIDE_W/2-2.5, y:SLIDE_H/2-2.5, w:5, h:5,
    fill:{ color:"FFFFFF", transparency:93 },
  });
  slide.addText(label, {
    x:MARGIN, y:0, w:SLIDE_W-2*MARGIN, h:SLIDE_H,
    fontSize:t.scale.hero+4, fontFace:t.fonts.display,
    color:"FFFFFF", bold:true, align:"center", valign:"middle",
  });
  addWatermark(slide, ctx);
}

/* ── QUOTE ── */
function createQuoteSlide(ctx, slideData) {
  const { pptx, t } = ctx;
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.bg };
  slide.addShape(pptx.shapes.RECTANGLE, { x:0, y:0, w:SLIDE_W, h:0.06, fill:{ color:t.palette.accent, transparency:30 } });

  const lines = slideData.content.split("\n");
  const quoteLines = []; let attribution = "";
  for (const l of lines) {
    const tl = l.trim();
    if (/^>\s*[—–-]\s/.test(tl)) attribution = tl.replace(/^>\s*[—–-]\s*/, "").trim();
    else if (/^>\s/.test(tl)) quoteLines.push(tl.replace(/^>\s*/, "").trim());
  }
  const quoteText = quoteLines.join("\n") || slideData.content.replace(/^>\s*/gm, "").trim();

  slide.addText("\u201C", {
    x:MARGIN+0.3, y:0.9, w:1.8, h:2.0,
    fontSize:t.scale.hero*1.4, fontFace:t.fonts.display,
    color:t.palette.accent, bold:true, align:"left", valign:"top",
  });
  slide.addText(quoteText, {
    x:MARGIN+1.5, y:1.6, w:SLIDE_W-2*MARGIN-1.5, h:3.2,
    fontSize:t.scale.title, fontFace:t.fonts.display,
    color:t.palette.text, italic:true, lineSpacingMultiple:1.5, valign:"middle",
  });
  if (attribution) {
    slide.addText(attribution, {
      x:MARGIN+1.5, y:5.0, w:SLIDE_W-2*MARGIN-1.5, h:0.5,
      fontSize:t.scale.meta, fontFace:t.fonts.body, color:t.palette.muted, align:"right",
    });
  }
  slide.addText(slideData.title, {
    x:MARGIN, y:SLIDE_H-0.65, w:SLIDE_W-2*MARGIN, h:0.35,
    fontSize:11, fontFace:t.fonts.body, color:t.palette.muted, align:"right",
  });
  addWatermark(slide, ctx);
  addProgressBar(slide, ctx);
  addSlideNumber(slide, ctx);
}

/* ── STATS ── */
function createStatsSlide(ctx, slideData, stats) {
  const { pptx, t } = ctx;
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.bg };
  slide.addShape(pptx.shapes.RECTANGLE, { x:0, y:0, w:SLIDE_W, h:0.06, fill:{ color:t.palette.accent, transparency:30 } });

  slide.addText(slideData.title, {
    x:MARGIN, y:0.45, w:SLIDE_W-2*MARGIN, h:0.7,
    fontSize:t.scale.title, fontFace:t.fonts.display, color:t.palette.text, bold:true,
  });

  const count = stats.length;
  const statY = 2.0, statH = 3.2;

  if (count===1) {
    slide.addText(stats[0].value, {
      x:MARGIN, y:statY, w:SLIDE_W-2*MARGIN, h:statH-0.6,
      fontSize:t.scale.hero*1.4, fontFace:t.fonts.display,
      color:t.palette.accent, bold:true, align:"center", valign:"middle",
    });
    slide.addText(stats[0].label, {
      x:MARGIN, y:statY+statH-0.8, w:SLIDE_W-2*MARGIN, h:0.5,
      fontSize:t.scale.meta, fontFace:t.fonts.body, color:t.palette.muted, align:"center",
    });
  } else {
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
        fontSize:t.scale.meta, fontFace:t.fonts.body, color:t.palette.muted, align:"center",
      });
      if (i<count-1) {
        slide.addShape(pptx.shapes.RECTANGLE, {
          x:cx+colW-0.01, y:statY+0.3, w:0.02, h:statH-1.0,
          fill:{ color:t.palette.muted, transparency:50 },
        });
      }
    });
  }
  addWatermark(slide, ctx);
  addProgressBar(slide, ctx);
  addSlideNumber(slide, ctx);
}

/* ── TABLE ── */
function createTableSlide(ctx, slideData, table) {
  const { pptx, t } = ctx;
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.bg };
  slide.addShape(pptx.shapes.RECTANGLE, { x:0, y:0, w:SLIDE_W, h:0.06, fill:{ color:t.palette.accent, transparency:30 } });

  slide.addText(slideData.title, {
    x:MARGIN, y:0.45, w:SLIDE_W-2*MARGIN, h:0.7,
    fontSize:t.scale.title, fontFace:t.fonts.display, color:t.palette.text, bold:true,
  });

  const colCount = table.headers.length;
  const colW = (SLIDE_W-2*MARGIN)/colCount;
  const rowH = 0.55;

  const tableData = [
    table.headers.map(h => ({
      text:h, options:{ bold:true, color:"FFFFFF", fill:{ color:t.palette.accent }, fontSize:t.scale.meta+1, fontFace:t.fonts.body, align:"center", valign:"middle" }
    })),
    ...table.rows.map((row, ri) =>
      row.map(cell => ({
        text:cell, options:{
          fontSize:t.scale.body-2, fontFace:t.fonts.body, color:t.palette.text,
          fill:{ color: ri%2===0 ? "FFFFFF" : t.palette.surface },
          valign:"middle", align:"center",
        }
      }))
    ),
  ];

  slide.addTable(tableData, {
    x:MARGIN, y:1.6, w:SLIDE_W-2*MARGIN,
    colW:Array(colCount).fill(colW), rowH,
    border:{ type:"solid", pt:0.5, color:t.palette.muted },
    autoPage:false,
  });

  addWatermark(slide, ctx);
  addProgressBar(slide, ctx);
  addSlideNumber(slide, ctx);
}

/* ── SPLIT (comparison) ── */
function createSplitSlide(ctx, slideData, st) {
  const { pptx, t } = ctx;
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.bg };
  slide.addShape(pptx.shapes.RECTANGLE, { x:0, y:0, w:SLIDE_W, h:0.06, fill:{ color:t.palette.accent, transparency:30 } });

  const midX = SLIDE_W/2;
  const colW = midX - MARGIN - 0.3;
  const colTitleY = 0.55;
  const colBodyY = 1.35;

  // Left title
  slide.addText(st.leftTitle, {
    x:MARGIN, y:colTitleY, w:colW, h:0.55,
    fontSize:t.scale.title-2, fontFace:t.fonts.display, color:t.palette.accent, bold:true, align:"center",
  });
  // Right title
  slide.addText(st.rightTitle, {
    x:midX+0.3, y:colTitleY, w:colW, h:0.55,
    fontSize:t.scale.title-2, fontFace:t.fonts.display, color:t.palette.accent, bold:true, align:"center",
  });

  // Vertical divider
  slide.addShape(pptx.shapes.RECTANGLE, {
    x:midX-0.015, y:colBodyY-0.2, w:0.03, h:5.5,
    fill:{ color:t.palette.muted, transparency:55 },
  });

  // Left content
  if (st.leftContent) {
    const leftItems = detectListItems(st.leftContent);
    if (leftItems.length>0) {
      renderListInColumn(slide, leftItems, t, MARGIN, colBodyY, colW, pptx);
    } else {
      slide.addText(st.leftContent, {
        x:MARGIN+0.2, y:colBodyY, w:colW-0.4, h:5.0,
        fontSize:t.scale.body, fontFace:t.fonts.body, color:t.palette.text, lineSpacingMultiple:1.6, valign:"top",
      });
    }
  }

  // Right content
  if (st.rightContent) {
    const rightItems = detectListItems(st.rightContent);
    if (rightItems.length>0) {
      renderListInColumn(slide, rightItems, t, midX+0.3, colBodyY, colW, pptx);
    } else {
      slide.addText(st.rightContent, {
        x:midX+0.5, y:colBodyY, w:colW-0.4, h:5.0,
        fontSize:t.scale.body, fontFace:t.fonts.body, color:t.palette.text, lineSpacingMultiple:1.6, valign:"top",
      });
    }
  }

  addWatermark(slide, ctx);
  addProgressBar(slide, ctx);
  addSlideNumber(slide, ctx);
}

/* ── CONTENT (with half-bleed panel for short prose) ── */
function createContentSlide(ctx, slideData) {
  const { pptx, t } = ctx;
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.bg };
  slide.addShape(pptx.shapes.RECTANGLE, { x:0, y:0, w:SLIDE_W, h:0.06, fill:{ color:t.palette.accent, transparency:30 } });

  const items = detectListItems(slideData.content);
  const isShortProse = !items.length && slideData.content && slideData.content.length < 180;

  if (isShortProse) {
    // Half-bleed accent panel (left 18%) + content (right 82%)
    const panelW = 2.0;
    slide.addShape(pptx.shapes.RECTANGLE, {
      x:0, y:0.06, w:panelW, h:SLIDE_H-0.06,
      fill:{ color:t.palette.accent, transparency:8 },
    });
    // Title on the panel side?
    slide.addText(slideData.title, {
      x:MARGIN, y:0.45, w:SLIDE_W-2*MARGIN, h:0.7,
      fontSize:t.scale.title, fontFace:t.fonts.display, color:t.palette.text, bold:true,
    });
    slide.addText(slideData.content, {
      x:MARGIN, y:1.8, w:SLIDE_W-2*MARGIN, h:4.8,
      fontSize:t.scale.body+1, fontFace:t.fonts.body, color:t.palette.text, lineSpacingMultiple:1.8, valign:"top",
    });
  } else {
    // Standard layout
    slide.addText(slideData.title, {
      x:MARGIN, y:0.45, w:SLIDE_W-2*MARGIN, h:0.7,
      fontSize:t.scale.title, fontFace:t.fonts.display, color:t.palette.text, bold:true,
    });
    if (items.length>0) {
      renderListFullWidth(slide, items, t, pptx);
    } else if (slideData.content) {
      slide.addText(slideData.content, {
        x:MARGIN, y:1.6, w:SLIDE_W-2*MARGIN, h:5.2,
        fontSize:t.scale.body, fontFace:t.fonts.body, color:t.palette.text, lineSpacingMultiple:1.75, valign:"top",
      });
    }
  }

  addWatermark(slide, ctx);
  addProgressBar(slide, ctx);
  addSlideNumber(slide, ctx);
}

/* ── LIST RENDERERS ── */

function renderListFullWidth(slide, items, t, pptx) {
  const startY = 1.55, itemH = 0.95, maxItems = Math.min(items.length, 6), circleSize = 0.48;
  items.slice(0, maxItems).forEach((item, idx) => {
    const y = startY + idx*itemH;
    if (item.type==="number") {
      slide.addShape(pptx.shapes.OVAL, { x:MARGIN, y:y+0.10, w:circleSize, h:circleSize, fill:{ color:t.palette.accent } });
      slide.addText(String(idx+1), { x:MARGIN, y:y+0.10, w:circleSize, h:circleSize, fontSize:15, fontFace:"Calibri", color:"FFFFFF", bold:true, align:"center", valign:"middle" });
    } else {
      slide.addShape(pptx.shapes.OVAL, { x:MARGIN+0.16, y:y+0.28, w:0.18, h:0.18, fill:{ color:t.palette.accent } });
    }
    slide.addText(item.text, { x:MARGIN+0.72, y:y, w:SLIDE_W-2*MARGIN-0.72, h:itemH, fontSize:t.scale.body, fontFace:t.fonts.body, color:t.palette.text, valign:"middle" });
  });
  if (items.length>maxItems) {
    slide.addText(`還有 ${items.length-maxItems} 項…`, { x:MARGIN+0.72, y:startY+maxItems*itemH, w:5, h:0.5, fontSize:t.scale.meta, fontFace:t.fonts.body, color:t.palette.muted, italic:true });
  }
}

function renderListInColumn(slide, items, t, colX, colY, colW, pptx) {
  const startY = colY, itemH = 0.75, maxItems = Math.min(items.length, 6), circleSize = 0.40;
  items.slice(0, maxItems).forEach((item, idx) => {
    const y = startY + idx*itemH;
    if (item.type==="number") {
      slide.addShape(pptx.shapes.OVAL, { x:colX, y:y+0.06, w:circleSize, h:circleSize, fill:{ color:t.palette.accent } });
      slide.addText(String(idx+1), { x:colX, y:y+0.06, w:circleSize, h:circleSize, fontSize:13, fontFace:"Calibri", color:"FFFFFF", bold:true, align:"center", valign:"middle" });
    } else {
      slide.addShape(pptx.shapes.OVAL, { x:colX+0.12, y:y+0.18, w:0.16, h:0.16, fill:{ color:t.palette.accent } });
    }
    slide.addText(item.text, { x:colX+0.55, y:y, w:colW-0.55, h:itemH, fontSize:t.scale.body-2, fontFace:t.fonts.body, color:t.palette.text, valign:"middle" });
  });
}

/* ── END SLIDE ── */
function createEndSlide(ctx) {
  const { pptx, t } = ctx;
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.accent };
  slide.addShape(pptx.shapes.OVAL, { x:-2, y:-2, w:6, h:6, fill:{ color:"FFFFFF", transparency:92 } });
  slide.addText("謝謝", { x:0, y:1.5, w:SLIDE_W, h:2, fontSize:t.scale.hero, fontFace:t.fonts.display, color:"FFFFFF", bold:true, align:"center", valign:"middle" });
  slide.addText("Thank You", { x:0, y:3.8, w:SLIDE_W, h:0.8, fontSize:18, fontFace:t.fonts.body, color:"FFFFFF", transparency:30, align:"center" });
  addWatermark(slide, ctx);
  addSlideNumber(slide, ctx);
}

/* ═══════════════════════════════════════════════════════════════
   HANDLER
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
    if (!THEMES[theme]) return res.status(400).json({ error:`無效的主題：${theme}` });

    const parsed = parseLessonPlan(lessonPlan);
    if (parsed.slides.length===0) return res.status(400).json({ error:"未偵測到任何 Slide。" });

    const pptx = generatePptx(parsed, theme);
    const buffer = await pptx.write({ outputType:"nodebuffer" });

    const safeTitle = (parsed.title||"lesson-plan").replace(/[<>:"/\\|?*]/g,"").replace(/\s+/g,"-").slice(0,60);
    res.setHeader("Access-Control-Allow-Origin","*");
    res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition",`attachment; filename*=UTF-8''${encodeURIComponent(safeTitle+".pptx")}`);
    res.setHeader("Content-Length", buffer.length);
    return res.status(200).send(buffer);
  } catch(err) {
    console.error("PPTX error:", err);
    return res.status(500).json({ error:"生成 PPTX 時發生錯誤", detail:err.message });
  }
};
