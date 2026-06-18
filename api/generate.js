const PptxGenJS = require("pptxgenjs");

/* ═══════════════════════════════════════════════════════════════
   12 THEMES — extracted from nexu-io/open-design (Apache 2.0)
   ═══════════════════════════════════════════════════════════════ */

const THEMES = {
  // ── Academic / Serious ──
  "academic-paper": {
    name:"學術論文", palette:{ bg:"FDFCF8", surface:"FFFFFF", text:"0A0A0A", muted:"707070", accent:"1A3A7A", accent2:"0A0A0A", accent3:"8A1A1A", good:"1A5A2A", warn:"8A6A1A", bad:"8A1A1A" },
    fonts:{ display:"Georgia", body:"Georgia", mono:"Courier New", cjk:"Microsoft JhengHei" },
    scale:{ hero:44, title:30, body:16, meta:14, small:11 }, radius:0, shadow:false, gradient:false,
    desc:"Serif 全場、零圓角、零陰影。適合論文答辯、學術會議。"
  },
  // ── Warm Education / Course Module (BEST for lesson plans) ──
  "course-module": {
    name:"教學模組 ★", palette:{ bg:"FBFBF6", surface:"FFFFFF", text:"2A2418", muted:"8A7F68", accent:"2D7D6E", accent2:"D88A3A", accent3:"C4593F", good:"1A5A2A", warn:"8A6A1A", bad:"C4593F" },
    fonts:{ display:"Georgia", body:"Calibri", mono:"Courier New", cjk:"Microsoft JhengHei" },
    scale:{ hero:48, title:32, body:16, meta:14, small:11 }, radius:12, shadow:true, gradient:false,
    desc:"暖米色底、Playfair 標題、左 sidebar 學習目標進度。最適合教案！"
  },
  // ── Corporate / Professional ──
  "corporate-clean": {
    name:"企業彙報", palette:{ bg:"FFFFFF", surface:"F8F9FC", text:"1A1D23", muted:"8A90A0", accent:"1A56DB", accent2:"7C3AED", accent3:"DB2777", good:"059669", warn:"D97706", bad:"DC2626" },
    fonts:{ display:"Calibri Light", body:"Calibri", mono:"Consolas", cjk:"Microsoft JhengHei" },
    scale:{ hero:44, title:30, body:16, meta:14, small:12 }, radius:12, shadow:true, gradient:false,
    desc:"藍色主調、乾淨俐落。適合行政彙報、學校會議。"
  },
  // ── Minimal / White ──
  "minimal-white": {
    name:"極簡白", palette:{ bg:"FFFFFF", surface:"FAFAFA", text:"0A0A0A", muted:"A3A3A3", accent:"18181B", accent2:"525252", accent3:"A3A3A3", good:"16A34A", warn:"CA8A04", bad:"DC2626" },
    fonts:{ display:"Calibri Light", body:"Calibri", mono:"Courier New", cjk:"Microsoft JhengHei" },
    scale:{ hero:40, title:28, body:15, meta:13, small:11 }, radius:8, shadow:true, gradient:false,
    desc:"黑白色系、大量留白。適合設計/藝術課，或全彩圖片主導嘅簡報。"
  },
  // ── Swiss Grid / Data ──
  "swiss-grid": {
    name:"瑞士網格", palette:{ bg:"F8F8F8", surface:"FFFFFF", text:"1A1A1A", muted:"808080", accent:"E30613", accent2:"1A1A1A", accent3:"404040", good:"006E4E", warn:"C44C00", bad:"E30613" },
    fonts:{ display:"Arial", body:"Arial", mono:"Consolas", cjk:"Microsoft JhengHei" },
    scale:{ hero:44, title:30, body:15, meta:13, small:11 }, radius:4, shadow:false, gradient:false,
    desc:"紅色 accent、強烈網格對齊。適合資料分析、統計課、圖表主導。"
  },
  // ── Editorial Serif / Humanities ──
  "editorial-serif": {
    name:"編輯社論", palette:{ bg:"FEFEFE", surface:"FFFFFF", text:"1A1513", muted:"8A8078", accent:"8B2F2F", accent2:"1A1513", accent3:"B06030", good:"2D6A4F", warn:"9A6A1A", bad:"8B2F2F" },
    fonts:{ display:"Georgia", body:"Calibri", mono:"Courier New", cjk:"Microsoft JhengHei" },
    scale:{ hero:48, title:32, body:16, meta:14, small:11 }, radius:12, shadow:true, gradient:false,
    desc:"Playfair 標題、深紅 accent。適合文學、歷史、哲學、語文課。"
  },
  // ── Tech / Dark ──
  "tech-sharing": {
    name:"技術暗色", palette:{ bg:"0D1117", surface:"161B22", text:"E6EDF3", muted:"6E7681", accent:"58A6FF", accent2:"A371F7", accent3:"F0883E", good:"3FB950", warn:"D29922", bad:"F85149" },
    fonts:{ display:"Calibri Light", body:"Calibri", mono:"Consolas", cjk:"Microsoft JhengHei" },
    scale:{ hero:44, title:30, body:15, meta:13, small:11 }, radius:12, shadow:true, gradient:false,
    desc:"GitHub dark 配色、code block 融合。適合編程、CS、STEM 課。"
  },
  // ── Pastel / Kids ──
  "pastel-warm": {
    name:"柔和馬卡龍", palette:{ bg:"FEF8F1", surface:"FFFFFF", text:"2D2418", muted:"A09080", accent:"E8927C", accent2:"7EC8B8", accent3:"B8A4D6", good:"7EC8B8", warn:"E8C87A", bad:"E8927C" },
    fonts:{ display:"Calibri Light", body:"Calibri", mono:"Courier New", cjk:"Microsoft JhengHei" },
    scale:{ hero:44, title:30, body:16, meta:14, small:11 }, radius:24, shadow:true, gradient:false,
    desc:"奶油色底、大圓角、柔和彩色卡。適合小學、幼教、通識。"
  },
  // ── Blueprint / Engineering ──
  "blueprint": {
    name:"藍圖工程", palette:{ bg:"1A3A5C", surface:"1E4D7A", text:"E8F0F8", muted:"7090B0", accent:"FFD700", accent2:"FF8C00", accent3:"FF4444", good:"44FF88", warn:"FFD700", bad:"FF4444" },
    fonts:{ display:"Calibri Light", body:"Calibri", mono:"Consolas", cjk:"Microsoft JhengHei" },
    scale:{ hero:44, title:30, body:15, meta:13, small:11 }, radius:2, shadow:false, gradient:false,
    desc:"深藍底 + 金黃 accent。適合物理、工程、數學、科學。"
  },
  // ── Nature / Biology ──
  "nature-green": {
    name:"森林自然", palette:{ bg:"F4F8F0", surface:"FFFFFF", text:"1A2E14", muted:"6A8A58", accent:"2D7A2D", accent2:"5A8A3A", accent3:"B06030", good:"2D7A2D", warn:"B08020", bad:"C04030" },
    fonts:{ display:"Calibri", body:"Calibri", mono:"Courier New", cjk:"Microsoft JhengHei" },
    scale:{ hero:44, title:30, body:16, meta:14, small:11 }, radius:16, shadow:true, gradient:false,
    desc:"綠色為主調。適合生物、地理、環境科學。"
  },
  // ── Ocean / Medical ──
  "ocean-blue": {
    name:"海洋冷色", palette:{ bg:"F0F4F8", surface:"FFFFFF", text:"0A1A2E", muted:"6080A0", accent:"1A6AA0", accent2:"4A8AC0", accent3:"C04060", good:"1A8A6A", warn:"B08020", bad:"C04040" },
    fonts:{ display:"Calibri Light", body:"Calibri", mono:"Consolas", cjk:"Microsoft JhengHei" },
    scale:{ hero:44, title:30, body:16, meta:14, small:11 }, radius:14, shadow:true, gradient:false,
    desc:"冷藍色系、冷靜專業。適合醫學、化學、物理、理科。"
  },
  // ── History / Amber ──
  "history-amber": {
    name:"歷史琥珀", palette:{ bg:"F5F0E0", surface:"FAF5E8", text:"2A2015", muted:"8A7A65", accent:"B5392A", accent2:"5A3A1A", accent3:"8A5A2A", good:"4A6A3A", warn:"8A6A1A", bad:"B5392A" },
    fonts:{ display:"Georgia", body:"Calibri", mono:"Courier New", cjk:"Microsoft JhengHei" },
    scale:{ hero:48, title:32, body:16, meta:14, small:11 }, radius:0, shadow:false, gradient:false,
    desc:"羊皮紙底色、磚紅 accent。適合歷史、社會學、人類學。"
  }
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

function parseCodeBlock(content) {
  const lines = content.split("\n");
  let inCode = false, codeLines = [], caption = "";
  for (const l of lines) {
    if (/^```/.test(l.trim())) {
      if (inCode) { inCode = false; }
      else { inCode = true; const m = l.trim().match(/^```(\w*)/); if (m) caption = m[1]; }
    } else if (inCode) { codeLines.push(l); }
  }
  if (codeLines.length>0) return { code: codeLines.join("\n"), lang: caption };
  return null;
}

function parseSpeakerNotes(content) {
  const notesMatch = content.match(/>\s*notes?[：:]\s*([\s\S]*?)(?=\n##|\n> \*\*|\n> 💡|\n> ✏️|\n> ❓|\n> ---|\n$)/i);
  if (notesMatch) return notesMatch[1].trim().replace(/\n> /g, "\n");
  return null;
}

function detectSlideType(slideData, slideIndex, allSlides) {
  const content = slideData.content;
  const title = slideData.title;

  // Divider
  if (title==="---" || /^---\s/.test(title)) {
    const label = title.replace(/^---\s*/, "").trim() || content.replace(/^[-*]\s*/, "").trim() || "";
    return { type:"divider", label };
  }

  // Split / Comparison
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

  // ── NEW: Callout block (💡 提示) ──
  if (/^>\s*💡/.test(content)) {
    const calloutLines = [];
    for (const l of content.split("\n")) {
      const t = l.replace(/^>\s*/, "").trim();
      if (t) calloutLines.push(t);
    }
    const calloutText = calloutLines.join(" ").replace(/^💡\s*/, "");
    return { type:"callout", calloutText };
  }

  // ── NEW: Exercise block (✏️ 練習) ──
  if (/^>\s*✏️/.test(content)) {
    const exerciseLines = [];
    for (const l of content.split("\n")) {
      const t = l.replace(/^>\s*/, "").trim();
      if (t) exerciseLines.push(t);
    }
    const exerciseTitle = exerciseLines[0] ? exerciseLines[0].replace(/^✏️\s*/, "") : "練習";
    const exerciseBody = exerciseLines.slice(1).join("\n");
    return { type:"exercise", exerciseTitle, exerciseBody };
  }

  // ── NEW: MCQ block (❓ 測驗) ──
  if (/^>\s*❓/.test(content)) {
    const lines = content.split("\n").map(l=>l.replace(/^>\s*/,"").trim()).filter(Boolean);
    const mcqTitle = lines[0] ? lines[0].replace(/^❓\s*/,"") : "";
    const options = []; let answerIdx = -1;
    let curLetter = "A";
    for (let i=1; i<lines.length; i++) {
      const optMatch = lines[i].match(/^([A-D])[.)]\s+(.+)/);
      if (optMatch) { options.push({ letter:optMatch[1], text:optMatch[2], correct:false }); curLetter = String.fromCharCode(optMatch[1].charCodeAt(0)+1); }
      else if (/^✅/.test(lines[i])) { const c = lines[i].replace(/^✅\s*/,""); if (options.length>0) { options[options.length-1].correct = true; } }
    }
    return { type:"mcq", mcqTitle, options };
  }

  // ── NEW: Concept-box (概念卡：> **概念**：xxx) ──
  if (/^>\s*\*\*[^*]+\*\*[：:]/.test(content)) {
    const concepts = [];
    for (const l of content.split("\n")) {
      const t = l.replace(/^>\s*/, "").trim();
      const m = t.match(/^\*\*([^*]+)\*\*[：:]\s*(.+)/);
      if (m) concepts.push({ name: m[1], desc: m[2] });
    }
    if (concepts.length>=2) return { type:"concept-box", concepts };
  }

  // Code block (```...```)
  const codeBlock = parseCodeBlock(content);
  if (codeBlock) return { type:"code", code:codeBlock.code, lang:codeBlock.lang };

  // Quote
  if (/^>\s/.test(content)) return { type:"quote" };

  // Table
  if (/\|.+\|/.test(content)) {
    const tbl = parseTable(content);
    if (tbl) return { type:"table", table:tbl };
  }

  // Stats: ONLY key-value pairs (1-4)
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

  // ── NEW: Timeline ──
  if (/時間線|timeline|roadmap|路線圖|進度/.test(title) && lines.length>=2) {
    const timelineItems = detectListItems(content);
    if (timelineItems.length>=2) return { type:"timeline", timelineItems };
  }

  // Default: content
  return { type:"content" };
}

/* ═══════════════════════════════════════════════════════════════
   GENERATOR
   ═══════════════════════════════════════════════════════════════ */

const SLIDE_W = 13.33;
const SLIDE_H = 7.5;
const MARGIN = 0.85;

function generatePptx(parsed, themeKey) {
  const t = THEMES[themeKey] || THEMES["course-module"];
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name:"WIDE", width:SLIDE_W, height:SLIDE_H });
  pptx.layout = "WIDE";
  pptx.author = "Lesson Plan to PPTX";
  pptx.title = parsed.title || "教案";

  const total = parsed.slides.length + 1; // +1 for end slide
  let n = 0;
  const ctx = { pptx, t, total, slideNum:()=>++n };

  if (parsed.slides.length===0) {
    createTitleSlide(ctx, parsed.title, null, 0, parsed.slides);
  } else {
    parsed.slides.forEach((sd, idx) => {
      if (idx===0) {
        createTitleSlide(ctx, parsed.title, sd, idx, parsed.slides);
      } else {
        const st = detectSlideType(sd, idx, parsed.slides);
        ctx.st = st; ctx.sd = sd; ctx.slideIdx = idx; ctx.allSlides = parsed.slides;
        switch(st.type) {
          case "divider":       createDividerSlide(ctx, st.label||sd.title); break;
          case "quote":         createQuoteSlide(ctx, sd); break;
          case "table":         createTableSlide(ctx, sd, st.table); break;
          case "stats":         createStatsSlide(ctx, sd, st.stats); break;
          case "split":         createSplitSlide(ctx, sd, st); break;
          case "concept-box":   createConceptBoxSlide(ctx, sd, st.concepts); break;
          case "callout":       createCalloutSlide(ctx, sd, st.calloutText); break;
          case "exercise":      createExerciseSlide(ctx, sd, st); break;
          case "mcq":           createMCQSlide(ctx, sd, st); break;
          case "code":          createCodeSlide(ctx, sd, st); break;
          case "timeline":      createTimelineSlide(ctx, sd, st.timelineItems); break;
          default:              createContentSlide(ctx, sd);
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

function addTopAccentBar(slide, ctx) {
  const { t } = ctx;
  slide.addShape(ctx.pptx.shapes.RECTANGLE, { x:0, y:0, w:SLIDE_W, h:0.06, fill:{ color:t.palette.accent, transparency:30 } });
}

function addSlideTitle(slide, ctx, title) {
  const { t } = ctx;
  slide.addText(title, {
    x:MARGIN, y:0.45, w:SLIDE_W-2*MARGIN, h:0.7,
    fontSize:t.scale.title, fontFace:t.fonts.display, color:t.palette.text, bold:true,
  });
}

/* ═══════════ SLIDE TYPES ═══════════ */

/* ── COVER ── */
function createTitleSlide(ctx, presTitle, slideData, idx, allSlides) {
  const { pptx, t } = ctx;
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.accent };

  // Decorative oval
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
    if (meta["姓名"]||meta["教師"]) mp.push(meta["姓名"]||meta["教師"]);
    if (mp.length>0) {
      slide.addText(mp.join("  ·  "), {
        x:MARGIN, y:3.65, w:9.5, h:0.6,
        fontSize:t.scale.meta, fontFace:t.fonts.body,
        color:"FFFFFF", transparency:25,
      });
    }
  }

  // Theme name badge
  slide.addText(t.name, {
    x:SLIDE_W-2.2, y:0.3, w:1.6, h:0.35,
    fontSize:9, fontFace:"Calibri", color:"FFFFFF", transparency:40, align:"right",
  });

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
  addTopAccentBar(slide, ctx);

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
  addTopAccentBar(slide, ctx);
  addSlideTitle(slide, ctx, slideData.title);

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
  addTopAccentBar(slide, ctx);
  addSlideTitle(slide, ctx, slideData.title);

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
  addTopAccentBar(slide, ctx);

  const midX = SLIDE_W/2;
  const colW = midX - MARGIN - 0.3;
  const colTitleY = 0.55;
  const colBodyY = 1.35;

  slide.addText(st.leftTitle, {
    x:MARGIN, y:colTitleY, w:colW, h:0.55,
    fontSize:t.scale.title-2, fontFace:t.fonts.display, color:t.palette.accent, bold:true, align:"center",
  });
  slide.addText(st.rightTitle, {
    x:midX+0.3, y:colTitleY, w:colW, h:0.55,
    fontSize:t.scale.title-2, fontFace:t.fonts.display, color:t.palette.accent, bold:true, align:"center",
  });

  slide.addShape(pptx.shapes.RECTANGLE, {
    x:midX-0.015, y:colBodyY-0.2, w:0.03, h:5.5,
    fill:{ color:t.palette.muted, transparency:55 },
  });

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

/* ═══════════ NEW SLIDE TYPES ═══════════ */

/* ── CONCEPT BOX (概念卡) ── */
function createConceptBoxSlide(ctx, slideData, concepts) {
  const { pptx, t } = ctx;
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.bg };
  addTopAccentBar(slide, ctx);
  addSlideTitle(slide, ctx, slideData.title);

  const count = concepts.length;
  const cols = Math.min(count, 3);
  const rows = Math.ceil(count / cols);
  const cardW = (SLIDE_W - 2*MARGIN - (cols-1)*0.25) / cols;
  const cardH = 4.5 / rows;
  const startY = 1.5;

  concepts.forEach((c, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const cx = MARGIN + col*(cardW + 0.25);
    const cy = startY + row*(cardH + 0.25);

    // Card background
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x:cx, y:cy, w:cardW, h:cardH,
      fill:{ color:t.palette.surface },
      line:{ color:t.palette.accent, width:0.5 },
      rectRadius: t.radius/72.0 || 0.1,
      shadow: t.shadow ? { type:"outer", blur:8, offset:2, color:"000000", opacity:0.08 } : undefined,
    });

    // Concept name
    slide.addText(c.name, {
      x:cx+0.3, y:cy+0.2, w:cardW-0.6, h:0.5,
      fontSize:t.scale.body+2, fontFace:t.fonts.display, color:t.palette.accent, bold:true, valign:"middle",
    });

    // Description
    slide.addText(c.desc, {
      x:cx+0.3, y:cy+0.75, w:cardW-0.6, h:cardH-1.0,
      fontSize:t.scale.body-2, fontFace:t.fonts.body, color:t.palette.muted, valign:"top", lineSpacingMultiple:1.4,
    });
  });

  addWatermark(slide, ctx);
  addProgressBar(slide, ctx);
  addSlideNumber(slide, ctx);
}

/* ── CALLOUT (提示框) ── */
function createCalloutSlide(ctx, slideData, calloutText) {
  const { pptx, t } = ctx;
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.bg };
  addTopAccentBar(slide, ctx);
  addSlideTitle(slide, ctx, slideData.title);

  // Callout box with left accent border
  const calloutX = MARGIN + 0.5;
  const calloutW = SLIDE_W - 2*MARGIN - 1.0;
  const calloutY = 1.8;
  const calloutH = 4.0;
  const calloutPad = 0.4;

  // Background
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x:calloutX, y:calloutY, w:calloutW, h:calloutH,
    fill:{ color:t.palette.surface },
    line:{ color:t.palette.accent2 || t.palette.accent, width:1.5 },
    rectRadius: (t.radius+4)/72.0 || 0.12,
  });

  // Left accent bar (thick left border)
  slide.addShape(pptx.shapes.RECTANGLE, {
    x:calloutX+0.15, y:calloutY+0.2, w:0.07, h:calloutH-0.4,
    fill:{ color:t.palette.accent2 || t.palette.accent },
  });

  // "💡 提示" label
  slide.addText("💡 提示", {
    x:calloutX+calloutPad+0.2, y:calloutY+0.2, w:2.0, h:0.4,
    fontSize:t.scale.meta, fontFace:t.fonts.body, color:t.palette.accent2 || t.palette.accent, bold:true,
  });

  // Content
  slide.addText(calloutText, {
    x:calloutX+calloutPad+0.2, y:calloutY+0.7, w:calloutW-2*calloutPad-0.2, h:calloutH-1.0,
    fontSize:t.scale.body, fontFace:t.fonts.body, color:t.palette.text, lineSpacingMultiple:1.7, valign:"top",
  });

  addWatermark(slide, ctx);
  addProgressBar(slide, ctx);
  addSlideNumber(slide, ctx);
}

/* ── EXERCISE (練習區) ── */
function createExerciseSlide(ctx, slideData, st) {
  const { pptx, t } = ctx;
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.bg };
  addTopAccentBar(slide, ctx);
  addSlideTitle(slide, ctx, slideData.title);

  const boxX = MARGIN + 0.3;
  const boxW = SLIDE_W - 2*MARGIN - 0.6;
  const boxY = 1.6;
  const boxH = 5.0;

  // Dashed border box (PPTX doesn't support dashed natively — use thin line)
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x:boxX, y:boxY, w:boxW, h:boxH,
    fill:{ color:t.palette.surface },
    line:{ color:t.palette.accent2 || t.palette.accent, width:1.5 },
    rectRadius: (t.radius+6)/72.0 || 0.12,
  });

  // "✏️ 練習" label
  slide.addText("✏️ 練習", {
    x:boxX+0.35, y:boxY+0.2, w:3.0, h:0.4,
    fontSize:t.scale.meta+1, fontFace:t.fonts.body, color:t.palette.accent2 || t.palette.accent, bold:true,
  });

  // Exercise content as list
  const exerciseItems = detectListItems(st.exerciseBody);
  if (exerciseItems.length>0) {
    renderListInBox(slide, exerciseItems, t, boxX+0.3, boxY+0.75, boxW-0.6, boxH-1.0, pptx, t.scale.body);
  } else if (st.exerciseBody) {
    slide.addText(st.exerciseBody, {
      x:boxX+0.5, y:boxY+0.75, w:boxW-1.0, h:boxH-1.2,
      fontSize:t.scale.body, fontFace:t.fonts.body, color:t.palette.text, lineSpacingMultiple:1.6, valign:"top",
    });
  }

  addWatermark(slide, ctx);
  addProgressBar(slide, ctx);
  addSlideNumber(slide, ctx);
}

/* ── MCQ (選擇題) ── */
function createMCQSlide(ctx, slideData, st) {
  const { pptx, t } = ctx;
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.bg };
  addTopAccentBar(slide, ctx);
  addSlideTitle(slide, ctx, slideData.title);

  const startY = 1.5;
  const optionH = 1.2;
  const optX = MARGIN + 0.5;
  const optW = SLIDE_W - 2*MARGIN - 1.0;

  st.options.forEach((opt, i) => {
    const y = startY + i*optionH;
    const isCorrect = opt.correct;
    const accentColor = isCorrect ? t.palette.good : t.palette.accent;

    // Letter circle
    slide.addShape(pptx.shapes.OVAL, {
      x:optX, y:y+0.15, w:0.55, h:0.55,
      fill:{ color: isCorrect ? t.palette.good : t.palette.bg },
      line:{ color:accentColor, width:1.5 },
    });
    slide.addText(opt.letter, {
      x:optX, y:y+0.15, w:0.55, h:0.55,
      fontSize:t.scale.body-1, fontFace:t.fonts.body, color: isCorrect ? "FFFFFF" : t.palette.accent,
      bold:true, align:"center", valign:"middle",
    });

    // Option text
    slide.addText(opt.text, {
      x:optX+0.8, y:y, w:optW-0.8, h:optionH,
      fontSize:t.scale.body, fontFace:t.fonts.body, color:t.palette.text, valign:"middle",
    });
  });

  addWatermark(slide, ctx);
  addProgressBar(slide, ctx);
  addSlideNumber(slide, ctx);
}

/* ── CODE (程式碼) ── */
function createCodeSlide(ctx, slideData, st) {
  const { pptx, t } = ctx;
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.bg };
  addTopAccentBar(slide, ctx);
  addSlideTitle(slide, ctx, slideData.title);

  // Dark code block background
  const codeX = MARGIN + 0.3;
  const codeW = SLIDE_W - 2*MARGIN - 0.6;
  const codeY = 1.6;
  const codeH = 5.0;

  // Determine if background should be dark
  const isDarkBg = parseInt(t.palette.bg, 16) < 0x444444;
  const codeBg = isDarkBg ? "1A1A2E" : "1E1E2E";

  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x:codeX, y:codeY, w:codeW, h:codeH,
    fill:{ color:codeBg },
    rectRadius: 0.08,
  });

  // Language badge
  if (st.lang) {
    slide.addText(st.lang, {
      x:codeX+0.2, y:codeY+0.1, w:1.5, h:0.3,
      fontSize:9, fontFace:"Consolas", color:t.palette.muted, align:"left",
    });
  }

  // Code content
  const codeLines = st.code.split("\n");
  const maxLines = Math.min(codeLines.length, 14);
  const lineH = 0.34;
  codeLines.slice(0, maxLines).forEach((line, i) => {
    slide.addText(line || " ", {
      x:codeX+0.3, y:codeY+0.45 + i*lineH, w:codeW-0.6, h:lineH,
      fontSize:t.scale.small+1, fontFace:t.fonts.mono, color:"D4D4D8", valign:"middle",
    });
  });
  if (codeLines.length > maxLines) {
    slide.addText(`... (${codeLines.length - maxLines} more lines)`, {
      x:codeX+0.3, y:codeY+0.45 + maxLines*lineH, w:codeW-0.6, h:0.3,
      fontSize:9, fontFace:t.fonts.mono, color:t.palette.muted, italic:true,
    });
  }

  addWatermark(slide, ctx);
  addProgressBar(slide, ctx);
  addSlideNumber(slide, ctx);
}

/* ── TIMELINE (時間線) ── */
function createTimelineSlide(ctx, slideData, items) {
  const { pptx, t } = ctx;
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.bg };
  addTopAccentBar(slide, ctx);
  addSlideTitle(slide, ctx, slideData.title);

  const count = Math.min(items.length, 6);
  const colW = (SLIDE_W - 2*MARGIN) / count;
  const tlY = 2.0;

  // Timeline line
  slide.addShape(pptx.shapes.RECTANGLE, {
    x:MARGIN+0.3, y:tlY+2.0, w:SLIDE_W-2*MARGIN-0.6, h:0.03,
    fill:{ color:t.palette.accent, transparency:50 },
  });

  items.slice(0, count).forEach((item, i) => {
    const cx = MARGIN + i*colW + colW/2;

    // Dot
    slide.addShape(pptx.shapes.OVAL, {
      x:cx-0.15, y:tlY+1.7, w:0.3, h:0.3,
      fill:{ color:t.palette.accent },
    });

    // Text
    slide.addText(item.text, {
      x:cx-colW/2+0.2, y:tlY+2.3, w:colW-0.4, h:2.0,
      fontSize:t.scale.body-2, fontFace:t.fonts.body, color:t.palette.text, align:"center", valign:"top",
    });
  });

  addWatermark(slide, ctx);
  addProgressBar(slide, ctx);
  addSlideNumber(slide, ctx);
}

/* ── CONTENT (with half-bleed panel for short prose + objectives sidebar) ── */
function createContentSlide(ctx, slideData) {
  const { pptx, t, slideIdx, allSlides } = ctx;
  const slide = pptx.addSlide();
  slide.background = { fill: t.palette.bg };
  addTopAccentBar(slide, ctx);

  const items = detectListItems(slideData.content);
  const isShortProse = !items.length && slideData.content && slideData.content.length < 180;

  // ── Objectives sidebar (slides 2-6 only, if first content slide has objectives) ──
  const showSidebar = slideIdx >= 1 && slideIdx <= 5;
  const firstContentObj = extractObjectives(allSlides);
  const hasObjectives = showSidebar && firstContentObj.length > 0;
  const sidebarW = hasObjectives ? 2.8 : 0;
  const mainX = hasObjectives ? MARGIN + sidebarW + 0.2 : MARGIN;
  const mainW = SLIDE_W - mainX - MARGIN;

  if (hasObjectives) {
    // Sidebar background
    slide.addShape(pptx.shapes.RECTANGLE, {
      x:MARGIN-0.1, y:0.06, w:sidebarW+0.2, h:SLIDE_H-0.06,
      fill:{ color:t.palette.surface },
    });

    // "學習目標" header
    slide.addText("學習目標", {
      x:MARGIN+0.2, y:0.25, w:sidebarW-0.4, h:0.35,
      fontSize:9, fontFace:t.fonts.body, color:t.palette.muted, bold:true,
    });

    // Objectives list with progress markers
    firstContentObj.forEach((obj, i) => {
      const isDone = i < slideIdx - 1; // objectives before current slide = done
      const isCurrent = i === slideIdx - 1; // this slide's objective = current
      const marker = isDone ? "●" : isCurrent ? "▸" : "○";
      const color = isDone ? t.palette.muted : isCurrent ? (t.palette.accent2||t.palette.accent) : t.palette.muted;
      const weight = isCurrent ? "bold" : "normal";

      const y = 0.8 + i*0.55;
      slide.addText(marker, {
        x:MARGIN+0.2, y:y, w:0.4, h:0.4,
        fontSize:12, fontFace:t.fonts.body, color:color, valign:"middle",
      });
      slide.addText(obj, {
        x:MARGIN+0.6, y:y, w:sidebarW-1.0, h:0.4,
        fontSize:10, fontFace:t.fonts.body, color:t.palette.text, valign:"middle",
      });
      if (isCurrent) {
        slide.addShape(pptx.shapes.RECTANGLE, {
          x:MARGIN-0.1, y:y-0.05, w:0.06, h:0.5,
          fill:{ color: t.palette.accent2 || t.palette.accent },
        });
      }
    });
  }

  if (isShortProse) {
    // Half-bleed accent panel
    const panelW = 2.0;
    slide.addShape(pptx.shapes.RECTANGLE, {
      x:0, y:0.06, w:panelW, h:SLIDE_H-0.06,
      fill:{ color:t.palette.accent, transparency:8 },
    });
    slide.addText(slideData.title, {
      x:mainX, y:0.45, w:mainW, h:0.7,
      fontSize:t.scale.title, fontFace:t.fonts.display, color:t.palette.text, bold:true,
    });
    slide.addText(slideData.content, {
      x:mainX, y:1.8, w:mainW, h:4.8,
      fontSize:t.scale.body+1, fontFace:t.fonts.body, color:t.palette.text, lineSpacingMultiple:1.8, valign:"top",
    });
  } else {
    slide.addText(slideData.title, {
      x:mainX, y:0.45, w:mainW, h:0.7,
      fontSize:t.scale.title, fontFace:t.fonts.display, color:t.palette.text, bold:true,
    });
    if (items.length>0) {
      renderListFullWidth(slide, items, t, pptx, mainX, mainW);
    } else if (slideData.content) {
      slide.addText(slideData.content, {
        x:mainX, y:1.6, w:mainW, h:5.2,
        fontSize:t.scale.body, fontFace:t.fonts.body, color:t.palette.text, lineSpacingMultiple:1.75, valign:"top",
      });
    }
  }

  // Speaker notes (invisible on slide, stored in PPTX notes)
  const notes = parseSpeakerNotes(slideData.content);
  if (notes) {
    try {
      const ns = slide.notesSlide;
      if (ns) ns.notes_text_frame.text = notes;
    } catch(e) { /* PPTXGenJS may not support notesSlide; silently ignore */ }
  }

  addWatermark(slide, ctx);
  addProgressBar(slide, ctx);
  addSlideNumber(slide, ctx);
}

/* ── Extract objectives from the 2nd slide (1-indexed) ── */
function extractObjectives(allSlides) {
  if (!allSlides || allSlides.length < 2) return [];
  const objSlide = allSlides[1]; // second slide (index 1)
  if (!objSlide || !objSlide.content) return [];
  const items = detectListItems(objSlide.content);
  if (items.length >= 2 && items.length <= 8) return items.map(i => i.text);
  return [];
}

/* ── LIST RENDERERS ── */

function renderListFullWidth(slide, items, t, pptx, adjX, adjW) {
  const startX = adjX || MARGIN;
  const maxW = adjW || (SLIDE_W - 2*MARGIN);
  const startY = 1.55, itemH = 0.95, maxItems = Math.min(items.length, 6), circleSize = 0.48;
  items.slice(0, maxItems).forEach((item, idx) => {
    const y = startY + idx*itemH;
    if (item.type==="number") {
      slide.addShape(pptx.shapes.OVAL, { x:startX, y:y+0.10, w:circleSize, h:circleSize, fill:{ color:t.palette.accent } });
      slide.addText(String(idx+1), { x:startX, y:y+0.10, w:circleSize, h:circleSize, fontSize:15, fontFace:"Calibri", color:"FFFFFF", bold:true, align:"center", valign:"middle" });
    } else {
      slide.addShape(pptx.shapes.OVAL, { x:startX+0.16, y:y+0.28, w:0.18, h:0.18, fill:{ color:t.palette.accent } });
    }
    slide.addText(item.text, { x:startX+0.72, y:y, w:maxW-0.72, h:itemH, fontSize:t.scale.body, fontFace:t.fonts.body, color:t.palette.text, valign:"middle" });
  });
  if (items.length>maxItems) {
    slide.addText(`還有 ${items.length-maxItems} 項…`, { x:startX+0.72, y:startY+maxItems*itemH, w:5, h:0.5, fontSize:t.scale.meta, fontFace:t.fonts.body, color:t.palette.muted, italic:true });
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

function renderListInBox(slide, items, t, boxX, boxY, boxW, boxH, pptx, fontSize) {
  const itemH = 0.55;
  const maxItems = Math.min(items.length, 6);
  items.slice(0, maxItems).forEach((item, idx) => {
    const y = boxY + idx*itemH;
    if (item.type==="number") {
      slide.addText(`${idx+1}.`, { x:boxX, y:y, w:0.4, h:itemH, fontSize:fontSize, fontFace:t.fonts.body, color:t.palette.accent, bold:true, valign:"middle" });
    } else {
      slide.addText("•", { x:boxX, y:y, w:0.4, h:itemH, fontSize:fontSize, fontFace:t.fonts.body, color:t.palette.accent, valign:"middle" });
    }
    slide.addText(item.text, { x:boxX+0.45, y:y, w:boxW-0.45, h:itemH, fontSize:fontSize, fontFace:t.fonts.body, color:t.palette.text, valign:"middle" });
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
    const { lessonPlan, theme="course-module" } = body || {};
    if (!lessonPlan || lessonPlan.trim().length<10) return res.status(400).json({ error:"請貼上教案內容（至少 10 個字）" });
    if (!THEMES[theme]) return res.status(400).json({ error:`無效的主題：${theme}。可用主題：${Object.keys(THEMES).join(", ")}` });

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
