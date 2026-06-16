/**
 * Lesson Plan → PPTX — Frontend Application
 */

const input = document.getElementById("input");
const charCount = document.getElementById("charCount");
const formatHint = document.getElementById("formatHint");
const themeGrid = document.getElementById("themeGrid");
const btnGenerate = document.getElementById("btnGenerate");
const statusEl = document.getElementById("status");
const downloadArea = document.getElementById("downloadArea");
const downloadLink = document.getElementById("downloadLink");
const downloadFilename = document.getElementById("downloadFilename");
const btnExample = document.getElementById("btnExample");
const btnClear = document.getElementById("btnClear");
const btnCopyExample = document.getElementById("btnCopyExample");
const btnCopyPrompt = document.getElementById("btnCopyPrompt");

/* ── AI Prompt ── */
const AI_PROMPT = `你是一個專業的教案簡報生成助手。請根據用戶提供的主題，生成一份可以直接轉換為 PowerPoint 的結構化 Markdown。

**格式規則：**

1. 第一行使用 \`# 主題名稱\` 作為整份簡報的標題
2. 每個 \`## \` 開頭的行代表一張新 Slide
3. 第一個 \`## \` 固定為封面 Slide

**5 種 Slide 類型：**

| 類型 | 寫法 | 效果 |
|------|------|------|
| 封面 | \`## 封面 — 主題\` + \`- **key**：value\` | 深色全幅標題頁 |
| 分隔頁 | \`## ---\` + 下一行寫 section 名稱 | 全色背景、置中大字的章節分隔 |
| 引用頁 | \`## 標題\` + \`> 引文\` + \`> — 出處\` | 大引號 + 金句 + attribution |
| 統計頁 | \`## 標題\` + \`- **標籤**：數值\` (1–4 組) | 超大數字 + label（1 組置中 / 多組並排） |
| 表格頁 | \`## 標題\` + \`| 欄A | 欄B |\` + \`|---|---|\` | 結構化表格 |
| 內容頁 | \`## 標題\` + 段落文字 / 列表 | 一般內容（支援 \`1. \` 編號 + \`- \` 項目） |

**完整範例：**

\`\`\`
# 教學主題：[主題名稱]

## 封面 — [主題名稱]
- **科目**：[科目]
- **年級**：[年級]
- **時間**：[分鐘] 分鐘

## 教學目標
1. [目標一]
2. [目標二]
3. [目標三]

## ---
核心教學

## [Slide 標題]
[段落文字…]

## 光合作用方程式
| 原料 | 條件 | 產物 |
|------|------|------|
| CO₂ + H₂O | 光 + 葉綠體 | C₆H₁₂O₆ + O₂ |

## 光合作用的重要性
> 植物係地球上最重要嘅生產者。
> — 生物學教科書

## ---
總結

## 本課重點
- **關鍵概念**：光合作用 = 光能 → 化學能
- **課時**：40 分鐘
- **實驗次數**：3
\`\`\`

**注意事項：**
- 分隔頁用 \`## ---\` + section 名稱做章節切換
- 統計頁只放 key-value pairs，唔好加其他文字
- 表格必須用 pipe \`|\` 格式，要有 header separator \`|---|---|\`
- 引用用 \`> \` 開頭，\`> — Name\` 係署名`;

/* ── Example ── */
const EXAMPLE = `# 教學主題：光合作用

## 封面 — 光合作用
- **科目**：生物科
- **年級**：中三
- **時間**：40 分鐘

## 教學目標
1. 說明光合作用的定義及化學方程式
2. 辨識光合作用所需的條件與產物
3. 理解光合作用對生態系統的重要性

## ---
課堂導入

## 課堂導入（5 分鐘）
播放一段植物生長的快鏡影片，提問：「植物為什麼不需要吃東西也能長大？」引導學生思考植物如何獲取能量。

## ---
核心教學

## 光合作用方程式
| 原料 | 條件 | 產物 |
|------|------|------|
| 二氧化碳 + 水 | 光 + 葉綠體 | 葡萄糖 + 氧氣 |
| 6CO₂ + 6H₂O | 光合作用 | C₆H₁₂O₆ + 6O₂ |

## 光合作用的重要性
> 植物係地球上最重要嘅生產者，佢哋將太陽能轉化為化學能，支撐住整個生態系統。
> — 生物學教科書

## ---
總結

## 本課重點
- **關鍵概念**：光合作用 = 光能 → 化學能
- **必需條件**：光、葉綠體、CO₂、H₂O
- **產物**：葡萄糖 + 氧氣`;

/* ── Format detection ── */
function checkFormat(text) {
  if (text.trim().length < 10) return { valid: false, hint: "" };
  const hasTitle = /^#\s+/m.test(text);
  const slides = text.match(/^##\s+/gm) || [];
  const slideCount = slides.length;
  const hasDivider = /^##\s+---/m.test(text);
  const hasTable = /\|.+\|/.test(text);
  const hasQuote = /^>\s/m.test(text);
  const features = [];
  if (hasDivider) features.push("分隔頁");
  if (hasTable) features.push("表格");
  if (hasQuote) features.push("引用");

  if (hasTitle && slideCount >= 2) {
    const extra = features.length ? ` · 偵測到 ${features.join("、")}` : "";
    return { valid: true, hint: `✅ 格式正確 — ${slideCount} 張 Slide${extra}` };
  }
  if (slideCount >= 1) return { valid: false, hint: "⚠️ 缺少主題標題（# 主題名稱）" };
  return { valid: false, hint: "💡 使用 `## Slide 標題` 分隔每張 Slide" };
}

/* ── Update UI ── */
function updateUI() {
  const text = input.value;
  charCount.textContent = `${text.length} 字`;
  const fmt = checkFormat(text);
  formatHint.textContent = fmt.hint;
  formatHint.className = "format-hint " + (fmt.valid ? "valid" : "invalid");
  btnGenerate.disabled = text.trim().length < 10;
}

/* ── Load / Clear ── */
function loadExample() { input.value = EXAMPLE; updateUI(); input.focus(); }
function clearInput() { input.value = ""; updateUI(); downloadArea.hidden = true; statusEl.hidden = true; input.focus(); }

/* ── Theme ── */
function getSelectedTheme() {
  const c = themeGrid.querySelector('input[name="theme"]:checked');
  return c ? c.value : "ocean";
}

/* ── Generate ── */
async function generatePptx() {
  const text = input.value.trim();
  const theme = getSelectedTheme();
  if (text.length < 10) { showStatus("請先貼上教案內容", "error"); return; }

  btnGenerate.classList.add("loading");
  btnGenerate.disabled = true;
  statusEl.hidden = true;
  downloadArea.hidden = true;

  try {
    const resp = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ lessonPlan: text, theme }),
    });
    if (!resp.ok) {
      const err = await resp.json().catch(() => ({}));
      throw new Error(err.error || `伺服器錯誤 (${resp.status})`);
    }
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const disp = resp.headers.get("Content-Disposition") || "";
    const fm = disp.match(/filename\*=UTF-8''(.+)/);
    const filename = fm ? decodeURIComponent(fm[1]) : "lesson-plan.pptx";

    downloadLink.href = url;
    downloadLink.download = filename;
    downloadFilename.textContent = filename;
    downloadArea.hidden = false;

    const a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);

    showStatus("✅ 簡報已成功生成並開始下載！", "success");
  } catch (err) {
    console.error(err);
    showStatus(`❌ ${err.message || "生成失敗"}`, "error");
  } finally {
    btnGenerate.classList.remove("loading");
    btnGenerate.disabled = false;
  }
}

function showStatus(msg, type) {
  statusEl.textContent = msg;
  statusEl.className = `status ${type}`;
  statusEl.hidden = false;
}

/* ── Copy ── */
async function copyToClipboard(text) {
  try { await navigator.clipboard.writeText(text); return true; }
  catch {
    const ta = document.createElement("textarea"); ta.value = text;
    ta.style.cssText = "position:fixed;opacity:0";
    document.body.appendChild(ta); ta.select();
    const ok = document.execCommand("copy");
    document.body.removeChild(ta); return ok;
  }
}

/* ── Events ── */
input.addEventListener("input", updateUI);
btnGenerate.addEventListener("click", generatePptx);
btnExample.addEventListener("click", loadExample);
btnClear.addEventListener("click", clearInput);

btnCopyExample.addEventListener("click", async () => {
  const t = document.getElementById("formatExample").textContent;
  const ok = await copyToClipboard(t);
  btnCopyExample.textContent = ok ? "✓ 已複製" : "複製失敗";
  setTimeout(() => { btnCopyExample.textContent = "複製"; }, 2000);
});

const aiPromptEl = document.getElementById("aiPromptText");
aiPromptEl.textContent = AI_PROMPT;

btnCopyPrompt.addEventListener("click", async () => {
  const ok = await copyToClipboard(AI_PROMPT);
  btnCopyPrompt.textContent = ok ? "✓ 已複製" : "複製失敗";
  setTimeout(() => { btnCopyPrompt.textContent = "📋 複製指令"; }, 2000);
});

document.addEventListener("keydown", (e) => {
  if ((e.ctrlKey||e.metaKey) && e.key==="Enter") {
    e.preventDefault();
    if (!btnGenerate.disabled) generatePptx();
  }
});

updateUI();
