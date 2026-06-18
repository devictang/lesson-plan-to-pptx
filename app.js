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
const btnCopyPrompt = document.getElementById("btnCopyPrompt");
const btnCopyInstruction = document.getElementById("btnCopyInstruction");

/* ── AI Format Appendix (append to YOUR instruction) ── */
const AI_PROMPT = `**輸出格式要求（必須嚴格遵守）：**

1. 第一行使用 \`# 主題名稱\`
2. 每個 \`## \` 開頭代表一張新 Slide
3. 第一個 \`## \` 固定為封面 Slide

**13 種 Slide 語法：**

| 類型 | 語法 | 效果 |
|------|------|------|
| 封面 | 第一個 \`## 封面 — 主題\` + \`- **key**：value\` | 深色全幅標題 |
| 分隔頁 | \`## ---\` + section 名稱 | 全色背景置中大字 |
| 引用頁 | \`> 引文\` + \`> — 署名\` | 大引號 + 斜體金句 |
| 統計頁 | \`- **標籤**：數值\` (1–4 組) | 超大數字 + label |
| 表格頁 | \`| A | B |\` + \`|---|---|\` | 結構化表格 |
| 對照頁 | \`## 左標題 || 右標題\` + 內容用 \`---\` 分隔 | 雙欄對比 |
| **概念卡** | \`> **概念名**：說明文字\` (2-3 組) | 多個 rounded card 並排 |
| **提示框** | \`> 💡 提示文字\` | left accent border 嘅 callout |
| **練習區** | \`> ✏️ 練習標題\` + 內容 | 邊框 box + 「練習」label |
| **選擇題** | \`> ❓ 題目\` + \`A. 選項\` + \`✅ 答案\` | circle letter + 綠色標記 |
| **程式碼** | \` \`\`\`語言\` + code + \` \`\`\` \` | 暗色 code block |
| **時間線** | 標題有「時間線/roadmap」+ 列表 | 橫向 timeline |
| **逐字稿** | \`> notes：\` + 口語提示稿 | 存入 PPTX speaker notes |
| 內容頁 | \`1. \` / \`- \` 列表 或 段落文字 | 一般內容 + objectives sidebar |

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

## ---
核心教學

## 光合作用方程式
| 原料 | 條件 | 產物 |
|------|------|------|
| CO₂ + H₂O | 光 + 葉綠體 | C₆H₁₂O₆ + O₂

## 核心概念
> **概念一**：說明文字
> **概念二**：說明文字
> **概念三**：說明文字
> notes：呢度用概念卡解釋三個要素。

## 💡 提示
> 💡 重要提示內容

## ✏️ 練習
> ✏️ 練習名稱
> 1. 步驟一
> 2. 步驟二

## ❓ 測驗
> ❓ 問題
> A. 選項
> B. 選項
> ✅ B

## 導入活動 || 總結活動
- 左欄內容
---
- 右欄內容

## 時間線
1. 事件一
2. 事件二

## ---
數據

## 重點
- **課時**：40 分鐘
- **實驗次數**：3
\`\`\`

**注意：**
- 分隔頁用 \`## ---\` + section 名切換章節
- 統計頁只放 key-value，唔好加其他文字
- 表格必須有 \`|---|---|\` separator row
- 對照頁用 \`||\` 分標題，內容用 \`---\` 分隔左右欄
- **概念卡**：連續 2-3 行 \`> **概念**：說明\` 自動變卡片
- **提示框**：用 \`> 💡\` 開頭
- **練習區**：用 \`> ✏️\` 開頭
- **選擇題**：用 \`> ❓\` 開頭，\`✅\` 標記正確答案
- **程式碼**：用 \` \`\`\` \` 包圍
- **逐字稿**：用 \`> notes：\` 開頭，每頁最後加`;

/* ── Example ── */
const EXAMPLE = `# 教學主題：光合作用

## 封面 — 光合作用
- **科目**：生物科
- **年級**：中三
- **時間**：40 分鐘

## 📋 教學目標
1. 說明光合作用的定義及化學方程式
2. 辨識光合作用所需的條件與產物
3. 理解光合作用對生態系統的重要性

## ---
認識光合作用

## 光合作用方程式
| 原料 | 條件 | 產物 |
|------|------|------|
| 二氧化碳 + 水 | 光 + 葉綠體 | 葡萄糖 + 氧氣 |
| 6CO₂ + 6H₂O | 光合作用 | C₆H₁₂O₆ + 6O₂

## 光合作用嘅核心概念
> **原料**：二氧化碳 + 水，由葉片氣孔吸收
> **條件**：陽光 + 葉綠體，缺一不可
> **產物**：葡萄糖 + 氧氣，為生態系提供能量
> notes：呢度用三張概念卡解釋光合作用嘅三個要素。

## 💡 常見誤解
> 💡 好多學生以為光合作用係響夜晚進行。實際上，光合作用必須有光！夜晚植物只進行呼吸作用。

## ✏️ 小組練習
> ✏️ 平衡方程式練習
> 1. 寫出光合作用嘅完整化學方程式
> 2. 標明每個物質嘅狀態（固/液/氣）
> 3. 解釋如果 CO₂ 濃度增加一倍會點樣

## ❓ 課堂測驗
> ❓ 以下邊個唔係光合作用所需嘅條件？
> A. 陽光
> B. 水
> C. 氧氣
> D. 葉綠體
> ✅ C

## ---
課堂數據

## 📈 本課統計
- **課時**：40 分鐘
- **實驗次數**：3 次
- **小組數量**：4 組`;

/* ── Format detection ── */
function checkFormat(text) {
  if (text.trim().length < 10) return { valid: false, hint: "" };
  const hasTitle = /^#\s+/m.test(text);
  const slides = text.match(/^##\s+/gm) || [];
  const slideCount = slides.length;
  const features = [];
  if (/^##\s+---/m.test(text)) features.push("分隔頁");
  if (/\|.+\|/.test(text)) features.push("表格");
  if (/^>\s/m.test(text)) features.push("引用");
  if (/\|\|/.test(text.match(/^##\s+(.+)/m)?.[1]||"")) features.push("對照頁");
  if (/^>\s*\*\*[^*]+\*\*[：:]/.test(text)) features.push("概念卡");
  if (/^>\s*💡/m.test(text)) features.push("提示");
  if (/^>\s*✏️/m.test(text)) features.push("練習");
  if (/^>\s*❓/m.test(text)) features.push("測驗");
  if (/\x60{3}/.test(text)) features.push("程式碼");
  if (hasTitle && slideCount >= 2) {
    const extra = features.length ? ` · ${features.join("、")}` : "";
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
  return c ? c.value : "course-module";
}

/* ── Generate ── */
async function generatePptx() {
  const text = input.value.trim();
  const theme = getSelectedTheme();
  if (text.length < 10) { showStatus("請先貼上教案內容", "error"); return; }
  btnGenerate.classList.add("loading"); btnGenerate.disabled = true;
  statusEl.hidden = true; downloadArea.hidden = true;
  try {
    const resp = await fetch("/api/generate", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ lessonPlan: text, theme }),
    });
    if (!resp.ok) {
      const err = await resp.json().catch(() => ({}));
      throw new Error(err.error || `伺服器錯誤 (${resp.status})`);
    }
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const fm = (resp.headers.get("Content-Disposition")||"").match(/filename\*=UTF-8''(.+)/);
    const filename = fm ? decodeURIComponent(fm[1]) : "lesson-plan.pptx";
    downloadLink.href = url; downloadLink.download = filename;
    downloadFilename.textContent = filename;
    downloadArea.hidden = false;
    const a = document.createElement("a"); a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    showStatus("✅ 簡報已成功生成並開始下載！", "success");
  } catch (err) {
    console.error(err);
    showStatus(`❌ ${err.message || "生成失敗"}`, "error");
  } finally {
    btnGenerate.classList.remove("loading"); btnGenerate.disabled = false;
  }
}

function showStatus(msg, type) {
  statusEl.textContent = msg; statusEl.className = `status ${type}`; statusEl.hidden = false;
}

/* ── Copy ── */
async function copyToClipboard(text) {
  try { await navigator.clipboard.writeText(text); return true; }
  catch {
    const ta = document.createElement("textarea"); ta.value = text;
    ta.style.cssText = "position:fixed;opacity:0";
    document.body.appendChild(ta); ta.select();
    const ok = document.execCommand("copy"); document.body.removeChild(ta); return ok;
  }
}

/* ── Events ── */
input.addEventListener("input", updateUI);
btnGenerate.addEventListener("click", generatePptx);
btnExample.addEventListener("click", loadExample);
btnClear.addEventListener("click", clearInput);

btnCopyInstruction.addEventListener("click", async () => {
  const text = "幫我設計一個中三生物科關於光合作用嘅教案，40 分鐘課時，包含導入、核心教學、小組討論同總結。";
  const ok = await copyToClipboard(text);
  btnCopyInstruction.textContent = ok ? "✓ 已複製" : "複製失敗";
  setTimeout(() => { btnCopyInstruction.textContent = "📋 複製"; }, 2000);
});

const aiPromptEl = document.getElementById("aiPromptText");
const formatRefEl = document.getElementById("formatExampleRef");
aiPromptEl.textContent = AI_PROMPT;
formatRefEl.textContent = EXAMPLE;

btnCopyPrompt.addEventListener("click", async () => {
  const ok = await copyToClipboard(AI_PROMPT);
  btnCopyPrompt.textContent = ok ? "✓ 已複製" : "複製失敗";
  setTimeout(() => { btnCopyPrompt.textContent = "📋 複製格式要求"; }, 2000);
});
document.addEventListener("keydown", (e) => {
  if ((e.ctrlKey||e.metaKey) && e.key==="Enter") {
    e.preventDefault(); if (!btnGenerate.disabled) generatePptx();
  }
});
updateUI();
