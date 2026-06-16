/**
 * Lesson Plan → PPTX — Frontend Application
 * Handles input, validation, API calls, and download
 */

// ── DOM refs ──
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

// ── AI Prompt (for slide-by-slide format) ──
const AI_PROMPT = `你是一個專業的教案簡報生成助手。請根據用戶提供的主題，生成一份可以直接轉換為 PowerPoint 的結構化 Markdown。

**格式規則（必須嚴格遵守）：**

1. 第一行使用 \`# 主題名稱\` 作為整份簡報的標題
2. 每個 \`## \` 開頭的行代表一張新 Slide
3. 第一個 \`## \` 是封面 Slide，可在內容中使用 \`- **key**：value\` 格式加入基本資訊（科目、年級、時間等）
4. 後續每個 \`## \` 都是一張內容 Slide，標題後面可以寫段落文字
5. 如需顯示列表，使用 \`1. \` （編號列表）或 \`- \` （項目符號列表），系統會自動加上精美設計

**格式範例：**

\`\`\`
# 教學主題：[主題名稱]

## 封面 — [主題名稱]
- **科目**：[科目]
- **年級**：[年級]
- **時間**：[分鐘] 分鐘

## [Slide 標題]
[內容段落…]

## [Slide 標題]
1. [列表項目一]
2. [列表項目二]
3. [列表項目三]

## [Slide 標題]
- [項目符號一]
- [項目符號二]
\`\`\`

**注意事項：**
- 每個 \`## \` 就是一張 Slide，請按你想要的簡報順序排列
- 不要使用 \`### \` 或其他層級的標題
- 列表（\`1. \` 或 \`- \`）會被自動美化為帶圖標的設計
- 非列表的內容會以段落文字形式顯示
- 輸出的簡報會自動在最後加上「謝謝」結尾頁`;

// ── Example lesson plan ──
const EXAMPLE = `# 教學主題：唐朝盛世

## 封面 — 唐朝盛世
- **科目**：中國歷史
- **年級**：中二
- **時間**：40 分鐘

## 教學目標
1. 了解唐朝建立的歷史背景
2. 分析貞觀之治的特點與影響
3. 培養學生對歷史的批判思考能力

## 教學資源
- 教科書第三章
- 唐朝疆域圖
- 多媒體簡報
- 工作紙

## 課堂導入（5 分鐘）
教師提問：「你心中最強盛的朝代是哪一個？」
引導學生思考並分享，帶出唐朝的主題。

## 唐朝建立背景（10 分鐘）
教師講解隋末農民起義、李淵起兵、
以及唐朝建立的歷史過程。

## 小組討論 — 貞觀之治（10 分鐘）
學生分 4 組，討論以下政策：
- 均田制
- 科舉制
- 三省六部制
- 對外政策

## 總結與評估（10 分鐘）
- 各組匯報討論成果
- 教師總結本課重點
- 完成工作紙`;

// ── Format detection ──
function checkFormat(text) {
  if (text.trim().length < 10) return { valid: false, hint: "" };

  const hasTitle = /^#\s+/m.test(text);
  const hasSlides = /^##\s+/m.test(text);
  // Count slide count
  const slideCount = (text.match(/^##\s+/gm) || []).length;

  if (hasTitle && hasSlides && slideCount >= 2) {
    return { valid: true, hint: `✅ 格式正確 — 偵測到 ${slideCount} 張 Slide` };
  }
  if (hasSlides && slideCount >= 1) {
    return { valid: false, hint: "⚠️ 缺少主題標題（# 主題名稱）" };
  }
  if (hasTitle && hasSlides) {
    return { valid: false, hint: "⚠️ 需要至少 2 張 Slide（封面 + 內容）" };
  }
  return { valid: false, hint: "💡 使用 `## Slide 標題` 分隔每張 Slide" };
}

// ── Update UI on input ──
function updateUI() {
  const text = input.value;
  const len = text.length;
  charCount.textContent = `${len} 字`;

  const fmt = checkFormat(text);
  formatHint.textContent = fmt.hint;
  formatHint.className = "format-hint " + (fmt.valid ? "valid" : "invalid");

  // Enable/disable generate button
  btnGenerate.disabled = len < 10;
}

// ── Load example ──
function loadExample() {
  input.value = EXAMPLE;
  updateUI();
  input.focus();
}

// ── Clear ──
function clearInput() {
  input.value = "";
  updateUI();
  downloadArea.hidden = true;
  statusEl.hidden = true;
  input.focus();
}

// ── Get selected theme ──
function getSelectedTheme() {
  const checked = themeGrid.querySelector('input[name="theme"]:checked');
  return checked ? checked.value : "ocean";
}

// ── Generate PPTX ──
async function generatePptx() {
  const text = input.value.trim();
  const theme = getSelectedTheme();

  if (text.length < 10) {
    showStatus("請先貼上教案內容", "error");
    return;
  }

  // Loading state
  btnGenerate.classList.add("loading");
  btnGenerate.disabled = true;
  statusEl.hidden = true;
  downloadArea.hidden = true;

  try {
    const response = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ lessonPlan: text, theme }),
    });

    if (!response.ok) {
      const errBody = await response.json().catch(() => ({}));
      throw new Error(errBody.error || `伺服器錯誤 (${response.status})`);
    }

    // Get blob and create download
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);

    // Extract filename from Content-Disposition or fallback
    const disposition = response.headers.get("Content-Disposition") || "";
    const filenameMatch = disposition.match(/filename\*=UTF-8''(.+)/);
    const filename = filenameMatch
      ? decodeURIComponent(filenameMatch[1])
      : "lesson-plan.pptx";

    // Set download link
    downloadLink.href = url;
    downloadLink.download = filename;
    downloadFilename.textContent = filename;
    downloadArea.hidden = false;

    // Auto-download
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);

    showStatus("✅ 簡報已成功生成並開始下載！", "success");
  } catch (err) {
    console.error("Generation failed:", err);
    showStatus(`❌ ${err.message || "生成失敗，請稍後再試"}`, "error");
  } finally {
    btnGenerate.classList.remove("loading");
    btnGenerate.disabled = false;
  }
}

// ── Show status ──
function showStatus(message, type) {
  statusEl.textContent = message;
  statusEl.className = `status ${type}`;
  statusEl.hidden = false;
}

// ── Copy to clipboard ──
async function copyToClipboard(text) {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    // Fallback for older browsers
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    const ok = document.execCommand("copy");
    document.body.removeChild(ta);
    return ok;
  }
}

// ── Event Listeners ──
input.addEventListener("input", updateUI);

btnGenerate.addEventListener("click", generatePptx);

btnExample.addEventListener("click", loadExample);

btnClear.addEventListener("click", clearInput);

// Copy example
btnCopyExample.addEventListener("click", async () => {
  const exampleText = document.getElementById("formatExample").textContent;
  const ok = await copyToClipboard(exampleText);
  btnCopyExample.textContent = ok ? "✓ 已複製" : "複製失敗";
  setTimeout(() => { btnCopyExample.textContent = "複製"; }, 2000);
});

// Copy AI prompt — load text once
const aiPromptEl = document.getElementById("aiPromptText");
aiPromptEl.textContent = AI_PROMPT;

btnCopyPrompt.addEventListener("click", async () => {
  const ok = await copyToClipboard(AI_PROMPT);
  btnCopyPrompt.textContent = ok ? "✓ 已複製" : "複製失敗";
  setTimeout(() => { btnCopyPrompt.textContent = "📋 複製指令"; }, 2000);
});

// Keyboard shortcut: Ctrl/Cmd + Enter to generate
document.addEventListener("keydown", (e) => {
  if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
    e.preventDefault();
    if (!btnGenerate.disabled) {
      generatePptx();
    }
  }
});

// ── Init ──
updateUI();
