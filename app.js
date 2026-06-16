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

// ── AI Prompt Text ──
const AI_PROMPT = `你是一個專業的教案撰寫助手。請根據以下要求，生成一份結構化的教案，**必須嚴格使用以下 Markdown 格式**：

# 教學主題：[在此填寫主題名稱]

## 基本資訊
- **科目**：[科目名稱]
- **年級**：[年級]
- **時間**：[分鐘數] 分鐘
- **課題**：[課題名稱]

## 教學目標
1. [目標一]
2. [目標二]
3. [目標三]

## 教學資源
- [資源一]
- [資源二]
- [資源三]

## 教學流程

### [導入活動名稱]（[分鐘] 分鐘）
[詳細描述導入活動的內容、教師行為和學生行為]

### [發展活動名稱]（[分鐘] 分鐘）
[詳細描述發展活動的內容、教師行為和學生行為]

### [總結活動名稱]（[分鐘] 分鐘）
[詳細描述總結活動的內容]

## 評估方法
- [評估方式一]
- [評估方式二]

---

格式規則：
- 必須使用 \`#\` 作為標題、\`##\` 作為區段標題、\`###\` 作為教學流程的子步驟
- 每個教學流程步驟必須包含時間
- 使用 \`- **key**：value\` 格式表示基本資訊
- 使用數字列表 \`1. \` 表示教學目標
- 使用 \`- \` 表示一般列表項
- 不要加入任何其他格式或多餘的說明文字`;

// ── Format detection ──
function checkFormat(text) {
  if (text.trim().length < 10) return { valid: false, hint: "" };

  const hasTitle = /^#\s+/m.test(text);
  const hasH2 = /^##\s+/m.test(text);
  const hasList = /^[-*]\s+|^\d+[.)]\s+/m.test(text);
  const hasKeyValue = /\*\*[^*]+\*\*[：:]/.test(text);

  if (hasTitle && hasH2 && hasList) {
    return { valid: true, hint: "✅ 格式正確" };
  }
  if (hasTitle || hasH2) {
    return { valid: false, hint: "⚠️ 格式不完整，建議參考下方指引" };
  }
  return { valid: false, hint: "💡 提示：請使用結構化 Markdown 格式" };
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
  input.value = `# 教學主題：唐朝盛世

## 基本資訊
- **科目**：中國歷史
- **年級**：中二
- **時間**：40 分鐘
- **課題**：唐朝的建立與貞觀之治

## 教學目標
1. 了解唐朝建立的歷史背景
2. 分析貞觀之治的特點與影響
3. 培養學生對歷史的批判思考能力

## 教學資源
- 教科書第三章
- 唐朝疆域圖
- 多媒體簡報
- 工作紙

## 教學流程

### 導入活動（5 分鐘）
教師提問：「你心中最強盛的朝代是哪一個？」引導學生思考並分享，帶出唐朝的主題。

### 發展活動（25 分鐘）
教師講解唐朝建立背景（10 分鐘），然後進行小組討論貞觀之治的政策（10 分鐘），最後各組匯報及教師總結（5 分鐘）。

### 總結活動（10 分鐘）
完成工作紙，鞏固所學內容。教師總結本課重點。

## 評估方法
- 課堂提問及回應
- 小組討論表現
- 工作紙完成情況`;
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
