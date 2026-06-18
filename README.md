# 📝 教案轉 PPTX

將結構化教案一鍵轉換成精美設計的 PowerPoint 簡報。

## ✨ 功能

- 🎯 **一鍵生成** — 貼上教案，選擇主題，即刻下載 .pptx
- 🎨 **12 款精美主題** — 從 Open Design (Apache 2.0) 提取的設計系統，涵蓋學術、教學、企業、藝術等場景
- 📋 **13 種 Slide 佈局** — 封面、分隔頁、引用頁、統計頁、表格頁、對照頁、概念卡、提示框、練習區、選擇題、程式碼、時間線、內容頁，自動偵測
- 🎤 **逐字稿支援** — 支援 `> notes：` 語法存入 PPTX speaker notes
- 📚 **學習目標側欄** — 自動提取學習目標，在內容頁左側顯示進度標記
- 🤖 **AI 生成支援** — 提供標準化 AI 指令，可讓 ChatGPT 生成含多種佈局的教案
- 📱 **響應式設計** — 支援桌面及手機瀏覽器

## 🚀 部署（Vercel）

### 一鍵部署

[![Deploy with Vercel](https://vercel.com/button)](https://vercel.com/new/clone?repository-url=https://github.com/devictang/lesson-plan-to-pptx)

### 手動部署

```bash
git clone https://github.com/devictang/lesson-plan-to-pptx.git
cd lesson-plan-to-pptx
npm install
vercel --prod
```

## 🛠 本地開發

```bash
npm install
vercel dev
# 開啟 http://localhost:3000
```

## 📖 教案格式

每個 `## ` 代表一張 Slide。13 種佈局自動偵測：

```markdown
# 教學主題：XXX

## 封面 — XXX              ← 封面（深色背景）
- **科目**：XXX
- **年級**：XXX

## ---                     ← 分隔頁（章節切換）
核心教學

## 核心概念                 ← 概念卡（> **概念**：說明 ×2-3）
> **概念一**：說明
> **概念二**：說明
> **概念三**：說明

## 💡 提示                  ← 提示框（> 💡 ...）
> 💡 重點提示內容

## ✏️ 練習                 ← 練習區（> ✏️ ...）
> ✏️ 練習名稱
> 1. 步驟一
> 2. 步驟二

## ❓ 測驗                  ← 選擇題（> ❓ ... + ✅）
> ❓ 問題
> A. 選項
> ✅ A

## 程式碼範例               ← 程式碼（``` ```）
```python
print("hello")
```

## Slide 標題               ← 表格頁（pipe table）
| A | B |
|---|---|
| 1 | 2 |

## 左標題 || 右標題         ← 對照頁
- 左欄
---
- 右欄

## 時間線                   ← 時間線
1. 事件一
2. 事件二

## 內容頁                   ← 內容頁
1. 列表項目
> notes：逐字稿，存入 PPTX speaker notes
```

📋 完整格式指引及 AI 指令：[PROMPT_INSTRUCTION.md](./PROMPT_INSTRUCTION.md)

## 🎨 12 款主題（從 Open Design 提取）

| 主題 | Accent | 風格 | 適用場景 |
|------|--------|------|---------|
| **教學模組 ★** | `#2D7D6E` | 暖米色、學習目標側欄 | **教案首選** |
| 學術論文 | `#1A3A7A` | 嚴肅 serif、零圓角 | 大學/論文答辯 |
| 企業彙報 | `#1A56DB` | 專業藍色 | 學校會議、行政 |
| 極簡白 | `#18181B` | 黑白留白 | 設計/藝術課 |
| 瑞士網格 | `#E30613` | 紅色網格 | 統計/數據課 |
| 編輯社論 | `#8B2F2F` | 優雅 serif | 語文/文學/歷史 |
| 技術暗色 | `#58A6FF` | GitHub dark | 編程/STEM |
| 柔和馬卡龍 | `#E8927C` | 大圓角、柔和 | 小學/幼教 |
| 藍圖工程 | `#FFD700` | 深藍工程 | 物理/工程/數學 |
| 森林自然 | `#2D7A2D` | 綠色自然 | 生物/地理 |
| 海洋冷色 | `#1A6AA0` | 冷藍專業 | 醫學/化學 |
| 歷史琥珀 | `#B5392A` | 羊皮紙懷舊 | 歷史/社會學 |

## 🏗 技術棧

- **Frontend**: HTML + CSS + Vanilla JS
- **Backend**: Vercel Serverless Functions (Node.js)
- **PPTX Generation**: [pptxgenjs](https://github.com/gitbrent/PptxGenJS)
- **Design System**: [nexu-io/open-design](https://github.com/nexu-io/open-design) (Apache 2.0)
- **Deployment**: Vercel

## 📂 項目結構

```
lesson-plan-to-pptx/
├── api/
│   └── generate.js         # PPTX 生成 API（含 12 主題 + 13 種佈局）
├── index.html              # 前端主頁
├── styles.css              # 樣式表
├── app.js                  # 前端邏輯
├── package.json
├── vercel.json             # Vercel 配置
├── PROMPT_INSTRUCTION.md   # AI 指令說明（含新佈局語法）
└── README.md
```

## 📄 授權

MIT License · Design system tokens extracted from [nexu-io/open-design](https://github.com/nexu-io/open-design) (Apache 2.0)
