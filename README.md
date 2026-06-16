# 📝 教案轉 PPTX

將結構化教案一鍵轉換成精美設計的 PowerPoint 簡報。

## ✨ 功能

- 🎯 **一鍵生成** — 貼上教案，選擇主題，即刻下載 .pptx
- 🎨 **3 款精美主題** — 海洋藍、森林綠、暖橙紅
- 📋 **結構化 Parsing** — 自動識別教案章節並生成對應 Slide
- 🤖 **AI 生成支援** — 提供標準化 AI 指令，可讓 ChatGPT 生成符合格式的教案
- 📱 **響應式設計** — 支援桌面及手機瀏覽器

## 🚀 部署（Vercel）

### 一鍵部署

[![Deploy with Vercel](https://vercel.com/button)](https://vercel.com/new/clone?repository-url=https://github.com/YOUR_USERNAME/lesson-plan-to-pptx)

### 手動部署

```bash
git clone https://github.com/YOUR_USERNAME/lesson-plan-to-pptx.git
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

教案必須使用以下結構化 Markdown 格式：

```markdown
# 教學主題：XXX

## 基本資訊
- **科目**：XXX
- **年級**：XXX
- **時間**：XX 分鐘
- **課題**：XXX

## 教學目標
1. XXX
2. XXX

## 教學資源
- XXX

## 教學流程

### 導入活動（X 分鐘）
XXX

### 發展活動（X 分鐘）
XXX

### 總結活動（X 分鐘）
XXX

## 評估方法
- XXX
```

📋 完整格式指引及 AI 指令：[PROMPT_INSTRUCTION.md](./PROMPT_INSTRUCTION.md)

## 🎨 主題預覽

| 海洋藍 | 森林綠 | 暖橙紅 |
|--------|--------|--------|
| 專業學術風格 | 清新自然風格 | 溫暖活力風格 |
| Navy + Teal + Gold | Forest + Moss + Tan | Coral + Rose + Navy |

## 🏗 技術棧

- **Frontend**: HTML + CSS + Vanilla JS
- **Backend**: Vercel Serverless Functions (Node.js)
- **PPTX Generation**: [pptxgenjs](https://github.com/gitbrent/PptxGenJS)
- **Deployment**: Vercel

## 📂 項目結構

```
lesson-plan-to-pptx/
├── api/
│   └── generate.js         # PPTX 生成 API
├── index.html              # 前端主頁
├── styles.css              # 樣式表
├── app.js                  # 前端邏輯
├── package.json
├── vercel.json             # Vercel 配置
├── PROMPT_INSTRUCTION.md   # AI 指令說明
└── README.md
```

## 📄 授權

MIT License
