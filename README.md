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

每個 `## ` 代表一張 Slide，第一個 `## ` 為封面：

```markdown
# 教學主題：XXX

## 封面 — XXX
- **科目**：XXX
- **年級**：XXX
- **時間**：XX 分鐘

## Slide 標題
1. 列表項目一
2. 列表項目二

## Slide 標題
段落文字內容…

## Slide 標題
- 項目符號一
- 項目符號二
```

📋 完整格式指引及 AI 指令：[PROMPT_INSTRUCTION.md](./PROMPT_INSTRUCTION.md)

## 🎨 8 款主題（來自真實 Design Systems）

| 主題 | 來源 | Accent | 風格 |
|------|------|--------|------|
| 海洋學術 | — | `#065A82` | 專業沉穩 |
| Notion 文青 | Notion | `#E16259` | 溫暖人文 |
| Stripe 藍紫 | Stripe | `#533AFD` | 專業科技 |
| Linear 極黑 | Linear | `#5E6AD2` | 極簡暗黑 |
| Apple 純白 | Apple | `#0071E3` | 簡約留白 |
| Airbnb 暖橙 | Airbnb | `#FF385C` | 溫暖活力 |
| 森林自然 | — | `#2C5F2D` | 清新平靜 |
| 日落暖暮 | — | `#B85042` | 濃郁溫度 |

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
