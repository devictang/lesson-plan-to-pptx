# 📝 教案轉 PPTX

將結構化教案一鍵轉換成精美設計的 PowerPoint 簡報。

## ✨ 功能

- 🎯 **一鍵生成** — 貼上教案，選擇主題，即刻下載 .pptx
- 🎨 **8 款精美主題** — 來自 Notion、Stripe、Linear、Apple、Airbnb 等真實 Design Systems
- 📋 **6 種 Slide 佈局** — 封面、分隔頁、引用頁、統計頁、表格頁、內容頁，自動偵測
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

每個 `## ` 代表一張 Slide。6 種佈局自動偵測：

```markdown
# 教學主題：XXX

## 封面 — XXX              ← 封面（深色背景）
- **科目**：XXX
- **年級**：XXX

## ---                     ← 分隔頁（章節切換）
核心教學

## Slide 標題               ← 內容頁（支援 1. / - 列表）
1. 列表項目

## Slide 標題               ← 表格頁（pipe table）
| A | B |
|---|---|
| 1 | 2 |

## Slide 標題               ← 引用頁
> 金句
> — 出處

## Slide 標題               ← 統計頁（key-value）
- **標籤**：數值
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
