# Gemini Canvas 前端協作 Prompt 模組（一）：`index.html` 語意化骨架生成

> **模組定位：** 本 Prompt 專門用於引導 Google Gemini (Canvas 模式) 或 Google AI Studio 生成活動網頁的結構與內容檔案 `index.html`。
> **核心原則：** 關注點分離（Separation of Concerns）。本檔案僅包含 HTML5 語意化結構與企劃文案， **絕不包含 `<style>` 內嵌 CSS 或 `<script>` 內嵌 JavaScript 程式碼**，所有樣式與動態行為均透過外部檔案 `style.css` 與 `script.js` 引用。

---

## 一、Prompt 使用說明（小組操作指南）

1. **複製與替換：** 將下方「第二部分：正式 Prompt 範本」全文複製，並將方括號 `[ ]` 內的內容替換為貴小組在第 4、5、7、9 週所產出的真實專案資產與企劃文案。
2. **在 Gemini Canvas 中執行：**
   - 開啟 Google Gemini，切換或啟動 **Canvas** 互動視窗。
   - 送出替換完成後的 Prompt。
   - 觀察 Canvas 側邊工作區產出的 `index.html` 程式碼。
3. **本機存檔：**
   - 在小組專案資料夾（如 `sdgs-campaign/`）中建立名為 `index.html` 的純文字檔案（編碼格式為 UTF-8）。
   - 將 Canvas 產出的程式碼完整貼入並存檔。

---

## 二、正式 Prompt 範本（可直接複製至 Gemini）

```markdown
# Role & Goal
你是一位精通現代網頁標準（W3C Web Standards）的高級前端架構師與永續倡議（SDGs）UX 設計專家。請為我們的小組專案生成符合語意化 HTML5 標準的活動單頁（One-Page Landing Page）檔案，檔名固定為 `index.html`。

# Architectural Constraints (關注點分離原則)
1. **嚴格禁止內嵌樣式與腳本**：本檔案內部絕對不得出現 `<style>` 標籤或元素內聯樣式（`style="..."`），亦不得出現 `<script>...</script>` 內嵌代碼。
2. **外部檔案引用**：
   - 在 `<head>` 標籤內必須包含 `<link rel="stylesheet" href="style.css">`。
   - 在 `<body>` 閉合標籤前必須包含 `<script src="script.js" defer></script>`。
3. **標準 Head 配置**：
   - 設定 `<!DOCTYPE html>`、`<html lang="zh-Hant-TW">`、`<meta charset="UTF-8">`。
   - 必須包含行動裝置自適應宣告：`<meta name="viewport" content="width=device-width, initial-scale=1.0">`。
   - 包含簡潔的 `<title>` 與 `<meta name="description">`。
4. **語意化標籤結構**：
   - 必須嚴格使用 `<header>`、`<nav>`、`<main>`、`<section>`、`<article>`、`<footer>` 等 HTML5 語意化標籤。
   - 每個 `<section>` 必須配有獨立且語意清晰的 `id` 屬性（如 `id="hero"`、`id="problem"`、`id="solution"`、`id="action"` 等），以支援單頁平滑滾動（Smooth Scroll）。
5. **無障礙與資產標記**：
   - 所有圖片標籤 `<img>` 必須提供完整具備說服力的 `alt` 屬性說明。
   - 所有按鈕 `<button>` 與超連結 `<a>` 必須具備清楚的文字或 `aria-label`。

# Project Upstream Assets (請將括號內的內容替換為本組資料)
- **選定之 SDGs 目標**：[例如：SDG 12 責任消費與生產 / Target 12.3 格外品剩食搶救]
- **專案名稱**：[例如：《格外美好》生活倡議專案]
- **第 4 週核心主視覺海報檔案路徑**：[例如：img/hero_poster.jpg]
- **第 5 週三大權威數據錨點**：
  1. 震撼型數據：[例如：台灣每年丟棄超過 380 萬噸格外蔬果，相當於每分鐘浪費 7.2 噸食物]
  2. 對比型數據：[例如：僅因外貌微疵被棄置的格外品，其維生素 C 與微量元素與頂級品完全相同，營養價值 100% 保留]
  3. 切身型數據：[例如：每位大學外宿生每年因蔬果昂貴減少攝取，但若採購格外品，每月可省下 \$1,200 元並減少 18 公斤碳排放]
- **第 9 週核心倡議主張（Slogan）**：[例如：外表不完美，風味更甜美！用格外品打造平價深夜食堂]
- **第 9 週 Persona 與專屬階梯式行動呼籲（CTA）**：
  - 代表人物畫像：[例如：阿豪，20 歲，大學外宿生，常吃外食、荷包緊張、想吃得健康卻下不了手]
  - 第一階微行動（Micro-CTA）：[例如：按鈕文字「30秒測測你的外宿蔬果赤字」]
  - 第二階體驗行動（Engagement-CTA）：[例如：按鈕文字「領取外宿格外品深夜食譜指南」]
  - 第三階核心承諾行動（Conversion-CTA）：[例如：按鈕文字「立即簽署外宿採購公約，領取合作小農優惠碼」]
- **第 7 週 SynthID 與 AI 透明揭露聲明**：[例如：本專案主視覺海報由 Google Imagen 3 生成並經 SynthID 數位浮水印簽署，全站文案與程式碼由 Google Gemini 協同完成，秉持負責任生成式 AI 原則。]

# Seven Golden Sections to Generate (7 大核心資訊架構模組)
請依序在 `<main>` 結構中建構以下 7 個區塊：

1. **頂部導覽列 (`<nav id="navbar">`)**：
   - 品牌 Logo 與專案名稱。
   - 錨點導航清單：包含「倡議痛點」、「數據真相」、「永續解方」、「立即行動」與「關於我們」。
   - 行動呼籲快捷按鈕（Mini-CTA）。

2. **首頁旗艦橫幅 (`<section id="hero">`)**：
   - 核心標語（`<h1>`，使用專案 Slogan）。
   - 副標題（副標說明與價值主張）。
   - 專案核心主視覺（`<img>` 引用第 4 週海報路徑）。
   - 主要行動按鈕（帶有 `.btn-primary` 類別，文字為第一階微行動）。

3. **痛點數據錨點區 (`<section id="problem">`)**：
   - 區塊標題（如：你看不見的日常浪費）。
   - 三張數據卡片（`.data-card`），分別展示「震撼數據」、「對比數據」與「切身數據」，包含突出的大數字標籤（`.data-number`）與說明文字。

4. **倡議解方路徑 (`<section id="solution">`)**：
   - 區塊標題（如：三步驟，輕鬆翻轉日常）。
   - 步驟導引卡（Step 1: 認知認同 ➔ Step 2: 友善選購 ➔ Step 3: 社群擴散），每步配有小標題與行動說明。

5. **受眾專屬轉換區 (`<section id="action">`)**：
   - 針對 Persona 設計的沉浸式互動區塊。
   - 倡議承諾線上表單（`<form id="pledge-form">`）：
     - 暱稱輸入框（`<input type="text" id="user-name" required>`）。
     - 身分或大學科系選單（`<select id="user-role">`）。
     - 承諾承諾勾選框（`<input type="checkbox" id="pledge-check" required>`）。
     - 送出承諾按鈕（`<button type="submit" class="btn-cta">`，文字為第三階核心承諾行動）。

6. **社會認同與承諾牆 (`<section id="social-proof">`)**：
   - 即時累計連署計數器（`<div class="counter-box">`，包含動態數字 `<span id="counter">1,248</span>` 人已加入）。
   - 滾動或卡片式展示 3 則同儕支持者的真實聲援留言。

7. **頁尾透明揭露與版權 (`<footer id="footer">`)**：
   - 小組資訊與 SDGs 專題名稱。
   - **AI 透明揭露專區（`.synthid-disclosure`）**：明確標示第 7 週 SynthID 水印與 AI 工具協同說明。
   - 版權年份與課程學期標記。

# Output Format
請僅輸出乾淨、語意嚴謹且格式化良好的完整 HTML5 原始碼。代碼中需附帶關鍵結構註解，便於初學者閱讀與後續維護。

```
