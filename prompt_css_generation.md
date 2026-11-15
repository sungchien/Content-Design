# Gemini Canvas 前端協作 Prompt 模組（二）：`style.css` 視覺系統與響應式排版生成

> **模組定位：** 本 Prompt 專門用於引導 Google Gemini (Canvas 模式) 或 Google AI Studio 生成活動網頁的樣式表檔案 `style.css`。
> **核心原則：** 關注點分離（Separation of Concerns）與設計系統（Design System）思維。本檔案負責全站的色彩變數、字體層級、Flexbox/Grid 佈局、元件外觀及行動裝置響應式排版（Mobile-First RWD）， **絕不包含 HTML 標籤或 JavaScript 程式碼**。
> **下游銜接重要性：** 在 `:root` 中定義的色彩變數，是第 11 週 Google Pomelli 自動掃描活動網頁、提煉 Business DNA 並生成社群跨通路行銷素材的核心依據！

---

## 一、Prompt 使用說明（小組操作指南）

1. **色系與風格設定：** 請先檢查第 4 週 Imagen 3 產出的專題核心主視覺海報。從中挑選出最具代表性的**品牌主色（Primary Color）**、**次要輔助色（Secondary Color）**與**行動強調色（Accent Color，通常用於 CTA 按鈕）**。
2. **複製與替換：** 將下方「第二部分：正式 Prompt 範本」複製，將方括號 `[ ]` 內的色碼與字體風格替換為貴小組的專案設定。
3. **在 Gemini Canvas 中執行：**
   - 在 Gemini Canvas 中，可直接提示：「請針對剛剛生成的 `index.html`，為其編寫對應的獨立 `style.css` 檔案」。
   - 送出替換後的完整 Prompt。
4. **本機存檔：**
   - 在專案資料夾中建立名為 `style.css` 的文字檔案（UTF-8 編碼）。
   - 將代碼完整貼入並存檔，重新整理 `index.html` 即可立即看見精美的視覺排版。

---

## 二、正式 Prompt 範本（可直接複製至 Gemini）

```markdown
# Role & Goal
你是一位精通現代 CSS3、設計系統（Design Tokens）與響應式網頁設計（RWD）的高級 UI/UX 前端工程師。請為我們剛剛定義的 `index.html` 架構，編寫一份現代化、優雅且具備極佳閱讀體驗的樣式表檔案，檔名固定為 `style.css`。

# Architectural Constraints (關注點分離與代碼規範)
1. **純 CSS 檔案**：檔案中絕對不得出現任何 HTML 標籤（如 `<style>`）或 JavaScript 腳本。
2. **CSS 變數（Design Tokens）**：
   - 必須在頂部 `:root` 選擇器中完整宣告全域 CSS 變數，包含：色彩變數（主色、副色、強調色、背景底色、深色文字、淺色邊框）、字型家族（Font Family）、陰影深度（Box Shadow）與圓角規格（Border Radius）。
   - 後續所有元件的色彩與間距必須優先調用 `var(--variable-name)`，方便後續一鍵抽換品牌風格與供第 11 週 Google Pomelli 掃描提煉。
3. **現代佈局技術**：
   - 採用 CSS Flexbox 與 CSS Grid 進行版面規劃。
   - 嚴格禁止使用過時的 `float` 浮動排版或寫死固定像素（`width: 1200px`）導致破版。所有寬度使用百分比（`%`）或彈性單位（`rem`, `fr`, `max-width`）。
4. **行動優先與自適應（Responsive Web Design）**：
   - 預設提供良好的現代桌面版排版。
   - 必須在檔案底部配置標準媒體查詢（Media Queries）：`@media (max-width: 768px)`，針對智慧型手機直向螢幕進行自適應優化（如：導覽列折疊或橫向排列微調、三欄數據卡片自動轉換為單欄直向堆疊、按鈕尺寸放大至適合拇指點擊）。
5. **微互動與平滑過渡（Micro-interactions）**：
   - 全域設定 `html { scroll-behavior: smooth; }` 實現錨點平滑滾動。
   - 所有可點擊按鈕（`.btn`、`a`）必須具備平滑的 `:hover` 與 `:active` 狀態（如微幅陰影加深、微小位移 `transform: translateY(-2px)` 或顏色漸層）。

# Project Design Tokens (請將括號內的內容替換為本組設計)
- **品牌主色（Primary Color）**：[例如：#2D6A4F（象徵永續農林深綠）]
- **次要輔助色（Secondary Color）**：[例如：#52B788（活力淺綠）]
- **行動強調色（Accent / CTA Color）**：[例如：#E76F51（引人注目的暖橙色，用於微行動與轉換按鈕）]
- **主背景底色（Background Color）**：[例如：#F8F9FA（乾淨柔和米白）或 #FFFFFF]
- **文字主色（Text Color）**：[例如：#2B2D42（高對比深炭灰，確保可讀性）]
- **專案調性關鍵字**：[例如：簡約俐落、溫暖人性、親近大學青年、不具教條感]

# Key Component Styling Requirements (各模組樣式規範)
請為 `index.html` 中的 7 大區塊精細編寫對應樣式：

1. **全域與 Reset**：包含標準 Box-sizing（`* { box-sizing: border-box; margin: 0; padding: 0; }`）、無襯線現代字體設定（優先微軟正黑體、PingFang TC、Noto Sans TC）。
2. **頂部導覽列 (`#navbar`)**：
   - 採用 `position: sticky; top: 0;` 或 `position: fixed;`，帶有微半透明毛玻璃效果（`backdrop-filter: blur(10px)`）與優雅底部陰影。
   - 使用 Flexbox 讓 Logo 與導覽超連結左右對齊分散。
3. **首頁旗艦橫幅 (`#hero`)**：
   - 沉浸式 Hero 區域，背景可使用主色系漸層或微漸變。
   - 主標題 `h1` 字級突出醒目（桌面版約 2.5rem - 3.2rem），段落行距寬鬆（line-height: 1.6）。
   - 海報圖片（`#hero img`）配置柔和圓角與立體陰影（`box-shadow`），且設定 `max-width: 100%; height: auto;` 防止圖片撐爆螢幕。
4. **痛點數據錨點區 (`#problem`)**：
   - 採用 CSS Grid 建立三欄並列的卡片佈局（`grid-template-columns: repeat(auto-fit, minmax(260px, 1fr))`）。
   - 數據卡片（`.data-card`）具備白底、邊框與懸浮上升效果（`transition: transform 0.3s ease`）。
   - 震撼數據大數字（`.data-number`）設定為強調色與醒目粗體（font-size: 2.2rem）。
5. **倡議解方路徑 (`#solution`)**：
   - 步驟導引卡以清晰的步驟編號標籤（Badge）引導，具備清晰的視覺流向感。
6. **受眾專屬轉換區 (`#action`)**：
   - 承諾線上表單（`#pledge-form`）置中排版，輸入框（`input`、`select`）具備清晰的 `:focus` 聚焦外框高亮效果。
   - 核心轉換按鈕（`.btn-cta`）視覺重量最重，具備微光澤或醒目背景色。
7. **社會認同牆 (`#social-proof`)**：
   - 計數器外框（`.counter-box`）以大字級展示累計支持人數，給予造訪者強烈的社會群體認同感。
8. **頁尾與 AI 揭露區 (`#footer`)**：
   - 深色背景或沉穩低對比背景。
   - SynthID 揭露標籤（`.synthid-disclosure`）設定微字級與半透明徽章框線，體現嚴謹科技倫理與合規專業度。

# Output Format
請僅輸出完整、語意排版優雅且具備模組化區塊註解的 `style.css` 代碼。

```
