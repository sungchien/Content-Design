# Gemini Canvas 前端協作 Prompt 模組（三）：`script.js` 輕量微互動與動態邏輯生成

> **模組定位：** 本 Prompt 專門用於引導 Google Gemini (Canvas 模式) 或 Google AI Studio 生成活動網頁的行為與邏輯檔案 `script.js`。
> **核心原則：** 關注點分離（Separation of Concerns）。本檔案僅負責使用純原生 JavaScript（Vanilla JS）實現網頁的動態互動、表單攔截與即時回饋， **絕不包含 HTML 標籤或 CSS 樣式代碼**。
> **技術要求：** 嚴禁依賴任何外部重型函式庫（如 jQuery、React 等），確保程式碼可在任何現代瀏覽器中零配置、直接本機運行！

---

## 一、Prompt 使用說明（小組操作指南）

1. **功能盤點：** 請確認在 `index.html` 中所定義的互動元件 ID（如 `#pledge-form` 表單、`#counter` 累計計數器、`#user-name` 姓名欄位等）。
2. **複製與替換：** 將下方「第二部分：正式 Prompt 範本」複製，並將方括號 `[ ]` 內的互動回饋文字替換為貴小組的專屬設計（如：承諾完成後的感謝彈窗文案）。
3. **在 Gemini Canvas 中執行：**
   - 在 Gemini Canvas 中發送提示：「請為剛剛生成的 `index.html` 與 `style.css`，編寫對應的純原生 `script.js` 檔案」。
   - 送出替換後的完整 Prompt。
4. **本機存檔：**
   - 在專案資料夾中建立名為 `script.js` 的檔案（UTF-8 編碼）。
   - 貼入代碼後，重新整理瀏覽器即可測試即時表單送出、計數器連動與彈窗互動！

---

## 二、正式 Prompt 範本（可直接複製至 Gemini）

```markdown
# Role & Goal
你是一位專注於使用者體驗（UX Interaction）的純原生 JavaScript 前端工程師。請為我們剛剛生成的 `index.html` 與 `style.css`，編寫一份輕量、高相容性且體驗流暢的腳本檔案，檔名固定為 `script.js`。

# Architectural Constraints (代碼原則與防禦性設計)
1. **純原生 Vanilla JavaScript (ES6+)**：
   - 絕對禁止引用或使用任何第三方外部套件（如 jQuery、Bootstrap JS、Axios 等）。
   - 代碼必須在原生現代瀏覽器環境下直接順暢執行。
2. **DOM 生命週期保護**：
   - 所有事件監聽與邏輯必須包裹在 `document.addEventListener('DOMContentLoaded', () => { ... });` 之中，確保 HTML 元件載入完成後才執行腳本。
3. **防禦性元件檢查（Defensive Programming）**：
   - 在為任何 DOM 元件附加監聽器前，必須先判斷該元件是否存在（例如：`if (pledgeForm) { ... }`），防止因元件缺漏引發控制台 JavaScript 報錯。
4. **樣式與邏輯分離**：
   - JS 僅負責處理資料、狀態切換與增加/移除 CSS 類別（如 `.active`、`.hidden`），絕不在 JS 中大量硬編寫 `element.style.color = '...'` 等內嵌樣式。

# Project Interaction Requirements (請將括號內的內容替換為本組設計)
- **初始支持者人數**：[例如：1248（預設基數）]
- **第三階核心承諾送出後的感謝訊息**：
  - 彈窗或提示標題：[例如：感謝你！永續倡議者]
  - 核心感謝內文：[例如：你的承諾已成功加入倡議！我們已將格外品選購地圖與專屬小農優惠碼發送至你的信箱。]
- **微行動（Micro-CTA）互動回饋**：
  - 第一階測驗/計算點擊反饋：[例如：點擊後平滑滾動至痛點數據區，並高亮顯示「外宿蔬果赤字計算結果」]

# Core Interactive Features to Implement (必備四大微互動功能)
請在 `script.js` 中完整實作下列四大互動模組：

1. **表單提交攔截與即時連動（Form Submission Handling）**：
   - 監聽倡議承諾表單 `#pledge-form` 的 `submit` 事件。
   - 使用 `event.preventDefault()` 阻止頁面重新整理跳轉。
   - 取得使用者填寫的姓名/稱謂與身分。
   - 驗證承諾勾選框是否已勾選；若通過，執行以下動作：
     a. 將社會認同區的計數器數字（`#counter`）即時加 1，並加入數字跳動的視覺動畫效果。
     b. 跳出自訂設計的感謝對話框（Modal）或在表單下方動態插入溫暖的成功回饋卡片。
     c. 表單重設（`form.reset()`）。

2. **平滑錨點滾動導航（Smooth Scroll Navigation）**：
   - 監聽頂部導覽列（`#navbar a`）的所有錨點點擊事件。
   - 針對固定導覽列的高度進行位移補償（Offset），避免跳轉後區塊標題被頂部導覽列遮擋。

3. **微行動按鈕快捷互動（Micro-CTA Trigger）**：
   - 監聽 Hero 區域的第一階微行動按鈕（如測驗或探索按鈕）。
   - 點擊後，以平滑動畫導引使用者視角滾動至數據痛點區（`#problem`）或行動區（`#action`）。

4. **本機瀏覽器狀態記憶（LocalStorage Persistence - 友善加分功能）**：
   - 當使用者完成一次承諾簽署後，將簽署狀態簡記於瀏覽器 `localStorage`（如 `localStorage.setItem('sdgs_pledged', 'true')`）。
   - 重新整理網頁時，若檢測到已簽署，可將按鈕狀態變更為「已簽署！感謝您的支持」，提供沉浸且連貫的數位體驗。

# Output Format
請僅輸出乾淨、註解清晰、邏輯分明的完整 `script.js` 原始碼。

```
