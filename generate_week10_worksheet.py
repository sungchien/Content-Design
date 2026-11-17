import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def create_document():
    doc = docx.Document()
    
    # 頁面邊界設定 (2.032 cm / 0.8 inch 全等邊界)
    section = doc.sections[0]
    section.top_margin = Cm(2.032)
    section.bottom_margin = Cm(2.032)
    section.left_margin = Cm(2.032)
    section.right_margin = Cm(2.032)
    section.page_width = Cm(21.59)
    section.page_height = Cm(27.94)
    
    # 全域樣式設定
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Microsoft JhengHei'
    normal_style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')
    normal_style.font.size = Pt(9.5)
    normal_style.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    normal_style.paragraph_format.line_spacing = 1.2
    normal_style.paragraph_format.space_after = Pt(3)

    # 樣式常數定義
    COLOR_PRIMARY = RGBColor(0x1a, 0x56, 0xb8)     # 皇家藍 (標題與章節標題)
    COLOR_MUTED = RGBColor(0x7f, 0x8c, 0x8d)       # 灰字 (頂部引導備註)
    COLOR_BODY = RGBColor(0x55, 0x55, 0x55)        # 說明次要字
    COLOR_REVIEW = RGBColor(0xc0, 0x39, 0x2b)      # 緋紅色 (教師審查欄)
    
    FILL_HEADER = "F0F4F8"                         # 欄位標籤淺藍灰底色
    BORDER_COLOR = "D0D7DE"                        # 表格細邊框淺灰色
    
    LABEL_WIDTH_DXA = 2880                         # 5.08 cm
    CONTENT_WIDTH_DXA = 7056                       # 12.446 cm
    
    def set_cell_format(cell, width_dxa=None, fill_hex=None, v_align="center", pad_top=100, pad_bottom=100, pad_left=130, pad_right=130):
        tcPr = cell._tc.get_or_add_tcPr()
        
        # 寬度
        if width_dxa is not None:
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = OxmlElement('w:tcW')
                tcPr.append(tcW)
            tcW.set(qn('w:type'), 'dxa')
            tcW.set(qn('w:w'), str(width_dxa))
            
        # 垂直對齊
        vAlign = tcPr.find(qn('w:vAlign'))
        if vAlign is None:
            vAlign = OxmlElement('w:vAlign')
            tcPr.append(vAlign)
        vAlign.set(qn('w:val'), v_align)
        
        # 內距 padding
        tcMar = tcPr.find(qn('w:tcMar'))
        if tcMar is None:
            tcMar = OxmlElement('w:tcMar')
            tcPr.append(tcMar)
        for side, pad in [('top', pad_top), ('bottom', pad_bottom), ('left', pad_left), ('right', pad_right)]:
            m = tcMar.find(qn(f'w:{side}'))
            if m is None:
                m = OxmlElement(f'w:{side}')
                tcMar.append(m)
            m.set(qn('w:w'), str(pad))
            m.set(qn('w:type'), 'dxa')
            
        # 邊框
        borders_xml = f'''
        <w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:top w:val="single" w:sz="4" w:space="0" w:color="{BORDER_COLOR}"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="{BORDER_COLOR}"/>
            <w:left w:val="single" w:sz="4" w:space="0" w:color="{BORDER_COLOR}"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="{BORDER_COLOR}"/>
        </w:tcBorders>
        '''
        existing_borders = tcPr.find(qn('w:tcBorders'))
        if existing_borders is not None:
            tcPr.remove(existing_borders)
        tcPr.append(parse_xml(borders_xml))
        
        # 底色
        if fill_hex:
            shd = tcPr.find(qn('w:shd'))
            if shd is None:
                shd = OxmlElement('w:shd')
                tcPr.append(shd)
            shd.set(qn('w:fill'), fill_hex)

    def set_table_props(table):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        tblPr = table._tbl.tblPr
        tblW = tblPr.find(qn('w:tblW'))
        if tblW is not None:
            tblW.set(qn('w:type'), 'auto')
            tblW.set(qn('w:w'), '0')
        for row in table.rows:
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml('<w:cantSplit xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))

    def add_heading_paragraph(text, color=COLOR_PRIMARY, size_pt=11.5, space_before=10, space_after=3):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Microsoft JhengHei'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')
        run.font.size = Pt(size_pt)
        run.font.bold = True
        run.font.color.rgb = color
        return p

    # ------------------ 文件頂部抬頭 ------------------
    p0 = doc.add_paragraph()
    p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p0.paragraph_format.space_after = Pt(2)
    r0 = p0.add_run('數位內容設計（Digital Content Design）｜ PBL 專題實作導引（第十週）')
    r0.font.name = 'Microsoft JhengHei'
    r0._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')
    r0.font.size = Pt(9.0)
    r0.font.color.rgb = COLOR_MUTED

    p1 = doc.add_paragraph()
    p1.paragraph_format.space_after = Pt(4)
    p1.paragraph_format.space_before = Pt(0)
    r1 = p1.add_run('第十週：SDGs 專題活動網頁原型開發、資訊架構與前端三檔案協作實作單')
    r1.font.name = 'Microsoft JhengHei'
    r1._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')
    r1.font.size = Pt(17.0)
    r1.font.bold = True
    r1.font.color.rgb = COLOR_PRIMARY

    p2 = doc.add_paragraph()
    p2.paragraph_format.space_after = Pt(8)
    p2.paragraph_format.line_spacing = 1.25
    r2 = p2.add_run(
        '本實作單用於第十週課堂工作坊（60 分鐘）。各小組正式將第九週《專題企劃書》轉譯為在瀏覽器中真實運行的 SDGs 活動單頁（One-Page Landing Page）原型。\n'
        '本實作嚴格遵循現代軟體工程「關注點分離（Separation of Concerns）」原則，嚴禁將樣式與腳本混雜於 HTML 中，必須獨立建立 index.html、style.css、script.js 三大檔案：\n'
        '1. 依序調用 prompt_html_generation.md、prompt_css_generation.md、prompt_js_generation.md 引導 Gemini Canvas 生成對應代碼；\n'
        '2. 完整縫合第 4 週核心海報、第 5 週三大數據錨點、第 7 週 SynthID 透明揭露聲明與第 9 週 Persona 階梯式 CTA；\n'
        '3. 透過 Chrome 開發者工具（F12）實施 390px 手機版 RWD 檢測，驗證表單防呆提交與計數器即時跳動互動；\n'
        '4. 完成同儕情境體驗測試互評與教師現場查核簽章，為第 11 週 Google Pomelli 品牌 DNA 提煉做好萬全準備！'
    )
    r2.font.name = 'Microsoft JhengHei'
    r2._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')
    r2.font.size = Pt(9.0)
    r2.font.color.rgb = COLOR_BODY

    # ------------------ 表格一：專案基本資訊與前期資產大回流 ------------------
    add_heading_paragraph('【模組一：專案基本資訊與前九週既有資產大回流】')
    
    t1 = doc.add_table(rows=4, cols=2)
    set_table_props(t1)
    
    fields_t1 = [
        ('小組組別與專案名稱', '組別：第 ______ 組 ｜ 專案名稱：《__________________________________________________》'),
        ('選定之 SDGs 目標', '第 ______ 號核心目標：________________________（中文完整名稱）'),
        ('前九週核心資產回流核驗\n（海報、數據、CTA 與倫理聲明）', 
         '• 【第 4 週核心主視覺海報檔案】：img/____________________（已存入專案資料夾）\n'
         '• 【第 5 週核心 Slogan 標語】：《________________________________________________》\n'
         '• 【第 5 週三大數據錨點】：震撼型（__________）、對比型（__________）、切身型（__________）\n'
         '• 【第 9 週核心承諾行動（Conversion-CTA）】：【 ____________________________________ 】\n'
         '• 【第 7 週 SynthID 揭露聲明】：□ 已備存，確認嵌入頁尾 Footer 倫理徽章區塊'),
        ('小組成員與前端協同分工', '專案經理（PM / 資訊架構審查）：______________ ｜ HTML 結構與文案工程師：__________\n'
         'CSS 視覺與 RWD 工程師：______________ ｜ JS 微互動與表單工程師：______________')
    ]
    
    for i, (label, val) in enumerate(fields_t1):
        c0 = t1.cell(i, 0)
        c1 = t1.cell(i, 1)
        set_cell_format(c0, width_dxa=LABEL_WIDTH_DXA, fill_hex=FILL_HEADER)
        set_cell_format(c1, width_dxa=CONTENT_WIDTH_DXA)
        
        p_lbl = c0.paragraphs[0]
        r_lbl = p_lbl.add_run(label)
        r_lbl.font.bold = True
        r_lbl.font.size = Pt(9.5)
        
        p_val = c1.paragraphs[0]
        r_val = p_val.add_run(val)
        r_val.font.size = Pt(9.0)

    # ------------------ 表格二：前端三檔案獨立分離架構檢驗 ------------------
    add_heading_paragraph('【模組二：前端三檔案獨立分離（關注點分離）工程檢驗】')
    
    t2 = doc.add_table(rows=4, cols=2)
    set_table_props(t2)
    
    fields_t2 = [
        ('專案目錄結構標準化\n（File Directory Audit）', 
         '【本機目錄路徑】：c:/.../_________________sdgs_web/\n'
         '□ index.html（主網頁結構檔，UTF-8 編碼）\n'
         '□ style.css（專屬樣式表，包含 :root 品牌變數宣告）\n'
         '□ script.js（原生 JavaScript 互動檔，無外部 CDN 依賴）\n'
         '□ img/ 資料夾（內含主視覺海報與圖示資產）'),
        ('index.html 關注點分離檢核\n（HTML5 語意與無內嵌樣式）', 
         '• □ 嚴格零內嵌 CSS（全文無任何 <style> 標籤或 style="..." 屬性）\n'
         '• □ 嚴格零內嵌 JS（全文無任何 <script>...</script> 邏輯，外部腳本置於結尾）\n'
         '• □ <head> 區塊正確引用外部樣式：<link rel="stylesheet" href="style.css">\n'
         '• □ 底部正確引用外部腳本：<script src="script.js" defer></script>'),
        ('style.css 品牌變數與設計規範\n（為第 11 週 Pomelli 鋪路）', 
         '【:root 頂部品牌色彩變數設定記錄】：\n'
         '• 品牌主色（--color-primary）：____________ ｜ 輔助色（--color-secondary）：____________\n'
         '• 行動強調色（--color-accent）：____________ ｜ 背景底色（--color-bg）：____________\n'
         '• 手機版 RWD 媒體查詢：□ 包含 @media (max-width: 768px) 自適應斷點'),
        ('script.js 防禦性原生互動規範\n（Vanilla JS 原則）', 
         '• □ 代碼全數包裹於 DOMContentLoaded 事件監聽器內\n'
         '• □ 嚴格零第三方外部相依套件（無引入 jQuery, React, Bootstrap 等）\n'
         '• □ 表單提交攔截具備 event.preventDefault()，支援平滑連動')
    ]
    
    for i, (label, val) in enumerate(fields_t2):
        c0 = t2.cell(i, 0)
        c1 = t2.cell(i, 1)
        set_cell_format(c0, width_dxa=LABEL_WIDTH_DXA, fill_hex=FILL_HEADER)
        set_cell_format(c1, width_dxa=CONTENT_WIDTH_DXA)
        
        p_lbl = c0.paragraphs[0]
        r_lbl = p_lbl.add_run(label)
        r_lbl.font.bold = True
        r_lbl.font.size = Pt(9.5)
        
        p_val = c1.paragraphs[0]
        r_val = p_val.add_run(val)
        r_val.font.size = Pt(9.0)

    # ------------------ 表格三：7 大黃金模組資訊架構與真實文案對照 ------------------
    add_heading_paragraph('【模組三：7 大核心模組資訊架構（IA）與企劃文案對照】')
    
    t3 = doc.add_table(rows=4, cols=2)
    set_table_props(t3)
    
    fields_t3 = [
        ('模組 1～2：導覽列與 Hero 橫幅\n（抓住黃金 5 秒注意力）', 
         '• 頂部導覽列（<nav id="navbar">）：Logo 文字【__________】 ｜ 導航錨點數：____ 個\n'
         '• 首頁旗艦橫幅（<section id="hero">）：\n'
         '  - 主標題 <h1> 完整文案：《________________________________________________》\n'
         '  - 副標題 <p> 價值主張：____________________________________________________\n'
         '  - 微行動按鈕（Micro-CTA）：【 ____________________________________ 】'),
        ('模組 3～4：痛點數據與解方路徑\n（客觀論據與低門檻路徑）', 
         '• 痛點數據區（<section id="problem">）三欄卡片文案：\n'
         '  - 震撼數據：指標數字【__________】 說明【________________________________】\n'
         '  - 對比數據：指標數字【__________】 說明【________________________________】\n'
         '  - 切身數據：指標數字【__________】 說明【________________________________】\n'
         '• 倡議解方區（<section id="solution">）三步驟：\n'
         '  Step 1: __________ ➔ Step 2: __________ ➔ Step 3: __________'),
        ('模組 5～6：受眾行動區與認同牆\n（核心轉換與群體效應）', 
         '• Persona 專屬行動區（<section id="action">）：\n'
         '  - 線上表單 ID：#pledge-form ｜ 核心轉換按鈕文字：【 ______________________ 】\n'
         '  - 消除受眾阻抗之文案設計：________________________________________________\n'
         '• 社會認同牆（<section id="social-proof">）：\n'
         '  - 預設基數：______ 人 ｜ 聲援者留言展示卡片數：____ 則'),
        ('模組 7：頁尾 AI 透明揭露與版權\n（科技倫理與法規誠信）', 
         '• 頁尾標籤（<footer id="footer">）：\n'
         '  - SynthID 揭露標籤（.synthid-disclosure）：□ 已標註 Imagen 3 浮水印與 Gemini 協作\n'
         '  - 小組學術專題著作權聲明：________________________________________________')
    ]
    
    for i, (label, val) in enumerate(fields_t3):
        c0 = t3.cell(i, 0)
        c1 = t3.cell(i, 1)
        set_cell_format(c0, width_dxa=LABEL_WIDTH_DXA, fill_hex=FILL_HEADER)
        set_cell_format(c1, width_dxa=CONTENT_WIDTH_DXA)
        
        p_lbl = c0.paragraphs[0]
        r_lbl = p_lbl.add_run(label)
        r_lbl.font.bold = True
        r_lbl.font.size = Pt(9.5)
        
        p_val = c1.paragraphs[0]
        r_val = p_val.add_run(val)
        r_val.font.size = Pt(9.0)

    # ------------------ 表格四：Canvas 調校軌跡與 Chrome DevTools (F12) 走查 ------------------
    add_heading_paragraph('【模組四：Gemini Canvas 提示調校軌跡與 Chrome DevTools (F12) 實機檢測】')
    
    t4 = doc.add_table(rows=3, cols=2)
    set_table_props(t4)
    
    fields_t4 = [
        ('Gemini Canvas 協同對話調校\n（記錄精準修正指令）', 
         '【記錄小組在 Canvas 中發送之關鍵微調提示詞（如：換色、排版、微動畫）】：\n'
         '• 目標檔案：□ index.html  □ style.css  □ script.js\n'
         '• 初版問題描述：________________________________________________________________\n'
         '• 下達之精準修正指令：\n'
         '  「____________________________________________________________________________」\n'
         '• 修正後之代碼效果：____________________________________________________________'),
        ('Chrome DevTools 跨裝置 RWD 檢測\n（手機版與桌機版走查）', 
         '• 【桌面版寬螢幕檢視（1200px）】：□ 導覽列與三欄卡片正常並列，排版寬鬆平衡\n'
         '• 【手機直向螢幕模擬（390px / iPhone 14）】：\n'
         '  - 佈局自適應：□ 三欄數據卡片自動轉換為垂直單欄堆疊，字體清晰適讀\n'
         '  - 橫向破版檢測：□ 螢幕完全無橫向滾動溢出捲軸（No Horizontal Scroll）\n'
         '  - 觸控友善性：□ 按鈕寬度大於 44px，利於手指輕鬆觸控點擊'),
        ('微互動功能與控制台零報錯檢驗\n（Console & Interactivity）', 
         '• 【表單提交攔截】：□ 點擊提交按鈕無重新整理跳轉，成功觸發感謝回饋訊息\n'
         '• 【計數器連動】：□ 連署數字即時跳動（+1 反饋正常）\n'
         '• 【平滑滾動】：□ 點擊導覽列錨點連結平滑滾動至對應區塊且標題未被遮擋\n'
         '• 【Console 控制台檢驗】：□ 零紅字拋錯（無 404 資源遺失或 JS TypeError 報錯）')
    ]
    
    for i, (label, val) in enumerate(fields_t4):
        c0 = t4.cell(i, 0)
        c1 = t4.cell(i, 1)
        set_cell_format(c0, width_dxa=LABEL_WIDTH_DXA, fill_hex=FILL_HEADER)
        set_cell_format(c1, width_dxa=CONTENT_WIDTH_DXA)
        
        p_lbl = c0.paragraphs[0]
        r_lbl = p_lbl.add_run(label)
        r_lbl.font.bold = True
        r_lbl.font.size = Pt(9.5)
        
        p_val = c1.paragraphs[0]
        r_val = p_val.add_run(val)
        r_val.font.size = Pt(9.0)

    # ------------------ 表格五：同儕情境體驗測試與教師審查簽章 ------------------
    add_heading_paragraph('【模組五：同儕情境體驗測試（User Testing）與教師審查簽章】')
    
    t5 = doc.add_table(rows=3, cols=2)
    set_table_props(t5)
    
    fields_t5 = [
        ('課堂 3 分鐘同儕情境體驗測試\n（受測對手組：第 _____ 組）', 
         '【受測組員化身第 9 週 Persona 進行 30 秒盲測體驗後真實作答】：\n'
         '1. 進入首頁 5 秒內，是否一眼能看出本專案在倡議什麼？ □ 清楚明確  □ 稍顯模糊\n'
         '2. 中間三大數據卡片，哪一張最能引發你的情緒共鳴？【 ________________________ 】\n'
         '3. 滾動至受眾行動區時，是否願意按下核心轉換按鈕？ □ 極有意願  □ 尚在猶豫\n'
         '   受測者真實回饋意見：__________________________________________________________'),
        ('同儕回饋與快速迭代優化方向', 
         '【本組依據受測回饋，預計在課後進行的一項即時代碼拋光優化】：\n'
         '__________________________________________________________________________________'),
        ('課堂即時諮詢與\n審查回饋欄\n（授課教師／助教專用）', 
         '【活動網頁原型工程規範與說服體驗審查判定】：\n'
         '□ 審核通過（標準三檔案獨立分離、關注點分離徹底、RWD 手機自適應良好、前九週資產無縫回流）\n'
         '□ 建議微調（手機版略有橫向溢出，或部分樣式未收斂至 :root 變數，請於課後修復以利 Pomelli 讀取）\n'
         '□ 需重新生成（代碼混雜未分離，或缺乏真實企劃數據支撐，需重新調用 Prompt 模組）\n\n'
         '教師／助教指導意見：\n'
         '__________________________________________________________________________________\n\n'
         '查核簽章：______________________ ｜ 審核日期：2026 年 11 月 18 日')
    ]
    
    for i, (label, val) in enumerate(fields_t5):
        c0 = t5.cell(i, 0)
        c1 = t5.cell(i, 1)
        if i == 2:
            set_cell_format(c0, width_dxa=LABEL_WIDTH_DXA, fill_hex="FDEDEC")
            set_cell_format(c1, width_dxa=CONTENT_WIDTH_DXA, fill_hex="FDFEFE")
            p_lbl = c0.paragraphs[0]
            r_lbl = p_lbl.add_run(label)
            r_lbl.font.bold = True
            r_lbl.font.size = Pt(9.5)
            r_lbl.font.color.rgb = COLOR_REVIEW
        else:
            set_cell_format(c0, width_dxa=LABEL_WIDTH_DXA, fill_hex=FILL_HEADER)
            set_cell_format(c1, width_dxa=CONTENT_WIDTH_DXA)
            p_lbl = c0.paragraphs[0]
            r_lbl = p_lbl.add_run(label)
            r_lbl.font.bold = True
            r_lbl.font.size = Pt(9.5)
        
        p_val = c1.paragraphs[0]
        r_val = p_val.add_run(val)
        r_val.font.size = Pt(9.0)

    # 存檔
    output_path = 'SDGs_專題活動網頁原型開發與資訊架構設計單.docx'
    doc.save(output_path)
    print(f"Successfully generated week 10 worksheet: {output_path}")

if __name__ == '__main__':
    create_document()
