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
    r0 = p0.add_run('數位內容設計（Digital Content Design）｜ PBL 專題實作導引（第九週）')
    r0.font.name = 'Microsoft JhengHei'
    r0._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')
    r0.font.size = Pt(9.0)
    r0.font.color.rgb = COLOR_MUTED

    p1 = doc.add_paragraph()
    p1.paragraph_format.space_after = Pt(4)
    p1.paragraph_format.space_before = Pt(0)
    r1 = p1.add_run('第九週：SDGs 專題面向語意分析（KJ法）、立體 Persona 與數位行銷企劃書')
    r1.font.name = 'Microsoft JhengHei'
    r1._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')
    r1.font.size = Pt(17.0)
    r1.font.bold = True
    r1.font.color.rgb = COLOR_PRIMARY

    p2 = doc.add_paragraph()
    p2.paragraph_format.space_after = Pt(8)
    p2.paragraph_format.line_spacing = 1.25
    r2 = p2.add_run(
        '本實作單用於第九週課堂工作坊（60 分鐘）。走過期中考，專案正式跨入「階段二：策略企劃與網頁原型開發（第 9–10 週）」，請各組整合前八週累積之核心主視覺海報、三大文獻數據錨點與負責任 AI 透明揭露聲明，深化為具備戰鬥力的數位行銷企劃書：\n'
        '首先調用第 5 週社群聆聽筆記，導入設計思考「KJ 法（親和圖法，Affinity Diagram）」，以卡片化方式由下而上聚類常民抱怨，打破僵化標籤，提煉出專屬於本組 SDGs 議題之客製化 ABSA 面向名稱；\n'
        '接著將最大的痛點卡片堆聚類為立體具象的代表性人物誌（Persona），深入同理其知易行難的內心妥協獨白；\n'
        '進而針對 Persona 量身定制由淺入深的「階梯式行動呼籲（Call-to-Action, CTA）」，並配置 Instagram、Threads、Facebook 與活動網頁之跨平台傳播矩陣；\n'
        '最後將 CTA 全面轉化為符合 SMART 原則的三層量化行銷目標，完成課堂 3 分鐘電梯簡報攻防與教師簽章，作為第 10 週活動原型網頁開發之作戰藍圖！'
    )
    r2.font.name = 'Microsoft JhengHei'
    r2._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')
    r2.font.size = Pt(9.0)
    r2.font.color.rgb = COLOR_BODY

    # ------------------ 表格一：專案基本資訊與前八週資產回流 ------------------
    add_heading_paragraph('【模組一：專案基本資訊與前八週既有資產大回流】')
    
    t1 = doc.add_table(rows=4, cols=2)
    set_table_props(t1)
    
    fields_t1 = [
        ('小組組別與專案名稱', '組別：第 ______ 組 ｜ 專案名稱：《__________________________________________________》'),
        ('選定之 SDGs 目標', '第 ______ 號核心目標：________________________（中文完整名稱）'),
        ('前八週專題核心資產回流\n（海報、Slogan 與數據錨點）', 
         '【第四週主視覺海報核心畫面意象與主角】：\n__________________________________________________________________________________\n'
         '【第五週 Campaign 一句話行動標語（20 字內）】：\n《________________________________________________________________》\n'
         '【第五週三大數據錨點】：震撼型（__________）、對比型（__________）、切身型（__________）\n'
         '【第七週負責任 AI 揭露聲明】：□ 已完成撰寫與備存，確認於第 10 週置於活動網頁頁尾 Footer'),
        ('小組成員與企劃編製分工', '專案經理（PM / 策略總監）：______________ ｜ KJ 法質化研究員（卡片聚類）：__________\n'
         'Persona 與文案策劃：______________ ｜ 跨平台矩陣與 CTA 設計：______________')
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

    # ------------------ 表格二：社群常民心聲之 KJ 法親和歸納 ------------------
    add_heading_paragraph('【模組二：社群常民心聲之 KJ 法（親和圖法）面向歸納】')
    
    t2 = doc.add_table(rows=3, cols=2)
    set_table_props(t2)
    
    fields_t2 = [
        ('提取事實卡片\n（調用第 5 週社群筆記）', 
         '【從第 5 週田野筆記中挑選 2～3 則最具爭議、抱怨或諷刺的社群真實留言（一卡一痛點）】：\n'
         '• 卡片 A（平台來源：[ Threads / Dcard / PTT ]）：\n'
         '  「__________________________________________________________________________________」\n'
         '• 卡片 B（平台來源：[ Threads / Dcard / PTT ]）：\n'
         '  「__________________________________________________________________________________」'),
        ('KJ 法親和聚類與\n專屬 ABSA 面向命名', 
         '【全組由下而上聚類，討論出專屬於本組 SDGs 主題的 2～3 個客製化面向名稱】：\n'
         '（例如：格外品食安疑慮面向、月經微羞恥面向、租屋動線限制面向、文青定價溢價面向）\n'
         '• 面向一：【____________________________________】（卡片數：____ 張）\n'
         '• 面向二：【____________________________________】（卡片數：____ 張）\n'
         '• 面向三：【____________________________________】（卡片數：____ 張）'),
        ('最大卡片堆之痛點剖析\n（鎖定第一大核心阻抗）', 
         '【針對卡片數量最多、痛感最深的那一堆，深入剖析受眾「知易行難」背後的生活委屈與無奈】：\n'
         '__________________________________________________________________________________\n'
         '__________________________________________________________________________________')
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

    # ------------------ 表格三：立體 Persona 人物誌建構 ------------------
    add_heading_paragraph('【模組三：從最大痛點堆收斂立體 Persona（人物誌）建構】')
    
    t3 = doc.add_table(rows=3, cols=2)
    set_table_props(t3)
    
    fields_t3 = [
        ('代表人物基本輪廓\n（具體化代表受眾模型）', 
         '姓名/稱謂：__________ ｜ 年齡：____ 歲 ｜ 身分/學校科系：__________\n'
         '居住型態：□ 住家  □ 外宿無廚房套房  □ 學校宿舍 ｜ 每日可支配餐費：______ 元'),
        ('數位生活軌跡與摩擦切片', 
         '每日最常滑的社群平台：□ Threads  □ Instagram  □ Dcard ｜ 作息最大壓力來源：__________\n'
         '生活情境描述（約 40 字）：________________________________________________________'),
        ('受眾心理阻抗與內心獨白\n（同理心錨點）', 
         '【受眾面對本專案 SDGs 議題時，最真實的內心妥協獨白（一句話總結）】：\n'
         '「不是我不支持永續，而是因為______________________________________________________\n'
         '逼得我不得不妥協。」')
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

    # ------------------ 表格四：階梯式行動呼籲 (CTA) 與跨平台傳播矩陣 ------------------
    add_heading_paragraph('【模組四：針對 Persona 之階梯式行動呼籲（CTA）與跨平台傳播矩陣】')
    
    t4 = doc.add_table(rows=3, cols=2)
    set_table_props(t4)
    
    fields_t4 = [
        ('階梯式行動呼籲設計\n（Ladder of CTA）', 
         '• 【第一階：微行動（Micro-CTA，30 秒零負擔參與）】：\n'
         '  具體動詞指令與訴求：____________________________________________________________\n'
         '  （例如：花 10 秒測驗你的外宿剩食指數、點擊領取月底格外品 50 元折價券）\n'
         '• 【第二階：體驗行動（Engagement-CTA，輕量回饋融入生活）】：\n'
         '  參與誘因與機制：________________________________________________________________\n'
         '• 【第三階：承諾行動（Conversion-CTA，第 10 週網頁核心轉化）】：\n'
         '  活動網頁具體轉換目標：__________________________________________________________\n'
         '  （例如：完成校園免洗循環聯盟連署、一鍵登記加入格外品共購群組）'),
        ('跨平台傳播矩陣配置\n（Platform Matrix Strategy）', 
         '• 【Instagram（視覺美感與生活風格展現）】：\n'
         '  定位：主視覺海報發布、多圖輪播懶人包 ｜ 內容語調：溫暖質感、啟發生活風格\n'
         '• 【Threads（生活摩擦與銳利吐槽迷因）】：\n'
         '  定位：直擊青年痛點短金句、吐苦水串文發酵 ｜ 內容語調：幽默自嘲、反說教、去標籤化\n'
         '• 【Facebook（深度論據與跨界破圈）】：\n'
         '  定位：轉發官方數據白皮書、長篇案例訪談 ｜ 內容語調：理性客觀、建立公共信任度'),
        ('活動專題網頁 Landing Page\n（第 10 週核心產出錨定）', 
         '【活動網頁作為承諾行動（CTA）轉換終點站之核心任務】：\n'
         '• 網頁主要轉換按鈕文字：【 ____________________________________ 】\n'
         '• 前端文案如何卸下受眾防備：____________________________________________________')
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

    # ------------------ 表格五：SMART 目標與課堂簡報互評簽章 ------------------
    add_heading_paragraph('【模組五：SMART 量化行銷目標設定、電梯簡報互評與教師審查簽章】')
    
    t5 = doc.add_table(rows=3, cols=2)
    set_table_props(t5)
    
    fields_t5 = [
        ('SMART 原則量化行銷目標\n（將 CTA 轉化為三層量化 KPI）', 
         '【全案數位傳播三層量化指標設定】：\n'
         '1. 【認知擴散指標（Awareness / Specific & Measurable）】：\n'
         '   活動網頁上線 14 天內，累積不重複瀏覽人次達 ______ PV\n'
         '2. 【情感互動指標（Engagement / Achievable & Relevant）】：\n'
         '   Threads 串文與 IG 限動累計獲得 ______ 次轉發與微行動（Micro-CTA）點擊參與\n'
         '3. 【行為轉換指標（Conversion / Time-bound）】：\n'
         '   在專案推動期間內，累計吸引 ______ 位目標受眾完成網頁核心承諾行動簽署'),
        ('課堂 3 分鐘電梯簡報互評\n（對手組別：第 _____ 組）', 
         '【對手組別紅隊稽核員挑刺意見（Persona 是否真實？CTA 行動門檻是否過高？SMART 目標是否合理？）】：\n'
         '__________________________________________________________________________________\n'
         '【本組藍隊企劃總監之抗辯與優化回應（如何依據 KJ 歸納面向分析捍衛企劃策略）】：\n'
         '__________________________________________________________________________________'),
        ('課堂即時諮詢與\n審查回饋欄\n（授課教師／助教專用）', 
         '【企劃書戰略完整度與可行性審查判定】：\n'
         '□ 審核通過（黃金圈架構完備、KJ 法面向歸納深刻切題、Persona 具體有血肉、階梯 CTA 切實可行）\n'
         '□ 建議微調（Persona 稍顯空泛，或微行動門檻仍然過高，建議微調 CTA 動詞與誘因機制）\n'
         '□ 需重新收斂（缺乏真實社群痛點支撐，仍帶有道德說教感，需重新梳理 KJ 歸納面向）\n\n'
         '教師／助教指導意見：\n'
         '__________________________________________________________________________________\n\n'
         '查核簽章：______________________ ｜ 審核日期：2026 年 11 月 11 日')
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
    output_path = 'SDGs_專題社群情緒極性分析與數位行銷企劃書.docx'
    doc.save(output_path)
    print(f"Successfully generated week 9 worksheet: {output_path}")

if __name__ == '__main__':
    create_document()
