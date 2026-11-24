import docx
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import qn


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

    def set_cell_format(cell, width_dxa=None, fill_hex=None, v_align="center",
                        pad_top=100, pad_bottom=100, pad_left=130, pad_right=130):
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

    def fill_two_col_table(table, fields, review_last_row=False):
        for i, (label, val) in enumerate(fields):
            c0 = table.cell(i, 0)
            c1 = table.cell(i, 1)
            is_review = review_last_row and (i == len(fields) - 1)
            if is_review:
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

    # ------------------ 文件頂部抬頭 ------------------
    p0 = doc.add_paragraph()
    p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p0.paragraph_format.space_after = Pt(2)
    r0 = p0.add_run('數位內容設計（Digital Content Design）｜ PBL 專題實作導引（第十一週）')
    r0.font.name = 'Microsoft JhengHei'
    r0._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')
    r0.font.size = Pt(9.0)
    r0.font.color.rgb = COLOR_MUTED

    p1 = doc.add_paragraph()
    p1.paragraph_format.space_after = Pt(4)
    p1.paragraph_format.space_before = Pt(0)
    r1 = p1.add_run('第十一週：SDGs 專題品牌 DNA 提煉與跨通路視覺自動化實作單')
    r1.font.name = 'Microsoft JhengHei'
    r1._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')
    r1.font.size = Pt(17.0)
    r1.font.bold = True
    r1.font.color.rgb = COLOR_PRIMARY

    p2 = doc.add_paragraph()
    p2.paragraph_format.space_after = Pt(8)
    p2.paragraph_format.line_spacing = 1.25
    r2 = p2.add_run(
        '本實作單用於第十一週課堂工作坊（60 分鐘）。各小組正式將第十週活動網頁原型公開部署並導入 Google Pomelli 引擎，完成品牌基因提煉與跨社群視覺爆發：\n'
        '1. 運用 Track A（Netlify Drop / GitHub Pages 公開 URL）或 Track B（手動覆寫校準），讓 Pomelli 讀取專案色彩與語意基因；\n'
        '2. 完成 Business DNA 四維校準（60-30-10 色票、字體層級、Tone of Voice 語氣、第 4 週海報與 Logo 資產）；\n'
        '3. 運用 Campaigns 結構化指令、Photoshoot 情境合成、Fix Layout 與 Animate 動態化功能，產出 1:1、4:5、9:16 與 1.91:1 四規格資產；\n'
        '4. 執行「同儕 3 秒滑動盲測」與手機安全區（Safe Zone）品牌一致性審計，完成教師現場查核簽章！'
    )
    r2.font.name = 'Microsoft JhengHei'
    r2._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')
    r2.font.size = Pt(9.0)
    r2.font.color.rgb = COLOR_BODY

    # ------------------ 模組一：專案基本資訊與公開化部署路徑 ------------------
    add_heading_paragraph('【模組一：專案基本資訊與第十週網頁公開化部署路徑】')

    t1 = doc.add_table(rows=4, cols=2)
    set_table_props(t1)
    fields_t1 = [
        ('小組組別與專案名稱', '組別：第 ______ 組 ｜ 專案名稱：《__________________________________________________》'),
        ('選定之 SDGs 目標', '第 ______ 號核心目標：________________________（中文完整名稱）'),
        ('網頁公開化路徑選擇\n（Track A / Track B）',
         '• 採用路徑：□ Track A 靜態託管公開 URL 掃描　□ Track B 手動補完校準（Manual Override）\n'
         '• 靜態部署平台：□ Netlify Drop　□ GitHub Pages　□ 其他：____________\n'
         '• 公開網站 URL：https://________________________________________________________\n'
         '• 手機實測檢驗：□ 第 4 週海報與 style.css 樣式於公網載入正常，無破圖或失聯'),
        ('小組成員與品牌分工',
         '專案經理（PM / 時程與查核）：______________ ｜ 視覺設計與品牌總監（Pomelli 主導）：__________\n'
         '內容策略師（Campaign Prompt 文案）：______________ ｜ 跨通路審計與動態驗證：______________')
    ]
    fill_two_col_table(t1, fields_t1)

    # ------------------ 模組二：Business DNA 四維校準與 60-30-10 色票對齊 ------------------
    add_heading_paragraph('【模組二：Business DNA 四維校準紀錄（與第 10 週 :root 色票對齊）】')

    t2 = doc.add_table(rows=4, cols=2)
    set_table_props(t2)
    fields_t2 = [
        ('色彩系統（Color Palette）\n60-30-10 配色法則檢核',
         '【對照第 10 週 style.css :root 變數，覆寫修正 Pomelli 誤抓色碼】：\n'
         '• 主背景色（60% 呼吸留白 / --color-bg）：#____________\n'
         '• 品牌輔助色（30% 區塊層次 / --color-primary & secondary）：#____________ / #____________\n'
         '• 行動強調色（10% CTA 與重點數字 / --color-accent）：#____________\n'
         '• □ 已確認以上色碼完全對齊第 10 週活動網頁，且文字對比度清晰易讀（WCAG AA）'),
        ('字體層級與視覺資產庫\n（Typography & Assets）',
         '• 標題／內文字體性格設定：________________________________________________\n'
         '• 視覺資產補全：□ 已上傳高清第 4 週主視覺海報 ｜ □ 已上傳小組專題 Logo（無裁切缺角）'),
        ('品牌語氣（Tone of Voice）\n與永續核心主張對齊',
         '• 呼應第 9 週 Persona 的三個品牌語氣關鍵詞：____________、____________、____________\n'
         '• 核心 Slogan 與倡議主張確認：《________________________________________________》'),
        ('Business DNA 校準截圖留存',
         '• □ 已擷取「校準後之 Business DNA 總覽畫面」並存入小組雲端資料夾\n'
         '• 人工修正記錄（AI 原先誤判了哪個顏色或語氣？我們如何修正？）：\n'
         '  __________________________________________________________________________________')
    ]
    fill_two_col_table(t2, fields_t2)

    # ------------------ 模組三：Campaign 結構化指令、Photoshoot 與精修軌跡 ------------------
    add_heading_paragraph('【模組三：Campaign 提示設計、Photoshoot 情境合成與自然語言精修軌跡】')

    t3 = doc.add_table(rows=3, cols=2)
    set_table_props(t3)
    fields_t3 = [
        ('結構化 Campaign 提示詞\n（受眾＋數據＋CTA＋安全區）',
         '• 鎖定之第 9 週 Persona 與生活痛點：________________________________________________\n'
         '• 注入之第 5 週數據錨點與 12 字內主標：《__________________________________________》\n'
         '• 導回第 10 週活動網頁之核心 CTA：【 ____________________________________ 】'),
        ('Photoshoot 情境攝影合成\n（賦予生活代入感）',
         '• 上傳之專題核心物件／主視覺元素：________________________________________________\n'
         '• 選定之真實生活情境模板與風格描述：____________________________________________\n'
         '• 合成效果評估（光影融合度與去漂綠檢核）：________________________________________'),
        ('Fix Layout 與自然語言精修\n（記錄版面救援指令）',
         '• 遭遇之初版排版問題（如文字壓到主角、中文字距過擠或背景對比不足）：\n'
         '  __________________________________________________________________________________\n'
         '• 使用功能與下達之精修指令：□ Fix Layout　□ More > Resize　□ 自然語言編修\n'
         '  指令內容：「______________________________________________________________________」')
    ]
    fill_two_col_table(t3, fields_t3)

    # ------------------ 模組四：跨通路視覺資產包清點與規格檢核 ------------------
    add_heading_paragraph('【模組四：跨通路視覺資產包產出清點（assets_pomelli/ 四大規格齊備確認）】')

    t4 = doc.add_table(rows=4, cols=2)
    set_table_props(t4)
    fields_t4 = [
        ('素材 A：IG 方形貼文\n（1:1 ｜ 1080 × 1080 px）',
         '• 歸檔檔名：assets_pomelli/________________________.png\n'
         '• 承載任務：核心 Slogan 與數據訴求 ｜ □ 四周保留 8% 呼吸白邊，主標 12 字以內'),
        ('素材 B：IG/FB 垂直動態\n（4:5 ｜ 1080 × 1350 px）',
         '• 歸檔檔名：assets_pomelli/________________________.png\n'
         '• 承載任務：Photoshoot 生活情境說服 ｜ □ 上數據下實景，無文字過載'),
        ('素材 C：限動／Reels 動態\n（9:16 ｜ 1080 × 1920 px）',
         '• 歸檔檔名：assets_pomelli/________________________（□ .mp4　□ .gif）\n'
         '• 動態與安全區檢核：□ 已執行 Animate 轉為動態 ｜ □ 上 14% / 下 20% 介面區無壓字遮擋'),
        ('素材 D：FB 分享與網頁橫幅\n（1.91:1 ｜ 1200 × 628 px）',
         '• 歸檔檔名：assets_pomelli/________________________.png\n'
         '• 承載任務：活動網址分享封面（OG Image）｜ □ 縮圖瀏覽時主標題仍清晰可辨')
    ]
    fill_two_col_table(t4, fields_t4)

    # ------------------ 模組五：同儕 3 秒滑動盲測與教師審查簽章 ------------------
    add_heading_paragraph('【模組五：同儕「3 秒滑動盲測」與教師品牌一致性審查簽章】')

    t5 = doc.add_table(rows=3, cols=2)
    set_table_props(t5)
    fields_t5 = [
        ('同儕 3 秒滑動盲測互評\n（受測對手組：第 _____ 組）',
         '【受測者以手機滑動速度瀏覽四款社群圖卡與第 10 週網頁後真實作答】：\n'
         '1. 品牌連貫性：把圖卡與網頁並排觀看，一眼辨識為同一品牌之程度（1–5 分）：______ 分\n'
         '2. 安全區與可讀性：9:16 限動按鈕是否避開遮擋區、文字對比是否清楚？ □ 清晰無誤  □ 需微調\n'
         '3. 導流點擊意願：看完這組圖卡後，是否願意點擊進入活動網頁？ □ 極有意願  □ 尚在猶豫\n'
         '   受測者真實建議：________________________________________________________________'),
        ('AI 倫理揭露與課後優化行動',
         '• □ 確認社群文案與圖卡素材皆符合第 7 週規範（引用第 5 週真實數據、不漂綠、誠實標註 AI 協作）\n'
         '• 依據同儕盲測回饋，預計於課後執行的最終微調項目：\n'
         '  __________________________________________________________________________________'),
        ('課堂即時諮詢與\n審查回饋欄\n（授課教師／助教專用）',
         '【品牌基因提煉與跨通路視覺自動化審查判定】：\n'
         '□ 審核通過（Business DNA 與 :root 色票對齊精準、四規格素材齊備、9:16 嚴守安全區且具動態輸出）\n'
         '□ 建議微調（部分圖卡文字過載或對比度不足，或 9:16 按鈕過於貼近底部，請於課後以 Fix Layout 修正）\n'
         '□ 需重新產出（色票嚴重偏離活動網頁，或缺乏 Photoshoot / Animate 實機應用，需重新校準生成）\n\n'
         '教師／助教指導意見：\n'
         '__________________________________________________________________________________\n\n'
         '查核簽章：______________________ ｜ 審核日期：2026 年 11 月 25 日')
    ]
    fill_two_col_table(t5, fields_t5, review_last_row=True)

    output_path = 'SDGs_專題品牌DNA提煉與跨通路視覺自動化實作單.docx'
    doc.save(output_path)
    print(f"Successfully generated week 11 worksheet: {output_path}")


if __name__ == '__main__':
    create_document()
