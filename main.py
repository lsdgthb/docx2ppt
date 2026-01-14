# -*- coding: utf-8 -*-
"""
Word → PPT 块级复制（保留模板首尾页）
"""
import os
import re
import win32com.client as win32
import pythoncom
from win32com.client import constants as c

# --------------  配置区  --------------
# WORD_FILE   = r"D:\pyproject\docx2ppt\2.2【审批部】审查意见-唐山350MW风电.docx"
WORD_FILE   = r"D:\pyproject\docx2ppt\2.2【审批部】审查意见-悦达集团.docx"
TEMPLATE    = r"D:\pyproject\docx2ppt\company_template.pptx"
MAX_CHAR    = 250          # 单页最多汉字数
# 输出文件名将在运行时从 Word 里抓取“项目名称：xxx”自动生成
# ---------------------------------------

ppLayoutText = 2           # 标题+内容
ppLayoutBlank = 12         # 空白（表格用）
ppPastePNG = 2             # PNG 粘贴

# ----------- 正则识别标题 -----------
LVL1_RE = re.compile(r'^[一二三四五六七八九十]+、')
LVL2_RE = re.compile(r'^[（(][一二三四五六七八九十]+[)）]')

def get_level(txt):
    txt = txt.strip()
    if LVL1_RE.match(txt):
        return 1
    if LVL2_RE.match(txt):
        return 2
    return 10

# ----------- Word 清洗（与你原函数一致） -----------
def clean_doc(doc):
    print("开始清理Word文档...")

    # 1. 删除头部
    target = "（二）租赁方案基本要素"
    rng = doc.Content.Duplicate
    rng.Find.ClearFormatting()
    if rng.Find.Execute(FindText=target, Forward=True, MatchCase=False):
        start_pos = rng.Start
        doc.Range(0, start_pos).Delete()
        print("删除头部完成")
    else:
        print(f"⚠️ 未找到目标文本: {target}")

    # 2. 循环替换手动换行符 ^l → ^p，直到干净
    while True:
        rng = doc.Content.Duplicate
        rng.Find.ClearFormatting()
        rng.Find.Text = "^l"                # 手动换行符
        rng.Find.Replacement.Text = "^p"    # 段落标记
        replaced = rng.Find.Execute(Replace=2, Forward=True)
        if not replaced:
            break
    print("手动换行符已全部替换")

    # 3. 重编号（不再手动加 \r）
    replacements = [
        ("（二）租赁方案基本要素", "（一）租赁方案基本要素"),
        ("（三）前置会议要求落实情况", "（二）前置会议要求落实情况"),
        ("（四）额度占用与有效期", "（三）额度占用与有效期"),
        ("（五）指导性标准事项说明", "（四）指导性标准事项说明"),
    ]
    for old, new in replacements:
        rng = doc.Content.Duplicate
        rng.Find.ClearFormatting()
        while rng.Find.Execute(FindText=old, Forward=True, MatchCase=False):
            if rng.Information(12):  # 在表格
                cell = rng.Cells(1)
                cell.Range.Text = cell.Range.Text.rstrip('\r\x07').replace(old, new)
            else:
                rng.Text = rng.Text.rstrip('\r\x07').replace(old, new)
    print("重编号完成")

    # 4. 删除签名行
    keys = ["主审员", "复核人", "部门负责人", "日 期", "日期", "日期：", "日 期：", "日  期",
            "主审员：", "复核人：", "部门负责人："]
    paragraphs = list(doc.Paragraphs)
    for para in paragraphs:
        if any(k in para.Range.Text.strip() for k in keys):
            para.Range.Delete()
    print("签名行删除完成")

    # 5. 二次检查是否还有手动换行符
    rng = doc.Content.Duplicate
    rng.Find.Text = "^l"
    rng.Find.ClearFormatting()
    cnt = 0
    while rng.Find.Execute(Forward=True):
        cnt += 1
    print(f"剩余手动换行符数量：{cnt}")

# ---------- 工具函数：幻灯片 / 块推送 ----------
def create_new_slide(insert_index):
    new_slide = prs.Slides(2).Duplicate()[0]
    if insert_index < prs.Slides.Count:
        new_slide.MoveTo(insert_index)
    return new_slide

def push_block(block_rng, insert_index):
    text = block_rng.Text.replace('\r', '').replace('\x07', '').strip()
    if not text or text.isdigit():
        return
    print(f"📄 推送第{insert_index}页: {text[:50]}...")
    new_slide = create_new_slide(insert_index)
    try:
        new_slide.Shapes.Placeholders(2).Delete()
    except:
        pass
    txt_box = new_slide.Shapes(1)
    tf = txt_box.TextFrame
    tf.TextRange.Font.Size = 15
    tf.TextRange.Font.Name = "仿宋"
    tf.TextRange.Font.Bold = False
    tf.TextRange.Font.Color.RGB = 0x000000
    block_rng.Copy()
    pythoncom.PumpWaitingMessages()
    tf.TextRange.Paste()
    pw, ph = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight
    txt_box.Left = (pw - txt_box.Width) / 2
    txt_box.Top = 70

def push_table_as_image(tbl_rng, insert_index):
    """在指定位置插入表格页"""
    print(f"📊 推送表格：位置{insert_index}")

    # 创建新幻灯片
    new_slide = create_new_slide(insert_index)

    # 移除原有文本框（如果有）
    for shape in list(new_slide.Shapes):
        if shape.HasTextFrame:
            shape.Delete()

    # 复制并粘贴表格
    tbl_rng.Copy()
    pythoncom.PumpWaitingMessages()
    shape = new_slide.Shapes.PasteSpecial(0)[0]  # 0 = 保持源格式

    # 居中定位
    pw, ph = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight
    shape.Left = (pw - shape.Width) / 2 - 80
    shape.Top = (ph - shape.Height) / 2 + 20

# ----------- 主流程 -----------
# ----------- 主流程 -----------
def main():
    global prs, insert_pos, buffer_rng, skip_until_end, done_tables, current_char
    done_tables = set()          # 已整表推送过的 Word 表格 ID 池

    word = win32.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(WORD_FILE)

    # 1. 提取项目名称
    content_before_clean = doc.Content.Text
    m = re.search(r'项目名称：([^\r\x07]+)', content_before_clean)
    project_name = m.group(1).strip() if m else "未知项目"
    OUTPUT = rf"D:\pyproject\docx2ppt\{project_name}.pptx"
    print(f"✅ 项目名称：{project_name}")

    # 2. 清洗
    clean_doc(doc)

    # 3. 打开PPT模板并回写项目名称
    ppt = win32.Dispatch("PowerPoint.Application")
    prs = ppt.Presentations.Open(TEMPLATE)
    for shp in prs.Slides(1).Shapes:
        if shp.Type == 17 and shp.Name == "TextBox 26":
            shp.TextFrame.TextRange.Text = project_name
            break
    else:
        for shp in prs.Slides(1).Shapes:
            if shp.HasTextFrame and shp.TextFrame.HasText and "项目名称" in shp.TextFrame.TextRange.Text:
                shp.TextFrame.TextRange.Text = f"项目名称：{project_name}"
                break

    # 4. 主循环：块级复制
    insert_pos = 2
    buffer_rng = None
    skip_until_end = 0
    paragraphs = list(doc.Paragraphs)
    i = 0
    while i < len(paragraphs):
        par = paragraphs[i]
        rng = par.Range
        txt = rng.Text.strip('\r\a\f\t\x07 ')

        if not txt or (txt and re.match(r'^\d+$', txt)):
            i += 1
            continue
        if skip_until_end and rng.End <= skip_until_end:
            i += 1
            continue

        # =========  表格统一入口  =========
        if rng.Information(12):
            tbl = rng.Tables(1)
            if tbl is None:
                i += 1
                continue
            tbl_key = (tbl.Range.Start, tbl.Range.End)
            if tbl_key not in done_tables:
                done_tables.add(tbl_key)
                # 1. 先 flush 文本缓冲区
                if buffer_rng:
                    push_block(buffer_rng, insert_pos)
                    insert_pos += 1
                    buffer_rng = None
                # 2. 整表一次性推成图片
                print(f'📊 推送表格：位置 {insert_pos}')
                push_table_as_image(tbl.Range, insert_pos)
                insert_pos += 1
                # 3. 跳过整张表
                skip_until_end = tbl.Range.End
                while i < len(paragraphs) and paragraphs[i].Range.End <= skip_until_end:
                    i += 1
                continue
            else:
                i += 1
                continue
        # =========  表格处理结束  =========

        lvl = get_level(txt)
        # 当前段落字符数（不含换行符）
        para_len = len(txt.replace('\r', '').replace('\x07', ''))

        # 一级标题 → 必须新页
        if lvl == 1:
            if buffer_rng:
                push_block(buffer_rng, insert_pos)
                insert_pos += 1
                buffer_rng = None
                current_char = 0  # ← 清零
            buffer_rng = rng.Duplicate
            current_char = len(txt.replace('\r', '').replace('\x07', ''))  # ← 重新算
            i += 1
            continue

        # 累加字符数
        if buffer_rng is None:
            buffer_rng = rng.Duplicate
            current_char = len(txt.replace('\r', '').replace('\x07', ''))
        else:
            buffer_rng.SetRange(buffer_rng.Start, rng.End)
            current_char += len(txt.replace('\r', '').replace('\x07', ''))

        # 超字符阈值 → 立即拆页
        if current_char >= MAX_CHAR:
            push_block(buffer_rng, insert_pos)
            insert_pos += 1
            buffer_rng = None
            current_char = 0  # ← 清零

        i += 1

    # 末尾 flush
    if buffer_rng:
        text = buffer_rng.Text.replace('\r', '').replace('\x07', '').strip()
        if text and not re.match(r'^\d+$', text):
            push_block(buffer_rng, insert_pos)
            insert_pos += 1

    # 删除倒数第二页
    try:
        if prs.Slides.Count > 2:
            second_last_index = prs.Slides.Count - 1
            prs.Slides(second_last_index).Delete()
            print(f"删除倒数第二页完成（原第{second_last_index}页）")
    except Exception as e:
        print(f"删除倒数第二页时出错: {e}")

    # 保存
    prs.SaveAs(OUTPUT)
    print(f"✅ 完成！共生成 {prs.Slides.Count} 页，文件：{OUTPUT}")

    doc.Close(SaveChanges=False)
    word.Quit()
    # ppt.Quit()

# ----------- 启动 -----------
if __name__ == "__main__":
    main()