# -*- coding: utf-8 -*-
"""
Word → PPT 块级复制（保留模板首尾页）
"""
import os
import re
import win32com.client as win32
import pythoncom

WORD_FILE = r"D:\pyproject\docx2ppt\2.2【审批部】审查意见-阿勒泰鼎风300MW.docx"
TEMPLATE = r"D:\pyproject\docx2ppt\company_template.pptx"
# OUTPUT = r"D:\pyproject\docx2ppt\评审意见_成品.pptx"

ppLayoutText = 2  # 标题+内容
ppLayoutBlank = 12  # 空白（表格用）
ppPastePNG = 2  # PNG 粘贴

# ----------- 正则识别标题 -----------
LVL1_RE = re.compile(r'^[一二三四五六七八九十]+、')  # 一、 二、 …
LVL2_RE = re.compile(r'^[（(][一二三四五六七八九十]+[)）]')  # （一） （二） …


def get_level(txt):
    txt = txt.strip()
    if LVL1_RE.match(txt):
        return 1
    if LVL2_RE.match(txt):
        return 2
    return 10


# ----------- Word 清洗 -----------
def clean_doc(doc):
    print("开始清理Word文档...")
    target = "（二）租赁方案基本要素"
    rng = doc.Content.Duplicate
    rng.Find.ClearFormatting()
    if rng.Find.Execute(FindText=target, Forward=True, MatchCase=False):
        start_pos = rng.Start
        doc.Range(0, start_pos).Delete()
        print("删除头部完成")
    else:
        print(f"⚠️ 未找到目标文本: {target}")

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
                cell.Range.Text = cell.Range.Text.rstrip('\r\x07').replace(old, new) + '\r'
            else:
                rng.Text = rng.Text.rstrip('\r\x07').replace(old, new) + '\r'[-1]
    print("重编号完成")

    keys = ["主审员", "复核人", "部门负责人", "日 期", "日期", "日期：", "日 期：", "日  期",
            "主审员：", "复核人：", "部门负责人："]
    paragraphs = list(doc.Paragraphs)
    for para in paragraphs:
        if any(k in para.Range.Text.strip() for k in keys):
            para.Range.Delete()
    print("签名行删除完成")


# ---------- 工具函数：全页统一模板 ----------
def create_new_slide(insert_index):
    """在指定位置创建新的内容页（基于模板第2页）"""
    # 复制模板第2页，默认会插入到第2页后面
    new_slide = prs.Slides(2).Duplicate()[0]

    # 如果插入位置不是最后，需要移动到指定位置
    if insert_index < prs.Slides.Count:
        new_slide.MoveTo(insert_index)

    return new_slide


def push_block(block_rng, insert_index):
    text = block_rng.Text.replace('\r', '').replace('\x07', '').strip()

    # 加强空内容检查
    if not text:
        print(f"⚠️ 跳过空文本块")
        return
    if text.isdigit():
        print(f"⚠️ 跳过纯数字块: '{text}'")
        return

    print(f"📄 推送第{insert_index}页: {text[:50]}...")

    # 1. 在指定位置创建新幻灯片
    new_slide = create_new_slide(insert_index)

    # 2. 清空内容占位符（如果有）
    try:
        new_slide.Shapes.Placeholders(2).Delete()
    except:
        pass

    # 3. 取新幻灯片的第1个形状（文本框）
    txt_box = new_slide.Shapes(1)

    # 4. 设置字体格式
    tf = txt_box.TextFrame
    tf.TextRange.Font.Size = 15
    tf.TextRange.Font.Name = "仿宋"
    tf.TextRange.Font.Bold = False
    tf.TextRange.Font.Color.RGB = 0x000000

    # 5. 复制并粘贴内容
    block_rng.Copy()
    pythoncom.PumpWaitingMessages()
    tf.TextRange.Paste()

    # 6. 居中定位
    pw, ph = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight
    txt_box.Left = (pw - txt_box.Width) / 2
    # txt_box.Top = (ph - txt_box.Height) / 2 + 20
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
# ppt.Visible = True
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
MAX_CHAR = 300  # 单页最多汉字数

insert_pos = 2
buffer_rng = None
skip_until_end = 0
prev_end = 0

# 获取所有段落
paragraphs = list(doc.Paragraphs)
i = 0

while i < len(paragraphs):
    par = paragraphs[i]
    rng = par.Range
    txt = rng.Text.strip('\r\a\f\t\x07 ')

    # 跳过空段落和纯数字段落（如页码）
    if not txt or (txt and re.match(r'^\d+$', txt)):
        i += 1
        continue

    # 跳过已处理的表格内容
    if skip_until_end and rng.Start < skip_until_end:
        i += 1
        continue

    # ----------- 表格处理 -----------
    if rng.Information(12):  # 在表格中
        tbl = rng.Tables(1)
        if tbl.Range.Start == rng.Start:
            # 先推送表格前的文本（包括表头）
            if buffer_rng:
                push_block(buffer_rng, insert_pos)
                insert_pos += 1
                buffer_rng = None

            # 推送表格
            push_table_as_image(tbl.Range, insert_pos)
            insert_pos += 1

            # 跳过整个表格
            skip_until_end = tbl.Range.End
            prev_end = tbl.Range.End

            # 跳过表格内的所有段落
            while i < len(paragraphs):
                if paragraphs[i].Range.End >= tbl.Range.End:
                    break
                i += 1
            continue
        else:
            i += 1
            continue

    # ----------- 文本处理 -----------
    lvl = get_level(txt)

    # 检查是否超过字数限制
    current_text = ""
    if buffer_rng:
        current_text = buffer_rng.Text.replace('\r', '').replace('\x07', '')

    # 特殊情况："汇报完毕，请审议"应该尽量和前面的内容放在一页
    is_ending = "汇报完毕，请审议" in txt

    # 如果当前块已经有一定长度，并且遇到结尾，先推送当前块
    if buffer_rng and len(current_text) > 0 and is_ending:
        # 把结尾加到当前块
        buffer_rng.SetRange(buffer_rng.Start, rng.End)
        push_block(buffer_rng, insert_pos)
        insert_pos += 1
        buffer_rng = None
        i += 1
        continue

    # 如果超过字数限制，推送当前块（但排除结尾）
    if buffer_rng and len(current_text) > MAX_CHAR and not is_ending:
        # 不包含当前段落推送
        push_block(buffer_rng, insert_pos)
        insert_pos += 1
        # 从当前段落开始新块
        buffer_rng = rng.Duplicate
        prev_end = rng.End
        i += 1
        continue

    # 一级标题强制分页
    if lvl == 1:
        if buffer_rng:
            push_block(buffer_rng, insert_pos)
            insert_pos += 1
        buffer_rng = rng.Duplicate
        prev_end = rng.End
        i += 1
        continue

    # 正常追加到当前块
    if buffer_rng is None:
        buffer_rng = rng.Duplicate
    else:
        buffer_rng.SetRange(buffer_rng.Start, rng.End)

    prev_end = rng.End
    i += 1

# ----------- 末尾处理 -----------
if buffer_rng:
    # 检查最后一个块是否有内容且不是纯数字
    text = buffer_rng.Text.replace('\r', '').replace('\x07', '').strip()
    if text and not re.match(r'^\d+$', text):
        push_block(buffer_rng, insert_pos)
        insert_pos += 1
    else:
        print(f"跳过空文本块或纯数字块: '{text}'")

# 5. 删除倒数第二页（直接删除，不检查内容）
try:
    if prs.Slides.Count > 2:
        # 倒数第二页的索引是 prs.Slides.Count - 1
        second_last_index = prs.Slides.Count - 1
        prs.Slides(second_last_index).Delete()
        print(f"删除倒数第二页完成（原第{second_last_index}页）")
except Exception as e:
    print(f"删除倒数第二页时出错: {e}")

# 6. 保存
prs.SaveAs(OUTPUT)
print(f"✅ 完成！共生成 {prs.Slides.Count} 页，文件：{OUTPUT}")

doc.Close(SaveChanges=False)
word.Quit()
# ppt.Quit()