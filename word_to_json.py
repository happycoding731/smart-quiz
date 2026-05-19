# -*- coding: utf-8 -*-
"""
Word 题库转 JSON 脚本
支持题型：判断、单选、画图（内嵌图片转Base64）
依赖安装：pip install python-docx
使用方法：python word_to_json.py <你的Word文件.docx>
输出：同级目录下的 questions.json
"""

import sys
import json
import base64
import re
from pathlib import Path
from docx import Document
from docx.image.exceptions import UnrecognizedImageError
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from io import BytesIO

def image_to_base64(image_blob, image_format='png'):
    """将图像二进制转换为 base64 字符串（带 data 前缀）"""
    b64 = base64.b64encode(image_blob).decode('utf-8')
    return f"data:image/{image_format};base64,{b64}"

def extract_images_from_paragraph(paragraph):
    """从段落中提取内联图片，返回 (文本, 图片base64列表)"""
    images = []
    # 方法1：通过 blip 元素提取（适用于内联图片）
    for blip in paragraph._element.xpath('.//a:blip'):
        rid = blip.get(qn('r:embed'))
        if rid:
            part = paragraph.part.related_parts.get(rid)
            if part and hasattr(part, 'image'):
                try:
                    img_blob = part.image.blob
                    img_format = part.image.content_type.split('/')[-1]
                    b64_img = image_to_base64(img_blob, img_format)
                    images.append(b64_img)
                except UnrecognizedImageError:
                    pass
    # 提取文本（移除图片占位符）
    text = paragraph.text
    return text, images

def parse_question(para, is_question_para=True):
    """
    解析单个题目段落，返回字典
    假设题目格式示例：
        判断题：1. 题目描述 (√)  或  题目描述 (×)
        单选题：2. 题目描述 (A)  后面可能有选项段落
        画图题：题目描述 + 图片
    实际解析需要根据你的Word结构定制，这里给出通用框架
    """
    # 先提取文本和图片
    text, imgs = extract_images_from_paragraph(para)
    # 预判题型：如果包含 <img 或检测到图片则优先认为是画图题
    if imgs:
        q_type = 'draw'
        # 构建 question 字段：原文本 + 图片HTML
        img_html = ''.join(f'<img src="{img}">' for img in imgs)
        question = f"<p>{text}</p>{img_html}"
        answer = ''  # 画图题的答案需要单独提取，此处留空，或根据后续段落处理
        options = []
        return {
            'type': q_type,
            'question': question,
            'answer': answer,
            'options': options,
            'correct': ''
        }
    # 判断题型：根据括号内的 √/× 或 A/B/C/D
    if re.search(r'[（(]\s*[√×]\s*[）)]', text):
        q_type = 'judge'
        question = text
        options = []
    elif re.search(r'[（(]\s*[A-D]\s*[）)]', text):
        q_type = 'single'
        question = text
        options = []  # 选项需要从后续段落收集
    else:
        # 默认识别为判断题？可根据需要调整
        q_type = 'judge'
        question = text
        options = []

    return {
        'type': q_type,
        'question': question,
        'answer': '',
        'options': options,
        'correct': ''
    }

def collect_options(paragraphs, start_idx):
    """从 start_idx 开始收集选项段落（格式如 A. xxx 或 A．xxx），返回选项列表和最后索引"""
    options = []
    i = start_idx
    while i < len(paragraphs):
        p = paragraphs[i]
        text = p.text.strip()
        # 匹配选项模式：A. 内容 或 A．内容 或 A、内容
        match = re.match(r'^([A-D])([\.、．])\s*(.*)', text)
        if match:
            opt_text = f"{match.group(1)}. {match.group(3)}"
            options.append(opt_text)
            i += 1
        else:
            break
    return options, i

def extract_draw_answer(paragraphs, start_idx):
    """提取画图题的答案图片（假设答案单独成一个段落，且包含图片）"""
    for i in range(start_idx, len(paragraphs)):
        p = paragraphs[i]
        text, imgs = extract_images_from_paragraph(p)
        if imgs:
            img_html = ''.join(f'<img src="{img}">' for img in imgs)
            return img_html, i+1
        # 也可以根据文本内容判断，比如"答案："等
        if '答案' in text and imgs:
            return img_html, i+1
    return '', start_idx

def convert_word_to_json(docx_path):
    doc = Document(docx_path)
    paragraphs = list(doc.paragraphs)
    questions = []
    i = 0
    total = len(paragraphs)

    while i < total:
        p = paragraphs[i]
        # 跳过空段落
        if not p.text.strip() and not any(p._element.xpath('.//a:blip')):
            i += 1
            continue

        # 解析题目（可能包含图片）
        q_data = parse_question(p)
        # 如果是单选题，需要收集后续的选项
        if q_data['type'] == 'single':
            opts, next_i = collect_options(paragraphs, i+1)
            q_data['options'] = opts
            # 题目文本可能已经包含答案字母，但也可以从选项中提取答案（后续可手动标注）
            # 这里先保留空白，后续可用正则从题目中提取
            # 跳过已处理的选项段落
            i = next_i
        elif q_data['type'] == 'draw':
            # 画图题的答案可能在后面
            ans, next_i = extract_draw_answer(paragraphs, i+1)
            q_data['answer'] = ans
            i = next_i
        else:
            # 判断题或其它，直接移动到下一段落
            i += 1

        questions.append(q_data)

    return questions

def main():
    if len(sys.argv) < 2:
        print("用法: python word_to_json.py <Word文件路径>")
        sys.exit(1)

    input_file = sys.argv[1]
    docx_path = Path(input_file)
    if not docx_path.exists():
        print(f"文件不存在: {input_file}")
        sys.exit(1)

    print("正在解析 Word 文档...")
    questions = convert_word_to_json(docx_path)
    output_file = docx_path.parent / "questions.json"
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(questions, f, ensure_ascii=False, indent=2)
    print(f"✅ 转换完成！共 {len(questions)} 道题，已保存至: {output_file}")

if __name__ == '__main__':
    main()