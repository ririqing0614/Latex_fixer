import re
import os
import sys
from docx import Document
from docx.oxml import OxmlElement
from lxml import etree
import latex2mathml.converter
from docx.oxml import OxmlElement, parse_xml  # 增加 parse_xml

# --- 核心配置 ---

# 获取当前脚本所在目录，确保离线 XSL 文件能被找到
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
XSL_PATH = os.path.join(BASE_DIR, "MML2OMML.XSL")

def latex_to_omml(latex_str):
    try:
        # 1. LaTeX -> MathML
        # 增加一些预处理，防止空内容或特殊转义字符导致崩溃
        if not latex_str.strip():
            return None
            
        mathml = latex2mathml.converter.convert(latex_str)
        
        if not os.path.exists(XSL_PATH):
            return "XSL_MISSING"
            
        xslt = etree.parse(XSL_PATH)
        transform = etree.XSLT(xslt)
        
        mathml_tree = etree.fromstring(mathml)
        omml_tree = transform(mathml_tree)
        
        return omml_tree.getroot()
        
    except Exception as e:
        # 这里修改：打印具体的公式内容和错误类型，方便排查
        print(f"  [DEBUG] 公式内容: {latex_str}")
        print(f"  [DEBUG] 错误详情: {type(e).__name__} - {e}")
        return None

def replace_latex_in_paragraph(paragraph):
    text = paragraph.text
    pattern = r'(\$\$.*?\$\$|\$.*?\$)'
    parts = re.split(pattern, text)
    
    if len(parts) <= 1:
        return

    paragraph.clear()
    
    for part in parts:
        if not part: continue
            
        if (part.startswith('$') and part.endswith('$')):
            clean_latex = part.strip('$').strip()
            
            if not clean_latex:
                paragraph.add_run(part)
                continue
            
            omml_element = latex_to_omml(clean_latex)
            
            if omml_element == "XSL_MISSING":
                paragraph.add_run(part)
            elif omml_element is not None:
                # 关键修正点：使用 etree 导出字符串，再用 parse_xml 导入
                xml_str = etree.tostring(omml_element, encoding='unicode')
                try:
                    # 使用 docx.oxml 提供的 parse_xml 函数
                    paragraph._element.append(parse_xml(xml_str))
                except Exception as e:
                    print(f"插入公式失败: {e}")
                    paragraph.add_run(part)
            else:
                paragraph.add_run(part)
        else:
            paragraph.add_run(part)

def process_document(input_path, output_path):
    if not os.path.exists(XSL_PATH):
        print(f"!!! 错误: 在脚本目录下未找到 MML2OMML.XSL 文件 !!!")
        print(f"请从 Office 目录复制或从 W3C 下载该文件到: {XSL_PATH}")
        return

    print(f"正在加载文档: {input_path}")
    try:
        doc = Document(input_path)
    except Exception as e:
        print(f"文件读取失败: {e}")
        return

    # 遍历所有段落
    for p in doc.paragraphs:
        if '$' in p.text:
            replace_latex_in_paragraph(p)

    # 遍历所有表格（针对实验题目中常见的函数表）
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if '$' in p.text:
                        replace_latex_in_paragraph(p)

    try:
        doc.save(output_path)
        print(f"\n--- 处理成功 ---")
        print(f"输出文件: {output_path}")
    except Exception as e:
        print(f"文件保存失败: {e}")

if __name__ == "__main__":
    print("========================================")
    print("   Word LaTeX 自动化修复工具 (离线版)   ")
    print("========================================\n")
    
    in_path = input("1. 请粘贴输入 Word 文档的完整路径: ").strip('"').strip()
    if not in_path:
        print("路径不能为空。")
        sys.exit()

    out_path = input("2. 请输入输出路径 (直接回车保存在同目录): ").strip('"').strip()
    if not out_path:
        base, ext = os.path.splitext(in_path)
        out_path = f"{base}_公式版{ext}"

    process_document(in_path, out_path)