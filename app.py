import streamlit as st
import os
import re
import io
from docx import Document
from docx.oxml import parse_xml
from lxml import etree
import latex2mathml.converter


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
XSL_PATH = os.path.join(BASE_DIR, "MML2OMML.XSL")

def latex_to_omml(latex_str):
    try:
        if not latex_str.strip(): return None
        mathml = latex2mathml.converter.convert(latex_str)
        if not os.path.exists(XSL_PATH):
            return "XSL_MISSING"
        xslt = etree.parse(XSL_PATH)
        transform = etree.XSLT(xslt)
        return transform(etree.fromstring(mathml)).getroot()
    except Exception:
        return None

def replace_latex_in_paragraph(paragraph):
    text = paragraph.text
    # 优先匹配 $$...$$
    pattern = r'(\$\$.*?\$\$|\$.*?\$)'
    parts = re.split(pattern, text)
    if len(parts) <= 1: return

    paragraph.clear()
    for part in parts:
        if not part: continue
        if part.startswith('$') and part.endswith('$'):
            clean = part.strip('$').strip()
            if not clean:
                paragraph.add_run(part)
                continue
            
            # 预处理：修复连续上标和积分结构
            if "'''" in clean: clean = clean.replace("'''", "^{\prime\prime\prime}")
            elif "''" in clean: clean = clean.replace("''", "^{\prime\prime}")
            elif "'" in clean: clean = re.sub(r"(?<=[a-zA-Z\)])'", r"^{\prime}", clean)
            
            if '\\int' in clean and '{' not in clean:
                match = re.search(r'(\\int[_^0-9a-zA-Z]*)', clean)
                if match:
                    prefix = match.group(0)
                    body = clean[match.end():].strip()
                    if body: clean = f"{prefix} {{{body}}}"

            omml = latex_to_omml(clean)
            if omml == "XSL_MISSING":
                st.error("❌ 严重错误：服务器端找不到 MML2OMML.XSL 文件！")
                paragraph.add_run(part)
            elif omml is not None:
                xml_str = etree.tostring(omml, encoding='unicode')
                try:
                    paragraph._element.append(parse_xml(xml_str))
                except:
                    paragraph.add_run(part)
            else:
                paragraph.add_run(part)
        else:
            paragraph.add_run(part)

st.set_page_config(page_title="LaTeX 转 Word 公式神器", layout="centered")

st.title("📄 Word LaTeX 乱码修复工具")
st.markdown("上传包含 `$E=mc^2$` 乱码的 Word 文档，自动转换为原生公式。")


if not os.path.exists(XSL_PATH):
    st.warning("⚠️ 警告：未检测到 MML2OMML.XSL 文件，功能将无法使用。")


uploaded_file = st.file_uploader("请选择 .docx 文件", type=["docx"])

if uploaded_file is not None:
    st.info(f"正在处理文件：{uploaded_file.name} ...")
    
    try:
        doc = Document(uploaded_file)
        
        progress_bar = st.progress(0)
        total_p = len(doc.paragraphs)
        
        for i, p in enumerate(doc.paragraphs):
            if '$' in p.text:
                replace_latex_in_paragraph(p)
            # 简单的进度更新
            if i % 10 == 0:
                progress_bar.progress(min(i / total_p, 1.0))
        
        # 处理表格
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if '$' in p.text:
                            replace_latex_in_paragraph(p)
        
        progress_bar.progress(100)
        st.success("✅ 处理完成！")

        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        
        # 下载按钮
        st.download_button(
            label="⬇️ 下载修复后的文档",
            data=bio,
            file_name=f"Fixed_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    except Exception as e:
        st.error(f"处理过程中发生错误: {e}")