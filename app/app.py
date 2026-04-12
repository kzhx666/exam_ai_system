import streamlit as st
import os
import requests
import docx
from docx.shared import Pt, RGBColor, Mm, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import json
import re
import uuid
import fitz
import subprocess
import hashlib
from PIL import Image, ImageOps, ImageFilter
import numpy as np
import google.generativeai as genai
from dotenv import load_dotenv

load_dotenv()

st.set_page_config(page_title="AI 试卷排版大师", layout="wide")
IMAGE_DIR = "/app/data/images"
os.makedirs(IMAGE_DIR, exist_ok=True)

if 'cart' not in st.session_state:
    st.session_state.cart = []
if 'section_order' not in st.session_state:
    st.session_state.section_order = []

with st.sidebar:
    st.header("🖼️ 漏图补丁站")
    st.info("如果原卷图表丢失，可在此上传截图获取代码，补入题干中。")
    patch_files = st.file_uploader("上传缺失的图片", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="patch_uploader")
    if patch_files:
        for pf in patch_files:
            file_bytes = pf.read()
            file_hash = hashlib.md5(file_bytes).hexdigest()[:6]
            img_ext = pf.name.split('.')[-1].lower()
            safe_name = f"patch_{file_hash}.{img_ext}"
            img_path = os.path.join(IMAGE_DIR, safe_name)
            if not os.path.exists(img_path):
                with open(img_path, "wb") as f:
                    f.write(file_bytes)
            st.image(img_path, width=150)
            st.code(f"[图片:{safe_name}]", language="text")
            st.divider()

st.title("📝 建筑材料与工程试卷 - AI 排版系统")

def normalize_type(raw_type):
    t = str(raw_type).lower()
    if 'single' in t or '单选' in t: return '单选题'
    if 'multiple' in t or '多选' in t: return '多选题'
    if 'judge' in t or '判断' in t: return '判断题'
    if 'fill' in t or '填空' in t: return '填空题'
    if 'short' in t or '简答' in t: return '简答题'
    if 'calc' in t or '计算' in t: return '计算题'
    return raw_type

def clean_option(opt_str):
    return re.sub(r'^[A-H][\.\、\s]+', '', str(opt_str)).strip()

def clean_stem(text):
    return str(text).strip()

def calculate_option_layout(question):
    q_type = normalize_type(question.get('type', ''))
    if q_type in ['判断题', '填空题', '简答题', '计算题']: return "none"
    options_list = question.get('options', [])
    if not options_list or len(options_list) == 0: return "none"
    if len(options_list) != 4: return "block"
    clean_opts = [clean_option(opt) for opt in options_list]
    max_length = max(len(opt) for opt in clean_opts)
    
    # 选项长度判定阈值收紧
    if max_length <= 5: return "inline"
    elif max_length <= 15: return "grid"
    else: return "block"

def group_questions(cart_list, section_configs, ordered_sections):
    groups = {}
    for q in cart_list:
        sec = q.get('clean_section', '未命名板块')
        if sec not in groups: groups[sec] = []
        groups[sec].append(q)
    cn_nums = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十']
    sections = []
    for i, sec_name in enumerate(ordered_sections):
        if sec_name not in groups: continue
        qs = groups[sec_name]
        pts = section_configs.get(sec_name, {}).get('score', 0)
        sec_score = len(qs) * pts
        
        # 智能大题号继承
        if re.match(r'^([一二三四五六七八九十]+|[0-9]+)[\、\.\s]+', sec_name):
            title_text = f"{sec_name}（本大题共 {len(qs)} 小题，每小题 {pts:g} 分，共 {sec_score:g} 分）"
        else:
            idx = cn_nums[i] if i < len(cn_nums) else str(i+1)
            title_text = f"{idx}、{sec_name}（本大题共 {len(qs)} 小题，每小题 {pts:g} 分，共 {sec_score:g} 分）"
            
        sections.append({"title": title_text, "questions": qs})
    return sections

def extract_text_and_images(uploaded_file):
    ext = uploaded_file.name.split('.')[-1].lower()
    text_content = []
    file_bytes = uploaded_file.getvalue()
    
    if ext == 'txt':
        return file_bytes.decode('utf-8', errors='ignore')

    elif ext == 'doc':
        temp_path = f"/tmp/doc_{uuid.uuid4().hex}.doc"
        with open(temp_path, "wb") as f:
            f.write(file_bytes)
        try:
            text = subprocess.check_output(['catdoc', '-d', 'utf-8', temp_path], stderr=subprocess.STDOUT)
            return text.decode('utf-8', errors='ignore')
        except Exception as e:
            st.error(f"❌ .doc 解析失败: {str(e)}")
            return None

    elif ext == 'pdf':
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            for page in doc:
                blocks = page.get_text("dict")["blocks"]
                blocks.sort(key=lambda b: b["bbox"][1]) 
                for b in blocks:
                    if b["type"] == 0:  
                        para_text = ""
                        for line in b["lines"]:
                            for span in line["spans"]:
                                para_text += span["text"]
                        if para_text.strip():
                            text_content.append(para_text)
                    elif b["type"] == 1: 
                        image_bytes = b.get("image")
                        img_ext = b.get("ext", "png")
                        if image_bytes:
                            img_name = f"img_{uuid.uuid4().hex[:6]}.{img_ext}"
                            img_path = os.path.join(IMAGE_DIR, img_name)
                            with open(img_path, "wb") as f:
                                f.write(image_bytes)
                            text_content.append(f"\n[图片:{img_name}]\n")
            return "\n".join(text_content)
        except Exception as e:
            st.error(f"❌ PDF 解析失败: {str(e)}")
            return None
        
    elif ext == 'docx':
        try:
            doc = docx.Document(io.BytesIO(file_bytes))
            for element in doc.element.body:
                if element.tag.endswith('p'):
                    para = docx.text.paragraph.Paragraph(element, doc)
                    para_text = ""
                    for run in para.runs:
                        para_text += run.text
                        blips = run._element.xpath('.//a:blip')
                        for blip in blips:
                            rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                            if rId and rId in doc.part.rels:
                                rel = doc.part.rels[rId]
                                if "image" in rel.target_ref:
                                    img_data = rel.target_part.blob
                                    img_ext = rel.target_part.content_type.split('/')[-1]
                                    img_name = f"img_{uuid.uuid4().hex[:6]}.{img_ext}"
                                    img_path = os.path.join(IMAGE_DIR, img_name)
                                    with open(img_path, "wb") as f:
                                        f.write(img_data)
                                    para_text += f"\n[图片:{img_name}]\n"
                    if para_text.strip():
                        text_content.append(para_text)
                elif element.tag.endswith('tbl'):
                    table = docx.table.Table(element, doc)
                    for row in table.rows:
                        row_data = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                        if row_data: text_content.append(" | ".join(row_data))
            return "\n".join(text_content)
        except Exception as e:
            st.error(f"❌ Word 读取失败: {str(e)}")
            return None
            
    elif ext in ['png', 'jpg', 'jpeg']:
        ocr_url = os.getenv("OCR_API_URL", "http://paddleocr:8866/predict/system")
        try:
            files = {"file": (uploaded_file.name, file_bytes, uploaded_file.type)}
            response = requests.post(ocr_url, files=files)
            return response.json().get("text", "") if response.status_code == 200 else None
        except: return None
    return None

def set_run_font(run, font_name, size_pt, bold=False, color=None):
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size_pt)
    run.bold = bold
    if color: run.font.color.rgb = color

def add_page_number_to_footer(doc):
    for section in doc.sections:
        footer = section.footer
        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run1 = p.add_run("第 ")
        set_run_font(run1, '宋体', 10)
        run2 = p.add_run()
        fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin')
        instr1 = OxmlElement('w:instrText'); instr1.set(qn('xml:space'), 'preserve'); instr1.text = 'PAGE \\* MERGEFORMAT'
        fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'separate')
        fld3 = OxmlElement('w:fldChar'); fld3.set(qn('w:fldCharType'), 'end')
        run2._r.extend([fld1, instr1, fld2, fld3])
        run3 = p.add_run(" 页  共 ")
        set_run_font(run3, '宋体', 10)
        run4 = p.add_run()
        fld4 = OxmlElement('w:fldChar'); fld4.set(qn('w:fldCharType'), 'begin')
        instr2 = OxmlElement('w:instrText'); instr2.set(qn('xml:space'), 'preserve'); instr2.text = 'NUMPAGES \\* MERGEFORMAT'
        fld5 = OxmlElement('w:fldChar'); fld5.set(qn('w:fldCharType'), 'separate')
        fld6 = OxmlElement('w:fldChar'); fld6.set(qn('w:fldCharType'), 'end')
        run4._r.extend([fld4, instr2, fld5, fld6])
        run5 = p.add_run(" 页")
        set_run_font(run5, '宋体', 10)

def sanitize_image_for_docx(img_path):
    try:
        clean_path = img_path + "_v3_super_clean.png" 
        
        if not os.path.exists(clean_path):
            with Image.open(img_path) as img:
                if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
                    bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
                    bg.paste(img, mask=img.convert('RGBA').split()[3])
                    img = bg.convert("RGB")
                else:
                    img = img.convert("RGB")
                    
                gray = img.convert("L")
                img_arr = np.array(gray)
                
                # CAD 等极黑背景反转
                if np.mean(img_arr) < 90:
                    img_arr = 255 - img_arr
                    
                blur_img = Image.fromarray(img_arr).filter(ImageFilter.GaussianBlur(radius=30))
                blur_arr = np.array(blur_img, dtype=np.float32) + 1e-5 
                img_arr_f = np.array(img_arr, dtype=np.float32)
                
                norm_arr = np.clip((img_arr_f / blur_arr) * 255.0, 0, 255)
                
                norm_arr = np.where(norm_arr > 200, 255, norm_arr) 
                norm_arr = np.where(norm_arr < 150, norm_arr * 0.7, norm_arr) 
                norm_arr = np.clip(norm_arr, 0, 255).astype(np.uint8)
                
                final_img = Image.fromarray(norm_arr)
                
                inv_for_bbox = ImageOps.invert(final_img)
                bbox = inv_for_bbox.getbbox()
                if bbox:
                    margin = 10
                    left = max(0, bbox[0] - margin)
                    top = max(0, bbox[1] - margin)
                    right = min(final_img.width, bbox[2] + margin)
                    bottom = min(final_img.height, bbox[3] + margin)
                    final_img = final_img.crop((left, top, right, bottom))
                    
                final_img.save(clean_path, "PNG")
        return clean_path
    except Exception as e:
        print(f"Sanitize error: {e}")
        return img_path 

def generate_word_direct(sections, show_answer, paper_size, meta_info):
    doc = docx.Document()
    section_cfg = doc.sections[0]
    
    if paper_size == "A3 双栏正式版":
        section_cfg.page_width = Mm(420)
        section_cfg.page_height = Mm(297)
        sectPr = section_cfg._sectPr
        cols = sectPr.xpath('./w:cols')
        if not cols:
            cols = OxmlElement('w:cols')
            sectPr.append(cols)
        else:
            cols = cols[0]
        cols.set(qn('w:num'), '2')     
        cols.set(qn('w:space'), '720') 
        max_img_width = Cm(7.5)
    else:
        section_cfg.page_width = Mm(210)
        section_cfg.page_height = Mm(297)
        max_img_width = Cm(14)

    add_page_number_to_footer(doc)
    style = doc.styles['Normal']
    style.paragraph_format.line_spacing = 1.25
    style.paragraph_format.space_after = Pt(0)
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(6)
    set_run_font(p_title.add_run('《建筑材料》智能化测试卷'), '黑体', 18, bold=True)
    
    meta_text = f"（考试时间：{meta_info['time']}分钟    满分：{meta_info['score']:g}分    考试形式：{meta_info['type']}）"
    p_meta = doc.add_paragraph(meta_text)
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_meta.paragraph_format.space_after = Pt(12)
    set_run_font(p_meta.runs[0], '宋体', 10.5)

    p_info = doc.add_paragraph('班级：__________  姓名：__________  学号：__________  得分：__________')
    p_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_info.paragraph_format.space_after = Pt(18)
    set_run_font(p_info.runs[0], '宋体', 12)
    
    for section in sections:
        p_sec = doc.add_paragraph()
        p_sec.paragraph_format.space_before = Pt(12)
        p_sec.paragraph_format.space_after = Pt(6)
        set_run_font(p_sec.add_run(section['title']), '黑体', 14, bold=True)
        
        for i, q in enumerate(section['questions']):
            q_type = normalize_type(q.get('type', ''))
            p_q = doc.add_paragraph()
            p_q.paragraph_format.space_before = Pt(8) 
            p_q.paragraph_format.left_indent = Cm(0.6)
            p_q.paragraph_format.first_line_indent = Cm(-0.6)
            
            raw_content = str(q.get('content', ''))
            img_tags = re.findall(r'\[图片:(.*?)\]', raw_content)
            clean_text = re.sub(r'\[图片:.*?\]', '', raw_content).strip()
            clean_text = clean_stem(clean_text)
            
            ans_raw = str(q.get('answer', '')).strip()
            if not ans_raw or ans_raw.lower() in ['null', 'none', '']:
                ans_raw = "待补全"
                
            q_id_val = str(q.get('id', '')).strip()
            q_id = q_id_val if q_id_val else str(i+1)
            clean_text = re.sub(r'^\d+[\.\、\s]+', '', clean_text)
            
            if q_type in ['单选题', '多选题', '判断题']:
                ans_text = ""
                if show_answer:
                    if q_type == '判断题' and ans_raw not in ["待补全"]:
                        ans_text = '√' if ans_raw in ['对', '正确', '√', 'T', '1'] else '×'
                    else:
                        ans_text = ans_raw
                    
                matches = list(re.finditer(r'[\(（]\s*[\)）]', clean_text))
                if matches:
                    last_match = matches[-1]
                    start, end = last_match.span()
                    prefix = clean_text[:start]
                    suffix = clean_text[end:]
                    
                    run_prefix = p_q.add_run(f"{q_id}. {prefix}")
                    set_run_font(run_prefix, '宋体', 12)
                    
                    set_run_font(p_q.add_run("（"), '宋体', 12)
                    if show_answer and ans_text:
                        run_ans = p_q.add_run(ans_text)
                        set_run_font(run_ans, '宋体', 12, bold=True, color=RGBColor(255, 0, 0))
                    else:
                        p_q.add_run("   ")
                    set_run_font(p_q.add_run("）"), '宋体', 12)
                    
                    if suffix:
                        run_suffix = p_q.add_run(suffix)
                        set_run_font(run_suffix, '宋体', 12)
                else:
                    run_q = p_q.add_run(f"{q_id}. {clean_text}")
                    set_run_font(run_q, '宋体', 12)
                    set_run_font(p_q.add_run(" （"), '宋体', 12)
                    if show_answer and ans_text:
                        run_ans = p_q.add_run(f" {ans_text} ")
                        set_run_font(run_ans, '宋体', 12, bold=True, color=RGBColor(255, 0, 0))
                    else:
                        p_q.add_run("   ")
                    set_run_font(p_q.add_run("）"), '宋体', 12)
            else:
                run_q = p_q.add_run(f"{q_id}. {clean_text}")
                set_run_font(run_q, '宋体', 12)
                if q_type == '填空题' and show_answer:
                    run_ans = p_q.add_run(f"  [答案：{ans_raw}]")
                    set_run_font(run_ans, '宋体', 12, color=RGBColor(255, 0, 0))
                
            for img_name in img_tags:
                img_path = os.path.join(IMAGE_DIR, img_name)
                if os.path.exists(img_path):
                    p_img = doc.add_paragraph()
                    p_img.paragraph_format.left_indent = Cm(0.6)
                    run_img = p_img.add_run()
                    try:
                        safe_img_path = sanitize_image_for_docx(img_path)
                        with Image.open(safe_img_path) as tmp_img:
                            w_px, h_px = tmp_img.size
                            ratio = w_px / float(h_px) if h_px > 0 else 1.0
                            
                        if ratio <= 1.3:
                            target_w = min(w_px * 0.02, 4.0)
                        elif ratio >= 2.0:
                            target_w = min(w_px * 0.02, max_img_width.cm)
                        else:
                            target_w = min(w_px * 0.02, 7.0)
                            
                        target_w = min(target_w, max_img_width.cm)
                        target_w = max(target_w, 1.5)
                        
                        run_img.add_picture(safe_img_path, width=Cm(target_w))
                    except Exception as e:
                        set_run_font(run_img, '宋体', 10, color=RGBColor(255, 0, 0))
                        run_img.text = f"\n[⚠️ 原图损坏，请在左侧补丁站截图补齐]"
            
            layout = q.get('layout', 'none')
            raw_opts = q.get('options', [])
            
            if layout != 'none' and raw_opts:
                opts = [clean_option(opt) for opt in raw_opts]
                
                tab_spacing = 2.5 if paper_size == "A3 双栏正式版" else 3.0 
                
                if layout == 'inline' and len(opts) == 4:
                    p_opt = doc.add_paragraph()
                    p_opt.paragraph_format.space_before = Pt(2) 
                    p_opt.paragraph_format.left_indent = Cm(0.6) 
                    p_opt.paragraph_format.tab_stops.add_tab_stop(Cm(0.6 + tab_spacing))
                    p_opt.paragraph_format.tab_stops.add_tab_stop(Cm(0.6 + tab_spacing * 2))
                    p_opt.paragraph_format.tab_stops.add_tab_stop(Cm(0.6 + tab_spacing * 3))
                    run_opt = p_opt.add_run(f"A. {opts[0]}\tB. {opts[1]}\tC. {opts[2]}\tD. {opts[3]}")
                    set_run_font(run_opt, '宋体', 12)
                elif layout == 'grid' and len(opts) == 4:
                    p_opt1 = doc.add_paragraph()
                    p_opt1.paragraph_format.space_before = Pt(2)
                    p_opt1.paragraph_format.left_indent = Cm(0.6)
                    p_opt1.paragraph_format.tab_stops.add_tab_stop(Cm(0.6 + tab_spacing * 2))
                    run1 = p_opt1.add_run(f"A. {opts[0]}\tB. {opts[1]}")
                    set_run_font(run1, '宋体', 12)
                    p_opt2 = doc.add_paragraph()
                    p_opt2.paragraph_format.left_indent = Cm(0.6)
                    p_opt2.paragraph_format.tab_stops.add_tab_stop(Cm(0.6 + tab_spacing * 2))
                    run2 = p_opt2.add_run(f"C. {opts[2]}\tD. {opts[3]}")
                    set_run_font(run2, '宋体', 12)
                else: 
                    for j, opt in enumerate(opts):
                        p_opt_block = doc.add_paragraph(f"{chr(65+j)}. {opt}")
                        p_opt_block.paragraph_format.left_indent = Cm(0.6)
                        if j == 0: p_opt_block.paragraph_format.space_before = Pt(2)
                        set_run_font(p_opt_block.runs[0], '宋体', 12)
            
            if q_type in ['简答题', '计算题', '论述题', '分析题']:
                if show_answer:
                    p_ans = doc.add_paragraph()
                    p_ans.paragraph_format.space_before = Pt(6)
                    run_ans = p_ans.add_run(f"【参考答案】：{ans_raw}")
                    set_run_font(run_ans, '宋体', 12, color=RGBColor(255, 0, 0))
                    doc.add_paragraph()
                else:
                    blank_lines = 8 if q_type == '计算题' else 5
                    for _ in range(blank_lines):
                        doc.add_paragraph()
                    
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

tab1, tab2, tab3 = st.tabs(["⚡ 全自动模式", "🌉 桥接模式 (免 API)", "🛒 试卷组装与导出"])

with tab1:
    st.header("全自动解析 (Gemini大模型)")
    uploaded_auto = st.file_uploader("上传杂乱的试卷 (支持 PDF/Word/Doc/图片/Txt)", type=["png", "jpg", "jpeg", "docx", "pdf", "doc", "txt"], key="auto_up")
    
    if st.button("一键提取与排版"):
        api_key = os.getenv("GEMINI_API_KEY")
        if not api_key or "your_api_key" in api_key:
            st.error("未检测到有效 API Key，请去 /opt/exam_ai/.env 配置。")
        elif uploaded_auto:
            with st.spinner("正在提取文字与图片素材，并请求大模型..."):
                raw_text = extract_text_and_images(uploaded_auto)
                if raw_text:
                    try:
                        genai.configure(api_key=api_key)
                        model = genai.GenerativeModel('gemini-2.5-flash')
                        prompt = f"""你是一个试卷排版员。将以下文本整理为JSON数组。只输出JSON。
包含字段: id(必须提取原卷题号), type(单选/多选/判断/填空/简答/计算), section(所属大题板块，必须在此分类), content(题干), options(字符串数组), answer(答案)。
【极其重要】：如果文本中存在类似 `[图片:img_xxxx.png]` 的占位标记，你必须将其原封不动地保留在对应题目的 content(题干) 中！绝对不能遗漏或删改代码！
文本：{raw_text}"""
                        response = model.generate_content(prompt)
                        json_str = response.text.strip().removeprefix("```json").removeprefix("```").removesuffix("```").strip()
                        questions = json.loads(json_str)
                        for q in questions:
                            q['layout'] = calculate_option_layout(q)
                            st.session_state.cart.append(q)
                        st.success(f"成功解析 {len(questions)} 道题！附带的图片已成功锁定，请前往第三页导出。")
                    except Exception as e:
                        st.error(f"解析失败：{e}")

with tab2:
    st.header("桥接模式：人机协同提取")
    uploaded_bridge = st.file_uploader("上传试卷文件", type=["png", "jpg", "jpeg", "docx", "pdf", "doc", "txt"], key="bridge_up")
    
    if st.button("提取文字并生成提示词"):
        if uploaded_bridge:
            raw_text = extract_text_and_images(uploaded_bridge)
            if raw_text:
                st.success("提取完毕！已将文档中的图片转化为标签。")
                prompt = f"""你是一个专业的试卷排版员。请将以下文字整理为严格的 JSON 数组。
字段：id(必须提取原卷题号), type, section(原卷所属大题标题), content, options(数组), answer。必须只输出JSON。
【极其重要】：如果你看到类似 `[图片:img_xxxx.png]` 的标记，请根据上下文将其完整拷贝到所属题目的 content 字段中！
--- 试卷内容 ---
{raw_text}"""
                st.code(prompt, language="text")
        else: st.warning("请先上传文件！")
        
    json_input = st.text_area("步骤 3：粘贴 AI 返回的 JSON 格式代码", height=150)
    
    if st.button("校验并加入试题库"):
        json_text = json_input.strip()
        if not json_text:
            st.warning("⚠️ 输入框为空，请先粘贴 JSON 代码！")
        else:
            try:
                # 智能清理 markdown
                clean_json = re.sub(r'^```(json)?\s*', '', json_text, flags=re.IGNORECASE)
                clean_json = re.sub(r'\s*```$', '', clean_json, flags=re.IGNORECASE).strip()
                
                # 【终极防崩修复】：智能补齐遗漏的大中小括号
                if clean_json.startswith('{') and clean_json.endswith(']'):
                    clean_json = '[' + clean_json
                elif clean_json.startswith('[') and clean_json.endswith('}'):
                    clean_json = clean_json + ']'
                elif clean_json.startswith('{') and clean_json.endswith('}'):
                    clean_json = '[' + clean_json + ']'
                
                if not clean_json:
                    st.warning("⚠️ 未检测到有效内容，请确保复制了完整的 [ ... ] 数组数据。")
                else:
                    questions = json.loads(clean_json)
                    if isinstance(questions, dict):
                        questions = [questions]
                        
                    for q in questions:
                        q['layout'] = calculate_option_layout(q)
                        st.session_state.cart.append(q)
                    st.success(f"✅ 成功加入 {len(questions)} 道题目。")
            except Exception as e:
                st.error(f"JSON 解析失败：{e}\n\n请检查代码是否复制完整（确保以 [ 开头，以 ] 结尾）。")

with tab3:
    st.header("🖨️ 最终渲染与导出 (教务处标准)")
    
    current_sections = []
    for q in st.session_state.cart:
        sec_name = str(q.get('section', q.get('type', '未命名板块'))).strip()
        if not sec_name: sec_name = '未命名板块'
        q['clean_section'] = sec_name
        if sec_name not in current_sections:
            current_sections.append(sec_name)
            
    for sec in current_sections:
        if sec not in st.session_state.section_order:
            st.session_state.section_order.append(sec)
    st.session_state.section_order = [s for s in st.session_state.section_order if s in current_sections]

    section_configs = {}
    total_calc_score = 0
    
    if len(st.session_state.cart) > 0:
        with st.expander("🛠️ 试题手术台 (直接双击修改，含图片标签)", expanded=True):
            edited_cart = st.data_editor(
                st.session_state.cart, 
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True
            )
            if st.button("💾 保存手术台修改"):
                st.session_state.cart = edited_cart
                st.success("修改已生效，请重新配置下方分值。")
                st.rerun()

    with st.expander("📝 试卷基础信息与大题设置"):
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            meta_time = st.number_input("考试时间 (分钟)", value=90, step=10)
        with col_m2:
            meta_type = st.selectbox("考试形式", ["闭卷", "开卷", "半开卷"])
            
        if st.session_state.section_order:
            st.markdown("##### 📐 大题排版顺序与分值设置")
            for i, sec in enumerate(st.session_state.section_order):
                q_count = sum(1 for q in st.session_state.cart if q['clean_section'] == sec)
                col_n, col_up, col_down, col_s = st.columns([3, 0.5, 0.5, 2])
                with col_n:
                    st.markdown(f"**{sec}** *(含 {q_count} 题)*")
                with col_up:
                    if st.button("⬆️", key=f"up_{sec}", disabled=(i == 0)):
                        st.session_state.section_order[i], st.session_state.section_order[i-1] = st.session_state.section_order[i-1], st.session_state.section_order[i]
                        st.rerun()
                with col_down:
                    if st.button("⬇️", key=f"down_{sec}", disabled=(i == len(st.session_state.section_order)-1)):
                        st.session_state.section_order[i], st.session_state.section_order[i+1] = st.session_state.section_order[i+1], st.session_state.section_order[i]
                        st.rerun()
                with col_s:
                    default_pt = 5.0 if '简答' in sec or '计算' in sec else 2.0
                    score = st.number_input(f"分值 (题)", value=default_pt, step=0.5, min_value=0.5, key=f"scr_{sec}", label_visibility="collapsed")
                section_configs[sec] = {'score': score}
                total_calc_score += score * q_count
            st.success(f"💯 系统已自动累加当前试卷满分为：**{total_calc_score:g} 分**")
    
    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        doc_type = st.radio("输出内容版本", ["纯净学生版 (留空)", "教师解析版 (红字填入答案、打勾叉)"])
    with col2:
        paper_size = st.radio("纸张与排版规格", ["A4 单栏日常版", "A3 双栏正式版"])
        
    st.subheader(f"当前试题量：{len(st.session_state.cart)} 道")
    col_btn1, col_btn2 = st.columns([1, 4])
    with col_btn1:
        if st.button("🗑️ 清空题库"):
            st.session_state.cart = []
            st.session_state.section_order = []
            st.rerun()
    with col_btn2:
        if st.button("🚀 生成教务处标准级 Word 试卷", type="primary"):
            if len(st.session_state.cart) == 0:
                st.warning("题库为空。")
            else:
                with st.spinner("正在智能合并题库并生成标准排版..."):
                    sections = group_questions(st.session_state.cart, section_configs, st.session_state.section_order)
                    show_answer = (doc_type == "教师解析版 (红字填入答案、打勾叉)")
                    meta_info = {"time": meta_time, "score": total_calc_score, "type": meta_type}
                    file_stream = generate_word_direct(sections, show_answer, paper_size, meta_info)
                    st.success("✅ 试卷生成完毕！")
                    st.download_button(
                        label=f"📥 下载《建筑材料试卷》_{paper_size}.docx",
                        data=file_stream,
                        file_name=f"建筑材料试卷_{'教师版' if show_answer else '学生版'}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
