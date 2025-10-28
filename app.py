# app.py
# -*- coding: utf-8 -*-
"""
Ứng dụng Streamlit tạo biên bản cuộc họp từ transcript (.docx) + Attendance (.csv/.xlsx).
- Giữ nguyên logic: validate bắt buộc, điền template, gửi email.
- Bổ sung:
    • Docling (nếu có): convert transcript .docx → Markdown (fallback python-docx)
    • Attendance .csv/.xlsx: thử Docling (nếu có/khả dụng), fallback pandas → bullets + bảng Markdown
    • Hợp nhất transcript + attendance vào prompt cho Gemini

Chạy:
    streamlit run app.py

Gợi ý requirements.txt:
    streamlit
    python-docx
    pandas
    openpyxl
    google-generativeai
    docling
"""

from __future__ import annotations
import io
import os
import re
import json
import zipfile
import ssl
import smtplib
import shutil
import tempfile
from typing import Dict, List, Optional, Tuple

import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.shared import Inches  # để sẵn nếu sau này cần chèn ảnh
import google.generativeai as genai

# =========================
# CẤU HÌNH BẢO MẬT / API
# =========================
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    APP_EMAIL      = st.secrets["APP_EMAIL"]
    APP_PASSWORD   = st.secrets["APP_PASSWORD"]
except Exception:
    st.warning("Không tìm thấy Streamlit Secrets. Đang dùng cấu hình local thử nghiệm!")
    GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY")
    APP_EMAIL      = os.getenv("APP_EMAIL", "your_email@example.com")
    APP_PASSWORD   = os.getenv("APP_PASSWORD", "your_app_or_email_password")

try:
    genai.configure(api_key=GEMINI_API_KEY)
except Exception as e:
    st.error(f"Lỗi cấu hình Gemini API: {e}. Vui lòng kiểm tra lại API Key.")

# =========================
# HẰNG SỐ & REGEX
# =========================
EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
REQUIRED_PLACEHOLDERS = ["TenCuocHop", "ThoiGianCuocHop", "DiaDiemCuocHop", "TenChuTri", "TenThuKy"]

COMMENT_RE     = re.compile(r"\{#.*?#\}")                # 1-run
COMMENT_ALL_RE = re.compile(r"\{#.*?#\}", re.DOTALL)     # đa-run
BOLD_RE        = re.compile(r"\*\*(.*?)\*\*")            # **bold**
TOKEN_RE       = re.compile(r"\{\{([^{}]+)\}\}")         # {{Key}}

# =========================
# VALIDATE BẮT BUỘC
# =========================
def validate_inputs(
    template_option: str,
    transcript_file,
    template_file,
    meeting_name: str,
    meeting_time: str,
    meeting_location: str,
    meeting_chair: str,
    meeting_secretary: str,
    recipient_email: str,
    default_template_path: str = None
) -> bool:
    """
    Trả về True nếu hợp lệ; ngược lại hiển thị thông báo đỏ và trả về False.
    """
    missing = []

    # File bắt buộc
    if not transcript_file:
        missing.append("File transcript (.docx)")

    if template_option == "Template VPI":
        if default_template_path and not os.path.exists(default_template_path):
            missing.append(f"Template mặc định không tồn tại: {default_template_path}")
    elif template_option == "Template tùy chỉnh":
        if not template_file:
            missing.append("File template tùy chỉnh (.docx)")

    # Trường bắt buộc
    if not meeting_name:
        missing.append("Tên cuộc họp")
    if not meeting_time:
        missing.append("Thời gian cuộc họp")
    if not meeting_location:
        missing.append("Địa điểm cuộc họp")
    if not meeting_chair:
        missing.append("Tên chủ trì")
    if not meeting_secretary:
        missing.append("Tên thư ký")
    if not recipient_email:
        missing.append("Email nhận kết quả")
    elif not EMAIL_RE.match(recipient_email.strip()):
        missing.append("Email nhận kết quả (không hợp lệ)")

    if missing:
        st.error("❌ **Chưa hoàn thành thông tin**:\n\n- " + "\n- ".join(missing) +
                 "\n\nVui lòng bổ sung/đính kèm đầy đủ rồi bấm lại **Tạo biên bản**.")
        return False

    return True

# =========================
# UTILITIES: WORD/Paragraph
# =========================
def _is_md_table(text: str) -> bool:
    lines = [l.strip() for l in (text or "").strip().splitlines() if l.strip()]
    return (
        len(lines) >= 2
        and "|" in lines[0]
        and set(lines[1].replace(" ", "").replace(":", "")) <= set("-|")
    )

def _parse_md_table(text: str) -> Tuple[List[str], List[List[str]]]:
    lines  = [l.strip() for l in (text or "").strip().splitlines() if l.strip()]
    header = [c.strip() for c in lines[0].split("|")]
    if header and header[0] == "":
        header = header[1:]
    if header and header[-1] == "":
        header = header[:-1]
    rows: List[List[str]] = []
    for ln in lines[2:]:  # Skip header + separator
        cols = [c.strip() for c in ln.split("|")]
        if cols and cols[0] == "":
            cols = cols[1:]
        if cols and cols[-1] == "":
            cols = cols[:-1]
        if cols:
            while len(cols) > len(header):
                cols.pop()
            while len(cols) < len(header):
                cols.append("")
            rows.append(cols)
    return header, rows

def _insert_paragraph_after(anchor_para: Paragraph, style: Optional[str] = None) -> Paragraph:
    new_p_ox = OxmlElement("w:p")
    anchor_para._p.addnext(new_p_ox)
    new_para = Paragraph(new_p_ox, anchor_para._parent)
    if style:
        try:
            new_para.style = style
        except Exception:
            pass
    return new_para

def add_formatted_text(paragraph: Paragraph, text: str, style_info: Optional[dict] = None) -> None:
    parts   = BOLD_RE.split(text or "")
    is_bold = False
    for part in parts:
        if part == "":
            is_bold = not is_bold
            continue
        lines = part.split("\n")
        for i, ln in enumerate(lines):
            if i > 0:
                paragraph.add_run().add_break()
            if ln == "":
                continue
            run = paragraph.add_run(ln)
            if style_info:
                try:
                    f = run.font
                    if style_info.get("size"):
                        f.size = style_info["size"]
                    if style_info.get("name"):
                        f.name = style_info["name"]
                    if style_info.get("bold") is not None:
                        f.bold = style_info["bold"]
                    if style_info.get("italic") is not None:
                        f.italic = style_info["italic"]
                except Exception:
                    pass
            run.bold = run.bold or is_bold
        is_bold = not is_bold

def _concat_runs(paragraph: Paragraph) -> Tuple[str, List[Tuple]]:
    meta, pos, buf = [], 0, []
    for r in paragraph.runs:
        t = r.text or ""
        start, end = pos, pos + len(t)
        meta.append((r, start, end))
        buf.append(t)
        pos = end
    return "".join(buf), meta

def _insert_table_after(paragraph: Paragraph, header: List[str], rows: List[List[str]], table_style: str = "New Table") -> None:
    if not header or not rows:
        return
    body = paragraph._parent
    tbl  = body.add_table(rows=len(rows)+1, cols=len(header))
    try:
        tbl.style = table_style
    except Exception:
        pass
    for i, h in enumerate(header):
        try:
            tbl.rows[0].cells[i].text = str(h)
        except Exception:
            pass
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, cell_val in enumerate(row):
            try:
                tbl.rows[r_idx].cells[c_idx].text = str(cell_val)
            except Exception:
                pass
    paragraph._p.addnext(tbl._tbl)

# =========================
# WORD TEMPLATE PROCESSING
# =========================
def extract_vars_and_desc(docx_file_or_buffer) -> Dict[str, str]:
    """Trích xuất {placeholder: mô tả} từ .docx (body/header/footer)."""
    xml_parts: List[str] = []
    with zipfile.ZipFile(docx_file_or_buffer) as z:
        for name in z.namelist():
            if name.startswith("word/") and name.endswith(".xml"):
                xml_parts.append(z.read(name).decode("utf8"))
    all_xml = "\n".join(xml_parts)
    texts = re.findall(r"<w:t[^>]*>(.*?)</w:t>", all_xml, flags=re.DOTALL)
    full_text = "".join(texts)
    pattern = re.compile(r"\{\{\s*([A-Za-z0-9_]+)\s*\}\}\s*\{#\s*(.*?)\s*#\}", flags=re.DOTALL)
    return dict(pattern.findall(full_text))

def replace_in_paragraph(paragraph: Paragraph, data: Dict[str, str]) -> None:
    if not paragraph.runs:
        return
    full_text, meta = _concat_runs(paragraph)
    if not full_text:
        return

    items = []
    for m in COMMENT_ALL_RE.finditer(full_text):
        items.append(("comment", m.start(), m.end(), None))
    for m in TOKEN_RE.finditer(full_text):
        key = (m.group(1) or "").strip()
        if key in data:
            items.append(("token", m.start(), m.end(), key))

    if not items:
        for r in paragraph.runs:
            if r.text and COMMENT_RE.search(r.text):
                r.text = COMMENT_RE.sub("", r.text)
        return

    items.sort(key=lambda x: x[1], reverse=True)

    bullet_queue: List[Tuple[str, str]] = []  # (text, style)
    table_queue:  List[Tuple[List[str], List[List[str]]]] = []

    for item_type, start, end, key in items:
        run_start_idx = next((i for i, (_, s, e) in enumerate(meta) if s <= start < e), None)
        run_end_idx   = next((i for i, (_, s, e) in enumerate(meta) if s <  end <= e), None)
        if run_start_idx is None or run_end_idx is None:
            continue

        run_start, s0, e0 = meta[run_start_idx]
        run_end,   s1, e1 = meta[run_end_idx]
        offset_start = start - s0
        offset_end   = end   - s1

        if item_type == "comment":
            if run_start_idx == run_end_idx:
                t = run_start.text or ""
                run_start.text = t[:offset_start] + t[offset_end:]
            else:
                run_start.text = (run_start.text or "")[:offset_start]
                for i in range(run_start_idx + 1, run_end_idx):
                    meta[i][0].text = ""
                run_end.text = (run_end.text or "")[offset_end:]
            continue

        value = data.get(key, "")

        if isinstance(value, str) and _is_md_table(value):
            try:
                header, rows = _parse_md_table(value)
                table_queue.append((header, rows))
                if run_start_idx == run_end_idx:
                    t = run_start.text or ""
                    run_start.text = t[:offset_start] + t[offset_end:]
                else:
                    run_start.text = (run_start.text or "")[:offset_start]
                    for i in range(run_start_idx + 1, run_end_idx):
                        meta[i][0].text = ""
                    run_end.text = (run_end.text or "")[offset_end:]
                continue
            except Exception:
                value = str(value)

        if isinstance(value, str) and any(line.strip().startswith(("-", "+")) for line in value.splitlines()):
            for line in value.splitlines():
                s = line.strip()
                if s.startswith("-"):
                    bullet_queue.append((s[1:].strip(), "List Bullet"))
                elif s.startswith("+"):
                    bullet_queue.append((s[1:].strip(), "List Bullet 2"))
            if run_start_idx == run_end_idx:
                t = run_start.text or ""
                run_start.text = t[:offset_start] + t[offset_end:]
            else:
                run_start.text = (run_start.text or "")[:offset_start]
                for i in range(run_start_idx + 1, run_end_idx):
                    meta[i][0].text = ""
                run_end.text = (run_end.text or "")[offset_end:]
            continue

        replacement_text = str(value)
        if run_start_idx == run_end_idx:
            t = run_start.text or ""
            run_start.text = t[:offset_start] + replacement_text + t[offset_end:]
        else:
            for i in range(run_start_idx + 1, run_end_idx):
                meta[i][0].text = ""
            start_text = (run_start.text or "")[:offset_start]
            run_start.text = start_text + replacement_text
            run_end.text = (run_end.text or "")[offset_end:]

    if bullet_queue or table_queue:
        current_para = paragraph
        for text, style in bullet_queue:
            current_para = _insert_paragraph_after(current_para, style=style)
            add_formatted_text(current_para, text)
        for header, rows in table_queue:
            try:
                _insert_table_after(current_para, header, rows)
            except Exception as e:
                print(f"Error inserting table: {e}")

def fill_template_to_buffer(template_file_or_path, data_input: Dict[str, str]) -> Optional[io.BytesIO]:
    try:
        doc = Document(template_file_or_path)
    except Exception as e:
        st.error(f"Lỗi mở template: {e}")
        return None

    for i, paragraph in enumerate(doc.paragraphs):
        try:
            replace_in_paragraph(paragraph, data_input)
        except Exception as e:
            print(f"Error processing paragraph {i}: {e}")

    for table_idx, table in enumerate(doc.tables):
        for row_idx, row in enumerate(table.rows):
            for cell_idx, cell in enumerate(row.cells):
                for para_idx, paragraph in enumerate(cell.paragraphs):
                    try:
                        replace_in_paragraph(paragraph, data_input)
                    except Exception as e:
                        print(f"Error processing table {table_idx}, row {row_idx}, cell {cell_idx}, paragraph {para_idx}: {e}")

    for section in doc.sections:
        if section.header:
            for paragraph in section.header.paragraphs:
                try:
                    replace_in_paragraph(paragraph, data_input)
                except Exception as e:
                    print(f"Error processing header paragraph: {e}")
        if section.footer:
            for paragraph in section.footer.paragraphs:
                try:
                    replace_in_paragraph(paragraph, data_input)
                except Exception as e:
                    print(f"Error processing footer paragraph: {e}")

    try:
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf
    except Exception as e:
        st.error(f"Đã xảy ra lỗi khi tạo file Word: {e}")
        return None

# =========================
# DOCLING + ATTENDANCE
# =========================
def extract_transcript_markdown(transcript_file) -> str:
    """
    Ưu tiên Docling để convert .docx → Markdown.
    Nếu lỗi/không có Docling → fallback python-docx → Markdown tối giản.
    """
    tmp_path = None
    try:
        try:
            from docling.document_converter import DocumentConverter
        except Exception:
            raise ImportError("Docling not available")

        suffix = os.path.splitext(getattr(transcript_file, "name", "") or "")[1] or ".docx"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp_path = tmp.name
            transcript_file.seek(0)
            shutil.copyfileobj(transcript_file, tmp)

        conv = DocumentConverter()
        res  = conv.convert(tmp_path)
        md   = res.document.export_markdown() if hasattr(res.document, "export_markdown") \
               else res.document.export_to_markdown()
        return (md or "").strip()

    except Exception:
        # Fallback python-docx
        try:
            transcript_file.seek(0)
        except Exception:
            pass
        try:
            doc = Document(transcript_file)
            text = "\n".join(p.text for p in doc.paragraphs)
            return ("## Transcript (fallback Docx)\n\n" + text).strip()
        except Exception as ee:
            st.error(f"Lỗi đọc transcript: {ee}")
            return ""
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try: os.remove(tmp_path)
            except Exception: pass

def _normalize(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s2 = s.strip().lower()
    rep = {
        "à":"a","á":"a","ả":"a","ã":"a","ạ":"a","ă":"a","ằ":"a","ắ":"a","ẳ":"a","ẵ":"a","ặ":"a",
        "â":"a","ầ":"a","ấ":"a","ẩ":"a","ẫ":"a","ậ":"a","è":"e","é":"e","ẻ":"e","ẽ":"e","ẹ":"e",
        "ê":"e","ề":"e","ế":"e","ể":"e","ễ":"e","ệ":"e","ì":"i","í":"i","ỉ":"i","ĩ":"i","ị":"i",
        "ò":"o","ó":"o","ỏ":"o","õ":"o","ọ":"o","ô":"o","ồ":"o","ố":"o","ổ":"o","ỗ":"o","ộ":"o",
        "ơ":"o","ờ":"o","ớ":"o","ở":"o","ỡ":"o","ợ":"o","ù":"u","ú":"u","ủ":"u","ũ":"u","ụ":"u",
        "ư":"u","ừ":"u","ứ":"u","ử":"u","ữ":"u","ự":"u","ỳ":"y","ý":"y","ỷ":"y","ỹ":"y","ỵ":"y",
        "đ":"d"
    }
    for a, b in rep.items():
        s2 = s2.replace(a, b)
    return s2

def _first_match(cols: List[str], candidates: List[str]) -> Optional[str]:
    cols_norm = {c: _normalize(c) for c in cols}
    for cand in candidates:
        for col, norm in cols_norm.items():
            if cand in norm:
                return col
    return None

def _looks_present(val) -> bool:
    if val is None:
        return True
    s = str(val).strip().lower()
    return s in {"1","x","✓","yes","y","true","present","co","có","tham du","attended"}

def attendance_df_to_struct(df: pd.DataFrame) -> Dict[str, str]:
    """Biến df attendance → bullets + bảng Markdown."""
    if df is None or df.empty:
        return {"participants_bullets":"", "participants_table_md":""}

    cols = list(df.columns)
    name_col = _first_match(cols, ["name","full name","fullname","ho va ten","ho ten","ten","họ và tên"])
    role_col = _first_match(cols, ["role","vai tro","chuc vu","title","position"])
    mail_col = _first_match(cols, ["email","mail"])
    dept_col = _first_match(cols, ["department","phong ban","don vi","unit","division"])
    att_col  = _first_match(cols, ["attendance","status","co mat","tham du","present","attended"])

    if att_col:
        df = df[df[att_col].apply(_looks_present)]

    # Bullets cấp 2
    bullets = []
    for _, r in df.iterrows():
        name = str(r.get(name_col, "")).strip()
        role = str(r.get(role_col, "")).strip()
        dept = str(r.get(dept_col, "")).strip()
        mail = str(r.get(mail_col, "")).strip()
        info = name
        tail = ", ".join([x for x in [role, dept] if x])
        if tail: info += f" — {tail}"
        if mail: info += f" ({mail})"
        if info:
            bullets.append(f"+ {info}")
    participants_bullets = "\n".join(bullets)

    # Bảng Markdown
    headers = []; rows = []
    def add_hdr(h):
        if h not in headers: headers.append(h)
    if name_col: add_hdr("Name")
    if role_col: add_hdr("Role/Title")
    if dept_col: add_hdr("Department")
    if mail_col: add_hdr("Email")
    if headers:
        for _, r in df.iterrows():
            row=[]
            if name_col: row.append(str(r.get(name_col,"")).strip())
            if role_col: row.append(str(r.get(role_col,"")).strip())
            if dept_col: row.append(str(r.get(dept_col,"")).strip())
            if mail_col: row.append(str(r.get(mail_col,"")).strip())
            rows.append(row)
        sep = "|" + "|".join(["---"]*len(headers)) + "|"
        table_md = "|" + "|".join(headers) + "|\n" + sep + "\n" + "\n".join(["|" + "|".join(r) + "|" for r in rows])
    else:
        table_md = ""

    return {
        "participants_bullets": participants_bullets,
        "participants_table_md": table_md,
    }

def attendance_via_docling(attendance_file) -> Optional[str]:
    """Thử convert attendance bằng Docling → Markdown (nếu hỗ trợ). Không đảm bảo cho CSV/XLSX."""
    tmp_path = None
    try:
        try:
            from docling.document_converter import DocumentConverter
        except Exception:
            return None
        suffix = os.path.splitext(getattr(attendance_file, "name","") or "")[1] or ".csv"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp_path = tmp.name
            attendance_file.seek(0)
            shutil.copyfileobj(attendance_file, tmp)
        conv = DocumentConverter()
        res  = conv.convert(tmp_path)
        md = res.document.export_markdown() if hasattr(res.document, "export_markdown") \
             else res.document.export_to_markdown()
        return (md or "").strip() or None
    except Exception:
        return None
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try: os.remove(tmp_path)
            except Exception: pass

def attendance_file_to_markdown(attendance_file) -> str:
    """
    Pipeline đọc attendance:
    1) Thử Docling → Markdown
    2) Fallback pandas → bullets + bảng Markdown
    """
    if not attendance_file:
        return ""

    # Docling trước
    md = attendance_via_docling(attendance_file)
    if md:
        return ("## Attendance (Docling)\n\n" + md).strip()

    # Fallback pandas
    try:
        attendance_file.seek(0)
    except Exception:
        pass

    name = getattr(attendance_file, "name", "") or ""
    ext = os.path.splitext(name.lower())[1]
    try:
        if ext in (".xlsx", ".xls"):
            df = pd.read_excel(attendance_file)
        else:
            last_err = None
            for enc in ["utf-8","utf-8-sig","cp1258","latin1"]:
                try:
                    attendance_file.seek(0)
                except Exception:
                    pass
                try:
                    df = pd.read_csv(attendance_file, encoding=enc)
                    break
                except Exception as e:
                    last_err = e
            else:
                raise last_err or RuntimeError("Không đọc được CSV.")
    except Exception as e:
        st.warning(f"Không thể đọc attendance bằng pandas: {e}")
        return ""

    struct = attendance_df_to_struct(df)
    bullets = struct.get("participants_bullets","").strip()
    tablemd = struct.get("participants_table_md","").strip()
    mk = ["## Attendance (normalized)"]
    if bullets:
        mk.append("\n### Participants (bullets)\n" + bullets)
    if tablemd:
        mk.append("\n### Participants (table)\n" + tablemd)
    return "\n".join(mk).strip()

# =========================
# LLM (Gemini)
# =========================
def call_gemini_model(transcript_markdown: str, placeholders: Dict[str, str], attendance_markdown: str = "") -> Optional[Dict[str, str]]:
    model = genai.GenerativeModel("gemini-2.5-pro")

    unified_md = f"""
# SOURCE PACKET
## 1) TRANSCRIPT (Markdown)
{transcript_markdown}

## 2) ATTENDANCE (Markdown)
{attendance_markdown or '*(Không có file attendance được cung cấp)*'}
""".strip()

    Prompt_word = f"""
# Vai trò
Bạn là trợ lý AI chuyên nghiệp, có nhiệm vụ trích xuất thông tin quan trọng từ *SOURCE PACKET* bên dưới để tạo nội dung cho biên bản cuộc họp (tiếng Việt, văn phong trang trọng).

# SOURCE PACKET (Markdown)
{unified_md}

# Placeholders (dict: key = tên trường, value = mô tả/định dạng):
```json
{json.dumps(placeholders, ensure_ascii=False)}
```

# Yêu cầu xuất
- **Chỉ trả về 1 JSON hợp lệ duy nhất**.
- **Keys trùng 100%** với placeholders (không thêm/bớt/đổi kiểu chữ).
- **Mọi value là chuỗi**.
- Tuân thủ **định dạng trong mô tả**: bullet 1 "- ", bullet 2 "+ ", bảng Markdown...
- Nếu thiếu thông tin → điền đúng chuỗi **"Chưa có thông tin"**.

# Kết quả
Trả về 1 chuỗi JSON duy nhất, không kèm giải thích.
"""

    try:
        response = model.generate_content(
            contents=Prompt_word,
            generation_config={"response_mime_type": "application/json"}
        )
        if response and hasattr(response, "text"):
            raw = response.text.strip()
            if raw.startswith("```"):
                raw = raw.split("```")[1].strip("json\n")
            return json.loads(raw)
        else:
            st.error("Phản hồi từ Gemini API bị thiếu hoặc không hợp lệ.")
            return None
    except Exception as e:
        st.error(f"Lỗi khi gọi Gemini API: {e}")
        return None

# =========================
# EMAIL
# =========================
def send_email_with_attachment(recipient_email: str, attachment_buffer: io.BytesIO, filename: str = "Bien_ban_cuoc_hop.docx") -> bool:
    SMTP_SERVER = "smtp.office365.com"
    SMTP_PORT = 587
    from email.message import EmailMessage

    msg = EmailMessage()
    msg["Subject"] = "Biên bản cuộc họp đã được tạo tự động"
    msg["From"] = APP_EMAIL
    msg["To"] = recipient_email
    msg.set_content(
        "Chào bạn,\n\nBiên bản cuộc họp đã được tạo thành công.\nVui lòng xem trong file đính kèm.\n\nTrân trọng,\nCông cụ tạo biên bản tự động."
    )
    msg.add_attachment(
        attachment_buffer.getvalue(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=filename,
    )

    try:
        ctx = ssl.create_default_context()
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
            s.starttls(context=ctx)
            s.login(APP_EMAIL, APP_PASSWORD)
            s.send_message(msg)
        return True
    except Exception as e:
        st.error(f"Lỗi khi gửi email: {e}. Vui lòng kiểm tra lại cấu hình email và mật khẩu ứng dụng.")
        return False

# =========================
# HELPERS (IO)
# =========================
def ensure_template_path(default_filename: str) -> Optional[str]:
    """Trả template path nếu tồn tại, ngược lại cảnh báo người dùng chọn custom."""
    if os.path.exists(default_filename):
        return default_filename
    st.error(f"Không tìm thấy template mặc định: {default_filename}. Hãy chọn 'Template tùy chỉnh' và tải file lên.")
    return None

def to_bytesio(file_or_path):
    """Đưa template (path hoặc UploadedFile) về BytesIO để dùng lặp nhiều lần."""
    if isinstance(file_or_path, (str, os.PathLike)):
        with open(file_or_path, "rb") as f:
            return io.BytesIO(f.read())
    else:
        try:
            file_or_path.seek(0)
        except Exception:
            pass
        return io.BytesIO(file_or_path.read())

# =========================
# STREAMLIT UI
# =========================
st.set_page_config(layout="wide", page_title="Công cụ tạo Biên bản cuộc họp")
st.title("🛠️ Công cụ tạo biên bản cuộc họp tự động")

with st.sidebar:
    st.info("📝 **Hướng dẫn sử dụng**")
    st.markdown("""
1. **Tải transcript (.docx)** và *(tuỳ chọn)* **attendance (.csv/.xlsx)**.
2. **Chọn Template:** "Template VPI" hoặc "Template tuỳ chỉnh (.docx)".
3. **Điền thông tin bắt buộc** (Tên họp, Thời gian, Địa điểm, Chủ trì, Thư ký, Email).
4. Nhấn **Tạo biên bản**.
    """)
    st.info("🧩 **Tạo template** — dùng {{Ten_bien}}{# mô tả #} cho các trường trích xuất. Bảng dùng Markdown, bullet: '- ' và '+ '.")

st.subheader("1. Nhập thông tin đầu vào")
col_in_1, col_in_2 = st.columns(2)
with col_in_1:
    transcript_file = st.file_uploader("Tải transcript (.docx) — bắt buộc", type=["docx"])
with col_in_2:
    attendance_file = st.file_uploader("Attendance (.csv/.xlsx) — tuỳ chọn", type=["csv","xlsx","xls"])

st.subheader("2. Lựa chọn Template")
template_option = st.selectbox(
    "Bạn muốn sử dụng loại template nào?",
    ("Template VPI", "Template tùy chỉnh"),
    help="Chọn 'Template VPI' để dùng mẫu có sẵn hoặc 'Template tùy chỉnh' để tải lên file của riêng bạn."
)
template_file = None
if template_option == "Template tùy chỉnh":
    template_file = st.file_uploader("Tải lên file template .docx của bạn", type=["docx"])

st.subheader("3. Thông tin cơ bản (bắt buộc)")
if template_option == "Template tùy chỉnh":
    st.info(
        "🔔 **Template tùy chỉnh** cần có các biến sau (không kèm mô tả `{# ... #}`): "
        "`{{TenCuocHop}}`, `{{ThoiGianCuocHop}}`, `{{DiaDiemCuocHop}}`, `{{TenChuTri}}`, `{{TenThuKy}}`."
    )
else:
    st.caption("Các trường bắt buộc đã có sẵn trong Template VPI (sẽ được ghi đè bằng input bạn nhập).")

col1, col2 = st.columns(2)
with col1:
    meeting_name      = st.text_input("Tên cuộc họp")
    meeting_time      = st.text_input("Thời gian cuộc họp (VD: 10/9/2025)")
    meeting_location  = st.text_input("Địa điểm cuộc họp")
with col2:
    meeting_chair     = st.text_input("Tên chủ trì")
    meeting_secretary = st.text_input("Tên thư ký")

recipient_email = st.text_input("4. Email nhận kết quả của bạn (bắt buộc)")

# Nút chạy
if st.button("🚀 Tạo biên bản", type="primary"):
    default_path = "2025.VPI_BB hop 2025 1.docx"

    # 1) Kiểm tra bắt buộc
    if not validate_inputs(
        template_option=template_option,
        transcript_file=transcript_file,
        template_file=template_file,
        meeting_name=meeting_name,
        meeting_time=meeting_time,
        meeting_location=meeting_location,
        meeting_chair=meeting_chair,
        meeting_secretary=meeting_secretary,
        recipient_email=recipient_email,
        default_template_path=default_path
    ):
        st.stop()

    # 2) Xác định template
    template_source = None
    if template_option == "Template VPI":
        template_source = ensure_template_path(default_path)
        if not template_source:
            st.stop()
    else:
        template_source = template_file

    # 2.1) Chuẩn về BytesIO (để dùng nhiều lần)
    template_stream = to_bytesio(template_source)

    with st.spinner("⏳ Hệ thống đang xử lý..."):
        try:
            st.info("1/5 - Đọc transcript (.docx) bằng Docling (fallback python-docx)...")
            transcript_markdown = extract_transcript_markdown(transcript_file)

            st.info("2/5 - Trích placeholders từ template...")
            p_stream_for_extract = io.BytesIO(template_stream.getvalue())
            placeholders = extract_vars_and_desc(p_stream_for_extract)

            # Kiểm tra placeholders bắt buộc với template tùy chỉnh
            missing_ph = []
            if template_option == "Template tùy chỉnh":
                try:
                    p_stream_for_scan = io.BytesIO(template_stream.getvalue())
                    tdoc = Document(p_stream_for_scan)
                    ttext = "\n".join([p.text for p in tdoc.paragraphs])
                    for ph in REQUIRED_PLACEHOLDERS:
                        if f"{{{{{ph}}}}}" not in ttext:
                            missing_ph.append(ph)
                except Exception:
                    pass
                if missing_ph:
                    st.error("❌ **Template tùy chỉnh thiếu các biến bắt buộc**: " + ", ".join(missing_ph) +
                             ".\nVui lòng cập nhật template rồi chạy lại.")
                    st.stop()

            st.info("3/5 - Chuẩn hoá attendance (nếu có)...")
            attendance_markdown = ""
            if attendance_file is not None:
                attendance_markdown = attendance_file_to_markdown(attendance_file)

            st.info("4/5 - Gọi AI để trích xuất nội dung (hợp nhất transcript + attendance)...")
            llm_result = call_gemini_model(transcript_markdown, placeholders, attendance_markdown)

            if llm_result:
                # Ghi đè các trường bắt buộc bằng input tay
                manual_inputs = {
                    'TenCuocHop':       meeting_name,
                    'ThoiGianCuocHop':  meeting_time,
                    'DiaDiemCuocHop':   meeting_location,
                    'TenChuTri':        meeting_chair,
                    'TenThuKy':         meeting_secretary
                }
                llm_result.update(manual_inputs)

                st.info("5/5 - Điền template và tạo file Word...")
                p_stream_for_fill = io.BytesIO(template_stream.getvalue())
                docx_buffer = fill_template_to_buffer(p_stream_for_fill, llm_result)
                if docx_buffer:
                    st.success("✅ Tạo biên bản thành công!")
                    st.download_button(
                        "⬇️ Tải về biên bản",
                        data=docx_buffer,
                        file_name="Bien_ban_cuoc_hop.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    if recipient_email:
                        if send_email_with_attachment(recipient_email, docx_buffer, filename="Bien_ban_cuoc_hop.docx"):
                            st.success("✉️ Đã gửi biên bản tới email của bạn.")
                else:
                    st.error("Không thể tạo file Word. Vui lòng kiểm tra lại file template.")
            else:
                st.error("Không thể lấy kết quả từ AI. Vui lòng thử lại.")
        except Exception as e:
            st.error(f"Đã xảy ra lỗi: {e}")
