# app_refactored.py
# -*- coding: utf-8 -*-
"""
Ứng dụng Streamlit tạo biên bản cuộc họp từ transcript (.docx) + CSV thành viên tham dự.
- Refactor mã gốc thành các hàm rõ ràng, dễ test, dễ tái sử dụng.
- Bổ sung đầu vào .csv để kết hợp với transcript trước khi gửi sang AI.

Yêu cầu thư viện (requirements.txt):
    streamlit
    pandas
    python-docx
    google-generativeai
    openpyxl   # (đọc Excel nếu cần trong tương lai)

Cách chạy (local):
    streamlit run app_refactored.py

Cấu trúc logic chính:
  1) Upload transcript .docx + CSV thành viên + chọn template .docx
  2) Trích placeholders từ template
  3) Đọc transcript + CSV → tạo participants_hint
  4) Gọi AI tạo JSON theo placeholders (ưu tiên dùng CSV cho trường liên quan thành viên)
  5) Ghi đè một số trường thủ công (Tên cuộc họp, Chủ trì, Thư ký...)
  6) Điền template → .docx → cho tải xuống và/hoặc gửi email
"""

from __future__ import annotations
import io
import os
import re
import json
import zipfile
import ssl
import smtplib
from typing import Dict, List, Optional, Tuple

import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.shared import Inches  # noqa: F401 (để sẵn nếu sau này cần chèn ảnh)
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
# HẰNG SỐ & REGEX PHỤ TRỢ
# =========================
COMMENT_RE     = re.compile(r"\{#.*?#\}")                # 1-run
COMMENT_ALL_RE = re.compile(r"\{#.*?#\}", re.DOTALL)     # đa-run
BOLD_RE        = re.compile(r"\*\*(.*?)\*\*")          # **bold**
TOKEN_RE       = re.compile(r"\{\{([^{}]+)\}\}")       # {{Key}}

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

        # token {{key}}
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
# CSV PARSER: THÀNH VIÊN
# =========================

def _normalize(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s2 = s.strip().lower()
    # bỏ dấu tiếng Việt đơn giản
    rep = {
        "à": "a", "á": "a", "ả": "a", "ã": "a", "ạ": "a",
        "ă": "a", "ằ": "a", "ắ": "a", "ẳ": "a", "ẵ": "a", "ặ": "a",
        "â": "a", "ầ": "a", "ấ": "a", "ẩ": "a", "ẫ": "a", "ậ": "a",
        "è": "e", "é": "e", "ẻ": "e", "ẽ": "e", "ẹ": "e",
        "ê": "e", "ề": "e", "ế": "e", "ể": "e", "ễ": "e", "ệ": "e",
        "ì": "i", "í": "i", "ỉ": "i", "ĩ": "i", "ị": "i",
        "ò": "o", "ó": "o", "ỏ": "o", "õ": "o", "ọ": "o",
        "ô": "o", "ồ": "o", "ố": "o", "ổ": "o", "ỗ": "o", "ộ": "o",
        "ơ": "o", "ờ": "o", "ớ": "o", "ở": "o", "ỡ": "o", "ợ": "o",
        "ù": "u", "ú": "u", "ủ": "u", "ũ": "u", "ụ": "u",
        "ư": "u", "ừ": "u", "ứ": "u", "ử": "u", "ữ": "u", "ự": "u",
        "ỳ": "y", "ý": "y", "ỷ": "y", "ỹ": "y", "ỵ": "y",
        "đ": "d",
    }
    for a, b in rep.items():
        s2 = s2.replace(a, b)
    return s2


def _first_match(cols: List[str], candidates: List[str]) -> Optional[str]:
    cols_norm = {c: _normalize(c) for c in cols}
    for c in candidates:
        for col, norm in cols_norm.items():
            if c in norm:
                return col
    return None


def _looks_present(val) -> bool:
    if val is None:
        return True  # nếu không có cột thì mặc định có mặt
    s = str(val).strip().lower()
    return s in {"1", "x", "✓", "yes", "y", "true", "present", "co", "có", "tham du", "attended"}


def parse_attendance_csv(file) -> Dict[str, str]:
    """Đọc CSV và trả về:
    {
      'participants_bullets': "+ Name — Chức vụ, Đơn vị (email)\n+ ...",
      'participants_table_md': "|Name|Title|Dept|Email|\n|---|---|---|---|\n|...|...|...|...|"
    }
    """
    df = pd.read_csv(file)
    if df.empty:
        return {"participants_bullets": "", "participants_table_md": ""}

    cols = list(df.columns)
    name_col = _first_match(cols, ["name", "full name", "fullname", "ho va ten", "ho ten", "ten", "hova ten", "ho-va-ten", "hvt", "họ và tên"])
    dept_col = _first_match(cols, ["don vi", "phong ban", "department", "unit", "division"])
    title_col= _first_match(cols, ["chuc vu", "title", "position", "role"])
    mail_col = _first_match(cols, ["email", "mail"])
    att_col  = _first_match(cols, ["attendance", "status", "co mat", "tham du", "present", "attended"])

    # Lọc hàng có mặt (nếu có cột attendance)
    if att_col:
        df = df[df[att_col].apply(_looks_present)]

    # Tạo bullets cấp 2 theo yêu cầu template VPI
    bullet_lines: List[str] = []
    for _, r in df.iterrows():
        parts = []
        name = str(r.get(name_col, "")).strip()
        if name:
            parts.append(name)
        title = str(r.get(title_col, "")).strip()
        dept  = str(r.get(dept_col,  "")).strip()
        email = str(r.get(mail_col,  "")).strip()
        tail_bits = []
        if title:
            tail_bits.append(title)
        if dept:
            tail_bits.append(dept)
        tail = ", ".join(tail_bits)
        shown = name
        if tail:
            shown += f" — {tail}"
        if email:
            shown += f" ({email})"
        if shown:
            bullet_lines.append(f"+ {shown}")

    participants_bullets = "\n".join(bullet_lines)

    # Tạo bảng markdown (dùng các cột còn lại nếu có)
    headers = []
    rows = []
    def add_hdr(h):
        if h not in headers:
            headers.append(h)

    if name_col: add_hdr("Name")
    if title_col: add_hdr("Title/Position")
    if dept_col: add_hdr("Department")
    if mail_col: add_hdr("Email")

    if headers:
        for _, r in df.iterrows():
            row = []
            if name_col: row.append(str(r.get(name_col, "")).strip())
            if title_col: row.append(str(r.get(title_col, "")).strip())
            if dept_col: row.append(str(r.get(dept_col, "")).strip())
            if mail_col: row.append(str(r.get(mail_col, "")).strip())
            rows.append(row)
        # markdown table
        sep = "|" + "|".join(["---" for _ in headers]) + "|"
        participants_table_md = "|" + "|".join(headers) + "|\n" + sep + "\n" + "\n".join(["|" + "|".join(r) + "|" for r in rows])
    else:
        participants_table_md = ""

    return {
        "participants_bullets": participants_bullets,
        "participants_table_md": participants_table_md,
    }


# =========================
# HỖ TRỢ ĐỌC CSV/XLSX
# =========================

def read_attendance_to_df(file) -> pd.DataFrame:
    """Cố gắng đọc file attendance dưới dạng Excel hoặc CSV.
    - Ưu tiên detect theo phần mở rộng file.name
    - Nếu là CSV: thử nhiều encoding.
    """
    name = getattr(file, "name", "") or ""
    ext = os.path.splitext(name.lower())[1]

    # Nếu là Excel
    if ext in (".xlsx", ".xls"):
        try:
            return pd.read_excel(file)
        except Exception:
            # Nếu đọc Excel thất bại, thử quay lại đầu và đọc dạng CSV (edge case export sai MIME)
            try:
                file.seek(0)
            except Exception:
                pass

    # Thử đọc CSV với nhiều encoding
    encodings = ["utf-8", "utf-8-sig", "cp1258", "latin1"]
    last_err = None
    for enc in encodings:
        try:
            file.seek(0)
        except Exception:
            pass
        try:
            return pd.read_csv(file, encoding=enc)
        except Exception as e:
            last_err = e
            continue

    # Thử chốt: nếu là Excel thật sự nhưng không có phần mở rộng
    try:
        file.seek(0)
    except Exception:
        pass
    try:
        return pd.read_excel(file)
    except Exception as e:
        raise RuntimeError(f"Không thể đọc file Attendance (CSV/Excel). Lỗi cuối: {last_err or e}")


def _attendance_df_to_struct(df: pd.DataFrame) -> Dict[str, str]:
    if df is None or df.empty:
        return {"participants_bullets": "", "participants_table_md": ""}

    cols = list(df.columns)
    name_col = _first_match(cols, ["name", "full name", "fullname", "ho va ten", "ho ten", "ten", "hova ten", "ho-va-ten", "hvt", "họ và tên"])
    dept_col = _first_match(cols, ["don vi", "phong ban", "department", "unit", "division"])
    title_col= _first_match(cols, ["chuc vu", "title", "position", "role"])
    mail_col = _first_match(cols, ["email", "mail"])
    att_col  = _first_match(cols, ["attendance", "status", "co mat", "tham du", "present", "attended"])

    if att_col:
        df = df[df[att_col].apply(_looks_present)]

    bullet_lines: List[str] = []
    for _, r in df.iterrows():
        parts = []
        name = str(r.get(name_col, "")).strip()
        if name:
            parts.append(name)
        title = str(r.get(title_col, "")).strip()
        dept  = str(r.get(dept_col,  "")).strip()
        email = str(r.get(mail_col,  "")).strip()
        tail_bits = []
        if title:
            tail_bits.append(title)
        if dept:
            tail_bits.append(dept)
        tail = ", ".join(tail_bits)
        shown = name
        if tail:
            shown += f" — {tail}"
        if email:
            shown += f" ({email})"
        if shown:
            bullet_lines.append(f"+ {shown}")

    participants_bullets = "".join(bullet_lines)

    headers = []
    rows = []
    def add_hdr(h):
        if h not in headers:
            headers.append(h)

    if name_col: add_hdr("Name")
    if title_col: add_hdr("Title/Position")
    if dept_col: add_hdr("Department")
    if mail_col: add_hdr("Email")

    if headers:
        for _, r in df.iterrows():
            row = []
            if name_col: row.append(str(r.get(name_col, "")).strip())
            if title_col: row.append(str(r.get(title_col, "")).strip())
            if dept_col: row.append(str(r.get(dept_col, "")).strip())
            if mail_col: row.append(str(r.get(mail_col, "")).strip())
            rows.append(row)
        sep = "|" + "|".join(["---" for _ in headers]) + "|"
        participants_table_md = "|" + "|".join(headers) + "|" + sep + "" + "".join(["|" + "|".join(r) + "|" for r in rows])
    else:
        participants_table_md = ""

    return {
        "participants_bullets": participants_bullets,
        "participants_table_md": participants_table_md,
    }


def parse_attendance_any(file) -> Dict[str, str]:
    """API hợp nhất: nhận file CSV/XLSX và trả về cấu trúc bullets + bảng markdown."""
    df = read_attendance_to_df(file)
    return _attendance_df_to_struct(df)

# =========================
# LLM CALL (Gemini)
# =========================

def call_gemini_model(transcript_content: str, placeholders: Dict[str, str], participants_hint: Dict[str, str] | None = None) -> Optional[Dict[str, str]]:
    model = genai.GenerativeModel("gemini-2.5-pro")

    # Chuẩn bị phần dữ liệu CSV cho prompt
    participants_block = ""
    if participants_hint:
        blt = participants_hint.get("participants_bullets", "").strip()
        tbl = participants_hint.get("participants_table_md", "").strip()
        participants_block = f"""
# Dữ liệu CSV thành viên (ưu tiên sử dụng khi điền các trường liên quan người tham dự)
- **Bullet cấp 2 (ưu tiên cho {{ThanhPhanThamGia}} nếu có trong placeholders):**\n{blt}
- **Bảng Markdown (nếu cần):**\n{tbl}
""".strip()

    # Prompt (kế thừa & mở rộng từ app gốc)
    Prompt_word = f"""
# Vai trò
Bạn là trợ lý AI chuyên nghiệp, nhiệm vụ: trích xuất/thể hiện nội dung cho biên bản cuộc họp từ transcript **và** dữ liệu CSV người tham dự (nếu có), đảm bảo chính xác và trình bày chuẩn mực.

# Đầu vào
1) **Bản ghi cuộc họp (transcript):** ```{transcript_content}```
2) **Danh sách placeholders cần điền** (dict: key = tên trường, value = mô tả/định dạng yêu cầu): ```{json.dumps(placeholders, ensure_ascii=False)}```
3) **Dữ liệu CSV về thành viên** (nếu có):
{participants_block}

# Yêu cầu quan trọng
- **Luôn trả về tiếng Việt**.
- **Chỉ trả về đúng một đối tượng JSON**: keys **trùng 100%** tên placeholders; values **chỉ là chuỗi** (string). **Không** thêm/bớt key, không lồng cấu trúc.
- **Tuân thủ chặt chẽ định dạng** ghi trong mô tả của từng placeholder (bullet 1: bắt đầu bằng "- ", bullet 2: bắt đầu bằng "+ ", bảng: Markdown...).
- **Ưu tiên sử dụng dữ liệu CSV** để điền các trường về **thành phần tham gia**, **vai trò/phụ trách**. Nếu transcript cũng có thông tin, **kết hợp hợp lý**.
- Nếu thiếu thông tin: ghi đúng chuỗi **"Chưa có thông tin"**.

# Kết quả
- Xuất **một chuỗi JSON hợp lệ duy nhất** theo đúng quy tắc trên.
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
# EMAIL SENDER
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
# HELPERS (IO/UI)
# =========================

def load_transcript_docx(file) -> str:
    """Đọc toàn bộ text từ .docx transcript."""
    try:
        doc = Document(file)
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        st.error(f"Lỗi đọc transcript .docx: {e}")
        return ""


def ensure_template_path(default_filename: str) -> Optional[str]:
    """Trả template path nếu tồn tại, ngược lại cảnh báo người dùng chọn custom."""
    if os.path.exists(default_filename):
        return default_filename
    st.error(f"Không tìm thấy template mặc định: {default_filename}. Hãy chọn 'Template tùy chỉnh' và tải file lên.")
    return None


# =========================
# STREAMLIT UI
# =========================

st.set_page_config(layout="wide", page_title="Công cụ tạo Biên bản cuộc họp (refactor)")
st.title("🛠️ Công cụ tạo biên bản cuộc họp tự động — Bản refactor")

with st.sidebar:
    st.info("**Hướng dẫn nhanh**")
    st.markdown(
        """
1) Tải **transcript (.docx)** và **CSV thành viên**
2) Chọn **Template VPI** hoặc **Template tùy chỉnh (.docx)**
3) Điền vài trường tay (nếu muốn)
4) Nhấn **Tạo biên bản**
        """
    )
    st.caption("Yêu cầu thư viện đã có trong requirements.txt của dự án.")

st.subheader("1) Nhập dữ liệu đầu vào")
colA, colB = st.columns(2)
with colA:
    transcript_file = st.file_uploader("Tải transcript (.docx)", type=["docx"], key="transcript")
with colB:
    csv_file = st.file_uploader("Tải CSV/Excel thành viên (Attendance)", type=["csv", "xlsx", "xls"], key="csv")

st.subheader("2) Lựa chọn Template")
template_option = st.selectbox(
    "Bạn muốn sử dụng loại template nào?",
    ("Template VPI", "Template tùy chỉnh"),
)

template_file = None
if template_option == "Template tùy chỉnh":
    template_file = st.file_uploader("Tải file template .docx của bạn", type=["docx"], key="tpl")

st.subheader("3) Thông tin cơ bản (ghi đè kết quả AI nếu nhập)")
col1, col2 = st.columns(2)
with col1:
    meeting_name      = st.text_input("Tên cuộc họp")
    meeting_time      = st.text_input("Thời gian cuộc họp (VD: 10/9/2025)")
    meeting_location  = st.text_input("Địa điểm cuộc họp")
with col2:
    meeting_chair     = st.text_input("Tên chủ trì")
    meeting_secretary = st.text_input("Tên thư ký")

recipient_email = st.text_input("4) Email nhận kết quả (tùy chọn)")

if st.button("🚀 Tạo biên bản", type="primary"):
    if not transcript_file:
        st.warning("Vui lòng tải lên file transcript .docx")
    else:
        # 1) Chọn template
        template_to_use = None
        if template_option == "Template VPI":
            # Giữ tên template mặc định y như repo gốc để tương thích
            default_path = "2025.VPI_BB hop 2025 1.docx"
            template_to_use = ensure_template_path(default_path)
        else:
            template_to_use = template_file

        if not template_to_use:
            st.stop()

        with st.spinner("⏳ Đang xử lý..."):
            try:
                st.info("1/5 - Đọc transcript .docx")
                transcript_content = load_transcript_docx(transcript_file)

                st.info("2/5 - Trích placeholders từ template")
                placeholders = extract_vars_and_desc(template_to_use)

                st.info("3/5 - Phân tích CSV thành viên")
                participants_hint = {"participants_bullets": "", "participants_table_md": ""}
if csv_file is not None:
    try:
        participants_hint = parse_attendance_any(csv_file)
    except Exception as e:
        st.warning(f"Không đọc được CSV/Excel: {e}")

                st.info("4/5 - Gọi AI tạo JSON theo placeholders (kết hợp transcript + CSV)")
                llm_result = call_gemini_model(transcript_content, placeholders, participants_hint)

                if llm_result:
                    # Ghi đè các input tay (nếu nhập)
                    manual_inputs = {
                        'TenCuocHop':       meeting_name,
                        'ThoiGianCuocHop':  meeting_time,
                        'DiaDiemCuocHop':   meeting_location,
                        'TenChuTri':        meeting_chair,
                        'TenThuKy':         meeting_secretary,
                    }
                    for k, v in manual_inputs.items():
                        if v and k in llm_result:
                            llm_result[k] = v

                    # Ưu tiên CSV cho thành phần tham gia nếu placeholder tồn tại
                    if 'ThanhPhanThamGia' in llm_result and participants_hint.get("participants_bullets"):
                        llm_result['ThanhPhanThamGia'] = participants_hint['participants_bullets']

                    st.info("5/5 - Điền template và tạo file Word")
                    docx_buffer = fill_template_to_buffer(template_to_use, llm_result)
                    if docx_buffer:
                        st.success("✅ Tạo biên bản thành công!")
                        st.download_button(
                            "⬇️ Tải về biên bản",
                            data=docx_buffer,
                            file_name="Bien_ban_cuoc_hop.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )
                        if recipient_email:
                            ok = send_email_with_attachment(recipient_email, docx_buffer)
                            if ok:
                                st.success("✉️ Đã gửi biên bản tới email của bạn.")
                    else:
                        st.error("Không thể tạo file Word. Kiểm tra lại template hoặc dữ liệu đầu vào.")
                else:
                    st.error("AI không trả về kết quả hợp lệ. Vui lòng thử lại.")
            except Exception as e:
                st.error(f"Đã xảy ra lỗi: {e}")
