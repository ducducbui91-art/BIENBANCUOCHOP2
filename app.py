# app.py
# -*- coding: utf-8 -*-
"""
Ứng dụng Streamlit tạo biên bản cuộc họp từ transcript (.docx) + CSV (docling).
- Đọc transcript Word như cũ bằng python-docx (KHÔNG thay đổi).
- THÊM: Nhận thêm file CSV danh sách/thông tin người tham dự và phân tích bằng docling.
  (Nếu docling không có/hoặc lỗi -> tự động fallback sang pandas.read_csv để không bị nghẽn).
- Kết hợp transcript + CSV (đổi ra bullets/bảng/timeline) đưa vào prompt cho AI (Gemini).
- Giữ nguyên phong cách/luồng UI/validate của code trước; CHỈ bổ sung logic CSV (docling).

Yêu cầu cài đặt:
  pip install streamlit python-docx pandas openpyxl google-generativeai docling

Chạy local:
  streamlit run app.py

Lưu ý secrets khi deploy:
  st.secrets["GEMINI_API_KEY"], st.secrets["APP_EMAIL"], st.secrets["APP_PASSWORD"]
(Chạy local có thể dùng biến môi trường GEMINI_API_KEY, APP_EMAIL, APP_PASSWORD).
"""

from __future__ import annotations
import io
import os
import re
import json
import zipfile
from typing import Dict, List, Optional, Tuple

import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.shared import Inches  # noqa: F401
import smtplib, ssl
from email.message import EmailMessage
import google.generativeai as genai

# =========================
# CẤU HÌNH BẢO MẬT / API
# =========================
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    APP_EMAIL      = st.secrets["APP_EMAIL"]
    APP_PASSWORD   = st.secrets["APP_PASSWORD"]
except Exception:
    st.warning("Không tìm thấy Streamlit Secrets. Đang sử dụng cấu hình local thử nghiệm.")
    GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY")
    APP_EMAIL      = os.getenv("APP_EMAIL", "your_email@example.com")
    APP_PASSWORD   = os.getenv("APP_PASSWORD", "your_app_or_email_password")

try:
    genai.configure(api_key=GEMINI_API_KEY)
except Exception as e:
    st.error(f"Lỗi cấu hình Gemini API: {e}. Kiểm tra lại API key.")

# =========================
# VALIDATION CƠ BẢN (giữ nguyên tinh thần code cũ)
# =========================
EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
REQUIRED_PLACEHOLDERS = ["TenCuocHop", "ThoiGianCuocHop", "DiaDiemCuocHop", "TenChuTri", "TenThuKy"]

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
        st.error(
            "❌ **Chưa hoàn thành thông tin**:\n\n- " + "\n- ".join(missing) +
            "\n\nVui lòng bổ sung/đính kèm đầy đủ rồi bấm lại **Tạo biên bản**."
        )
        return False

    return True

# =========================
# REGEX & WORD HELPERS
# =========================
COMMENT_RE     = re.compile(r"\{#.*?#\}")                 # 1-run
COMMENT_ALL_RE = re.compile(r"\{#.*?#\}", re.DOTALL)      # đa-run
BOLD_RE        = re.compile(r"\*\*(.*?)\*\*")             # **bold**
TOKEN_RE       = re.compile(r"\{\{([^{}]+)\}\}")          # {{Key}}

def _is_md_table(text: str) -> bool:
    lines = [l.strip() for l in (text or "").strip().splitlines() if l.strip()]
    return (
        len(lines) >= 2 and "|" in lines[0] and
        set(lines[1].replace(" ", "").replace(":", "")) <= set("-|")
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

def _insert_paragraph_after(anchor_para: Paragraph, style=None) -> Paragraph:
    new_p_ox = OxmlElement("w:p")
    anchor_para._p.addnext(new_p_ox)
    new_para = Paragraph(new_p_ox, anchor_para._parent)
    if style:
        try:
            new_para.style = style
        except Exception:
            pass
    return new_para

def add_formatted_text(paragraph: Paragraph, text: str, style_info=None):
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
                    if style_info.get("size"):   f.size = style_info["size"]
                    if style_info.get("name"):   f.name = style_info["name"]
                    if style_info.get("bold") is not None:   f.bold = style_info["bold"]
                    if style_info.get("italic") is not None: f.italic = style_info["italic"]
                except Exception:
                    pass
            run.bold = run.bold or is_bold
        is_bold = not is_bold

def _concat_runs(paragraph: Paragraph):
    meta, pos, buf = [], 0, []
    for r in paragraph.runs:
        t = r.text or ""
        start, end = pos, pos + len(t)
        meta.append((r, start, end))
        buf.append(t)
        pos = end
    return "".join(buf), meta

def _insert_table_after(paragraph: Paragraph, header, rows, table_style="New Table"):
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

def extract_vars_and_desc(docx_file_or_buffer) -> Dict[str, str]:
    """Trích xuất placeholders {{Key}} {# mô tả #} từ .docx (body/header/footer)."""
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

def replace_in_paragraph(paragraph: Paragraph, data: Dict[str, str]):
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

    bullet_queue = []  # (text, style)
    table_queue  = []  # (header, rows)

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

        # Bảng Markdown
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

        # Bullets
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

        # Text thường
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

    # Body
    for i, paragraph in enumerate(doc.paragraphs):
        try:
            replace_in_paragraph(paragraph, data_input)
        except Exception as e:
            print(f"Error processing paragraph {i}: {e}")

    # Tables
    for table_idx, table in enumerate(doc.tables):
        for row_idx, row in enumerate(table.rows):
            for cell_idx, cell in enumerate(row.cells):
                for para_idx, paragraph in enumerate(cell.paragraphs):
                    try:
                        replace_in_paragraph(paragraph, data_input)
                    except Exception as e:
                        print(f"Error processing table {table_idx}, row {row_idx}, cell {cell_idx}, paragraph {para_idx}: {e}")

    # Headers & Footers
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
# CSV (DOC LING) PARSER
# =========================
def _normalize(v: str) -> str:
    if not isinstance(v, str):
        v = str(v or "")
    s2 = v.strip().lower()
    rep = {
        "à":"a","á":"a","ả":"a","ã":"a","ạ":"a",
        "ă":"a","ằ":"a","ắ":"a","ẳ":"a","ẵ":"a","ặ":"a",
        "â":"a","ầ":"a","ấ":"a","ẩ":"a","ẫ":"a","ập":"a","ậ":"a",
        "è":"e","é":"e","ẻ":"e","ẽ":"e","ẹ":"e",
        "ê":"e","ề":"e","ế":"e","ể":"e","ễ":"e","ệ":"e",
        "ì":"i","í":"i","ỉ":"i","ĩ":"i","ị":"i",
        "ò":"o","ó":"o","ỏ":"o","õ":"o","ọ":"o",
        "ô":"o","ồ":"o","ố":"o","ổ":"o","ỗ":"o","ộ":"o",
        "ơ":"o","ờ":"o","ớ":"o","ở":"o","ỡ":"o","ợ":"o",
        "ù":"u","ú":"u","ủ":"u","ũ":"u","ụ":"u",
        "ư":"u","ừ":"u","ứ":"u","ử":"u","ữ":"u","ự":"u",
        "ỳ":"y","ý":"y","ỷ":"y","ỹ":"y","ỵ":"y",
        "đ":"d"
    }
    for a,b in rep.items():
        s2 = s2.replace(a,b)
    return s2

def _first_match(cols: List[str], candidates: List[str]) -> Optional[str]:
    cols_norm = {c: _normalize(c) for c in cols}
    for cand in candidates:
        for col, norm in cols_norm.items():
            if cand in norm:
                return col
    return None

def _truthy_attended(val) -> bool:
    s = str(val).strip().lower()
    return s in {"1","x","✓","yes","y","true","present","co","có","tham du","attended","attend","có mặt","co mat"}

def _df_from_csv_with_docling(uploaded_file) -> pd.DataFrame:
    """
    Cố gắng đọc CSV bằng docling; nếu lỗi sẽ fallback sang pandas.read_csv.
    Trả về DataFrame (không raise).
    """
    # Đọc toàn bộ bytes của file uploader để có thể tua lại
    content = uploaded_file.read()
    # Reset stream cho các hàm khác nếu dùng lại
    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    # ƯU TIÊN: DOC LING
    try:
        # Tạo file tạm vì đa số converter cần đường dẫn
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix=".csv") as tmp:
            tmp.write(content)
            tmp_path = tmp.name

        try:
            # API docling thường dùng DocumentConverter.convert(input_path)
            # Mục tiêu: lấy table -> đưa về pandas.
            from docling.document_converter import DocumentConverter  # type: ignore
            conv = DocumentConverter()
            result = conv.convert(tmp_path)
            # Thử quét các artifacts/table
            # (API của docling có thể khác giữa phiên bản; dùng cách "best-effort")
            rows = []
            headers = None
            # result có thể có thuộc tính "tables" hoặc "artifacts"
            tables = getattr(result, "tables", None)
            if tables is None and hasattr(result, "artifacts"):
                # artifacts là list, mỗi phần tử có thể là bảng
                tables = [a for a in result.artifacts if getattr(a, "type", "") == "table"]

            if tables:
                # Dùng bảng đầu tiên làm attendance
                tbl = tables[0]
                # Thử các cách lấy dữ liệu phổ biến
                # 1) tbl.to_list()
                if hasattr(tbl, "to_list"):
                    data = tbl.to_list()
                # 2) tbl.data / tbl.rows
                elif hasattr(tbl, "data"):
                    data = getattr(tbl, "data")
                elif hasattr(tbl, "rows"):
                    data = getattr(tbl, "rows")
                else:
                    data = None

                if data:
                    # data có thể là list[list[str]]
                    if isinstance(data, list) and data and isinstance(data[0], list):
                        headers = [str(x) for x in data[0]]
                        rows = [list(map(lambda x: str(x) if x is not None else "", r)) for r in data[1:]]
                        df = pd.DataFrame(rows, columns=headers)
                        return df

            # Nếu không thu được bảng từ docling -> rơi xuống fallback pandas
        finally:
            try:
                os.unlink(tmp_path)
            except Exception:
                pass
    except Exception as e:
        # Không hiển thị dài dòng: chỉ cảnh báo nhẹ
        st.warning(f"Docling không đọc được CSV (sẽ dùng pandas): {e}")

    # FALLBACK: PANDAS
    encs = ["utf-8", "utf-8-sig", "cp1258", "latin1"]
    for enc in encs:
        try:
            from io import BytesIO
            return pd.read_csv(BytesIO(content), encoding=enc)
        except Exception:
            continue
    # chốt
    from io import BytesIO
    return pd.read_csv(BytesIO(content), engine="python", error_bad_lines=False)  # best-effort

def attendance_struct_from_df(df: pd.DataFrame) -> Dict[str, str]:
    """
    Chuẩn hoá DF -> bullets + bảng markdown + timeline markdown.
    Heuristic nhận dạng cột: name, email, dept, title, role, join, leave, duration, spoke/remark...
    """
    if df is None or df.empty:
        return {"participants_bullets": "", "participants_table_md": "", "participants_timeline_md": ""}

    cols = list(df.columns)
    name_col = _first_match(cols, ["name","full name","fullname","ho va ten","ho ten","ten","hvt","họ và tên"])
    dept_col = _first_match(cols, ["don vi","phong ban","department","unit","division"])
    title_col= _first_match(cols, ["chuc vu","title","position","role","vai tro"])
    mail_col = _first_match(cols, ["email","mail"])
    att_col  = _first_match(cols, ["attendance","status","co mat","co-mat","tham du","present","attended"])

    join_col = _first_match(cols, ["join","check in","gio vao","bat dau","start","join time","time joined","joined"])
    leave_col= _first_match(cols, ["leave","check out","gio ra","ket thuc","end","leave time","time left","left"])
    dur_col  = _first_match(cols, ["duration","thoi luong","tg tham gia","tgian"])
    spoke_col= _first_match(cols, ["spoke","phat bieu","noi","intervention","remark","ykien","ý kiến","content"])

    df2 = df.copy()
    # Lọc người có mặt
    if att_col:
        try:
            df2 = df2[df2[att_col].apply(_truthy_attended)]
        except Exception:
            pass

    # Bullets cấp 2
    bullets = []
    for _, r in df2.iterrows():
        name = str(r.get(name_col, "")).strip()
        title = str(r.get(title_col, "")).strip()
        dept  = str(r.get(dept_col,  "")).strip()
        email = str(r.get(mail_col,  "")).strip()
        tail_bits = []
        if title: tail_bits.append(title)
        if dept:  tail_bits.append(dept)
        tail = ", ".join(tail_bits)
        shown = name or ""
        if tail:  shown += f" — {tail}"
        if email: shown += f" ({email})"
        if shown: bullets.append(f"+ {shown}")
    participants_bullets = "\n".join(bullets)

    # Bảng markdown danh bạ
    headers = []
    rows = []
    def add_hdr(h):
        if h not in headers:
            headers.append(h)
    if name_col:  add_hdr("Name")
    if title_col: add_hdr("Title/Position")
    if dept_col:  add_hdr("Department")
    if mail_col:  add_hdr("Email")
    if headers:
        for _, r in df2.iterrows():
            row = []
            if name_col:  row.append(str(r.get(name_col, "")))
            if title_col: row.append(str(r.get(title_col, "")))
            if dept_col:  row.append(str(r.get(dept_col, "")))
            if mail_col:  row.append(str(r.get(mail_col, "")))
            rows.append(row)
        sep = "|" + "|".join(["---" for _ in headers]) + "|"
        participants_table_md = "|" + "|".join(headers) + "|\n" + sep + "\n" + "\n".join(["|" + "|".join(map(str, r)) + "|" for r in rows])
    else:
        participants_table_md = ""

    # Timeline markdown (nếu có cột thời gian/phát biểu)
    t_headers = []
    t_rows = []
    if name_col:  t_headers.append("Name")
    if join_col:  t_headers.append("Join")
    if leave_col: t_headers.append("Leave")
    if dur_col:   t_headers.append("Duration")
    if spoke_col: t_headers.append("Spoke/Remarks")
    if t_headers:
        for _, r in df2.iterrows():
            row = []
            if name_col:  row.append(str(r.get(name_col, "")))
            if join_col:  row.append(str(r.get(join_col, "")))
            if leave_col: row.append(str(r.get(leave_col, "")))
            if dur_col:   row.append(str(r.get(dur_col, "")))
            if spoke_col: row.append(str(r.get(spoke_col, "")))
            t_rows.append(row)
        if t_rows:
            sep2 = "|" + "|".join(["---" for _ in t_headers]) + "|"
            participants_timeline_md = "|" + "|".join(t_headers) + "|\n" + sep2 + "\n" + "\n".join(["|" + "|".join(map(str, r)) + "|" for r in t_rows])
        else:
            participants_timeline_md = ""
    else:
        participants_timeline_md = ""

    return {
        "participants_bullets": participants_bullets,
        "participants_table_md": participants_table_md,
        "participants_timeline_md": participants_timeline_md,
    }

def parse_attendance_csv_docling(uploaded_csv_file) -> Dict[str, str]:
    """
    Đọc file CSV bằng docling (ưu tiên) rồi chuẩn hoá ra bullets/table/timeline.
    Lưu ý: chỉ nhận CSV ở uploader CSV (Excel để uploader khác nếu muốn).
    """
    if uploaded_csv_file is None:
        return {"participants_bullets": "", "participants_table_md": "", "participants_timeline_md": ""}
    try:
        df = _df_from_csv_with_docling(uploaded_csv_file)
        return attendance_struct_from_df(df)
    except Exception as e:
        st.warning(f"Không thể phân tích CSV Attendance: {e}")
        return {"participants_bullets": "", "participants_table_md": "", "participants_timeline_md": ""}

# =========================
# LLM (GEMINI)
# =========================
def call_gemini_model(transcript_content: str, placeholders: Dict[str, str], csv_hint: Dict[str, str] | None = None) -> Optional[Dict[str, str]]:
    model = genai.GenerativeModel("gemini-2.5-pro")

    participants_block = ""
    if csv_hint:
        blt = (csv_hint.get("participants_bullets", "") or "").strip()
        tbl = (csv_hint.get("participants_table_md", "") or "").strip()
        tml = (csv_hint.get("participants_timeline_md", "") or "").strip()
        participants_block = f"""
# Dữ liệu CSV thành viên (đã chuẩn hoá)
- **Bullet cấp 2** (ưu tiên cho {{ThanhPhanThamGia}} nếu có): 
{blt}

- **Bảng danh bạ (Markdown)**:
{tbl}

- **Timeline tham dự & phát biểu (Markdown nếu có)**:
{tml}
""".strip()

    Prompt_word = f"""
# Vai trò
Bạn là trợ lý AI chuyên nghiệp, nhiệm vụ: trích xuất/thể hiện nội dung cho biên bản cuộc họp từ transcript **và** dữ liệu CSV người tham dự (nếu có), đảm bảo chính xác và trình bày chuẩn mực.

# Đầu vào
1) **Bản ghi cuộc họp (transcript):** ```{transcript_content}```
2) **Danh sách placeholders cần điền** (dict: key = tên trường, value = mô tả/định dạng yêu cầu): ```{json.dumps(placeholders, ensure_ascii=False)}```
3) **Dữ liệu người tham dự từ CSV (đã chuẩn hoá)**:
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
# EMAIL
# =========================
def send_email_with_attachment(recipient_email, attachment_buffer, filename="Bien_ban_cuoc_hop.docx"):
    SMTP_SERVER = "smtp.office365.com"
    SMTP_PORT = 587

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
        filename=filename
    )

    try:
        ctx = ssl.create_default_context()
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
            s.starttls(context=ctx)
            s.login(APP_EMAIL, APP_PASSWORD)
            s.send_message(msg)
        return True
    except Exception as e:
        st.error(f"Lỗi khi gửi email: {e}. Vui lòng kiểm tra lại cấu hình email/mật khẩu ứng dụng.")
        return False

# =========================
# HELPERS (IO/UI)
# =========================
def load_transcript_docx(file) -> str:
    """Đọc toàn bộ text từ .docx transcript (giữ như code cũ)."""
    try:
        doc = Document(file)
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        st.error(f"Lỗi đọc transcript .docx: {e}")
        return ""

def ensure_template_path(default_filename: str) -> Optional[str]:
    if os.path.exists(default_filename):
        return default_filename
    st.error(f"Không tìm thấy template mặc định: {default_filename}. Hãy chọn 'Template tùy chỉnh' và tải file lên.")
    return None

# =========================
# UI (GIỮ NGUYÊN PHẦN HƯỚNG DẪN CŨ + bổ sung uploader CSV docling)
# =========================
st.set_page_config(layout="wide", page_title="Công cụ tạo Biên bản cuộc họp")
st.title("🛠️ Công cụ tạo biên bản cuộc họp tự động")

with st.sidebar:
    st.info("📝 **Hướng dẫn sử dụng**")
    st.markdown("""
    1.  **Tải file transcript:** Tải lên file `.docx` chứa nội dung cuộc họp.
    2.  **Chọn Template:**
        * Sử dụng mẫu có sẵn bằng cách chọn "Template VPI".
        * Hoặc "Template tùy chỉnh" và tải file của bạn lên.
    3.  **Điền thông tin:** Nhập các thông tin cơ bản của cuộc họp.
    4.  **Nhập email:** Điền địa chỉ email bạn muốn nhận kết quả.
    5.  **Chạy:** Nhấn nút 'Tạo biên bản'.
    """)
    st.info("📝 **Hướng dẫn tạo template**")
    st.markdown("""
📂 File nhận đầu vào là file có đuôi `.docx`
Khi tạo template cho biên bản cuộc họp, bạn cần mô tả rõ từng biến để đảm bảo hệ thống hiểu đúng và điền thông tin chính xác. Mỗi biến cần tuân thủ cấu trúc sau: 
{{Ten_bien}}{# Mo_ta_chi_tiet #}
🔍 Trong đó:
- ✅ {{Ten_bien}}: Tên biến **viết bằng tiếng Anh hoặc tiếng Việt không dấu**, **không dùng khoảng trắng** (nếu cần dùng `_`). Ví dụ: {{Thanh_phan_tham_du}}
- ✅ {# Mo_ta_chi_tiet #}: nêu rõ dữ liệu cần điền + yêu cầu trình bày (bullet/bảng...), tối đa **hai cấp** bullet.
- 📍 Bullet cấp 1: **List Bullet**
- 📍 Bullet cấp 2: **List Bullet 2**
- 📍 Bảng: tạo Table Style `"New Table"` trong template.
    """)
    st.markdown("---")
    st.success("Ứng dụng được phát triển bởi VPI.")

st.subheader("1. Nhập thông tin đầu vào")
transcript_file = st.file_uploader("1) Tải lên file transcript (.docx)", type=["docx"])

# MỚI: Chỉ đọc CSV qua docling
csv_file = st.file_uploader("1b) Tải CSV Attendance (sẽ đọc bằng docling)", type=["csv"], help="Chỉ CSV. Công cụ sẽ ưu tiên đọc bằng docling, lỗi sẽ tự fallback pandas.")

st.subheader("2. Lựa chọn Template")
template_option = st.selectbox(
    "Bạn muốn sử dụng loại template nào?",
    ("Template VPI", "Template tùy chỉnh"),
    help="Chọn 'Template VPI' để dùng mẫu có sẵn hoặc 'Template tùy chỉnh' để tải lên file của riêng bạn."
)
template_file = None
if template_option == "Template tùy chỉnh":
    template_file = st.file_uploader("Tải lên file template .docx của bạn", type=["docx"])

st.subheader("3. Thông tin cơ bản")
if template_option == "Template tùy chỉnh":
    st.info(
        "🔔 **Lưu ý đối với Template tùy chỉnh**\n\n"
        "- File template **bắt buộc** có đúng và đủ: `{{TenCuocHop}}`, `{{ThoiGianCuocHop}}`, `{{DiaDiemCuocHop}}`, `{{TenChuTri}}`, `{{TenThuKy}}` (không kèm `{# ... #}`)."
    )
else:
    st.caption("Các trường bắt buộc đã có sẵn trong Template VPI.")

col1, col2 = st.columns(2)
with col1:
    meeting_name      = st.text_input("Tên cuộc họp")
    meeting_time      = st.text_input("Thời gian cuộc họp (VD: 10/9/2025)")
    meeting_location  = st.text_input("Địa điểm cuộc họp")
with col2:
    meeting_chair     = st.text_input("Tên chủ trì")
    meeting_secretary = st.text_input("Tên thư ký")

recipient_email = st.text_input("4. Email nhận kết quả của bạn")

# Nút chạy
if st.button("🚀 Tạo biên bản", type="primary"):
    default_path = "2025.VPI_BB hop 2025 1.docx"

    # Kiểm tra bắt buộc
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

    # Xác định template
    if template_option == "Template VPI":
        template_to_use = ensure_template_path(default_path)
        if not template_to_use:
            st.stop()
    else:
        template_to_use = template_file

    with st.spinner("⏳ Hệ thống đang xử lý..."):
        try:
            st.info("1/5 - Đang đọc và phân tích transcript (.docx)...")
            transcript_content = load_transcript_docx(transcript_file)

            st.info("2/5 - Đang trích placeholders từ template...")
            placeholders = extract_vars_and_desc(template_to_use)

            # Kiểm tra template có đủ placeholders cơ bản (mềm)
            try:
                tdoc = Document(template_to_use)
                ttext = "\n".join([p.text for p in tdoc.paragraphs])
                missing_ph = []
                for ph in REQUIRED_PLACEHOLDERS:
                    if f\"{{{{{ph}}}}}\" not in ttext:
                        missing_ph.append(ph)
                if missing_ph and template_option == "Template tùy chỉnh":
                    st.error("❌ **Template tùy chỉnh thiếu biến bắt buộc**: " + ", ".join(missing_ph))
                    st.stop()
            except Exception:
                pass

            st.info("3/5 - Đang đọc CSV Attendance bằng docling...")
            csv_hint = parse_attendance_csv_docling(csv_file) if csv_file else {"participants_bullets":"", "participants_table_md":"", "participants_timeline_md":""}

            st.info("4/5 - Đang gọi AI (Gemini) để tổng hợp nội dung...")
            llm_result = call_gemini_model(transcript_content, placeholders, csv_hint)

            if llm_result:
                # Ghi đè bằng input tay (trường bắt buộc)
                manual_inputs = {
                    'TenCuocHop':       meeting_name,
                    'ThoiGianCuocHop':  meeting_time,
                    'DiaDiemCuocHop':   meeting_location,
                    'TenChuTri':        meeting_chair,
                    'TenThuKy':         meeting_secretary
                }
                for k, v in manual_inputs.items():
                    if v is not None and v != "":
                        llm_result[k] = v

                # Ưu tiên CSV cho thành phần tham gia nếu placeholder tương ứng tồn tại
                if 'ThanhPhanThamGia' in llm_result and (csv_hint.get("participants_bullets") or "").strip():
                    llm_result['ThanhPhanThamGia'] = csv_hint['participants_bullets']

                st.info("5/5 - Đang điền template và tạo file Word...")
                docx_buffer = fill_template_to_buffer(template_to_use, llm_result)
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
                    st.error("Không thể tạo file Word. Vui lòng kiểm tra lại template hoặc dữ liệu đầu vào.")
            else:
                st.error("AI không trả về kết quả hợp lệ. Vui lòng thử lại.")
        except Exception as e:
            st.error(f"Đã xảy ra lỗi: {e}")
