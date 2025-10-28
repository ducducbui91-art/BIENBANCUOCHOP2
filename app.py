# app.py
# -*- coding: utf-8 -*-
"""
Ứng dụng Streamlit tạo biên bản cuộc họp từ transcript (.docx) + Attendance (.csv/.xlsx).
- Parse Excel kiểu Microsoft Teams Attendance (1. Summary → 2. Participants → 3. In-Meeting Activities).
- Sinh bullets + bảng Markdown cho thành phần tham dự và đưa vào prompt AI.
- Tự điền một số trường từ Summary nếu AI không trả về.
- Bắt buộc người dùng điền đủ thông tin (báo đỏ nếu thiếu).

Chạy:
    streamlit run app.py
"""

from __future__ import annotations
import io, os, re, json, zipfile, ssl, smtplib
from typing import Dict, List, Optional, Tuple

import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.shared import Inches  # noqa: F401
import google.generativeai as genai
from email.message import EmailMessage

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
COMMENT_RE     = re.compile(r"\{#.*?#\}")                 # 1-run
COMMENT_ALL_RE = re.compile(r"\{#.*?#\}", re.DOTALL)      # đa-run
BOLD_RE        = re.compile(r"\*\*(.*?)\*\*")             # **bold**
TOKEN_RE       = re.compile(r"\{\{([^{}]+)\}\}")          # {{Key}}

# =========================
# TIỆN ÍCH WORD
# =========================
def _is_md_table(text: str) -> bool:
    lines = [l.strip() for l in (text or "").strip().splitlines() if l.strip()]
    return (len(lines) >= 2 and "|" in lines[0] and set(lines[1].replace(" ", "").replace(":", "")) <= set("-|"))

def _parse_md_table(text: str) -> Tuple[List[str], List[List[str]]]:
    lines  = [l.strip() for l in (text or "").strip().splitlines() if l.strip()]
    header = [c.strip() for c in lines[0].split("|")]
    if header and header[0] == "": header = header[1:]
    if header and header[-1] == "": header = header[:-1]
    rows: List[List[str]] = []
    for ln in lines[2:]:
        cols = [c.strip() for c in ln.split("|")]
        if cols and cols[0] == "": cols = cols[1:]
        if cols and cols[-1] == "": cols = cols[:-1]
        if cols:
            while len(cols) > len(header): cols.pop()
            while len(cols) < len(header): cols.append("")
            rows.append(cols)
    return header, rows

def _insert_paragraph_after(anchor_para: Paragraph, style: Optional[str] = None) -> Paragraph:
    new_p_ox = OxmlElement("w:p")
    anchor_para._p.addnext(new_p_ox)
    new_para = Paragraph(new_p_ox, anchor_para._parent)
    if style:
        try: new_para.style = style
        except Exception: pass
    return new_para

def add_formatted_text(paragraph: Paragraph, text: str, style_info: Optional[dict] = None) -> None:
    parts, is_bold = BOLD_RE.split(text or ""), False
    for part in parts:
        if part == "":
            is_bold = not is_bold
            continue
        lines = part.split("\n")
        for i, ln in enumerate(lines):
            if i > 0: paragraph.add_run().add_break()
            if ln == "": continue
            run = paragraph.add_run(ln)
            if style_info:
                try:
                    f = run.font
                    if style_info.get("size"):  f.size   = style_info["size"]
                    if style_info.get("name"):  f.name   = style_info["name"]
                    if style_info.get("bold") is not None:   f.bold   = style_info["bold"]
                    if style_info.get("italic") is not None: f.italic = style_info["italic"]
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
    if not header or not rows: return
    body = paragraph._parent
    tbl  = body.add_table(rows=len(rows)+1, cols=len(header))
    try: tbl.style = table_style
    except Exception: pass
    for i, h in enumerate(header):
        try: tbl.rows[0].cells[i].text = str(h)
        except Exception: pass
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, cell_val in enumerate(row):
            try: tbl.rows[r_idx].cells[c_idx].text = str(cell_val)
            except Exception: pass
    paragraph._p.addnext(tbl._tbl)

# =========================
# XỬ LÝ TEMPLATE .DOCX
# =========================
def extract_vars_and_desc(docx_file_or_buffer) -> Dict[str, str]:
    xml_parts: List[str] = []
    with zipfile.ZipFile(docx_file_or_buffer) as z:
        for name in z.namelist():
            if name.startswith("word/") and name.endswith(".xml"):
                xml_parts.append(z.read(name).decode("utf8"))
    all_xml   = "\n".join(xml_parts)
    texts     = re.findall(r"<w:t[^>]*>(.*?)</w:t>", all_xml, flags=re.DOTALL)
    full_text = "".join(texts)
    pattern   = re.compile(r"\{\{\s*([A-Za-z0-9_]+)\s*\}\}\s*\{#\s*(.*?)\s*#\}", flags=re.DOTALL)
    return dict(pattern.findall(full_text))

def replace_in_paragraph(paragraph: Paragraph, data: Dict[str, str]) -> None:
    if not paragraph.runs: return
    full_text, meta = _concat_runs(paragraph)
    if not full_text: return

    items = []
    for m in COMMENT_ALL_RE.finditer(full_text):
        items.append(("comment", m.start(), m.end(), None))
    for m in TOKEN_RE.finditer(full_text):
        key = (m.group(1) or "").strip()
        if key in data: items.append(("token", m.start(), m.end(), key))
    if not items:
        for r in paragraph.runs:
            if r.text and COMMENT_RE.search(r.text):
                r.text = COMMENT_RE.sub("", r.text)
        return

    items.sort(key=lambda x: x[1], reverse=True)
    bullet_queue: List[Tuple[str, str]] = []
    table_queue:  List[Tuple[List[str], List[List[str]]]] = []

    for item_type, start, end, key in items:
        run_start_idx = next((i for i, (_, s, e) in enumerate(meta) if s <= start < e), None)
        run_end_idx   = next((i for i, (_, s, e) in enumerate(meta) if s <  end <= e), None)
        if run_start_idx is None or run_end_idx is None: continue

        run_start, s0, _ = meta[run_start_idx]
        run_end,   _, e1 = meta[run_end_idx]
        offset_start, offset_end = start - s0, end - e1

        if item_type == "comment":
            if run_start_idx == run_end_idx:
                t = run_start.text or ""
                run_start.text = t[:offset_start] + t[offset_end:]
            else:
                run_start.text = (run_start.text or "")[:offset_start]
                for i in range(run_start_idx + 1, run_end_idx): meta[i][0].text = ""
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
                    for i in range(run_start_idx + 1, run_end_idx): meta[i][0].text = ""
                    run_end.text = (run_end.text or "")[offset_end:]
                continue
            except Exception:
                value = str(value)

        # Bullet
        if isinstance(value, str) and any(line.strip().startswith(("-", "+")) for line in value.splitlines()):
            for line in value.splitlines():
                s = line.strip()
                if s.startswith("-"): bullet_queue.append((s[1:].strip(), "List Bullet"))
                elif s.startswith("+"): bullet_queue.append((s[1:].strip(), "List Bullet 2"))
            if run_start_idx == run_end_idx:
                t = run_start.text or ""
                run_start.text = t[:offset_start] + t[offset_end:]
            else:
                run_start.text = (run_start.text or "")[:offset_start]
                for i in range(run_start_idx + 1, run_end_idx): meta[i][0].text = ""
                run_end.text = (run_end.text or "")[offset_end:]
            continue

        # Văn bản thường
        replacement_text = str(value)
        if run_start_idx == run_end_idx:
            t = run_start.text or ""
            run_start.text = t[:offset_start] + replacement_text + t[offset_end:]
        else:
            for i in range(run_start_idx + 1, run_end_idx): meta[i][0].text = ""
            start_text = (run_start.text or "")[:offset_start]
            run_start.text = start_text + replacement_text
            run_end.text   = (run_end.text or "")[offset_end:]

    # chèn bullet/bảng sau paragraph
    if bullet_queue or table_queue:
        current_para = paragraph
        for text, style in bullet_queue:
            current_para = _insert_paragraph_after(current_para, style=style)
            add_formatted_text(current_para, text)
        for header, rows in table_queue:
            try: _insert_table_after(current_para, header, rows)
            except Exception as e: print(f"Error inserting table: {e}")

def fill_template_to_buffer(template_file_or_path, data_input: Dict[str, str]) -> Optional[io.BytesIO]:
    try:
        doc = Document(template_file_or_path)
    except Exception as e:
        st.error(f"Lỗi mở template: {e}")
        return None

    for i, paragraph in enumerate(doc.paragraphs):
        try: replace_in_paragraph(paragraph, data_input)
        except Exception as e: print(f"Error processing paragraph {i}: {e}")

    for table_idx, table in enumerate(doc.tables):
        for row_idx, row in enumerate(table.rows):
            for cell_idx, cell in enumerate(row.cells):
                for para_idx, paragraph in enumerate(cell.paragraphs):
                    try: replace_in_paragraph(paragraph, data_input)
                    except Exception as e: print(f"Error processing table {table_idx}, row {row_idx}, cell {cell_idx}, paragraph {para_idx}: {e}")

    for section in doc.sections:
        if section.header:
            for paragraph in section.header.paragraphs:
                try: replace_in_paragraph(paragraph, data_input)
                except Exception as e: print(f"Error processing header paragraph: {e}")
        if section.footer:
            for paragraph in section.footer.paragraphs:
                try: replace_in_paragraph(paragraph, data_input)
                except Exception as e: print(f"Error processing footer paragraph: {e}")

    try:
        buf = io.BytesIO(); doc.save(buf); buf.seek(0); return buf
    except Exception as e:
        st.error(f"Đã xảy ra lỗi khi tạo file Word: {e}")
        return None

# =========================
# ATTENDANCE PARSERS
# =========================
def _normalize(s: str) -> str:
    if not isinstance(s, str): return ""
    s2 = s.strip().lower()
    rep = {"à":"a","á":"a","ả":"a","ã":"a","ạ":"a","ă":"a","ằ":"a","ắ":"a","ẳ":"a","ẵ":"a","ặ":"a","â":"a","ầ":"a","ấ":"a","ẩ":"a","ẫ":"a","ậ":"a",
           "è":"e","é":"e","ẻ":"e","ẽ":"e","ẹ":"e","ê":"e","ề":"e","ế":"e","ể":"e","ễ":"e","ệ":"e",
           "ì":"i","í":"i","ỉ":"i","ĩ":"i","ị":"i",
           "ò":"o","ó":"o","ỏ":"o","õ":"o","ọ":"o","ô":"o","ồ":"o","ố":"o","ổ":"o","ỗ":"o","ộ":"o","ơ":"o","ờ":"o","ớ":"o","ở":"o","ỡ":"o","ợ":"o",
           "ù":"u","ú":"u","ủ":"u","ũ":"u","ụ":"u","ư":"u","ừ":"u","ứ":"u","ử":"u","ữ":"u","ự":"u",
           "ỳ":"y","ý":"y","ỷ":"y","ỹ":"y","ỵ":"y","đ":"d"}
    for a,b in rep.items(): s2 = s2.replace(a,b)
    return s2

def _first_match(cols: List[str], candidates: List[str]) -> Optional[str]:
    cols_norm = {c: _normalize(c) for c in cols}
    for c in candidates:
        for col, norm in cols_norm.items():
            if c in norm: return col
    return None

def _looks_present(val) -> bool:
    if val is None: return True
    s = str(val).strip().lower()
    return s in {"1","x","✓","yes","y","true","present","co","có","tham du","attended"}

def read_attendance_to_df(file) -> pd.DataFrame:
    name = getattr(file, "name", "") or ""
    ext  = os.path.splitext(name.lower())[1]
    if ext in (".xlsx", ".xls"):
        try: return pd.read_excel(file)
        except Exception:
            try: file.seek(0)
            except Exception: pass
    encodings = ["utf-8","utf-8-sig","cp1258","latin1"]
    last_err = None
    for enc in encodings:
        try:
            try: file.seek(0)
            except Exception: pass
            return pd.read_csv(file, encoding=enc)
        except Exception as e:
            last_err = e
            continue
    try:
        try: file.seek(0)
        except Exception: pass
        return pd.read_excel(file)
    except Exception as e:
        raise RuntimeError(f"Không thể đọc Attendance (CSV/Excel). Lỗi cuối: {last_err or e}")

def _attendance_df_to_struct(df: pd.DataFrame) -> Dict[str, str]:
    if df is None or df.empty:
        return {"participants_bullets": "", "participants_table_md": ""}

    cols = list(df.columns)
    name_col = _first_match(cols, ["name","full name","fullname","ho va ten","ho ten","ten","họ và tên"])
    dept_col = _first_match(cols, ["don vi","phong ban","department","unit","division"])
    title_col= _first_match(cols, ["chuc vu","title","position","role"])
    mail_col = _first_match(cols, ["email","mail"])
    att_col  = _first_match(cols, ["attendance","status","co mat","tham du","present","attended"])

    if att_col: df = df[df[att_col].apply(_looks_present)]

    bullet_lines: List[str] = []
    for _, r in df.iterrows():
        name  = str(r.get(name_col, "")).strip()
        title = str(r.get(title_col, "")).strip()
        dept  = str(r.get(dept_col,  "")).strip()
        email = str(r.get(mail_col,  "")).strip()
        if not name: continue
        tail_bits: List[str] = []
        if title: tail_bits.append(title)
        if dept:  tail_bits.append(dept)
        shown = name + (f" — {', '.join(tail_bits)}" if tail_bits else "")
        if email: shown += f" ({email})"
        bullet_lines.append(f"+ {shown}")

    participants_bullets = "\n".join(bullet_lines)

    headers, rows = [], []
    def add_hdr(h): 
        if h not in headers: headers.append(h)
    if name_col:  add_hdr("Name")
    if title_col: add_hdr("Title/Position")
    if dept_col:  add_hdr("Department")
    if mail_col:  add_hdr("Email")
    if headers:
        for _, r in df.iterrows():
            row = []
            if name_col:  row.append(str(r.get(name_col, "")).strip())
            if title_col: row.append(str(r.get(title_col, "")).strip())
            if dept_col:  row.append(str(r.get(dept_col, "")).strip())
            if mail_col:  row.append(str(r.get(mail_col, "")).strip())
            rows.append(row)
        sep = "|" + "|".join(["---" for _ in headers]) + "|"
        participants_table_md = "|" + "|".join(headers) + "|\n" + sep + "\n" + "\n".join(["|" + "|".join(r) + "|" for r in rows])
    else:
        participants_table_md = ""

    return {"participants_bullets": participants_bullets, "participants_table_md": participants_table_md}

# ---- Parser Excel kiểu Microsoft Teams Attendance ----
def _df_values(df, r, c):
    try:
        v = df.iat[r, c]
        if pd.isna(v): return ""
        return str(v).strip()
    except Exception:
        return ""

def parse_teams_attendance_excel(file) -> Dict[str, str]:
    df_raw = pd.read_excel(file, header=None)
    df = df_raw.fillna("")

    idx_summary = idx_part = idx_acts = None
    for i in range(len(df)):
        first = _df_values(df, i, 0)
        if first.startswith("1. Summary"): idx_summary = i
        elif first.startswith("2. Participants"): idx_part = i
        elif first.startswith("3. In-Meeting Activities"): idx_acts = i

    # Summary
    summary = {}
    if idx_summary is not None:
        r = idx_summary + 1
        while r < len(df):
            k, v = _df_values(df, r, 0), _df_values(df, r, 1)
            if k.startswith("2. ") or k.startswith("3. "): break
            if k or v:
                if k: summary[k] = v
            r += 1
    meeting_title = summary.get("Meeting title", "")
    start_time    = summary.get("Start time", "")
    end_time      = summary.get("End time", "")
    attended_cnt  = summary.get("Attended participants", "")

    # Participants
    participants: List[Dict[str, str]] = []
    if idx_part is not None:
        r = idx_part + 1
        while r < len(df) and not any(_df_values(df, r, c) for c in range(8)): r += 1
        header_row = r
        headers = [_df_values(df, header_row, c) for c in range(8)]
        hmap = {h.strip().lower(): c for c, h in enumerate(headers) if h}

        r = header_row + 1
        while r < len(df):
            first_cell = _df_values(df, r, 0)
            if first_cell.startswith("3. In-Meeting Activities"): break
            if not any(_df_values(df, r, c) for c in range(0, 7)):
                r += 1
                next_first = _df_values(df, r, 0) if r < len(df) else ""
                if next_first.startswith("3. In-Meeting Activities"): break
                continue
            rec = {
                "name":       _df_values(df, r, hmap.get("name", 0)),
                "first_join": _df_values(df, r, hmap.get("first join", 1)),
                "last_leave": _df_values(df, r, hmap.get("last leave", 2)),
                "duration":   _df_values(df, r, hmap.get("in-meeting duration", 3)),
                "email":      _df_values(df, r, hmap.get("email", 4)),
                "upn":        _df_values(df, r, hmap.get("participant id (upn)", 5)),
                "role":       _df_values(df, r, hmap.get("role", 6)),
            }
            if rec["name"]: participants.append(rec)
            r += 1

    # Fallback time from participants
    if not start_time and participants:
        fj = [p["first_join"] for p in participants if p.get("first_join")]
        start_time = min(fj) if fj else start_time
    if not end_time and participants:
        ll = [p["last_leave"] for p in participants if p.get("last_leave")]
        end_time = max(ll) if ll else end_time

    # Bullets & Markdown
    bullet_lines: List[str] = []
    for p in participants:
        tail = []
        if p.get("role"):  tail.append(p["role"])
        if p.get("email"): tail.append(p["email"])
        times = " ".join(x for x in [
            p.get("first_join"), "→" if (p.get("first_join") and p.get("last_leave")) else None,
            p.get("last_leave"), f"({p.get('duration')})" if p.get("duration") else None
        ] if x)
        if times: tail.append(times)
        line = p["name"] + (" — " + ", ".join(tail) if tail else "")
        bullet_lines.append("+ " + line)
    participants_bullets = "\n".join(bullet_lines)

    md_headers = ["Name", "Role", "Email", "First Join", "Last Leave", "Duration"]
    md_sep = "|" + "|".join(["---"]*len(md_headers)) + "|"
    md_rows = []
    for p in participants:
        md_rows.append("|" + "|".join([
            p.get("name",""), p.get("role",""), p.get("email",""),
            p.get("first_join",""), p.get("last_leave",""), p.get("duration","")
        ]) + "|")
    participants_table_md = "|" + "|".join(md_headers) + "|\n" + md_sep + "\n" + "\n".join(md_rows)

    return {
        "participants_bullets": participants_bullets,
        "participants_table_md": participants_table_md,
        "meta": {
            "meeting_title": meeting_title,
            "start_time": start_time,
            "end_time": end_time,
            "attended_participants": attended_cnt,
            "participants": participants,
        }
    }

def parse_attendance_any(file) -> Dict[str, str]:
    """Ưu tiên Excel kiểu Teams; nếu không, dùng parser CSV chung."""
    name = getattr(file, "name", "") or ""
    ext  = os.path.splitext(name.lower())[1]
    if ext in (".xlsx", ".xls"):
        try:
            return parse_teams_attendance_excel(file)
        except Exception:
            try: file.seek(0)
            except Exception: pass
    df = read_attendance_to_df(file)
    return _attendance_df_to_struct(df)

# =========================
# LLM CALL (Gemini)
# =========================
def call_gemini_model(transcript_content: str,
                      placeholders: Dict[str, str],
                      participants_hint: Dict[str, str] | None = None) -> Optional[Dict[str, str]]:
    model = genai.GenerativeModel("gemini-2.5-pro")

    # block dữ liệu tham dự
    blt = (participants_hint or {}).get("participants_bullets", "").strip()
    tbl = (participants_hint or {}).get("participants_table_md", "").strip()
    meta = (participants_hint or {}).get("meta", {}) or {}

    participants_block = f"""
# Dữ liệu thành viên (ưu tiên dùng cho các trường liên quan người tham dự)
- Tóm tắt: Tiêu đề = {meta.get('meeting_title','')}, Bắt đầu = {meta.get('start_time','')}, Kết thúc = {meta.get('end_time','')}, Số người tham dự = {meta.get('attended_participants','')}
- Bullet cấp 2 (gợi ý cho {{ThanhPhanThamGia}}): 
{blt}
- Bảng Markdown (nếu template có trường dạng bảng): 
{tbl}
""".strip()

    prompt = f"""
# Vai trò
Bạn là trợ lý AI chuyên nghiệp, nhiệm vụ: trích xuất/thể hiện nội dung cho biên bản cuộc họp từ transcript **và** dữ liệu Attendance (nếu có), đảm bảo chính xác và trình bày chuẩn mực.

# Đầu vào
1) Transcript: ```{transcript_content}```
2) Placeholders (dict: key = tên trường, value = mô tả/định dạng): ```{json.dumps(placeholders, ensure_ascii=False)}```
3) Attendance:
{participants_block}

# Yêu cầu
- Luôn trả về **TIẾNG VIỆT**.
- **Chỉ trả về một đối tượng JSON hợp lệ**: key trùng 100% tên placeholders; value là **chuỗi**; không thêm/bớt key; không lồng cấu trúc.
- Tuân thủ đúng định dạng mong muốn trong mô tả của từng placeholder (bullet 1: '- ', bullet 2: '+ ', bảng: Markdown).
- Ưu tiên dữ liệu Attendance cho trường thành phần tham gia/role, kết hợp transcript khi cần.
- Nếu thiếu thông tin: dùng đúng chuỗi **"Chưa có thông tin"**.

# Kết quả
Trả về **duy nhất** một chuỗi JSON hợp lệ.
"""

    try:
        response = model.generate_content(
            contents=prompt,
            generation_config={"response_mime_type": "application/json"}
        )
        if response and hasattr(response, "text"):
            raw = response.text.strip()
            if raw.startswith("```"):
                raw = raw.split("```")[1].strip("json\n")
            return json.loads(raw)
        st.error("Phản hồi từ Gemini API bị thiếu hoặc không hợp lệ.")
        return None
    except Exception as e:
        st.error(f"Lỗi khi gọi Gemini API: {e}")
        return None

# =========================
# EMAIL
# =========================
def send_email_with_attachment(recipient_email: str,
                               attachment_buffer: io.BytesIO,
                               filename: str = "Bien_ban_cuoc_hop.docx") -> bool:
    SMTP_SERVER = "smtp.office365.com"; SMTP_PORT = 587
    msg = EmailMessage()
    msg["Subject"] = "Biên bản cuộc họp đã được tạo tự động"
    msg["From"] = APP_EMAIL; msg["To"] = recipient_email
    msg.set_content("Chào bạn,\n\nBiên bản cuộc họp đã được tạo thành công.\nVui lòng xem file đính kèm.\n\nTrân trọng,\nVPI.")
    msg.add_attachment(attachment_buffer.getvalue(),
                       maintype="application",
                       subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
                       filename=filename)
    try:
        ctx = ssl.create_default_context()
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
            s.starttls(context=ctx)
            s.login(APP_EMAIL, APP_PASSWORD)
            s.send_message(msg)
        return True
    except Exception as e:
        st.error(f"Lỗi khi gửi email: {e}.")
        return False

# =========================
# HELPERS (IO/UI)
# =========================
def load_transcript_docx(file) -> str:
    try:
        doc = Document(file)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception as e:
        st.error(f"Lỗi đọc transcript .docx: {e}")
        return ""

def ensure_template_path(default_filename: str) -> Optional[str]:
    if os.path.exists(default_filename): return default_filename
    st.error(f"Không tìm thấy template mặc định: {default_filename}. Hãy chọn 'Template tùy chỉnh' và tải file lên.")
    return None

def validate_required(transcript_file, attendance_file, template_to_use,
                      meeting_name, meeting_time, meeting_location,
                      meeting_chair, meeting_secretary, recipient_email) -> bool:
    """Hiển thị lỗi đỏ nếu thiếu mục bắt buộc."""
    missing = []
    if not transcript_file:   missing.append("• File transcript (.docx)")
    if not attendance_file:   missing.append("• File Attendance (.csv/.xlsx)")
    if not template_to_use:   missing.append("• File template (.docx)")
    if not meeting_name:      missing.append("• Tên cuộc họp")
    if not meeting_time:      missing.append("• Thời gian cuộc họp")
    if not meeting_location:  missing.append("• Địa điểm cuộc họp")
    if not meeting_chair:     missing.append("• Tên chủ trì")
    if not meeting_secretary: missing.append("• Tên thư ký")
    if not recipient_email:   missing.append("• Email nhận kết quả")
    if missing:
        st.error("❌ **Chưa hoàn thành thông tin. Vui lòng bổ sung các mục sau:**\n\n" + "\n".join(missing))
        return False
    return True

# =========================
# STREAMLIT UI
# =========================
st.set_page_config(layout="wide", page_title="Công cụ tạo Biên bản cuộc họp")
st.title("🛠️ Công cụ tạo biên bản cuộc họp tự động")

with st.sidebar:
    st.info("**Hướng dẫn nhanh**")
    st.markdown("1) Tải transcript (.docx) **và** Attendance (.csv/.xlsx)\n2) Chọn Template\n3) Điền thông tin\n4) Nhấn **Tạo biên bản**")
    st.caption("Yêu cầu: streamlit, pandas, python-docx, google-generativeai, openpyxl")

st.subheader("1) Tải dữ liệu đầu vào")
colA, colB = st.columns(2)
with colA:
    transcript_file = st.file_uploader("Transcript (.docx) *bắt buộc*", type=["docx"], key="transcript")
with colB:
    attendance_file = st.file_uploader("Attendance (.csv/.xlsx) *bắt buộc*", type=["csv", "xlsx", "xls"], key="attendance")

st.subheader("2) Lựa chọn Template")
template_option = st.selectbox("Loại template", ("Template VPI", "Template tùy chỉnh"))
template_file = None
if template_option == "Template tùy chỉnh":
    template_file = st.file_uploader("Tải template .docx", type=["docx"], key="tpl")

st.subheader("3) Thông tin cơ bản (sẽ ghi đè kết quả AI)")
col1, col2 = st.columns(2)
with col1:
    meeting_name      = st.text_input("Tên cuộc họp *", placeholder="Ví dụ: Đánh giá sản phẩm biên bản cuộc họp")
    meeting_time      = st.text_input("Thời gian cuộc họp *", placeholder="Ví dụ: 10/21/2025, 10:17–11:06")
    meeting_location  = st.text_input("Địa điểm cuộc họp *", placeholder="VPI Hà Nội / Teams / ...")
with col2:
    meeting_chair     = st.text_input("Tên chủ trì *")
    meeting_secretary = st.text_input("Tên thư ký *")

recipient_email = st.text_input("4) Email nhận kết quả *", placeholder="you@company.com")

# ====== Action ======
if st.button("🚀 Tạo biên bản", type="primary"):
    # Chọn template
    template_to_use = None
    if template_option == "Template VPI":
        default_path = "2025.VPI_BB hop 2025 2.docx"
        template_to_use = ensure_template_path(default_path)
    else:
        template_to_use = template_file

    # Kiểm tra bắt buộc
    if not validate_required(transcript_file, attendance_file, template_to_use,
                             meeting_name, meeting_time, meeting_location,
                             meeting_chair, meeting_secretary, recipient_email):
        st.stop()

    with st.spinner("⏳ Đang xử lý..."):
        try:
            # 1) Đọc transcript
            st.info("1/6 - Đọc transcript .docx")
            transcript_content = load_transcript_docx(transcript_file)

            # 2) Trích placeholders
            st.info("2/6 - Trích placeholders từ template")
            placeholders = extract_vars_and_desc(template_to_use)

            # 3) Parse Attendance
            st.info("3/6 - Phân tích Attendance (CSV/Excel)")
            participants_hint = {"participants_bullets": "", "participants_table_md": "", "meta": {}}
            try:
                participants_hint = parse_attendance_any(attendance_file)
            except Exception as e:
                st.error(f"Không đọc được Attendance: {e}")
                st.stop()

            # 4) Gọi AI
            st.info("4/6 - Gọi AI tạo JSON theo placeholders")
            llm_result = call_gemini_model(transcript_content, placeholders, participants_hint)
            if not llm_result:
                st.error("AI không trả về kết quả hợp lệ.")
                st.stop()

            # 5) Ghi đè các input tay + ưu tiên dữ liệu Attendance cho thành phần tham gia
            st.info("5/6 - Áp dụng dữ liệu nhập tay/Attendance")
            manual_inputs = {
                "TenCuocHop":       meeting_name,
                "ThoiGianCuocHop":  meeting_time,
                "DiaDiemCuocHop":   meeting_location,
                "TenChuTri":        meeting_chair,
                "TenThuKy":         meeting_secretary,
            }
            for k, v in manual_inputs.items():
                if k in llm_result and v: llm_result[k] = v

            # Prefill từ meta nếu key tồn tại mà rỗng
            meta = participants_hint.get("meta", {}) or {}
            if "TenCuocHop" in llm_result and not llm_result["TenCuocHop"] and meta.get("meeting_title"):
                llm_result["TenCuocHop"] = meta["meeting_title"]
            if "ThoiGianCuocHop" in llm_result and not llm_result["ThoiGianCuocHop"]:
                st_ = meta.get("start_time", ""); en_ = meta.get("end_time", "")
                if st_ or en_:
                    llm_result["ThoiGianCuocHop"] = (st_ + (" - " + en_ if en_ else "")).strip(" -")

            # Ưu tiên Attendance cho thành phần tham gia
            if "ThanhPhanThamGia" in llm_result and participants_hint.get("participants_bullets"):
                llm_result["ThanhPhanThamGia"] = participants_hint["participants_bullets"]
            # Nếu template có trường dạng bảng
            for key in ["BangThanhPhanThamGia", "BangNguoiThamDu", "ParticipantsTable"]:
                if key in llm_result and participants_hint.get("participants_table_md"):
                    llm_result[key] = participants_hint["participants_table_md"]

            # 6) Điền template
            st.info("6/6 - Tạo file biên bản Word")
            docx_buffer = fill_template_to_buffer(template_to_use, llm_result)
            if not docx_buffer:
                st.error("Không thể tạo file Word. Kiểm tra lại template hoặc dữ liệu.")
                st.stop()

            st.success("✅ Tạo biên bản thành công!")
            st.download_button("⬇️ Tải về biên bản",
                               data=docx_buffer,
                               file_name="Bien_ban_cuoc_hop.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

            if recipient_email:
                ok = send_email_with_attachment(recipient_email, docx_buffer)
                if ok: st.success("✉️ Đã gửi biên bản tới email của bạn.")

        except Exception as e:
            st.error(f"Đã xảy ra lỗi: {e}")
