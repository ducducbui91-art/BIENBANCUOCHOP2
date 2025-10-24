# app.py
import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.shared import Inches
import re
import os
import json
import zipfile
from typing import Dict
import io
import smtplib, ssl
from email.message import EmailMessage
import google.generativeai as genai
import csv

# --- CẤU HÌNH BẢO MẬT ---
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    APP_EMAIL      = st.secrets["APP_EMAIL"]
    APP_PASSWORD   = st.secrets["APP_PASSWORD"]
except Exception:
    st.warning("Không tìm thấy Streamlit Secrets. Đang sử dụng cấu hình local. Đừng quên thiết lập Secrets khi deploy!")
    GEMINI_API_KEY = "YOUR_GEMINI_API_KEY"
    APP_EMAIL      = "your_email@example.com"
    APP_PASSWORD   = "your_app_or_email_password"

# Cấu hình API key cho Gemini
try:
    genai.configure(api_key=GEMINI_API_KEY)
except Exception as e:
    st.error(f"Lỗi cấu hình Gemini API: {e}. Vui lòng kiểm tra lại API Key.")

#======================================================================
# PHẦN 1: HÀM XỬ LÝ
#======================================================================

# Regex
COMMENT_RE     = re.compile(r"\{#.*?#\}")                 # 1-run
COMMENT_ALL_RE = re.compile(r"\{#.*?#\}", re.DOTALL)      # đa-run
BOLD_RE        = re.compile(r"\*\*(.*?)\*\*")             # **bold**
TOKEN_RE       = re.compile(r"\{\{([^{}]+)\}\}")          # {{Key}}

def _is_md_table(text: str) -> bool:
    lines = [l.strip() for l in (text or "").strip().splitlines() if l.strip()]
    return (
        len(lines) >= 2
        and "|" in lines[0]
        and set(lines[1].replace(" ", "").replace(":", "")) <= set("-|")
    )

def _parse_md_table(text: str):
    lines  = [l.strip() for l in (text or "").strip().splitlines() if l.strip()]
    header = [c.strip() for c in lines[0].split("|")]
    if header and header[0] == "":
        header = header[1:]
    if header and header[-1] == "":
        header = header[:-1]
    rows   = []
    for ln in lines[2:]:
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
    xml_parts = []
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

    bullet_queue = []
    table_queue  = []

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

def fill_template_to_buffer(template_file_or_path, data_input: Dict[str, str]):
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

# ---------- NEW: đọc CSV attendance thành text an toàn ----------
def read_uploaded_csv_as_text(uploaded_file, max_rows=1000, max_chars=200_000):
    """
    Đọc file CSV đã upload và trả về chuỗi CSV (tối đa max_rows dòng, tối đa max_chars ký tự).
    """
    if uploaded_file is None:
        return ""
    raw = uploaded_file.getvalue()
    try:
        text = raw.decode("utf-8-sig")
    except Exception:
        text = raw.decode("latin-1", errors="ignore")

    reader = csv.reader(io.StringIO(text))
    rows = []
    for idx, row in enumerate(reader):
        if idx >= max_rows:
            break
        rows.append(row)

    output = io.StringIO()
    writer = csv.writer(output)
    for r in rows:
        writer.writerow(r)
    csv_text = output.getvalue()

    if len(csv_text) > max_chars:
        csv_text = csv_text[:max_chars] + "\n...[TRUNCATED]"
    return csv_text

# ---------- CHANGED: gộp transcript + CSV khi gọi AI ----------
def call_gemini_model(transcript_content, csv_text, placeholders):
    """Gửi yêu cầu đến Gemini và nhận về kết quả JSON (đã gộp transcript + CSV)."""
    model = genai.GenerativeModel("gemini-2.5-pro")
    Prompt_word = """
# Vai trò
Bạn là một trợ lý AI chuyên nghiệp, có nhiệm vụ trích xuất thông tin quan trọng từ tư liệu cuộc họp (transcript + dữ liệu attendance CSV) để tạo nội dung cho biên bản, đảm bảo tính chính xác và trình bày chuyên nghiệp.

# Đầu vào
1.  **Bản ghi cuộc họp (transcript):**
{0}

1b. **Dữ liệu attendance dạng CSV** (ví dụ từ Google Meet/Teams; có thể gồm tên người tham gia, giờ vào/ra, thời lượng, email, v.v.):
```csv
{1}
Danh sách các trường thông tin cần trích xuất (placeholders):
{2}
(Là một đối tượng/dictionary nơi mỗi key là tên trường cần trích xuất và value là mô tả/yêu cầu định dạng.)

Nhiệm vụ

Đọc & hiểu toàn bộ transcript và CSV attendance.

Với từng key trong placeholders:

Tìm thông tin tương ứng từ transcript/CSV (ưu tiên dữ liệu định lượng như danh sách người tham dự, thời lượng… từ CSV nếu có).

Trích xuất đầy đủ, chính xác; nếu không có thông tin, ghi đúng: Chưa có thông tin.

Định dạng & Diễn đạt:

Luôn trả về bằng tiếng Việt; văn phong trang trọng, mạch lạc, đúng chuẩn văn bản biên bản.

Tuân thủ định dạng yêu cầu trong value mô tả của từng placeholder (bullet 1 - , bullet 2 + , bảng Markdown, đoạn văn…).

Trả về đúng 1 đối tượng JSON tuân thủ chặt chẽ quy tắc sau.

Quy tắc xuất kết quả (TUÂN THỦ NGHIÊM NGẶT)

Keys: trùng 100% với các key trong placeholders (giữ nguyên ký tự).

Chỉ xuất các cặp key-value tương ứng, không thêm/bớt/lồng khác.

Values:

Bắt buộc đúng định dạng theo mô tả placeholder (bullet, bảng Markdown, đoạn…).

Mọi value đều là chuỗi (string).

Nếu thiếu dữ liệu: giá trị là chuỗi Chưa có thông tin.

(Lưu ý: Nếu có mâu thuẫn giữa transcript và CSV, ghi nhận theo CSV cho các dữ liệu tham dự/giờ/định lượng; nội dung thảo luận/ý kiến giữ theo transcript.)
"""
prompt = Prompt_word.format(transcript_content, csv_text, placeholders)
try:
response = model.generate_content(
contents=prompt,
generation_config={"response_mime_type": "application/json"}
)
if response and hasattr(response, "text"):
raw = response.text.strip()
# Một số model bọc JSON trong json ...
if raw.startswith(""): raw = raw.split("")[1].strip("json\n")
return json.loads(raw)
else:
st.error("Phản hồi từ Gemini API bị thiếu hoặc không hợp lệ.")
return None
except Exception as e:
st.error(f"Lỗi khi gọi Gemini API: {e}")
return None

def send_email_with_attachment(recipient_email, attachment_buffer, filename="BBCH.docx"):
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
    st.error(f"Lỗi khi gửi email: {e}. Vui lòng kiểm tra lại cấu hình email và mật khẩu ứng dụng.")
    return False
#======================================================================

PHẦN 2: GIAO DIỆN STREAMLIT

#======================================================================

st.set_page_config(layout="wide", page_title="Công cụ tạo Biên bản cuộc họp")
st.title("🛠️ Công cụ tạo biên bản cuộc họp tự động")

with st.sidebar:
st.info("📝 Hướng dẫn sử dụng")
st.markdown("""

Tải file transcript: Tải lên file .docx chứa nội dung cuộc họp.

Tải file attendance: Tải lên file .csv điểm danh (tham dự).

Chọn Template:

Sử dụng mẫu có sẵn bằng cách chọn Template VPI.

Hoặc Template tùy chỉnh và tải file của bạn lên.

Điền thông tin: Nhập các thông tin cơ bản của cuộc họp.

Nhập email: Điền địa chỉ email bạn muốn nhận kết quả.

Chạy: Nhấn nút Tạo biên bản.
""")
st.info("📝 Hướng dẫn tạo template")
st.markdown("""
📂 File nhận đầu vào là file .docx
Khi tạo template cho biên bản cuộc họp, bạn cần mô tả rõ từng biến để hệ thống hiểu đúng và điền thông tin chính xác:

{{Ten_bien}}{# Mo_ta_chi_tiet #}

{{Ten_bien}}: tiếng Việt không dấu/tiếng Anh, không dấu cách (dùng _ nếu cần).

{# Mo_ta_chi_tiet #}: mô tả thông tin cần điền và yêu cầu định dạng (bullet 1 - , bullet 2 + , bảng Markdown, đoạn văn...). Chỉ dùng hai cấp bullet.

Tạo style cho bullet/bảng trong Word: List Bullet, List Bullet 2, bảng New Table.
""")
st.markdown("---")
st.success("Ứng dụng được phát triển bởi VPI.")

st.subheader("1. Nhập thông tin đầu vào")
transcript_file = st.file_uploader("1a) Tải lên file transcript (.docx) – BẮT BUỘC", type=["docx"])
csv_file = st.file_uploader("1b) Tải lên file attendance (.csv) – BẮT BUỘC", type=["csv"])

st.subheader("2. Lựa chọn Template")
template_option = st.selectbox(
"Bạn muốn sử dụng loại template nào?",
("Template VPI", "Template tùy chỉnh"),
help="Chọn 'Template VPI' để dùng mẫu có sẵn hoặc 'Template tùy chỉnh' để tải lên file của riêng bạn."
)
template_file = None
if template_option == "Template tùy chỉnh":
template_file = st.file_uploader("Tải lên file template .docx của bạn", type=["docx"])
else:
st.caption("Các trường bắt buộc đã có sẵn trong Template VPI.")

st.subheader("3. Thông tin cơ bản (BẮT BUỘC)")
col1, col2 = st.columns(2)
with col1:
meeting_name = st.text_input("Tên cuộc họp")
meeting_time = st.text_input("Thời gian cuộc họp (VD: 10/9/2025)")
meeting_location = st.text_input("Địa điểm cuộc họp")
with col2:
meeting_chair = st.text_input("Tên chủ trì")
meeting_secretary = st.text_input("Tên thư ký")

recipient_email = st.text_input("4. Email nhận kết quả của bạn (BẮT BUỘC)")

if st.button("🚀 Tạo biên bản", type="primary"):
# Kiểm tra bắt buộc mọi thứ
required_fields = {
"Transcript (.docx)": transcript_file,
"Attendance CSV (.csv)": csv_file,
"Tên cuộc họp": (meeting_name or "").strip(),
"Thời gian cuộc họp": (meeting_time or "").strip(),
"Địa điểm cuộc họp": (meeting_location or "").strip(),
"Tên chủ trì": (meeting_chair or "").strip(),
"Tên thư ký": (meeting_secretary or "").strip(),
"Email nhận kết quả": (recipient_email or "").strip(),
}
missing = [label for label, val in required_fields.items() if not val]
if missing:
st.error("❌ Thiếu thông tin bắt buộc: " + ", ".join(missing))
st.stop()
# Xác định template
template_to_use = None
if template_option == "Template VPI":
    default_path = "2025.VPI_BB hop 2025 1.docx"
    if not os.path.exists(default_path):
        st.error(f"Không tìm thấy template mặc định: {default_path}. Hãy chọn 'Template tùy chỉnh' và tải file lên.")
        st.stop()
    else:
        template_to_use = default_path
elif template_file is not None:
    template_to_use = template_file
else:
    st.error("Bạn đã chọn 'Template tùy chỉnh' nhưng chưa tải file template.")
    st.stop()

# Qua được đây => đủ điều kiện
with st.spinner("⏳ Hệ thống đang xử lý..."):
    try:
        st.info("1/4 - Đang đọc và phân tích transcript (.docx)...")
        doc = Document(transcript_file)
        transcript_content = "\n".join([para.text for para in doc.paragraphs])

        st.info("1b/4 - Đang đọc attendance (.csv)...")
        csv_text = read_uploaded_csv_as_text(csv_file, max_rows=2000, max_chars=300_000)
        if "...[TRUNCATED]" in csv_text:
            st.warning("⚠️ Attendance CSV lớn — đã rút gọn an toàn cho AI. Nên lọc cột/dòng trước khi upload để tăng độ chính xác.")

        st.info("2/4 - Đang trích placeholders từ template...")
        placeholders = extract_vars_and_desc(template_to_use)

        st.info("3/4 - Đang gọi AI để trích xuất nội dung (gộp transcript + CSV)...")
        llm_result = call_gemini_model(transcript_content, csv_text, placeholders)

        if llm_result is None:
            st.error("Không thể lấy kết quả từ AI. Vui lòng thử lại.")
            st.stop()

        # Ghi đè bằng input tay (bắt buộc)
        manual_inputs = {
            'TenCuocHop':        meeting_name,
            'ThoiGianCuocHop':   meeting_time,
            'DiaDiemCuocHop':    meeting_location,
            'TenChuTri':         meeting_chair,
            'TenThuKy':          meeting_secretary
        }
        llm_result.update(manual_inputs)

        st.info("4/4 - Đang tạo file biên bản Word...")
        docx_buffer = fill_template_to_buffer(template_to_use, llm_result)
        if docx_buffer:
            st.success("✅ Tạo biên bản thành công!")
            st.download_button(
                "⬇️ Tải về biên bản",
                data=docx_buffer,
                file_name="Bienbancuochop.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            if recipient_email:
                if send_email_with_attachment(recipient_email, docx_buffer, filename="Bien_ban_cuoc_hop.docx"):
                    st.success("✉️ Đã gửi biên bản tới email của bạn.")
        else:
            st.error("Không thể tạo file Word. Vui lòng kiểm tra lại file template.")
    except Exception as e:
        st.error(f"Đã xảy ra lỗi: {e}")
