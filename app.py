import os
import re
import json
import time
import hashlib
from io import BytesIO
from datetime import datetime, date
from pathlib import Path

import streamlit as st
import google.generativeai as genai

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# ==========================================
# CẤU HÌNH LƯU TRỮ & API KEY
# ==========================================
st.set_page_config(
    page_title="Trợ Lý Soạn Giáo Án Mầm Non",
    page_icon="👶",
    layout="centered",
    initial_sidebar_state="expanded",
)

CONFIG_DIR = Path.home() / ".giao_an_mam_non"
USAGE_FILE = CONFIG_DIR / "usage.json"
CONFIG_FILE = CONFIG_DIR / "config.json"

AGE_OPTIONS = [
    "Nhà trẻ (24 - 36 tháng)",
    "Mẫu giáo bé (3 - 4 tuổi)",
    "Mẫu giáo nhỡ (4 - 5 tuổi)",
    "Mẫu giáo lớn (5 - 6 tuổi)",
]

DOMAIN_OPTIONS = [
    "Phát triển thể chất (Tổng hợp)", "Phát triển nhận thức (Tổng hợp)", "Phát triển ngôn ngữ (Tổng hợp)",
    "Phát triển tình cảm và kỹ năng xã hội (Tổng hợp)", "Phát triển thẩm mỹ (Tổng hợp)",
    "Thể chất - Thể dục", "Thể chất - Vận động cơ bản", "Thể chất - Vận động tinh", 
    "Thể chất - Dinh dưỡng và sức khỏe", "Thể chất - Vệ sinh cá nhân", "Thể chất - An toàn",
    "Nhận thức - Làm quen với Toán", "Nhận thức - Khám phá khoa học", "Nhận thức - Khám phá môi trường xung quanh",
    "Nhận thức - Tư duy logic", "Nhận thức - Không gian và thời gian", "Nhận thức - Hình dạng và kích thước",
    "Nhận thức - Số lượng", "Nhận thức - Giáo dục STEM/STEAM",
    "Ngôn ngữ - Văn học (Thơ, Truyện)", "Ngôn ngữ - Kể chuyện sáng tạo", "Ngôn ngữ - Đọc thơ", "Ngôn ngữ - Đồng dao",
    "Ngôn ngữ - Giao tiếp", "Ngôn ngữ - Nghe hiểu", "Ngôn ngữ - Nói mạch lạc", "Ngôn ngữ - Làm quen chữ viết",
    "Ngôn ngữ - Làm quen chữ cái", "Ngôn ngữ - Tập tô",
    "TC-XH - Kỹ năng sống", "TC-XH - Kỹ năng tự phục vụ", "TC-XH - Giao tiếp ứng xử", "TC-XH - Quản lý cảm xúc",
    "TC-XH - Tình cảm gia đình", "TC-XH - Tình bạn", "TC-XH - Kỷ luật", "TC-XH - Bảo vệ môi trường", "TC-XH - Giáo dục giới tính",
    "Thẩm mỹ - Tạo hình (Vẽ)", "Thẩm mỹ - Tạo hình (Nặn)", "Thẩm mỹ - Tạo hình (Cắt dán)", "Thẩm mỹ - Âm nhạc (Hát)",
    "Thẩm mỹ - Âm nhạc (Vận động theo nhạc)", "Thẩm mỹ - Âm nhạc (Nghe nhạc)", "Thẩm mỹ - Cảm thụ cái đẹp", "Thẩm mỹ - Sáng tạo nghệ thuật",
    "Giáo dục Montessori", "Giáo dục Reggio Emilia", "Trò chơi đóng vai", "Trò chơi dân gian"
]

DURATION_OPTIONS = ["15 - 20 phút", "20 - 25 phút", "25 - 30 phút", "30 - 35 phút", "35 - 40 phút"]

MODEL_LIMITS = {
    "gemini-3.5-flash": 1500, "gemini-2.5-flash": 1500, "gemini-1.5-flash-latest": 1500,
    "gemini-1.5-flash": 1500, "gemini-1.5-pro": 50, "gemini-pro": 50
}
AUTO_MODELS = list(MODEL_LIMITS.keys())

TOPIC_OPTIONS = [
    "Trường Mầm non của bé", "Lớp học của bé", "Cô giáo của bé", "Đồ dùng, đồ chơi trong trường mầm non", "Bé và các bạn",
    "Cơ thể của bé", "Các giác quan của bé", "Bé cần gì để lớn lên và khỏe mạnh", "Trang phục của bé", "Bé biết tự phục vụ",
    "Gia đình của bé", "Ngôi nhà của bé", "Đồ dùng trong gia đình", "Những người thân trong gia đình", "Ngày hội của gia đình",
    "Thế giới thực vật", "Cây xanh và môi trường sống", "Các loại hoa", "Rau củ quả quanh bé", "Một số loại quả",
    "Thế giới động vật", "Động vật nuôi trong gia đình", "Động vật sống trong rừng", "Động vật sống dưới nước", "Côn trùng và chim",
    "Giao thông", "Phương tiện giao thông đường bộ", "Phương tiện giao thông đường hàng không", "Phương tiện giao thông đường thủy", "Luật lệ giao thông",
    "Nghề nghiệp", "Nghề giáo viên", "Nghề bác sĩ", "Nghề nông", "Nghề xây dựng", 
    "Nghề bộ đội", "Hiện tượng tự nhiên", "Các mùa trong năm", "Mùa xuân", "Mùa hè", 
    "Mùa thu", "Mùa đông", "Nước và các hiện tượng tự nhiên", "Ngày và đêm - Mặt trời và Mặt trăng", "Đất, đá, cát, sỏi",
    "Quê hương - Đất nước - Bác Hồ", "Quê hương của bé", "Thủ đô Hà Nội", "Bác Hồ kính yêu", "Bé chuẩn bị vào lớp 1"
]

FAMILY_TOPICS = [
    "Kể về gia đình của bé", "Bé yêu ông bà", "Bé yêu bố mẹ", "Bé chăm sóc em nhỏ", "Bé giúp mẹ việc nhà",
    "Bé dọn dẹp nhà cửa cùng bố mẹ", "Bữa cơm ấm áp của gia đình", "Bé làm thiệp tặng mẹ ngày 8/3", "Bé làm quà tặng bố", "Bé kể chuyện về ngôi nhà của bé",
    "Bé biết vâng lời bố mẹ", "Bé giúp bà tưới cây", "Bé sắp xếp quần áo gọn gàng", "Bé gọi điện hỏi thăm ông bà", "Bé an ủi khi mẹ buồn",
    "Bé chơi cùng anh chị em trong nhà", "Ngày Tết sum họp gia đình", "Bé nhận biết đồ dùng trong gia đình", "Bé giữ gìn đồ chơi trong gia đình", "Gia đình bé đi dã ngoại (đóng vai)",
    "Bé yêu thương thú cưng trong nhà", "Bé cùng bố mẹ nấu ăn", "Bé nhận biết công việc của bố mẹ", "Bé chia sẻ cảm xúc với gia đình", "Trò chuyện về những người thân yêu"
]

LIFE_SKILLS_TOPICS = [
    "Bé tự rửa tay bằng xà phòng", "Bé tự đánh răng đúng cách", "Bé tự xúc cơm ăn gọn gàng", "Bé tự mặc quần áo", "Bé tự đi giày dép đúng chiều",
    "Bé tự cất đồ chơi sau khi chơi", "Bé biết xếp hàng chờ đến lượt", "Bé biết nói lời cảm ơn", "Bé biết nói lời xin lỗi", "Bé biết chào hỏi lễ phép",
    "Bé biết xin phép người lớn", "Bé phòng tránh nguy hiểm (ổ điện, bếp lửa)", "Bé không đi theo người lạ", "Bé biết gọi người lớn khi gặp nguy hiểm", "Bé biết bảo vệ bản thân",
    "Bé biết chia sẻ đồ chơi với bạn", "Bé biết lắng nghe người khác nói", "Bé biết nói lên cảm xúc của mình", "Bé kiềm chế cơn giận (không đánh bạn)", "Bé giữ trật tự nơi công cộng",
    "Bé bỏ rác đúng nơi quy định", "Bé tiết kiệm nước khi rửa tay", "Bé phân loại rác đơn giản", "Bé tự chuẩn bị đồ dùng đi học", "Bé sắp xếp góc học tập gọn gàng"
]

TOPICS_BY_DOMAIN = {
    "Phát triển thể chất": [
        "Vận động: Đi trong đường hẹp", "Bò chui qua cổng", "Ném bóng vào rổ", "Chạy theo hiệu lệnh của cô", "Bật nhảy qua vật cản nhỏ",
        "Tung và bắt bóng tại chỗ", "Đi trên ghế thể dục", "Bò ziczac qua chướng ngại vật", "Chuyền bóng qua đầu", "Thể dục sáng với vòng",
        "Đi khuỵu gối", "Bật xa tại chỗ", "Bò bằng bàn tay và bàn chân", "Ném xa bằng một tay", "Chuyền bóng qua chân",
        "Đi trên ván dốc", "Nhảy lò cò một chân", "Trườn sấp kết hợp trèo qua ghế", "Chạy nhanh về đích", "Vận động với dây thừng",
    ],
    "Phát triển nhận thức (Khám phá khoa học / Làm quen với toán)": [
        "Khám phá rau củ quả quanh bé", "So sánh to - nhỏ", "Đếm số lượng từ 1 đến 5", "Nhận biết hình tròn và hình vuông", "Khám phá các loại hoa",
        "Tìm hiểu về nước và ích lợi của nước", "Nhận biết con vật nuôi trong gia đình", "Khám phá nghề nông", "Phân biệt màu đỏ - vàng - xanh", "Sắp xếp theo quy tắc đơn giản",
        "Phân biệt cao - thấp", "Nhận biết số lượng 1 - 3", "Nhận biết hình tam giác và hình chữ nhật", "Khám phá các mùa trong năm", "Tìm hiểu về phương tiện giao thông đường bộ",
        "Khám phá cơ thể của bé", "So sánh dài - ngắn", "Tìm hiểu về các loại quả", "Khám phá nghề bác sĩ", "Nhận biết buổi sáng - trưa - tối",
    ],
    "Phát triển ngôn ngữ (Thơ, truyện, kể chuyện)": [
        "Đọc thơ 'Bắp cải xanh'", "Kể chuyện 'Chú dê đen'", "Kể chuyện 'Cáo và Thỏ'", "Đồng dao 'Dung dăng dung dẻ'", "Kể chuyện 'Gà con tìm mẹ'",
        "Đọc thơ 'Cái bát'", "Truyện 'Sự tích hoa hồng'", "Kể chuyện 'Thỏ và Rùa'", "Đọc thơ 'Yêu mẹ'", "Kể chuyện 'Ba chú lợn con'",
        "Đọc thơ 'Cái bống'", "Kể chuyện 'Sói và Cừu'", "Đọc thơ 'Ông mặt trời'", "Kể chuyện 'Rùa con tìm nhà'", "Đọc thơ 'Bác nông dân'",
        "Kể chuyện 'Ai đáng khen nhiều hơn'", "Đồng dao 'Chi chi chành chành'", "Đọc thơ 'Đi học'", "Kể chuyện 'Gấu con chia quà'", "Kể chuyện theo tranh",
    ],
    "Phát triển tình cảm và kỹ năng xã hội": [
        "Bé chào hỏi người lớn", "Chia sẻ đồ chơi với bạn", "Bé biết cảm ơn và xin lỗi", "Bé yêu mẹ", "Bé quan tâm ông bà",
        "Bé giúp đỡ bạn", "Bé giữ gìn đồ chơi chung", "Bé yêu thiên nhiên", "Bé xếp hàng chờ đợi", "Bé giữ trật tự khi ăn",
        "Bé biết lắng nghe người khác", "Bé quan tâm bạn bè", "Bé giữ lời hứa", "Bé biết chờ đến lượt", "Bé yêu quý cô giáo",
        "Bé chăm sóc cây xanh", "Bé biết chia sẻ cảm xúc", "Bé biết tự phục vụ", "Bé yêu thương em nhỏ", "Bé biết bảo vệ môi trường",
    ],
    "Phát triển thẩm mỹ (Tạo hình, Âm nhạc)": [
        "Vẽ mưa rơi", "Nặn quả cam", "Xé dán đám mây", "Vẽ hoa tặng mẹ", "Dạy hát 'Cả nhà thương nhau'",
        "Dạy hát 'Bắp cải xanh'", "Vận động theo nhạc 'Đàn gà con'", "Tô màu bức tranh phong cảnh", "In hình từ củ quả", "Dạy hát 'Cháu yêu bà'",
        "Vẽ ông mặt trời", "Nặn con vật yêu thích", "Cắt dán ngôi nhà", "Vẽ cây xanh", "Dạy hát 'Cô giáo miền xuôi'",
        "Vận động theo nhạc 'Em yêu cây xanh'", "Tô màu con vật", "Vẽ chân dung bạn", "Nặn cái bát", "Gấp giấy hình con thuyền",
    ],
}

def get_topics_for_domain(domain):
    d_lower = domain.lower()
    if "gia đình" in d_lower: return FAMILY_TOPICS
    elif "kỹ năng sống" in d_lower or "kỹ năng tự phục vụ" in d_lower: return LIFE_SKILLS_TOPICS
    elif "tình cảm" in d_lower or "tc-xh" in d_lower: return FAMILY_TOPICS + LIFE_SKILLS_TOPICS + TOPICS_BY_DOMAIN["Phát triển tình cảm và kỹ năng xã hội"]
    elif "thể chất" in d_lower or "thể dục" in d_lower: return TOPICS_BY_DOMAIN["Phát triển thể chất"]
    elif "nhận thức" in d_lower or "toán" in d_lower or "khoa học" in d_lower or "stem" in d_lower: return TOPICS_BY_DOMAIN["Phát triển nhận thức (Khám phá khoa học / Làm quen với toán)"]
    elif "ngôn ngữ" in d_lower or "văn học" in d_lower or "chữ" in d_lower: return TOPICS_BY_DOMAIN["Phát triển ngôn ngữ (Thơ, truyện, kể chuyện)"]
    elif "thẩm mỹ" in d_lower or "tạo hình" in d_lower or "âm nhạc" in d_lower: return TOPICS_BY_DOMAIN["Phát triển thẩm mỹ (Tạo hình, Âm nhạc)"]
    elif "montessori" in d_lower: return ["Hoạt động thực hành cuộc sống", "Giác quan", "Ngôn ngữ", "Toán học", "Văn hóa"]
    elif "reggio" in d_lower: return ["Khám phá ánh sáng", "Khám phá thiên nhiên", "Nghệ thuật tái chế", "Dự án nhóm"]
    elif "dân gian" in d_lower: return ["Rồng rắn lên mây", "Lộn cầu vồng", "Chi chi chành chành", "Dung dăng dung dẻ", "Nu na nu nống"]
    elif "đóng vai" in d_lower: return ["Bác sĩ", "Cô giáo", "Đầu bếp", "Bán hàng", "Kỹ sư xây dựng"]
    else: return ["Hoạt động phát triển toàn diện", "Khám phá môi trường xung quanh", "Trò chơi học tập", "Hoạt động trải nghiệm"]

# ============ HÀM XỬ LÝ HỆ THỐNG ============
def get_api_key():
    # 1. Ưu tiên Streamlit Secrets (Dùng khi deploy web giấu key)
    try:
        if "GEMINI_API_KEY" in st.secrets:
            return st.secrets["GEMINI_API_KEY"]
    except Exception:
        pass
    
    # 2. Biến môi trường máy tính
    env_key = os.environ.get("GEMINI_API_KEY")
    if env_key: return env_key

    # 3. Lưu từ file cá nhân
    try:
        if CONFIG_FILE.exists():
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8")).get("api_key", "")
    except Exception: pass
    return ""

def get_today_usage() -> dict:
    today_str = date.today().isoformat()
    try:
        if USAGE_FILE.exists():
            data = json.loads(USAGE_FILE.read_text(encoding="utf-8"))
            if data.get("date") == today_str: return data.get("usage", {})
    except Exception: pass
    return {}

def increment_usage(model_name: str):
    today_str = date.today().isoformat()
    usage = get_today_usage()
    usage[model_name] = usage.get(model_name, 0) + 1
    try:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        USAGE_FILE.write_text(json.dumps({"date": today_str, "usage": usage}, ensure_ascii=False), encoding="utf-8")
    except Exception: pass

# ============ XUẤT FILE WORD CHUYÊN NGHIỆP ============
def safe_slug(text: str) -> str:
    text = re.sub(r"[^\w\s\-]", "", text, flags=re.UNICODE)
    text = re.sub(r"\s+", "_", text.strip())
    return text[:80] or "giao_an"

def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def add_styled_run(paragraph, text, bold=False, color=None, size=13):
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor(*color)
    return run

def markdown_to_docx(md_text: str, title: str) -> bytes:
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(13)
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_styled_run(p_title, title.upper(), bold=True, color=(0, 51, 102), size=16)
    
    md_text = md_text.replace("<br>", "\n").replace("<br/>", "\n").replace("<br />", "\n")
    
    lines = md_text.split("\n")
    i = 0
    while i < len(lines):
        s = lines[i].strip()

        # Xử lý Bảng
        if s.startswith("|") and s.endswith("|"):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
                    rows.append(cells)
                i += 1
            if rows:
                n = max(len(r) for r in rows)
                table = doc.add_table(rows=len(rows), cols=n)
                table.style = 'Table Grid'
                table.autofit = True
                
                for r_idx, row in enumerate(rows):
                    for c_idx in range(n):
                        cell_text = row[c_idx] if c_idx < len(row) else ""
                        cell = table.cell(r_idx, c_idx)
                        p = cell.paragraphs[0]
                        
                        if r_idx == 0:
                            set_cell_background(cell, "D9E2F3") # Màu nền xanh nhạt cho tiêu đề
                            add_styled_run(p, cell_text, bold=True)
                        else:
                            add_styled_run(p, cell_text)
            continue

        # Xử lý Tiêu đề
        if s.startswith("# ") or s.startswith("## ") or s.startswith("### "):
            h_text = s.replace("#", "").strip()
            p = doc.add_paragraph()
            add_styled_run(p, h_text, bold=True, color=(0, 102, 204), size=14)
        
        # Xử lý Gạch đầu dòng & Đánh số
        elif s.startswith(("- ", "* ")):
            p = doc.add_paragraph(style="List Bullet")
            add_styled_run(p, s[2:])
        elif re.match(r"^\d+\.\s", s):
            p = doc.add_paragraph(style="List Number")
            add_styled_run(p, re.sub(r"^\d+\.\s+", "", s))
        
        # Đoạn văn bản thường
        elif s:
            p = doc.add_paragraph()
            parts = re.split(r"(\*\*.+?\*\*)", s)
            for part in parts:
                if part.startswith("**") and part.endswith("**") and len(part) > 4:
                    add_styled_run(p, part[2:-2], bold=True)
                elif part:
                    add_styled_run(p, part)
        else:
            doc.add_paragraph("")
        i += 1

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

def build_prompt(topic, age_group, domain, activity_name, duration, extra_notes):
    extra = (extra_notes or "").strip() or "(không có)"
    return f"""Bạn là chuyên gia phương pháp Giáo dục Mầm non Việt Nam, am hiểu Chương trình GDMN hiện hành của Bộ GD&ĐT và Thông tư 17/2009/TT-BGDĐT.

NHIỆM VỤ: Soạn 01 giáo án hoạt động học hoàn chỉnh, dạy được ngay, theo đúng phom giáo án mầm non Việt Nam.

THÔNG TIN ĐẦU VÀO:
- Chủ đề: {topic}
- Độ tuổi: {age_group}
- Lĩnh vực phát triển: {domain}
- Tên hoạt động / Đề tài: {activity_name}
- Thời lượng: {duration}
- Yêu cầu bổ sung: {extra}

NGUYÊN TẮC BẮT BUỘC:
1. Nội dung phù hợp đặc điểm tâm sinh lý lứa tuổi (thời gian chú ý, ngôn ngữ, vận động).
2. Lấy trẻ làm trung tâm: tối thiểu 60% thời gian là trẻ hoạt động, cô chỉ hướng dẫn.
3. Tích hợp nhẹ các lĩnh vực khác vào hoạt động chính.
4. KHÔNG bịa tên bài hát, bài thơ, trò chơi không tồn tại.
5. Có phương án dự phòng khi trẻ không hứng thú hoặc hoạt động kéo dài quá giờ.
6. Ngôn ngữ sư phạm, ngắn gọn, dễ đọc, dễ in.
7. ĐẶC BIỆT: KHÔNG sử dụng thẻ HTML như <br>, <p>, <div> để xuống dòng. Chỉ dùng dấu Enter (xuống dòng) thuần túy. Nếu muốn xuống dòng trong bảng, hãy dùng dấu chấm phẩy (;).

ĐỊNH DẠNG ĐẦU RA (Markdown, giữ đúng thứ tự):

# GIÁO ÁN: {activity_name}

- **Lĩnh vực:** {domain}
- **Độ tuổi:** {age_group}
- **Chủ đề:** {topic}
- **Thời gian:** {duration}

## I. MỤC ĐÍCH - YÊU CẦU
### 1. Kiến thức
### 2. Kỹ năng
### 3. Thái độ

## II. CHUẨN BỊ
### 1. Đồ dùng của cô
### 2. Đồ dùng của trẻ
### 3. Địa điểm, không gian lớp học

## III. TIẾN TRÌNH HOẠT ĐỘNG

### 1. Ổn định tổ chức, gây hứng thú (khoảng x phút)
| Hoạt động của cô | Hoạt động của trẻ |
|---|---|
| ... | ... |

### 2. Nội dung trọng tâm (khoảng x phút)
| Hoạt động của cô | Hoạt động của trẻ |
|---|---|
| ... | ... |

### 3. Luyện tập, củng cố (khoảng x phút)
| Hoạt động của cô | Hoạt động của trẻ |
|---|---|
| ... | ... |

### 4. Kết thúc hoạt động (khoảng x phút)
| Hoạt động của cô | Hoạt động của trẻ |
|---|---|
| ... | ... |

## IV. ĐÁNH GIÁ CUỐI GIỜ
- **Ưu điểm dự kiến:**
- **Hạn chế dự kiến:**
- **Điều chỉnh cho lần dạy sau:**

## V. LƯU Ý AN TOÀN & PHƯƠNG ÁN DỰ PHÒNG
"""

# ============ LẤY KEY TỰ ĐỘNG ============
SECRET_KEY = get_api_key()

# ============ SIDEBAR ============
with st.sidebar:
    st.header("⚙️ Cài đặt hệ thống")
    
    # TỰ ĐỘNG ẨN Ô NHẬP NẾU ĐÃ CÓ KEY TRÊN MÂY
    if SECRET_KEY:
        api_key = SECRET_KEY
        st.success("🔒 Hệ thống đã cấu hình sẵn API Key. Bạn không cần nhập!")
    else:
        if "gemini_key" not in st.session_state:
            st.session_state["gemini_key"] = ""
        api_key = st.text_input("Google Gemini API Key", type="password", key="gemini_key")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("💾 Lưu key", use_container_width=True):
                if api_key.strip():
                    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
                    CONFIG_FILE.write_text(json.dumps({"api_key": api_key.strip()}), encoding="utf-8")
                    st.success("Đã lưu key!")
        with col2:
            if st.button("🗑️ Xóa key", use_container_width=True):
                st.session_state["gemini_key"] = ""
                if CONFIG_FILE.exists(): CONFIG_FILE.unlink()
                st.rerun()

    st.markdown("---")
    
    st.subheader("📊 Hạn mức hôm nay")
    today_usage = get_today_usage()
    if not today_usage: st.caption("Chưa có dữ liệu sử dụng.")
    else:
        for m_name, limit in MODEL_LIMITS.items():
            used = today_usage.get(m_name, 0)
            remain = limit - used
            if used > 0:
                if remain > (limit * 0.2): st.success(f"🟢 **{m_name}**: {used} / {limit}")
                elif remain > 0: st.warning(f"🟡 **{m_name}**: {used} / {limit}")
                else: st.error(f"🔴 **{m_name}**: Đã hết lượt")

    with st.expander("Tùy chọn nâng cao"):
        st.info("Hệ thống Tự động dò tìm Model.")
        temperature = st.slider("Độ sáng tạo", 0.0, 1.0, 0.6, 0.05)
        max_tokens = st.number_input("Giới hạn tokens", min_value=1024, max_value=8192, value=4096, step=512)

# ============ HEADER & FORM ============
st.title("👶 Trợ Lý AI Soạn Giáo Án Mầm Non")
st.write("Tạo nhanh bản nháp giáo án chi tiết, chuẩn phom mẫu GDMN Việt Nam.")

st.subheader("📋 Thông tin bài dạy")
col1, col2 = st.columns(2)

with col1:
    chosen_topic = st.selectbox("Chủ đề bài học *", TOPIC_OPTIONS + ["✏️ Tự nhập chủ đề khác..."], key="topic_select")
    if chosen_topic.startswith("✏️"):
        topic_name = st.text_input("Nhập chủ đề của bạn *", key="topic_custom", max_chars=200)
    else: 
        topic_name = chosen_topic
    st.selectbox("Độ tuổi *", AGE_OPTIONS, key="age_input")

with col2:
    domain = st.selectbox("Lĩnh vực phát triển *", DOMAIN_OPTIONS, key="domain_input")
    st.selectbox("Thời lượng dự kiến", DURATION_OPTIONS, key="duration_input")

domain_key = hashlib.md5(domain.encode("utf-8")).hexdigest()[:8]
topics = get_topics_for_domain(domain)
topic_choices = list(topics) + ["✏️ Tự nhập đề tài khác..."]

chosen_activity = st.selectbox("Tên đề tài / hoạt động cụ thể *", topic_choices, index=0, key=f"activity_select_{domain_key}")

if chosen_activity.startswith("✏️"):
    activity_name = st.text_input("Nhập tên đề tài của bạn *", key=f"activity_custom_{domain_key}", max_chars=200)
else:
    activity_name = chosen_activity

st.text_area("Yêu cầu bổ sung hoặc ý tưởng riêng (không bắt buộc)", key="extra_input", placeholder="Ví dụ: Thêm trò chơi vận động ở giữa, dùng học liệu trực quan...", max_chars=1000, height=100)

submitted = st.button("✨ Soạn Giáo Án Ngay", type="primary", use_container_width=True)

# ============ XỬ LÝ ============
if submitted:
    errors = []
    if not api_key or not api_key.strip(): errors.append("Chưa có API Key.")
    if not topic_name or not topic_name.strip(): errors.append("Chưa nhập Chủ đề.")
    if not activity_name or not activity_name.strip(): errors.append("Chưa nhập Tên đề tài.")

    if errors:
        for e in errors: st.error("❌ " + e)
    else:
        extra_val = st.session_state.get("extra_input") or ""
        prompt = build_prompt(
            topic_name.strip(), 
            st.session_state["age_input"], 
            st.session_state["domain_input"], 
            activity_name.strip(), 
            st.session_state["duration_input"], 
            extra_val
        )

        st.session_state["lesson_plan"] = None
        st.markdown("---")
        st.subheader("📝 Giáo án của bạn")
        
        placeholder = st.empty()
        collected = []
        success = False
        used_model = ""
        
        try:
            genai.configure(api_key=api_key.strip())
            with st.status("🚀 Đang kết nối với AI...", expanded=True) as status:
                for model_name in AUTO_MODELS:
                    status.write(f"🔄 Đang kiểm tra model: `{model_name}`...")
                    try:
                        model = genai.GenerativeModel(model_name)
                        start_time = time.time()
                        increment_usage(model_name)
                        
                        response = model.generate_content(
                            prompt,
                            generation_config=genai.types.GenerationConfig(
                                temperature=float(temperature), 
                                max_output_tokens=int(max_tokens)
                            ),
                            stream=True
                        )
                        status.write(f"✅ Model `{model_name}` phản hồi tốt. Đang soạn...")
                        
                        for chunk in response:
                            if chunk.text:
                                collected.append(chunk.text)
                                placeholder.markdown("".join(collected))
                                
                        success = True
                        used_model = model_name
                        status.update(label=f"✅ Xong! Dùng model: **{model_name}**", state="complete", expanded=False)
                        break
                    except Exception as e:
                        err_msg = str(e).lower()
                        if any(err in err_msg for err in ["429", "quota", "exhausted", "404", "not found", "500", "503"]):
                            status.write(f"⚠️ Model `{model_name}` quá tải/hết lượt. Chuyển model khác...")
                            placeholder.empty()
                            collected.clear()
                            time.sleep(1)
                            continue
                        else: 
                            raise e
                            
            if not success: 
                raise RuntimeError("Tất cả các model đều đã bị giới hạn (hết token). Vui lòng thử lại sau.")
                
            full_text = "".join(collected).strip()
            if not full_text: 
                raise RuntimeError("AI không trả về nội dung.")
                
            st.session_state["lesson_plan"] = full_text
            st.session_state["lesson_meta"] = {
                "activity_name": activity_name.strip(), 
                "age_group": st.session_state["age_input"],
                "domain": st.session_state["domain_input"], 
                "model": used_model,
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            }
            st.success(f"🎉 Đã soạn xong bằng model: **{used_model}**")
            st.rerun() 

        except Exception as e:
            msg = str(e)
            st.error(f"❌ Lỗi: {msg}")
            with st.expander("Chi tiết lỗi (debug)"): st.code(msg)

# ============ HIỂN THỊ + TẢI XUỐNG ============
if "lesson_plan" in st.session_state and st.session_state["lesson_plan"]:
    plan = st.session_state["lesson_plan"]
    meta = st.session_state["lesson_meta"]

    if not submitted:
        st.markdown("---")
        st.subheader("📝 Giáo án đã tạo")
        st.caption(f"Tạo lúc {meta.get('created_at')} · **Model: {meta.get('model', '?')}** · Lĩnh vực: {meta.get('domain', '?')}")
        st.markdown(plan)

    st.markdown("---")
    st.subheader("📥 Tải xuống / Sao chép")
    
    base = safe_slug(meta.get("activity_name", "giao_an"))
    tag = datetime.now().strftime("%Y%m%d")
    
    c1, c2, c3 = st.columns(3)
    with c1:
        if DOCX_AVAILABLE:
            st.download_button(
                "📄 Tải .DOCX (Word)",
                data=markdown_to_docx(plan, meta.get("activity_name", "Giáo án mầm non")),
                file_name=f"GiaoAn_{base}_{tag}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        else:
            st.caption("Cài `python-docx` để xuất Word.")
    with c2:
        st.download_button("📝 Tải .MD", data=plan, file_name=f"GiaoAn_{base}_{tag}.md", mime="text/markdown", use_container_width=True)
    with c3:
        st.download_button("📃 Tải .TXT", data=plan, file_name=f"GiaoAn_{base}_{tag}.txt", mime="text/plain", use_container_width=True)

    with st.expander("📋 Xem raw để copy thủ công"):
        st.code(plan, language="markdown")

    if st.button("🗑️ Xóa giáo án hiện tại để soạn bài mới"):
        st.session_state["lesson_plan"] = None
        st.session_state["lesson_meta"] = {}
        st.rerun()

# ============ FOOTER ============
st.markdown("---")
st.caption("💡 **Mẹo:** Đọc lại và điều chỉnh chi tiết chưa phù hợp với lớp mình trước khi in/nộp.")