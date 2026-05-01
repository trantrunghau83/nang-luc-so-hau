import streamlit as st
import io
import os
import re
from docx import Document

# 1. GIAO DIỆN TRẮNG - ĐEN SANG TRỌNG
st.set_page_config(page_title="AI Tích Hợp NLS Sáng Tạo - Thầy Hậu", layout="wide")
st.markdown("""
    <style>
    .stApp { background-color: #FFFFFF !important; }
    h1, h2, h3, h4, p, span, div, label, li, .stMarkdown { color: #000000 !important; }
    .stButton>button { background-color: #047857 !important; color: white !important; font-weight: bold !important; border-radius: 8px; }
    .status-box { padding: 15px; border-radius: 8px; border-left: 5px solid #047857; background-color: #F0FDF4; margin-bottom: 20px;}
    </style>
    """, unsafe_allow_html=True)

st.title("🤖 TRỢ LÝ AI: TÍCH HỢP NLS SÁNG TẠO & SINH ĐỘNG")
st.info("💡 Chế độ: AI tự viết lại nội dung, gợi ý công cụ số dựa trên câu hỏi và nhiệm vụ của giáo viên.")

# Thư viện gợi ý công cụ để giáo án sinh động
TOOLS_MAPPING = {
    "5.1": "Sử dụng sơ đồ Mindmap/Canva để hệ thống hóa thiết bị",
    "5.2": "Thực hành trực tiếp trên MS Excel/Google Sheets",
    "2.5": "Thảo luận nhóm qua Padlet/Zalo/Google Classroom",
    "4.3": "Xem video tình huống trên YouTube/Tiktok giáo dục",
    "3.1": "Sử dụng phần mềm mô phỏng hoặc bảng tính tự động",
    "1.1": "Tìm kiếm nâng cao trên Google Search/Wikipedia",
}

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Dữ liệu nguồn")
    file_phuluc = st.file_uploader("Tải Phụ lục 3 (Bảng mã NLS)", type=["docx"])
with col2:
    st.subheader("2. Giáo án cần xử lý")
    file_giaolan = st.file_uploader("Tải Giáo án gốc", type=["docx"])

def get_nls_data(pl_doc, lesson_name):
    results = []
    found_text = ""
    for table in pl_doc.tables:
        for row in table.rows:
            if len(row.cells) >= 7:
                col3 = row.cells[2].text.lower()
                if lesson_name in col3:
                    found_text += row.cells[6].text + "\n"
    
    pattern = r'(\d+\.\d+\.[A-Za-z0-9]+)\s*[:\-]\s*(.*?)(?=\d+\.\d+\.[A-Za-z0-9]+\s*[:\-]|$)'
    matches = re.findall(pattern, found_text, re.DOTALL)
    for ma, nd in matches:
        results.append({"ma": ma.strip(), "nd": nd.strip()})
    return results

if st.button("🚀 BẮT ĐẦU BIÊN SOẠN GIÁO ÁN"):
    if file_giaolan and file_phuluc:
        doc = Document(file_giaolan)
        pl_doc = Document(file_phuluc)
        
        # Nhận diện bài học
        bai_hoc = ""
        for p in doc.paragraphs[:30]:
            match = re.search(r'(bài\s+\d+)', p.text, re.IGNORECASE)
            if match:
                bai_hoc = match.group(1).lower()
                break
        
        nls_list = get_nls_data(pl_doc, bai_hoc)
        
        if not nls_list:
            st.error(f"Không tìm thấy dữ liệu cho {bai_hoc} trong Phụ lục 3!")
        else:
            # --- PHẦN 1: CHỈNH SỬA MỤC TIÊU (KHÔNG XÁO TRỘN) ---
            target_idx = -1
            for i, p in enumerate(doc.paragraphs):
                if "2.2." in p.text and "tin học" in p.text.lower():
                    target_idx = i; break
            
            if target_idx != -1:
                # Chèn theo đúng thứ tự mảng nls_list
                ref_p = doc.paragraphs[target_idx + 1]
                p_head = ref_p.insert_paragraph_before("2.4. Năng lực số:")
                p_head.runs[0].bold = True
                for item in nls_list:
                    p_item = ref_p.insert_paragraph_before(f"- {item['ma']}: {item['nd']}")
                    p_item.runs[0].italic = True

            # --- PHẦN 2: TÍCH HỢP SÁNG TẠO VÀO TIẾN TRÌNH ---
            inserted_codes = set()
            all_paras = []
            
            # Hàm gom tất cả đoạn văn (kể cả trong bảng)
            def collect_all(doc_obj):
                items = []
                for p in doc_obj.paragraphs: items.append(p)
                for t in doc_obj.tables:
                    for r in t.rows:
                        for c in r.cells:
                            for p in c.paragraphs: items.append(p)
                return items

            all_paras = collect_all(doc)

            for p in all_paras:
                txt_low = p.text.lower()
                # AI tìm cả Chuyển giao nhiệm vụ và các Câu hỏi giáo viên (?)
                if "chuyển giao nhiệm vụ" in txt_low or "?" in txt_low or "câu hỏi:" in txt_low:
                    if len(inserted_codes) < len(nls_list):
                        # Lấy mã NLS chưa dùng
                        available = [it for it in nls_list if it['ma'] not in inserted_codes]
                        if not available: available = nls_list # Tái sử dụng nếu hết
                        
                        item = available[0]
                        prefix = item['ma'][:3] # Lấy "5.1", "2.5"...
                        tool_suggest = TOOLS_MAPPING.get(prefix, "Sử dụng ứng dụng học tập trực tuyến")
                        
                        # AI BIÊN SOẠN LẠI CÂU LỆNH SINH ĐỘNG
                        creative_text = f"\n   => Tích hợp NLS ({item['ma']}): Thông qua việc trả lời câu hỏi và thực hiện nhiệm vụ trên, HS được rèn luyện kỹ năng '{item['nd']}'. GV khuyến khích: {tool_suggest}."
                        
                        run = p.add_run(creative_text)
                        run.italic = True; run.bold = True; run.underline = True
                        inserted_codes.add(item['ma'])

            # Xuất file
            bio = io.BytesIO()
            doc.save(bio)
            st.markdown(f"<div class='status-box'>✅ <b>HOÀN THÀNH:</b> AI đã sắp xếp lại mục tiêu và biên soạn lại {len(inserted_codes)} câu tích hợp kèm gợi ý công cụ số.</div>", unsafe_allow_html=True)
            st.download_button("📥 Tải giáo án 'Sinh động' ngay", bio.getvalue(), f"Giao_an_{bai_hoc}_Creative_NLS.docx")
    else:
        st.warning("Anh Hậu ơi, tải file lên để em 'phù phép' cho giáo án nhé!")
