import streamlit as st
import io
import os
import re
from docx import Document

# --- 1. GIAO DIỆN CHUẨN: TRẮNG NỀN - ĐEN CHỮ ---
st.set_page_config(page_title="AI Tích Hợp NLS - Thầy Hậu (Final)", layout="wide")
st.markdown("""
    <style>
    /* Ép nền trắng và chữ đen toàn diện */
    .stApp { background-color: #FFFFFF !important; }
    h1, h2, h3, p, span, div, label, li, .stMarkdown { color: #000000 !important; font-family: 'Times New Roman', serif; }
    .stButton>button { 
        background-color: #1E40AF !important; 
        color: white !important; 
        font-weight: bold !important; 
        border-radius: 8px; 
        width: 100%;
        height: 50px;
    }
    .status-box { 
        padding: 15px; 
        border-radius: 8px; 
        border: 1px solid #1E40AF; 
        background-color: #F0F4FF; 
        margin-bottom: 20px;
        color: #000000 !important;
    }
    /* Đảm bảo file uploader cũng dễ nhìn */
    .stFileUploader label { color: #000000 !important; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.title("🤖 TRỢ LÝ AI: TÍCH HỢP NLS (BẢN HOÀN THIỆN)")
st.info("🎯 Trạng thái: Đã fix màu chữ, vị trí 2.4, lấy đúng Cột 7 và giữ tên file gốc.")

# Thư viện công cụ số sinh động
DIGITAL_TOOLS = {
    "1.1": "Google Search, Wikipedia (tra cứu)",
    "2.5": "Padlet, Zalo, Google Classroom (thảo luận)",
    "3.1": "Canva, MS Word, PowerPoint (sáng tạo)",
    "4.3": "Youtube, Quizizz (an toàn & củng cố)",
    "5.1": "Sơ đồ MindMap, Lucidchart (hệ thống hóa)",
    "5.2": "MS Excel, Google Sheets (xử lý dữ liệu)",
    "5.3": "Phần mềm mô phỏng, BKAV/Kaspersky (bảo mật)"
}

col1, col2 = st.columns(2)
with col1:
    file_phuluc = st.file_uploader("📂 Tải Phụ lục 3 (Mã NLS)", type=["docx"])
with col2:
    file_giaolan = st.file_uploader("📂 Tải Giáo án gốc", type=["docx"])

def get_nls_data(pl_doc, lesson_name):
    """Lấy dữ liệu Cột 7 dựa trên tên bài Cột 3"""
    results = []
    for table in pl_doc.tables:
        for row in table.rows:
            if len(row.cells) >= 7:
                # Cột 3 (Index 2): Tên bài | Cột 7 (Index 6): Nội dung NLS
                c3 = row.cells[2].text.lower()
                if lesson_name in c3:
                    text_c7 = row.cells[6].text
                    # Tìm mã (vd: 5.1.TC1a) và nội dung phía sau
                    matches = re.findall(r'(\d+\.\d+\.[A-Za-z0-9]+)\s*[:\-]\s*(.*?)(?=\d+\.\d+\.[A-Za-z0-9]+\s*[:\-]|$)', text_c7, re.DOTALL)
                    for ma, nd in matches:
                        if ma.strip() not in [x['ma'] for x in results]:
                            results.append({"ma": ma.strip(), "nd": nd.strip()})
    return results

if st.button("🚀 BẮT ĐẦU XỬ LÝ LẦN CUỐI"):
    if file_giaolan and file_phuluc:
        doc = Document(file_giaolan)
        pl_doc = Document(file_phuluc)
        
        # 1. Tên file mới: Giữ nguyên tên gốc + NLS
        original_name = os.path.splitext(file_giaolan.name)[0]
        new_filename = f"{original_name} NLS.docx"

        # 2. Nhận diện bài học để đối chiếu Phụ lục
        bai_hoc = ""
        for p in doc.paragraphs[:30]:
            m = re.search(r'(bài\s+\d+)', p.text, re.IGNORECASE)
            if m: bai_hoc = m.group(1).lower(); break
        
        nls_list = get_nls_data(pl_doc, bai_hoc)
        
        if not nls_list:
            st.error(f"❌ Không tìm thấy thông tin bài '{bai_hoc}' trong Phụ lục 3!")
        else:
            # --- PHẦN 1: CHÈN 2.4 TRƯỚC MỤC 3 ---
            idx_3 = -1
            idx_22 = -1
            for i, p in enumerate(doc.paragraphs):
                txt = p.text.strip().lower()
                if "3." in txt and "phẩm chất" in txt: idx_3 = i
                if "2.2." in txt: idx_22 = i
            
            # Nếu không thấy mục 3, chèn trước phần II
            if idx_3 == -1:
                for i, p in enumerate(doc.paragraphs):
                    if "ii." in p.text.lower(): idx_3 = i; break

            if idx_3 != -1:
                ref_p = doc.paragraphs[idx_3]
                
                # Đảm bảo có mục 2.3
                has_23 = any("2.3." in p.text for p in doc.paragraphs)
                if not has_23 and idx_22 != -1:
                    p23 = ref_p.insert_paragraph_before("2.3. Năng lực khác:")
                    p23.runs[0].bold = True

                # Chèn 2.4. Năng lực số
                p24 = ref_p.insert_paragraph_before("2.4. Năng lực số:")
                p24.runs[0].bold = True
                for item in nls_list:
                    p_item = ref_p.insert_paragraph_before(f"- {item['ma']}: {item['nd']}")
                    p_item.runs[0].italic = True
                ref_p.insert_paragraph_before("") # Khoảng trống

            # --- PHẦN 2: DÀN TRẢI SÁNG TẠO VÀO TIẾN TRÌNH ---
            tasks = []
            def collect_tasks(container):
                for p in container.paragraphs:
                    t = p.text.lower()
                    if "?" in t or "chuyển giao nhiệm vụ" in t: tasks.append(p)
                if hasattr(container, 'tables'):
                    for t in container.tables:
                        for r in t.rows:
                            for c in r.cells: collect_tasks(c)
            collect_tasks(doc)

            if tasks:
                for i, item in enumerate(nls_list):
                    t_idx = (i * len(tasks)) // len(nls_list)
                    p_target = tasks[t_idx]
                    
                    # Lấy công cụ gợi ý
                    tool = DIGITAL_TOOLS.get(item['ma'][:3], "phần mềm học tập")
                    
                    # Viết lại câu tích hợp sinh động
                    note = f"\n   => Tích hợp NLS ({item['ma']}): Thông qua hoạt động trả lời/thảo luận trên, HS phát triển năng lực '{item['nd']}'. GV hướng dẫn sử dụng {tool} để thực hiện sinh động hơn."
                    
                    run = p_target.add_run(note)
                    run.bold = True; run.italic = True; run.underline = True

            # --- XUẤT FILE ---
            bio = io.BytesIO()
            doc.save(bio)
            st.markdown(f"<div class='status-box'><b>✅ HOÀN TẤT:</b> Đã xử lý xong bài {bai_hoc.upper()}. Chữ màu đen, nền trắng, vị trí chuẩn 2.4.</div>", unsafe_allow_html=True)
            st.download_button(
                label=f"📥 TẢI GIÁO ÁN: {new_filename}",
                data=bio.getvalue(),
                file_name=new_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning("Anh Hậu vui lòng tải đủ Phụ lục 3 và Giáo án gốc nhé!")
