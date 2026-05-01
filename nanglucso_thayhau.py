import streamlit as st
import io
import os
import re
from docx import Document

# 1. GIAO DIỆN CHUẨN TRẮNG - ĐEN
st.set_page_config(page_title="AI Tích Hợp NLS - Thầy Hậu", layout="wide")
st.markdown("""
    <style>
    .stApp { background-color: #FFFFFF !important; }
    h1, h2, h3, h4, p, span, div, label, li, .stMarkdown { color: #000000 !important; }
    .stButton>button { background-color: #1E40AF !important; color: white !important; font-weight: bold !important; border-radius: 8px; }
    .status-box { padding: 15px; border-radius: 8px; border-left: 5px solid #1E40AF; background-color: #F8F9FA; margin-bottom: 20px;}
    </style>
    """, unsafe_allow_html=True)

st.title("🤖 TRỢ LÝ AI: TÍCH HỢP NLS THÔNG MINH")
st.info("💡 Chế độ: Đối chiếu nội dung phù hợp & Đảm bảo số lượng tối thiểu theo Mục tiêu.")

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Dữ liệu nguồn")
    file_phuluc = st.file_uploader("Tải Phụ lục 3 (Bảng mã NLS)", type=["docx"])
    file_khung = st.file_uploader("Tải Khung NLS (Chi tiết)", type=["docx"])
with col2:
    st.subheader("2. Giáo án cần xử lý")
    file_giaolan = st.file_uploader("Tải Giáo án gốc (File Word)", type=["docx"])

def get_nls_from_pl3(pl_doc, lesson_name):
    """Hàm lấy NLS từ cột 3 và cột 7 của Phụ lục 3"""
    results = []
    found_text = ""
    for table in pl_doc.tables:
        for row in table.rows:
            if len(row.cells) >= 7:
                col3 = row.cells[2].text.lower()
                if lesson_name in col3:
                    found_text += row.cells[6].text + "\n"
    
    # Tách mã và nội dung bằng Regex
    pattern = r'(\d+\.\d+\.[A-Za-z0-9]+)\s*[:\-]\s*(.*?)(?=\d+\.\d+\.[A-Za-z0-9]+\s*[:\-]|$)'
    matches = re.findall(pattern, found_text, re.DOTALL)
    for ma, nd in matches:
        results.append({"ma": ma.strip(), "nd": nd.strip()})
    return results

if st.button("🚀 BẮT ĐẦU TÍCH HỢP"):
    if file_giaolan and file_phuluc:
        doc = Document(file_giaolan)
        pl_doc = Document(file_phuluc)
        
        # 1. Tìm tên bài
        bai_hoc = ""
        for p in doc.paragraphs[:30]:
            match = re.search(r'(bài\s+\d+)', p.text, re.IGNORECASE)
            if match:
                bai_hoc = match.group(1).lower()
                break
        
        # 2. Lấy danh sách NLS mục tiêu
        nls_list = get_nls_from_pl3(pl_doc, bai_hoc)
        
        if not nls_list:
            st.error(f"Không tìm thấy dữ liệu cho {bai_hoc} trong Phụ lục 3!")
        else:
            # --- BƯỚC 1: ĐIỀN MỤC TIÊU ---
            # (Giữ nguyên logic chèn vào mục 2.3/2.4 như bản trước)
            target_idx = -1
            for i, p in enumerate(doc.paragraphs):
                if "2.2." in p.text and "tin học" in p.text.lower():
                    target_idx = i; break
            
            if target_idx != -1:
                ref_p = doc.paragraphs[target_idx + 1]
                p_head = ref_p.insert_paragraph_before("2.4. Năng lực số:")
                p_head.runs[0].bold = True
                for item in nls_list:
                    p_item = ref_p.insert_paragraph_before(f"- {item['ma']}: {item['nd']}")
                    p_item.runs[0].italic = True

            # --- BƯỚC 2: TÍCH HỢP VÀO TIẾN TRÌNH (CÓ ĐỐI CHIẾU) ---
            inserted_codes = set()
            all_task_paras = []

            # Thu thập tất cả các đoạn "Chuyển giao nhiệm vụ"
            def collect_tasks(container):
                for p in container.paragraphs:
                    if "chuyển giao nhiệm vụ" in p.text.lower():
                        all_task_paras.append(p)
                if hasattr(container, 'tables'):
                    for table in container.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                collect_tasks(cell)

            collect_tasks(doc)

            # Lặp qua các nhiệm vụ để đối chiếu từ khóa (Matching sơ bộ)
            for p in all_task_paras:
                task_text = p.text.lower()
                for item in nls_list:
                    # Nếu nội dung NLS có từ khóa xuất hiện trong nhiệm vụ (ví dụ: 'thiết bị', 'dữ liệu', 'mạng'...)
                    keywords = [k for k in item['nd'].lower().split() if len(k) > 3]
                    if any(k in task_text for k in keywords) and item['ma'] not in inserted_codes:
                        run = p.add_run(f"\n   => Tích hợp NLS: {item['ma']}: {item['nd']}")
                        run.italic = True; run.bold = True; run.underline = True
                        inserted_codes.add(item['ma'])
                        break

            # BƯỚC 3: ĐẢM BẢO SỐ LƯỢNG TỐI THIỂU
            # Nếu số lượng đã chèn ít hơn số lượng NLS trong mục tiêu
            if len(inserted_codes) < len(nls_list):
                remaining_nls = [it for it in nls_list if it['ma'] not in inserted_codes]
                for p in all_task_paras:
                    if not remaining_nls: break
                    # Nếu đoạn này chưa được gán NLS nào
                    if "Tích hợp NLS:" not in p.text:
                        item = remaining_nls.pop(0)
                        run = p.add_run(f"\n   => Tích hợp NLS: {item['ma']}: {item['nd']}")
                        run.italic = True; run.bold = True; run.underline = True
                        inserted_codes.add(item['ma'])

            # Xuất file
            bio = io.BytesIO()
            doc.save(bio)
            st.success(f"✅ Đã xử lý xong! Tổng cộng đã tích hợp {len(inserted_codes)} vị trí NLS.")
            st.download_button("📥 Tải giáo án đã tích hợp", bio.getvalue(), f"{bai_hoc}_NLS_Chuan.docx")
    else:
        st.warning("Vui lòng tải đủ file Phụ lục 3 và Giáo án.")
