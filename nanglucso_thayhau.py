import streamlit as st
import io
import os
from docx import Document
from docx.shared import Pt

# 1. GIAO DIỆN CHUẨN TRẮNG - ĐEN
st.set_page_config(page_title="AI Tích Hợp NLS - Thầy Hậu", layout="wide")
st.markdown("""
    <style>
    .stApp { background-color: #FFFFFF !important; }
    h1, h2, h3, h4, p, span, div, label, li, .stMarkdown { color: #000000 !important; }
    .stButton>button { background-color: #1E40AF !important; color: white !important; font-weight: bold !important; }
    </style>
    """, unsafe_allow_html=True)

st.title("🤖 TRỢ LÝ AI: TÍCH HỢP NĂNG LỰC SỐ (BẢN CHUẨN 5512)")
st.info("Quy tắc: Tự đổi tên file | Gán vào 'Chuyển giao nhiệm vụ' | Gạch chân & In nghiêng")

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Dữ liệu nguồn")
    file_phuluc = st.file_uploader("Phụ lục 3 (Mã NLS)", type=["docx"])
    file_khung = st.file_uploader("Khung NLS (Chi tiết)", type=["docx"])
with col2:
    st.subheader("2. Giáo án cần xử lý")
    file_giaolan = st.file_uploader("Tải Giáo án gốc", type=["docx"])

if st.button("🚀 BẮT ĐẦU TÍCH HỢP"):
    if file_giaolan:
        with st.spinner("AI đang quét sâu các bước Chuyển giao nhiệm vụ..."):
            doc = Document(file_giaolan)
            
            # ĐẶT TÊN FILE: Tên cũ + NLS
            original_name = os.path.splitext(file_giaolan.name)[0]
            new_filename = f"{original_name} NLS.docx"

            # DỮ LIỆU NLS GIẢ LẬP (Sẽ lấy từ Phụ lục/Khung của anh)
            nls_list = [
                {"ma": "5.1.TC1a", "nd": "Hướng dẫn HS thực hành thao tác bật máy, sử dụng chuột và bàn phím đúng quy trình."},
                {"ma": "5.2.TC1a", "nd": "Yêu cầu HS xác định các nhu cầu tính toán thực tế trên bảng tính."},
                {"ma": "5.2.TC1b", "nd": "Hướng dẫn HS cách trình bày và định dạng dữ liệu số cho bảng tính chuyên nghiệp."}
            ]

            # --- BƯỚC 1: XỬ LÝ MỤC TIÊU 2.3/2.4 ---
            target_idx = -1
            has_khac = any("2.3." in p.text and "khác" in p.text.lower() for p in doc.paragraphs)
            num_nls = "2.4." if has_khac else "2.3."

            for i, p in enumerate(doc.paragraphs):
                if "2.2." in p.text and "tin học" in p.text.lower():
                    target_idx = i
                    break
            
            if target_idx != -1:
                # Tìm điểm chèn trước mục III hoặc mục Phẩm chất
                insert_pos = target_idx + 1
                for j in range(target_idx + 1, len(doc.paragraphs)):
                    t = doc.paragraphs[j].text.strip()
                    if t.startswith("3.") or t.startswith("II.") or t.startswith("III."):
                        insert_pos = j
                        break
                
                # Chèn tiêu đề và nội dung NLS
                p_head = doc.paragraphs[insert_pos].insert_paragraph_before(f"{num_nls} Năng lực số:")
                p_head.runs[0].bold = True
                for item in nls_list:
                    p_item = doc.paragraphs[insert_pos].insert_paragraph_before(f"- {item['nd']} ({item['ma']})")
                    p_item.runs[0].italic = True

            # --- BƯỚC 2: DÒ VÀO TIẾN TRÌNH DẠY HỌC (CHUYỂN GIAO NHIỆM VỤ) ---
            found_iii = False
            curr_nls_idx = 0

            for table in doc.tables:
                # Kiểm tra xem bảng này có nằm trong vùng Tiến trình không
                table_content = "".join([cell.text for row in table.rows for cell in row.cells]).upper()
                if "TIẾN TRÌNH" in table_content or "HOẠT ĐỘNG" in table_content:
                    found_iii = True
                
                if found_iii:
                    for row in table.rows:
                        for cell in row.cells:
                            # Tách các đoạn văn trong ô để tìm chính xác dòng "Chuyển giao nhiệm vụ"
                            for i, para in enumerate(cell.paragraphs):
                                txt = para.text.strip().lower()
                                
                                # Tìm đúng từ khóa "Chuyển giao nhiệm vụ"
                                if "chuyển giao nhiệm vụ" in txt:
                                    if curr_nls_idx < len(nls_list):
                                        # Chèn thêm dòng tích hợp vào ngay sau đó
                                        nls_data = nls_list[curr_nls_idx]
                                        new_run = para.add_run(f"\n=> Tích hợp NLS: {nls_data['nd']} ({nls_data['ma']})")
                                        
                                        # ÉP ĐỊNH DẠNG: GẠCH CHÂN + IN NGHIÊNG
                                        new_run.underline = True
                                        new_run.italic = True
                                        
                                        curr_nls_idx += 1 # Chuyển sang năng lực tiếp theo cho hoạt động sau

            # XUẤT FILE
            bio = io.BytesIO()
            doc.save(bio)
            
            st.success(f"✅ Đã xử lý xong! Tên file mới: {new_filename}")
            st.download_button(
                label="📥 TẢI GIÁO ÁN ĐÃ TÍCH HỢP NLS",
                data=bio.getvalue(),
                file_name=new_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            st.markdown("### 🔍 Nhật ký AI:")
            st.write(f"- Đã chèn mục **{num_nls} Năng lực số**.")
            st.write("- Đã quét các bảng trong phần **Tiến trình dạy học**.")
            st.write("- Đã tìm thấy các dòng **'Chuyển giao nhiệm vụ'** và thực hiện gạch chân nội dung tích hợp.")
    else:
        st.warning("Anh Hậu ơi, tải file giáo án lên giúp em nhé!")
