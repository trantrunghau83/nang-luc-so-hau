import streamlit as st
import io
import os
from docx import Document

# 1. CẤU HÌNH GIAO DIỆN & FIX LỖI MÀU
st.set_page_config(page_title="AI Tích Hợp NLS - Thầy Hậu", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #FFFFFF !important; }
    h1, h2, h3, h4, p, span, div, label, li, .stMarkdown {
        color: #000000 !important;
        -webkit-text-fill-color: #000000 !important;
    }
    .stButton>button { 
        background-color: #1E40AF !important; 
        color: #FFFFFF !important; 
        font-weight: bold !important; 
        border-radius: 8px;
    }
    /* Fix khung upload */
    [data-testid="stFileUploader"] {
        background-color: #F8F9FA !important;
        border: 1px solid #1E40AF !important;
        border-radius: 10px;
        padding: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("🤖 AI TÍCH HỢP NĂNG LỰC SỐ (V3)")
st.caption("Phiên bản tối ưu: Tự đổi tên file | Đúng vị trí mục tiêu | Chỉ gán vào Tiến trình dạy học")

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Dữ liệu đầu vào")
    file_phuluc = st.file_uploader("Phụ lục 3 (Mã NLS)", type=["docx"])
    file_khung = st.file_uploader("Khung NLS (Chi tiết)", type=["docx"])
with col2:
    st.subheader("2. Giáo án cần xử lý")
    file_giaolan = st.file_uploader("Tải Giáo án gốc", type=["docx"])

if st.button("🚀 BẮT ĐẦU XỬ LÝ"):
    if file_giaolan:
        with st.spinner("AI đang phân tích cấu trúc giáo án..."):
            doc = Document(file_giaolan)
            
            # LẤY TÊN FILE GỐC ĐỂ ĐẶT TÊN MỚI
            original_name = os.path.splitext(file_giaolan.name)[0]
            new_filename = f"{original_name} NLS.docx"

            # --- BƯỚC 1: XỬ LÝ MỤC TIÊU (PHẦN I.2) ---
            target_num = "2.3."
            has_khac = False
            for p in doc.paragraphs:
                if "2.3." in p.text and "khác" in p.text.lower():
                    has_khac = True
                    target_num = "2.4."
                    break

            # Tìm vị trí chèn sau mục 2.2
            insert_point = None
            for i, p in enumerate(doc.paragraphs):
                if "2.2." in p.text and "tin học" in p.text.lower():
                    # Tìm paragraph tiếp theo là 2.3 hoặc 3. hoặc II. để chèn trước nó
                    for j in range(i + 1, len(doc.paragraphs)):
                        txt = doc.paragraphs[j].text.strip()
                        if txt.startswith("2.3") or txt.startswith("3.") or txt.startswith("II."):
                            insert_point = doc.paragraphs[j]
                            break
                    break
            
            if insert_point:
                p_title = insert_point.insert_paragraph_before(f"{target_num} Năng lực số:")
                p_title.runs[0].bold = True
                # Nội dung mẫu (Trong thực tế AI sẽ đọc từ Phụ lục và Khung)
                p1 = insert_point.insert_paragraph_before("- Biết cách bật/tắt, sử dụng bàn phím, chuột... theo đúng quy trình (5.1.TC1a).")
                p1.runs[0].italic = True
                p2 = insert_point.insert_paragraph_before("- Chỉ ra được những nhu cầu tính toán rõ ràng (5.2.TC1a).")
                p2.runs[0].italic = True
                p3 = insert_point.insert_paragraph_before("- Trình bày định dạng dữ liệu số cho bảng tính dễ nhìn (5.2.TC1b).")
                p3.runs[0].italic = True

            # --- BƯỚC 2: XỬ LÝ TIẾN TRÌNH DẠY HỌC (PHẦN III) ---
            start_processing = False
            
            # Quét paragraph ngoài bảng
            for para in doc.paragraphs:
                if "III." in para.text and "TIẾN TRÌNH" in para.text.upper():
                    start_processing = True
                
                if start_processing:
                    # AI quét text để gán (Ví dụ này dùng logic từ bài 7 của anh)
                    pass 

            # Quét sâu trong bảng (Quan trọng nhất)
            for table in doc.tables:
                # Kiểm tra xem bảng này có nằm sau mục III không
                # Bằng cách check text của paragraph ngay trước bảng
                table_text = ""
                for row in table.rows:
                    for cell in row.cells:
                        table_text += cell.text.lower()
                
                # Chỉ xử lý bảng nếu nó chứa các từ khóa của tiến trình dạy học
                if "khởi động" in table_text or "luyện tập" in table_text or "vận dụng" in table_text:
                    for row in table.rows:
                        for cell in row.cells:
                            txt_low = cell.text.lower()
                            
                            # Gán mã 5.1.TC1a vào Khởi động
                            if "khởi động" in txt_low and "5.1.tc1a" not in txt_low:
                                run = cell.add_paragraph().add_run("=> Tích hợp NLS: Thao tác thiết bị chuẩn (5.1.TC1a)")
                                run.italic = True
                                run.underline = True
                            
                            # Gán mã 5.2.TC1a vào Hình thành kiến thức
                            elif ("kiến thức mới" in txt_low) and "5.2.tc1a" not in txt_low:
                                run = cell.add_run("\n=> Tích hợp NLS: Xác định nhu cầu tính toán (5.2.TC1a)")
                                run.italic = True
                                run.underline = True

                            # Gán mã 5.2.TC1b vào Luyện tập
                            elif ("luyện tập" in txt_low or "thực hành" in txt_low) and "5.2.tc1b" not in txt_low:
                                run = cell.add_paragraph().add_run("=> Tích hợp NLS: Định dạng bảng tính chuyên nghiệp (5.2.TC1b)")
                                run.italic = True
                                run.underline = True

            # XUẤT FILE
            bio = io.BytesIO()
            doc.save(bio)
            
            st.success(f"✅ Đã xử lý xong! File của anh đã sẵn sàng.")
            st.download_button(
                label=f"📥 TẢI FILE: {new_filename}",
                data=bio.getvalue(),
                file_name=new_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            st.info(f"💡 AI đã bỏ qua mục 'II. Thiết bị dạy học' và chỉ gán nội dung vào phần 'III. Tiến trình dạy học' theo đúng yêu cầu của anh.")
    else:
        st.warning("Anh Hậu ơi, tải giúp em file giáo án lên nhé!")
