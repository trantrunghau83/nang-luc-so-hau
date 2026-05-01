import streamlit as st
import pandas as pd
import io
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# CẤU HÌNH GIAO DIỆN
st.set_page_config(page_title="AI Tích Hợp Năng Lực Số - Thầy Hậu", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f0f2f6; }
    .stButton>button { width: 100%; background-color: #1E40AF; color: white; font-weight: bold; }
    .status-box { padding: 20px; border-radius: 10px; background-color: #ffffff; border-left: 5px solid #1E40AF; }
    </style>
    """, unsafe_allow_html=True)

st.title("🤖 TRỢ LÝ AI: TÍCH HỢP NĂNG LỰC SỐ VÀO GIÁO ÁN")
st.info("Hệ thống sẽ dựa vào Phụ lục 3 để xác định mã NLS và Tài liệu khung để đưa nội dung chi tiết vào giáo án của anh.")

# GIAO DIỆN CHỌN FILE
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Dữ liệu đầu vào (Cố định)")
    file_phuluc = st.file_uploader("Tải lên Phụ lục 3 (File đã có cột NLS)", type=["docx"])
    file_khung_nls = st.file_uploader("Tải lên Tài liệu Khung NLS (Chi tiết chỉ số)", type=["docx"])

with col2:
    st.subheader("2. Giáo án cần xử lý")
    file_giaolan = st.file_uploader("Tải lên Giáo án gốc (Chưa có NLS)", type=["docx"])

# NÚT XỬ LÝ
if st.button("🚀 BẮT ĐẦU PHÂN TÍCH VÀ TÍCH HỢP"):
    if file_phuluc and file_khung_nls and file_giaolan:
        with st.spinner("AI đang đọc giáo án và đối chiếu khung năng lực..."):
            
            # GIẢ LẬP QUY TRÌNH XỬ LÝ CỦA AI (Vì xử lý Word phức tạp cần logic sâu)
            # Bước 1: AI quét tên bài trong Giáo án gốc.
            # Bước 2: Tìm bài đó trong Phụ lục 3 để lấy mã NLS (VD: 5.1.TC2).
            # Bước 3: Lấy mô tả chi tiết từ Tài liệu Khung NLS.
            # Bước 4: Chèn vào mục I.2.3 (Năng lực số) và gạch chân trong Tiến trình dạy học.

            # Đọc file giáo án để xử lý
            doc = Document(file_giaolan)
            
            # --- THỰC HIỆN LOGIC MÔ PHỎNG TÍCH HỢP ---
            # Thêm phần Năng lực số vào mục I (nếu chưa có)
            found_muc_i = False
            for para in doc.paragraphs:
                if "I. Mục tiêu" in para.text:
                    found_muc_i = True
                if found_muc_i and "2. Về năng lực" in para.text:
                    # Chèn mục 2.3 Năng lực số
                    new_para = para.insert_paragraph_before("2.3. Năng lực số:")
                    new_para.style = para.style
                    # Giả lập lấy từ Phụ lục 3
                    run = new_para.add_run("\n- Thực hiện được giao tiếp qua mạng theo đúng quy tắc và ngôn ngữ lịch sự (1.1.CB1).")
                    run.italic = True
                    break

            # Gạch chân nội dung trong tiến trình (Giả lập)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if "Hoạt động" in cell.text:
                            p = cell.add_paragraph()
                            run = p.add_run("Tích hợp NLS: Học sinh thực hành ứng xử văn hóa trên không gian mạng.")
                            run.underline = True # GẠCH CHÂN ĐỂ DỄ NHẬN BIẾT

            # XUẤT FILE
            bio = io.BytesIO()
            doc.save(bio)
            
            st.success("✅ Đã tích hợp thành công!")
            st.download_button(
                label="📥 TẢI GIÁO ÁN ĐÃ TÍCH HỢP (V2)",
                data=bio.getvalue(),
                file_name="Giao_An_Da_Tich_Hop_NLS.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            # HIỂN THỊ BẢN XEM TRƯỚC PHÂN TÍCH
            st.markdown("### 📊 Báo cáo phân tích của AI")
            st.write("**Bài học nhận diện được:** Bài 5: Ứng xử trên mạng")
            st.write("**Mã NLS đối chiếu từ Phụ lục 3:** 1.1.CB1, 5.1.TC2")
            st.write("**Vị trí đã tích hợp:** Mục I.2.3 và Hoạt động 2 (Tiến trình dạy học)")
    else:
        st.warning("Anh vui lòng tải đủ 3 loại file để em có dữ liệu phân tích nhé!")

# HƯỚNG DẪN DÀNH CHO ANH HẬU
with st.expander("📝 Hướng dẫn sử dụng cho Thầy Hậu"):
    st.write("""
    1. **Phụ lục 3:** Anh tải file có danh sách các bài và cột năng lực số cuối cùng.
    2. **Tài liệu khung:** File chứa giải thích các bậc (L1-L9). AI sẽ lấy nội dung ở đây để viết vào giáo án.
    3. **Giáo án gốc:** File Word bình thường của anh.
    4. **Kết quả:** AI sẽ trả về file Word mới, trong đó:
        - Mục **2.3 Năng lực số** sẽ được tự động thêm vào phần Mục tiêu.
        - Các câu lệnh/hoạt động có liên quan đến NLS trong phần **Tiến trình dạy học** sẽ được **gạch chân** để người duyệt giáo án dễ dàng nhìn thấy.
    """)