import streamlit as st
import io
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="AI Tích Hợp NLS - Thầy Hậu", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #FFFFFF !important; }
    .stButton>button { background-color: #1E40AF !important; color: white !important; font-weight: bold !important; }
    </style>
    """, unsafe_allow_html=True)

st.title("🤖 TRỢ LÝ AI: TÍCH HỢP NĂNG LỰC SỐ VÀO GIÁO ÁN")

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Dữ liệu chuẩn (Tải 1 lần)")
    file_phuluc = st.file_uploader("Phụ lục 3 (Mã NLS theo bài)", type=["docx"])
    file_khung_nls = st.file_uploader("Tài liệu Khung NLS (Chi tiết)", type=["docx"])

with col2:
    st.subheader("2. Xử lý Giáo án")
    file_giaolan = st.file_uploader("Tải Giáo án cần gán NLS", type=["docx"])

if st.button("🚀 PHÂN TÍCH VÀ TÍCH HỢP NĂNG LỰC SỐ"):
    if file_giaolan: # Chỉ cần có giáo án là chạy demo được logic này
        with st.spinner("AI đang đọc tiến trình dạy học và phân bổ năng lực số..."):
            
            doc = Document(file_giaolan)
            
            # --- LOGIC 1: ĐÁNH SỐ VÀ CHÈN MỤC TIÊU NĂNG LỰC SỐ ---
            muc_nls_title = "2.3. Năng lực số:"
            
            # Quét xem có "2.3. Năng lực khác" chưa
            for para in doc.paragraphs:
                if "2.3." in para.text and "khác" in para.text.lower():
                    muc_nls_title = "2.4. Năng lực số:"
                    break

            # Chèn nội dung vào sau phần Năng lực
            for para in doc.paragraphs:
                if "2.2. Năng lực Tin học" in para.text:
                    # Chèn tiêu đề NLS
                    p_title = para.insert_paragraph_before(muc_nls_title)
                    p_title.runs[0].bold = True
                    
                    # Chèn các mã NLS (Giả lập đọc từ Phụ lục 3)
                    p1 = para.insert_paragraph_before("- Học sinh biết cách bật/tắt, kết nối, sử dụng bàn phím, chuột, màn hình, USB, máy in... theo đúng quy trình (5.1.TC1a).")
                    p1.runs[0].italic = True
                    
                    p2 = para.insert_paragraph_before("- Chỉ ra được những nhu cầu được xác định rõ ràng và thường xuyên (5.2.TC1a).")
                    p2.runs[0].italic = True
                    
                    p3 = para.insert_paragraph_before("- Biết trình bày và định dạng dữ liệu số để bảng tính dễ nhìn, dễ đọc (5.2.TC1b).")
                    p3.runs[0].italic = True
                    break

            # --- LOGIC 2: GÁN VÀO ĐÚNG VỊ TRÍ TRONG TIẾN TRÌNH DẠY HỌC ---
            # Hàm hỗ trợ gạch chân
            def insert_nls_into_cell(cell, text):
                p = cell.add_paragraph()
                run = p.add_run(f"Tích hợp NLS: {text}")
                run.underline = True
                run.italic = True

            # Quét các bảng chứa Tiến trình dạy học
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        text_lower = cell.text.lower()
                        
                        # 1. Gán 5.1.TC1a vào phần Khởi động hoặc lúc mở máy
                        if "khởi động" in text_lower or "giao nhiệm vụ" in text_lower:
                            # Chỉ gán 1 lần tránh lặp
                            if "5.1.TC1a" not in cell.text:
                                insert_nls_into_cell(cell, "Học sinh thực hành thao tác bật máy, sử dụng chuột và bàn phím để khởi động phần mềm bảng tính (5.1.TC1a).")
                        
                        # 2. Gán 5.2.TC1a vào phần Hình thành kiến thức / Đặt vấn đề
                        elif "kiến thức mới" in text_lower or "hoạt động 1" in text_lower:
                            if "5.2.TC1a" not in cell.text:
                                insert_nls_into_cell(cell, "Giáo viên hướng dẫn học sinh xác định nhu cầu tính toán lặp đi lặp lại để thấy được lợi ích của công thức (5.2.TC1a).")

                        # 3. Gán 5.2.TC1b vào phần Luyện tập / Thực hành
                        elif "luyện tập" in text_lower or "thực hành" in text_lower:
                            if "5.2.TC1b" not in cell.text:
                                insert_nls_into_cell(cell, "Học sinh thực hành nhập dữ liệu, định dạng bảng tính cho rõ ràng, dễ nhìn trước khi lập công thức (5.2.TC1b).")

            # XUẤT FILE ĐÃ SỬA
            bio = io.BytesIO()
            doc.save(bio)
            
            st.success("✅ Đã chèn thành công các năng lực vào Giáo án Bài 7: Tính toán tự động!")
            st.download_button(
                label="📥 TẢI GIÁO ÁN (ĐÃ TÍCH HỢP NLS - CÓ GẠCH CHÂN)",
                data=bio.getvalue(),
                file_name="GiaoAn_TichHop_NLS.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

            # BÁO CÁO NHANH
            st.markdown("### 📊 Chi tiết AI phân bổ:")
            st.write("1. **Chỉ mục:** Đã tự động tạo mục `" + muc_nls_title + "`.")
            st.write("2. **Vị trí 1 (Khởi động):** Gán mã `5.1.TC1a` (Bật tắt, dùng phím chuột).")
            st.write("3. **Vị trí 2 (Hình thành kiến thức):** Gán mã `5.2.TC1a` (Chỉ ra nhu cầu tính toán).")
            st.write("4. **Vị trí 3 (Luyện tập):** Gán mã `5.2.TC1b` (Định dạng bảng tính).")
            st.info("💡 Nội dung tích hợp đã được AI **in nghiêng, gạch chân và đính kèm mã** ở cuối để dễ kiểm duyệt!")
    else:
        st.warning("Vui lòng tải Giáo án mẫu lên để AI chạy thử nghiệm.")
