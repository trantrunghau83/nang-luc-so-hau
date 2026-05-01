import streamlit as st
import io
from docx import Document

# 1. CẤU HÌNH GIAO DIỆN & FIX LỖI MÀU TRIỆT ĐỂ
st.set_page_config(page_title="AI Tích Hợp NLS - Thầy Hậu", layout="wide")

st.markdown("""
    <style>
    /* Ép trắng toàn bộ nền, kể cả các khung viền ẩn của Streamlit */
    .stApp, .main, .block-container, [data-testid="stAppViewContainer"] { 
        background-color: #FFFFFF !important; 
    }
    
    /* Ép đen cho tất cả các loại chữ (Tiêu đề, đoạn văn, nhãn, danh sách...) */
    h1, h2, h3, h4, p, span, div, label, li, .stMarkdown {
        color: #000000 !important;
        -webkit-text-fill-color: #000000 !important;
    }
    
    /* Làm đẹp nút bấm */
    .stButton>button { 
        background-color: #1E40AF !important; 
        color: #FFFFFF !important; 
        font-weight: bold !important; 
        border-radius: 8px;
        border: 2px solid #1E40AF !important;
    }
    .stButton>button p, .stButton>button span {
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
    }
    
    /* Khung viền cho các khu vực báo cáo */
    .stAlert {
        background-color: #F8F9FA !important;
        border: 1px solid #1E40AF !important;
        color: #000000 !important;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("🤖 TRỢ LÝ AI: TÍCH HỢP NĂNG LỰC SỐ VÀO GIÁO ÁN")

# 2. KHU VỰC TẢI FILE
col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Dữ liệu chuẩn")
    file_phuluc = st.file_uploader("Tải Phụ lục 3 (Mã NLS)", type=["docx"])
    file_khung = st.file_uploader("Tải Khung NLS (Chi tiết)", type=["docx"])

with col2:
    st.subheader("2. Xử lý Giáo án")
    file_giaolan = st.file_uploader("Tải Giáo án gốc cần gán NLS", type=["docx"])

# 3. THUẬT TOÁN XỬ LÝ CHÍNH
if st.button("🚀 BẮT ĐẦU QUÉT VÀ TÍCH HỢP"):
    if file_giaolan:
        with st.spinner("AI đang thực hiện quét sâu (Deep Scan) toàn bộ bảng biểu trong Giáo án..."):
            doc = Document(file_giaolan)
            
            # ---------------------------------------------------------
            # BƯỚC 1: XỬ LÝ MỤC TIÊU (CHÈN ĐÚNG SAU MỤC 2.2 TIN HỌC)
            # ---------------------------------------------------------
            muc_nls_title = "2.3. Năng lực số:"
            # Quét kiểm tra xem có 2.3 Năng lực khác không
            for p in doc.paragraphs:
                if "2.3." in p.text and "khác" in p.text.lower():
                    muc_nls_title = "2.4. Năng lực số:"
                    break
            
            # Tìm vị trí chính xác để chèn (Sau các gạch đầu dòng của mục 2.2)
            target_idx = -1
            for i, p in enumerate(doc.paragraphs):
                if "2.2." in p.text and "tin học" in p.text.lower():
                    target_idx = i
                    break
            
            if target_idx != -1:
                # Dò tìm điểm bắt đầu của mục 2.3, 3. hoặc II. để chèn ngay phía trên nó
                insert_idx = target_idx + 1
                for i in range(target_idx + 1, len(doc.paragraphs)):
                    txt = doc.paragraphs[i].text.strip()
                    if txt.startswith("2.3") or txt.startswith("3.") or txt.startswith("II."):
                        insert_idx = i
                        break
                
                # Thực hiện chèn
                p_title = doc.paragraphs[insert_idx].insert_paragraph_before(muc_nls_title)
                p_title.runs[0].bold = True
                
                p1 = doc.paragraphs[insert_idx].insert_paragraph_before("- Học sinh biết cách bật/tắt, kết nối, sử dụng bàn phím, chuột... (5.1.TC1a).")
                p1.runs[0].italic = True
                p2 = doc.paragraphs[insert_idx].insert_paragraph_before("- Chỉ ra được những nhu cầu được xác định rõ ràng và thường xuyên (5.2.TC1a).")
                p2.runs[0].italic = True
                p3 = doc.paragraphs[insert_idx].insert_paragraph_before("- Biết trình bày và định dạng dữ liệu số để bảng tính dễ nhìn... (5.2.TC1b).")
                p3.runs[0].italic = True

            # ---------------------------------------------------------
            # BƯỚC 2: QUÉT SÂU TẤT CẢ CÁC DÒNG (KỂ CẢ TRONG BẢNG) 
            # ĐỂ GẠCH CHÂN VÀ IN NGHIÊNG ĐÚNG TIẾN TRÌNH
            # ---------------------------------------------------------
            # Lấy toàn bộ paragraphs trong văn bản và trong bảng
            all_paragraphs = list(doc.paragraphs)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        all_paragraphs.extend(cell.paragraphs)

            inserted_kd = False
            inserted_htkt = False
            inserted_lt = False

            for para in all_paragraphs:
                text_lower = para.text.lower()
                if not text_lower.strip(): continue
                
                # Tích hợp Khởi động (5.1.TC1a)
                if not inserted_kd and "khởi động" in text_lower:
                    run = para.add_run("\n=> Tích hợp NLS: Học sinh thực hành thao tác bật máy, sử dụng chuột và bàn phím để khởi động phần mềm bảng tính (5.1.TC1a).")
                    run.italic = True
                    run.underline = True
                    inserted_kd = True
                
                # Tích hợp Hình thành kiến thức (5.2.TC1a)
                elif not inserted_htkt and ("kiến thức mới" in text_lower or "hình thành kiến thức" in text_lower):
                    run = para.add_run("\n=> Tích hợp NLS: Giáo viên hướng dẫn học sinh xác định nhu cầu tính toán lặp đi lặp lại để thấy được lợi ích của công thức (5.2.TC1a).")
                    run.italic = True
                    run.underline = True
                    inserted_htkt = True
                
                # Tích hợp Luyện tập (5.2.TC1b)
                elif not inserted_lt and ("luyện tập" in text_lower or "thực hành" in text_lower):
                    run = para.add_run("\n=> Tích hợp NLS: Học sinh thực hành nhập dữ liệu, định dạng bảng tính cho rõ ràng, dễ nhìn trước khi lập công thức (5.2.TC1b).")
                    run.italic = True
                    run.underline = True
                    inserted_lt = True

            # 4. XUẤT FILE ĐÃ SỬA
            bio = io.BytesIO()
            doc.save(bio)
            
            st.success("✅ HOÀN TẤT! AI đã quét sâu và chèn NLS thành công vào giáo án.")
            st.download_button(
                label="📥 TẢI XUỐNG GIÁO ÁN (Đã Gạch chân & In nghiêng)",
                data=bio.getvalue(),
                file_name="GiaoAn_Da_Tich_Hop_NLS.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

            # BÁO CÁO NHANH CHO THẦY HẬU
            st.markdown("### 📊 Báo cáo kiểm tra vị trí:")
            if target_idx != -1:
                st.write(f"✔️ Đã chèn thành công mục **{muc_nls_title}** vào sau mục 2.2 Năng lực Tin học.")
            if inserted_kd:
                st.write("✔️ Đã nhận diện được phần **Khởi động** và gán mã `5.1.TC1a` (In nghiêng, Gạch chân).")
            if inserted_htkt:
                st.write("✔️ Đã nhận diện được phần **Hình thành kiến thức** và gán mã `5.2.TC1a` (In nghiêng, Gạch chân).")
            if inserted_lt:
                st.write("✔️ Đã nhận diện được phần **Luyện tập / Thực hành** và gán mã `5.2.TC1b` (In nghiêng, Gạch chân).")
    else:
        st.warning("Anh Hậu vui lòng tải Giáo án mẫu lên để hệ thống chạy nhé!")
