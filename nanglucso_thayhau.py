import streamlit as st
import io
import os
from docx import Document

# 1. GIAO DIỆN CHUẨN TRẮNG - ĐEN
st.set_page_config(page_title="AI Tích Hợp NLS - Thầy Hậu", layout="wide")
st.markdown("""
    <style>
    .stApp { background-color: #FFFFFF !important; }
    h1, h2, h3, h4, p, span, div, label, li, .stMarkdown { color: #000000 !important; }
    .stButton>button { background-color: #1E40AF !important; color: white !important; font-weight: bold !important; border-radius: 8px; }
    </style>
    """, unsafe_allow_html=True)

st.title("🤖 TRỢ LÝ AI: TÍCH HỢP NĂNG LỰC SỐ (BẢN FIX VỊ TRÍ & ĐỊNH DẠNG)")
st.info("💡 Đã sửa lỗi đảo ngược dòng. Định dạng chuẩn xác: [Mã NLS]: [Nội dung] theo Phụ lục 3.")

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Dữ liệu nguồn")
    file_phuluc = st.file_uploader("Phụ lục 3 (Mã NLS)", type=["docx"])
    file_khung = st.file_uploader("Khung NLS (Chi tiết)", type=["docx"])
with col2:
    st.subheader("2. Giáo án cần xử lý")
    file_giaolan = st.file_uploader("Tải Giáo án gốc (File Word)", type=["docx"])

if st.button("🚀 BẮT ĐẦU TÍCH HỢP"):
    if file_giaolan:
        with st.spinner("AI đang quét và sắp xếp đúng vị trí..."):
            doc = Document(file_giaolan)
            
            # Đặt tên file: Tên cũ + NLS
            original_name = os.path.splitext(file_giaolan.name)[0]
            new_filename = f"{original_name} NLS.docx"

            # Dữ liệu NLS (Đã chuẩn hóa để ráp đúng định dạng của anh)
            nls_list = [
                {"ma": "5.1.TC1a", "nd": "Hướng dẫn HS thực hành thao tác thiết bị, bật/tắt đúng quy trình."},
                {"ma": "5.2.TC1a", "nd": "Yêu cầu HS xác định nhu cầu thực tế cần giải quyết."},
                {"ma": "5.2.TC1b", "nd": "Hướng dẫn HS cách trình bày và định dạng dữ liệu số."}
            ]

            # --- BƯỚC 1: XỬ LÝ MỤC TIÊU (2.3 hoặc 2.4) ---
            has_khac = any("2.3." in p.text and "khác" in p.text.lower() for p in doc.paragraphs)
            num_nls = "2.4." if has_khac else "2.3."

            target_idx = -1
            for i, p in enumerate(doc.paragraphs):
                if "2.2." in p.text and "tin học" in p.text.lower():
                    target_idx = i
                    break
            
            if target_idx != -1:
                insert_pos = target_idx + 1
                for j in range(target_idx + 1, len(doc.paragraphs)):
                    t = doc.paragraphs[j].text.strip()
                    if t.startswith("3.") or t.startswith("II.") or t.startswith("III."):
                        insert_pos = j
                        break
                
                # SỬA LỖI ĐẢO NGƯỢC: Dùng ref_p để đóng đinh vị trí chèn
                ref_p = doc.paragraphs[insert_pos]
                
                # 1. Chèn Tiêu đề TRƯỚC
                p_head = ref_p.insert_paragraph_before(f"{num_nls} Năng lực số:")
                p_head.runs[0].bold = True
                
                # 2. Chèn các mục NLS theo đúng thứ tự (Định dạng: Mã: Nội dung)
                for item in nls_list:
                    # In nghiêng phần nội dung để đồng bộ giáo án
                    p_item = ref_p.insert_paragraph_before(f"- {item['ma']}: {item['nd']}")
                    p_item.runs[0].italic = True

            # --- BƯỚC 2: RÀ SOÁT TÌM "CHUYỂN GIAO NHIỆM VỤ" ---
            tracker = {"idx": 0, "count": 0}

            def inject_nls(para):
                txt = para.text.lower()
                
                if "chuyển giao nhiệm vụ" in txt or "giao nhiệm vụ:" in txt:
                    if len(nls_list) > 0:
                        nls = nls_list[tracker["idx"] % len(nls_list)]
                        
                        # Chèn thêm 1 dòng mới (Định dạng: Mã: Nội dung)
                        run = para.add_run(f"\n   => Tích hợp NLS: {nls['ma']}: {nls['nd']}")
                        
                        # Ép In Đậm, In Nghiêng, Gạch Chân
                        run.italic = True
                        run.underline = True
                        run.bold = True 
                        
                        tracker["idx"] += 1
                        tracker["count"] += 1

            # 1. Quét mọi dòng bên ngoài bảng
            for p in doc.paragraphs:
                inject_nls(p)
                
            # 2. Quét mọi dòng bên trong TẤT CẢ các bảng
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            inject_nls(p)

            # XUẤT FILE
            bio = io.BytesIO()
            doc.save(bio)
            
            # THÔNG BÁO KẾT QUẢ RÕ RÀNG
            if tracker["count"] > 0:
                st.success(f"✅ TUYỆT VỜI! Đã chèn {tracker['count']} Năng lực số đúng thứ tự và định dạng.")
            else:
                st.warning("⚠️ Quét xong nhưng không thấy chữ 'Chuyển giao nhiệm vụ'.")

            st.download_button(
                label=f"📥 TẢI FILE: {new_filename}",
                data=bio.getvalue(),
                file_name=new_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning("Anh Hậu ơi, tải file giáo án lên giúp em nhé!")
