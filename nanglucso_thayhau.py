import streamlit as st
import io
import os
import re
from docx import Document

# --- CẤU HÌNH GIAO DIỆN ---
st.set_page_config(page_title="AI Tích Hợp NLS V3 - Thầy Hậu", layout="wide")
st.markdown("""
    <style>
    .stApp { background-color: #FFFFFF !important; }
    .stButton>button { background-color: #1E40AF !important; color: white !important; font-weight: bold; }
    .status-box { padding: 15px; border-radius: 8px; border-left: 5px solid #1E40AF; background-color: #F8F9FA; margin-bottom: 20px;}
    </style>
    """, unsafe_allow_html=True)

st.title("🤖 AI TỔ CHỨC GIÁO ÁN NLS CHUẨN (V3)")
st.info("💡 Vị trí: 2.4 Năng lực số đặt trước mục 3. Về phẩm chất. AI tự động dàn trải nội dung và gợi ý phần mềm.")

# Thư viện công cụ số để giáo án sinh động
TOOLS_DICT = {
    "1.1": "Google Search, Wikipedia để tra cứu dữ liệu thực tế",
    "2.5": "Padlet, Zalo Group hoặc Google Classroom để thảo luận và nộp sản phẩm",
    "3.1": "Canva, MS Word hoặc PowerPoint để thiết kế nội dung số",
    "4.3": "Xem video trên Youtube về kỹ năng an toàn mạng, chơi Quizizz để củng cố",
    "5.1": "Sử dụng sơ đồ tư duy (MindMap) để hệ thống hóa kiến thức bài học",
    "5.2": "Phần mềm bảng tính (MS Excel/Google Sheets) để xử lý dữ liệu tự động",
}

col1, col2 = st.columns(2)
with col1:
    file_phuluc = st.file_uploader("1. Tải Phụ lục 3 (Mã NLS)", type=["docx"])
with col2:
    file_giaolan = st.file_uploader("2. Tải Giáo án gốc", type=["docx"])

def get_nls_from_pl3(pl_doc, lesson_name):
    """Hàm lấy dữ liệu cột 7 so với tên bài ở cột 3"""
    results = []
    for table in pl_doc.tables:
        for row in table.rows:
            if len(row.cells) >= 7:
                col3_text = row.cells[2].text.lower()
                if lesson_name in col3_text:
                    col7_text = row.cells[6].text
                    matches = re.findall(r'(\d+\.\d+\.[A-Za-z0-9]+)\s*[:\-]\s*(.*?)(?=\d+\.\d+\.[A-Za-z0-9]+\s*[:\-]|$)', col7_text, re.DOTALL)
                    for ma, nd in matches:
                        if ma.strip() not in [x['ma'] for x in results]:
                            results.append({"ma": ma.strip(), "nd": nd.strip()})
    return results

if st.button("🚀 CẬP NHẬT GIÁO ÁN NGAY"):
    if file_giaolan and file_phuluc:
        doc = Document(file_giaolan)
        pl_doc = Document(file_phuluc)
        
        # 1. Nhận diện bài học
        bai_hoc = ""
        for p in doc.paragraphs[:30]:
            match = re.search(r'(bài\s+\d+)', p.text, re.IGNORECASE)
            if match:
                bai_hoc = match.group(1).lower()
                break
        
        nls_list = get_nls_from_pl3(pl_doc, bai_hoc)
        
        if not nls_list:
            st.error(f"⚠️ Không tìm thấy bài '{bai_hoc}' trong cột 3 của Phụ lục 3.")
        else:
            # --- PHẦN 1: CHÈN 2.4 VÀO TRƯỚC MỤC 3 ---
            idx_3 = -1
            idx_22 = -1
            for i, p in enumerate(doc.paragraphs):
                txt = p.text.strip()
                if "3." in txt and "về phẩm chất" in txt.lower(): idx_3 = i
                if "2.2." in txt: idx_22 = i
            
            # Nếu không tìm thấy mục 3, chèn vào trước phần II. Thiết bị
            if idx_3 == -1:
                for i, p in enumerate(doc.paragraphs):
                    if "II." in p.text: idx_3 = i; break

            if idx_3 != -1:
                ref_p = doc.paragraphs[idx_3]
                
                # Kiểm tra/Chèn 2.3 nếu chưa có để đảm bảo thứ tự 2.1 -> 2.4
                has_23 = any("2.3." in p.text for p in doc.paragraphs)
                if not has_23 and idx_22 != -1:
                    p23 = ref_p.insert_paragraph_before("2.3. Năng lực khác:")
                    p23.runs[0].bold = True

                # Chèn 2.4. Năng lực số
                p24_head = ref_p.insert_paragraph_before("2.4. Năng lực số:")
                p24_head.runs[0].bold = True
                for item in nls_list:
                    p_item = ref_p.insert_paragraph_before(f"- {item['ma']}: {item['nd']}")
                    p_item.runs[0].italic = True
                # Chèn dòng trống để ngăn cách với mục 3
                ref_p.insert_paragraph_before("")

            # --- PHẦN 2: DÀN TRẢI SÁNG TẠO VÀO TIẾN TRÌNH ---
            all_candidate_paras = []
            def find_tasks(container):
                for p in container.paragraphs:
                    t = p.text.lower()
                    # Tìm các câu hỏi (?) hoặc lệnh chuyển giao nhiệm vụ
                    if "?" in t or "chuyển giao nhiệm vụ" in t or "câu hỏi:" in t:
                        all_candidate_paras.append(p)
                if hasattr(container, 'tables'):
                    for table in container.tables:
                        for row in table.rows:
                            for cell in row.cells: find_tasks(cell)

            find_tasks(doc)

            if all_candidate_paras:
                # Dàn trải nls_list vào các vị trí đã tìm được
                for i, item in enumerate(nls_list):
                    # Chia đều vị trí dựa trên số lượng mã NLS
                    target_idx = (i * len(all_candidate_paras)) // len(nls_list)
                    target_p = all_candidate_paras[target_idx]
                    
                    # Lấy gợi ý công cụ
                    prefix = item['ma'][:3]
                    tool = TOOLS_DICT.get(prefix, "phần mềm học tập chuyên dụng")
                    
                    # Biên soạn câu văn sáng tạo dựa trên câu hỏi của GV
                    creative_note = f"\n   => Tích hợp NLS ({item['ma']}): Thông qua hoạt động trả lời/thảo luận trên, HS rèn luyện: {item['nd']}. GV hướng dẫn sử dụng {tool} để tăng tính sinh động."
                    
                    run = target_p.add_run(creative_note)
                    run.bold = True; run.italic = True; run.underline = True

            # Xuất file
            bio = io.BytesIO()
            doc.save(bio)
            st.markdown(f"<div class='status-box'>✅ <b>HOÀN THÀNH:</b> Đã đặt 2.4 trước mục 3 và dàn trải {len(nls_list)} vị trí tích hợp sáng tạo vào giáo án.</div>", unsafe_allow_html=True)
            st.download_button(f"📥 Tải Giáo án {bai_hoc.upper()} chuẩn vị trí", bio.getvalue(), f"Giao_an_Chuan_V3_{bai_hoc}.docx")
    else:
        st.warning("Vui lòng tải đủ 2 file để em xử lý ạ!")
