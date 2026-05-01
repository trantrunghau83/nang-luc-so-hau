import streamlit as st
import io
import os
import re
from docx import Document

# --- CẤU HÌNH GIAO DIỆN ---
st.set_page_config(page_title="AI Tích Hợp NLS Toàn Diện - Thầy Hậu", layout="wide")
st.markdown("""
    <style>
    .stApp { background-color: #FFFFFF !important; }
    h1, h2, h3 { color: #1E40AF !important; }
    .stButton>button { background-color: #1E40AF !important; color: white !important; width: 100%; border-radius: 10px; height: 50px; font-size: 18px;}
    .status-card { padding: 20px; border-radius: 10px; background-color: #F1F5F9; border-left: 6px solid #1E40AF; margin-bottom: 20px; color: #1E293B;}
    </style>
    """, unsafe_allow_html=True)

st.title("🤖 AI TÍCH HỢP NLS: PHÂN BỔ DÀN TRẢI & TỐI ƯU CÔNG CỤ")
st.info("💡 Hướng dẫn: AI sẽ tự động dàn trải NLS vào các nhiệm vụ và gợi ý công cụ số (Canva, Padlet, Excel...) dựa trên nội dung bài học.")

# Thư viện công cụ số thông minh theo từng nhóm năng lực
DIGITAL_TOOLS = {
    "1.1": "Google Search, Wikipedia (để tìm kiếm dữ liệu)",
    "2.5": "Padlet, Zalo Group, Google Classroom (để thảo luận nhóm)",
    "3.1": "MS Excel, Google Sheets, Canva (để tạo nội dung số)",
    "4.3": "Youtube, Quizizz (để học tập và giải trí an toàn)",
    "5.1": "Sơ đồ tư duy MindMap, MS Word (để hệ thống hóa kiến thức)",
    "5.2": "Phần mềm bảng tính, các công cụ đo lường trực tuyến",
    "5.3": "Các phần mềm mô phỏng, công cụ bảo mật máy tính"
}

col1, col2 = st.columns(2)
with col1:
    file_phuluc = st.file_uploader("1. Tải Phụ lục 3 (Mã NLS)", type=["docx"])
with col2:
    file_giaolan = st.file_uploader("2. Tải Giáo án gốc", type=["docx"])

def get_nls_from_pl3(pl_doc, lesson_name):
    """Lấy dữ liệu NLS từ cột 3 và cột 7"""
    nls_list = []
    for table in pl_doc.tables:
        for row in table.rows:
            if len(row.cells) >= 7:
                c3 = row.cells[2].text.lower()
                if lesson_name in c3:
                    text = row.cells[6].text
                    matches = re.findall(r'(\d+\.\d+\.[A-Za-z0-9]+)\s*[:\-]\s*(.*?)(?=\d+\.\d+\.[A-Za-z0-9]+\s*[:\-]|$)', text, re.DOTALL)
                    for ma, nd in matches:
                        if ma.strip() not in [x['ma'] for x in nls_list]:
                            nls_list.append({"ma": ma.strip(), "nd": nd.strip()})
    return nls_list

if st.button("🚀 THỰC HIỆN TÍCH HỢP DÀN TRẢI"):
    if file_giaolan and file_phuluc:
        doc = Document(file_giaolan)
        pl_doc = Document(file_phuluc)
        
        # 1. Nhận diện bài học
        bai_hoc = ""
        for p in doc.paragraphs[:30]:
            m = re.search(r'(bài\s+\d+)', p.text, re.IGNORECASE)
            if m: bai_hoc = m.group(1).lower(); break
        
        nls_data = get_nls_from_pl3(pl_doc, bai_hoc)
        
        if not nls_data:
            st.error(f"Không tìm thấy bài '{bai_hoc}' trong Phụ lục 3. Vui lòng kiểm tra lại cột 3.")
        else:
            # --- PHẦN 1: CHỈNH SỬA MỤC TIÊU 2.1 -> 2.4 ---
            idx_22 = -1
            idx_23 = -1
            for i, p in enumerate(doc.paragraphs):
                if "2.2." in p.text: idx_22 = i
                if "2.3." in p.text: idx_23 = i
            
            # Nếu không có 2.3, tạo 2.3 sau 2.2
            if idx_23 == -1 and idx_22 != -1:
                ref_p = doc.paragraphs[idx_22 + 1]
                p23 = ref_p.insert_paragraph_before("2.3. Năng lực khác:")
                p23.runs[0].bold = True
                idx_23 = idx_22 + 1
            
            # Chèn 2.4 sau 2.3
            if idx_23 != -1:
                # Tìm vị trí bắt đầu mục 3 hoặc phần II lớn để chèn vào cuối mục 2
                insert_pos = idx_23 + 1
                for k in range(idx_23 + 1, len(doc.paragraphs)):
                    if doc.paragraphs[k].text.strip().startswith(("3.", "II.", "III.")):
                        insert_pos = k; break
                
                ref_p = doc.paragraphs[insert_pos]
                p24_head = ref_p.insert_paragraph_before("2.4. Năng lực số:")
                p24_head.runs[0].bold = True
                for item in nls_data:
                    p_item = ref_p.insert_paragraph_before(f"- {item['ma']}: {item['nd']}")
                    p_item.runs[0].italic = True

            # --- PHẦN 2: DÀN TRẢI VÀO TIẾN TRÌNH ---
            all_task_paras = []
            # Gom tất cả "Chuyển giao nhiệm vụ" hoặc "Câu hỏi"
            for p in doc.paragraphs:
                if any(x in p.text.lower() for x in ["chuyển giao nhiệm vụ", "câu hỏi", "nhiệm vụ 1", "nhiệm vụ 2"]):
                    all_task_paras.append(p)
            for t in doc.tables:
                for r in t.rows:
                    for c in r.cells:
                        for p in c.paragraphs:
                            if any(x in p.text.lower() for x in ["chuyển giao nhiệm vụ", "câu hỏi"]):
                                all_task_paras.append(p)

            if all_task_paras:
                # Thuật toán dàn trải: Chia đều nls_data vào all_task_paras
                num_tasks = len(all_task_paras)
                num_nls = len(nls_data)
                step = max(1, num_tasks // num_nls)
                
                for i in range(num_nls):
                    target_idx = min(i * step, num_tasks - 1)
                    para = all_task_paras[target_idx]
                    nls = nls_data[i]
                    
                    # Lấy gợi ý công cụ
                    prefix = nls['ma'][:3]
                    tool = DIGITAL_TOOLS.get(prefix, "các ứng dụng học tập số phù hợp")
                    
                    # Biên soạn lại câu văn sinh động
                    creative_msg = (f"\n   => Tích hợp NLS ({nls['ma']}): Thông qua nhiệm vụ/câu hỏi trên, "
                                   f"HS rèn luyện kỹ năng {nls['nd']}. GV khuyến khích HS sử dụng {tool} "
                                   f"để hoàn thành nội dung sinh động hơn.")
                    
                    run = para.add_run(creative_msg)
                    run.bold = True; run.italic = True; run.underline = True

            # Xuất file
            bio = io.BytesIO()
            doc.save(bio)
            st.markdown(f"<div class='status-card'>✅ <b>THÀNH CÔNG:</b> Đã chuẩn hóa mục tiêu 2.1-2.4 và dàn trải {len(nls_data)} năng lực số kèm gợi ý công cụ vào tiến trình dạy học.</div>", unsafe_allow_html=True)
            st.download_button(f"📥 TẢI GIÁO ÁN {bai_hoc.upper()}", bio.getvalue(), f"Giao_an_NLS_Chuan_{bai_hoc}.docx")
    else:
        st.warning("Vui lòng tải lên cả Phụ lục 3 và Giáo án gốc.")
