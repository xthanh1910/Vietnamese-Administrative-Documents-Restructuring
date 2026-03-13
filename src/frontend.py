"""
==========================================================================
GIAO DIỆN STREAMLIT: TÁI CẤU TRÚC VĂN BẢN PHÁP LUẬT VIỆT NAM
==========================================================================
- st.session_state lưu kết quả → không mất khi bấm download
- Thanh tiến trình (progress bar)
- Hiển thị Extended Title info
==========================================================================
"""

import streamlit as st
import cv2
import numpy as np
from io import BytesIO
from PIL import Image

from backend import process_multiple_images

# ==========================================================================
# CẤU HÌNH TRANG
# ==========================================================================
st.set_page_config(
    page_title="Tái cấu trúc Văn bản Pháp luật VN",
    page_icon="📜",
    layout="wide"
)

st.title("TÁI CẤU TRÚC VĂN BẢN PHÁP LUẬT VIỆT NAM")
st.markdown("""
**Pipeline:** Ảnh chụp → PaddleOCR → Model AI sửa lỗi tiếng Việt + Phân tích layout → Xuất Word giống ảnh input

**Hướng dẫn:**
1. Upload ảnh các trang văn bản pháp luật
2. **Trang đầu tiên** (có tiêu đề cơ quan, quốc hiệu): tick ✅ **"Trang đầu"**
3. Các trang còn lại: không cần tick gì
4. Bấm "Bắt đầu xử lý" và chờ đợi kết quả
""")

st.divider()

# ==========================================================================
# UPLOAD ẢNH
# ==========================================================================
uploaded_files = st.file_uploader(
    "Upload ảnh các trang văn bản pháp luật (JPG, PNG, JPEG)",
    type=["jpg", "jpeg", "png", "bmp", "tiff"],
    accept_multiple_files=True,
    help="Có thể upload nhiều trang cùng lúc"
)

if not uploaded_files:
    st.info("👆 Hãy upload ảnh để bắt đầu!")
    st.stop()

# ==========================================================================
# HIỂN THỊ ẢNH + CHECKBOX "TRANG ĐẦU"
# ==========================================================================
st.subheader(f"Đã upload {len(uploaded_files)} trang")

cols_per_row = 4
for row_start in range(0, len(uploaded_files), cols_per_row):
    cols = st.columns(cols_per_row)
    for col_idx, file_idx in enumerate(
        range(row_start, min(row_start + cols_per_row, len(uploaded_files)))
    ):
        with cols[col_idx]:
            file = uploaded_files[file_idx]
            pil_img = Image.open(file)
            st.image(pil_img, caption=file.name, use_container_width=True)
            st.checkbox(
                "Trang đầu tiên",
                key=f"first_page_{file_idx}",
                help="Tick nếu đây là trang đầu tiên"
            )
            file.seek(0)

st.divider()

# ==========================================================================
# NÚT XỬ LÝ
# ==========================================================================
if st.button("Bắt đầu xử lý", type="primary", use_container_width=True):

    first_page_flags = []
    for i in range(len(uploaded_files)):
        first_page_flags.append(st.session_state.get(f"first_page_{i}", False))

    image_list = []
    for file in uploaded_files:
        file_bytes = np.frombuffer(file.read(), dtype=np.uint8)
        img_bgr = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)
        image_list.append(img_bgr)
        file.seek(0)

    progress_bar = st.progress(0, text="Đang khởi tạo...")
    status_text = st.empty()

    def update_progress(current, total, message):
        if total > 0:
            pct = min(current / total, 1.0)
            progress_bar.progress(pct, text=f"{int(pct * 100)}% — {message}")
            status_text.text(message)

    doc, debug_images, all_info = process_multiple_images(
        image_list, first_page_flags, progress_cb=update_progress
    )

    progress_bar.progress(1.0, text="100% — Hoàn tất!")
    status_text.text("✅ Xử lý hoàn tất!")

    # Lưu vào session_state (tránh mất khi bấm download)
    word_buffer = BytesIO()
    doc.save(word_buffer)
    word_buffer.seek(0)
    st.session_state['word_bytes'] = word_buffer.getvalue()

    debug_png_list = []
    for dbg_img in debug_images:
        _, encoded = cv2.imencode('.png', dbg_img)
        debug_png_list.append(encoded.tobytes())
    st.session_state['debug_png_list'] = debug_png_list
    st.session_state['all_info'] = all_info
    st.session_state['processing_done'] = True

# ==========================================================================
# HIỂN THỊ KẾT QUẢ (ĐỌC TỪ SESSION_STATE → KHÔNG MẤT KHI BẤM DOWNLOAD)
# ==========================================================================
if st.session_state.get('processing_done', False):
    st.divider()
    st.success("Kết quả đã sẵn sàng!")

    word_bytes = st.session_state['word_bytes']
    debug_png_list = st.session_state['debug_png_list']
    all_info = st.session_state['all_info']

    # --- NÚT TẢI FILE WORD ---
    st.subheader("📥 Tải file Word")
    st.download_button(
        label="Tải file Word (.docx)",
        data=word_bytes,
        file_name="van_ban_phap_luat.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True
    )

    st.divider()

    # --- ẢNH DEBUG + THÔNG TIN + NÚT TẢI ---
    st.subheader("Ảnh Debug (Box OCR + Title)")

    for page_idx, (png_bytes, info) in enumerate(zip(debug_png_list, all_info)):
        # Tạo label cho expander
        if info['title_text']:
            label = (f"Trang {page_idx+1} có "
                     f"{info['total_boxes']} box | Title: {info['title_text']}")
        else:
            label = (f"Trang {page_idx+1} có "
                     f"{info['total_boxes']} box | Không có Title")

        with st.expander(label, expanded=True):
            # Hiển thị ảnh debug
            st.image(
                png_bytes,
                caption=f"Trang {page_idx+1} - Debug",
                use_container_width=True
            )

            # Metrics dạng 4 cột
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Tổng box", info['total_boxes'])
            col2.metric("Dòng Header", info['num_header_lines'])
            col3.metric("Dòng Body", info['num_body_lines'])

            # Extended title info
            ext_start, ext_end = info.get('ext_title_range', (0, 0))
            ext_count = ext_end - ext_start
            col4.metric("Dòng Title dài", ext_count)

            # Hiển thị title nếu có
            if info['title_text']:
                st.info(
                    f"**Title chính:** {info['title_text']} "
                    f"(dòng thứ {info['title_line_idx'] + 1})"
                )

            # Hiển thị extended title nếu có
            if ext_count > 0:
                st.warning(
                    f"**Title dài:** {ext_count} dòng "
                    f"(dòng {ext_start + 1} → {ext_end})"
                )

            # Nút tải ảnh debug
            st.download_button(
                label=f"📥 Tải ảnh debug trang {page_idx+1}",
                data=png_bytes,
                file_name=f"debug_trang_{page_idx+1}.png",
                mime="image/png",
                key=f"dl_debug_{page_idx}"
            )