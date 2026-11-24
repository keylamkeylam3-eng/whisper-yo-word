# ==========================================================
# MP3 語音轉文字 Word 產生器
# Streamlit Web App
# ==========================================================

import os
import io
import tempfile
from datetime import datetime

import streamlit as st
import whisper
from docx import Document
from docx.shared import Pt

# ==========================================================
# 頁面設定
# ==========================================================
st.set_page_config(
    page_title="MP3 語音轉文字 Word 產生器",
    page_icon="🎧",
    layout="centered"
)

st.title("🎧 MP3 語音轉文字 Word 產生器")
st.write("上傳 MP3 檔案，自動轉錄成文字，並下載 Word（.docx）檔。")

# ==========================================================
# 快取 Whisper 模型，避免每次上傳都重新載入
# ==========================================================
@st.cache_resource(show_spinner="正在載入 Whisper 模型，請稍候……")
def load_whisper_model(model_name):
    return whisper.load_model(model_name)

# ==========================================================
# 側邊欄設定
# ==========================================================
with st.sidebar:
    st.header("轉錄設定")

    model_name = st.selectbox(
        "選擇 Whisper 模型",
        options=[
            "tiny",
            "base",
            "small",
            "medium",
            "large-v3",
            "turbo"
        ],
        index=5
    )

    language = st.selectbox(
        "選擇語音語言",
        options=[
            ("廣東話／粵語", "yue"),
            ("普通話／國語", "zh"),
            ("英文", "en"),
            ("自動偵測語言", None)
        ],
        format_func=lambda x: x[0],
        index=0
    )[1]

    st.info(
        "建議：\n\n"
        "- 廣東話／粵語：選擇「yue」\n"
        "- 普通話／國語：選擇「zh」\n"
        "- 想要更高準確度：使用「large-v3」\n"
        "- 想要較快速度：使用「tiny」或「base」"
    )

# ==========================================================
# 上傳 MP3
# ==========================================================
uploaded_file = st.file_uploader(
    "請上傳 MP3 檔案",
    type=["mp3", "wav", "m4a", "flac"]
)

if uploaded_file is None:
    st.info("請先上傳音訊檔案。")
    st.stop()

st.success(f"已上傳檔案：{uploaded_file.name}")

# ==========================================================
# 開始轉錄
# ==========================================================
if st.button("開始轉錄並產生 Word 檔", type="primary"):
    model = load_whisper_model(model_name)

    # 建立暫存檔案，供 Whisper 讀取
    file_extension = os.path.splitext(uploaded_file.name)[1].lower()
    
    with tempfile.NamedTemporaryFile(
        delete=False,
        suffix=file_extension
    ) as temp_audio:
        temp_audio.write(uploaded_file.getvalue())
        temp_audio_path = temp_audio.name

    try:
        with st.spinner("正在轉錄語音，請稍候……"):
            if language:
                result = model.transcribe(
                    temp_audio_path,
                    language=language,
                    verbose=False
                )
            else:
                result = model.transcribe(
                    temp_audio_path,
                    verbose=False
                )

        transcript = result["text"].strip()

        if not transcript:
            st.error("沒有辨識到文字，請確認音訊內容、音量或語言設定。")
            st.stop()

        st.subheader("轉錄結果")
        st.text_area(
            "文字內容",
            value=transcript,
            height=350
        )

        # ======================================================
        # 建立 Word 文件
        # ======================================================
        document = Document()
        document.add_heading("語音轉文字結果", level=0)

        document.add_paragraph(
            f"來源音訊檔：{uploaded_file.name}"
        )
        document.add_paragraph(
            f"轉錄時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        )
        document.add_paragraph(
            f"使用模型：Whisper {model_name}"
        )
        document.add_paragraph("")

        document.add_heading("轉錄內容", level=1)

        paragraph = document.add_paragraph(transcript)
        paragraph.runs[0].font.size = Pt(12)

        # 將 Word 檔寫入記憶體
        docx_buffer = io.BytesIO()
        document.save(docx_buffer)
        docx_buffer.seek(0)

        # 產生下載檔名
        base_name = os.path.splitext(uploaded_file.name)[0]
        output_filename = f"{base_name}_語音轉文字結果.docx"

        st.download_button(
            label="⬇️ 下載 Word 檔",
            data=docx_buffer,
            file_name=output_filename,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            )
        )

    except Exception as e:
        st.error(f"轉錄或建立 Word 檔時發生錯誤：{e}")

    finally:
        # 刪除暫存音訊檔
        if os.path.exists(temp_audio_path):
            os.remove(temp_audio_path)
