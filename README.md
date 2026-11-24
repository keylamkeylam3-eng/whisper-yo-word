# MP3 語音轉文字 Word 產生器

這是一個使用 Streamlit 和 OpenAI Whisper 的語音轉文字網頁應用程式。

## 功能

- 上傳 MP3、WAV、M4A 或 FLAC 音訊檔
- 使用 Whisper 轉錄語音為文字
- 選擇廣東話、普通話、英文或自動偵測語言
- 產生 Word（.docx）檔
- 直接下載 Word 檔到本機電腦

## 使用方法

1. 開啟 Streamlit 網頁
2. 上傳音訊檔案
3. 選擇語言與 Whisper 模型
4. 點選「開始轉錄並產生 Word 檔」
5. 轉錄完成後，點選「下載 Word 檔」

## 技術

- Python
- Streamlit
- OpenAI Whisper
- python-docx
