# app.py (シート保持・最終版)

import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2.service_account import Credentials
from datetime import datetime
import time

# --- ページ設定 ---
st.set_page_config(
    page_title="Excel直接更新システム",
    page_icon="📎",
    layout="centered"
)

# --- CSS ---
st.markdown("""
<style>
body { font-family: 'Noto Sans JP', sans-serif; }
.main .block-container { padding-top: 2rem; }
h1 { border-bottom: 2px solid #2563eb; padding-bottom: 0.5rem; }
</style>
""", unsafe_allow_html=True)

st.title("📎 Excel直接更新システム")
st.markdown("更新元ファイル（CSV/Excel）をアップロードすると、Google Drive上の指定のExcelファイルの1枚目のシートを上書きします。")
st.warning("**注意:** この操作はDrive上のファイルを直接変更します。2枚目以降のシートは保持されますが、念のためバックアップを取ることを推奨します。")

# --- Google API 認証 ---
def get_google_creds():
    try:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/drive"]
        )
        return creds
    except Exception as e:
        st.error(f"Googleへの認証に失敗しました: {e}")
        return None

# --- メインのUI ---
uploaded_file = st.file_uploader(
    "更新元となるファイル（CSVまたはExcel）を選択してください",
    type=['csv', 'xlsx', 'xls'],
    label_visibility="collapsed"
)

if uploaded_file:
    st.info(f"**選択中のファイル:** {uploaded_file.name}")

is_pressed = st.button("Drive上のExcelを更新実行", type="primary", use_container_width=True, disabled=(not uploaded_file))

st.markdown("---")
result_placeholder = st.empty()

# --- ボタンが押された後の処理ロジック ---
if is_pressed:
    start_time = time.time()
    creds = get_google_creds()

    if creds:
        try:
            with st.spinner('処理を実行中... しばらくお待ちください。'):
                # 1. Drive APIサービスを構築
                drive_service = build('drive', 'v3', credentials=creds)
                file_id = st.secrets["target_excel_file_id"]

                # 2. アップロードされたファイル（A）からデータをDataFrameとして読み込む
                if uploaded_file.name.endswith('.csv'):
                    source_df = pd.read_csv(uploaded_file)
                else:
                    source_df = pd.read_excel(uploaded_file, sheet_name=0)

                # 3. Drive上のExcelファイル（B）をダウンロード
                st.write("ステップ1/3: Drive上の既存ファイルをダウンロード中...")
                request = drive_service.files().get_media(fileId=file_id)
                file_content_bytes = request.execute()
                fh = io.BytesIO(file_content_bytes)
                
                # 4. openpyxlでExcelワークブックとして読み込む
                st.write("ステップ2/3: Excelデータをメモリ上で編集中...")
                workbook = openpyxl.load_workbook(fh)
                
                # 5. 1枚目のシートを取得し、既存のデータをクリア
                sheet_to_update = workbook.worksheets[0]
                sheet_to_update.delete_rows(2, sheet_to_update.max_row + 1) # ヘッダーを残し、2行目以降を全削除

                # 6. 新しいデータを書き込む
                for row in dataframe_to_rows(source_df, index=False, header=False):
                    sheet_to_update.append(row)

                # 7. 変更をメモリ上で保存
                output_buffer = io.BytesIO()
                workbook.save(output_buffer)
                output_buffer.seek(0)
                
                # 8. 再構築したファイルで、Drive上のファイルBを上書き更新
                st.write("ステップ3/3: Drive上のファイルを新しい内容で上書き中...")
                media = MediaIoBaseUpload(output_buffer, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                drive_service.files().update(
                    fileId=file_id,
                    media_body=media
                ).execute()

            # --- 正常終了時のメッセージ ---
            end_time = time.time()
            processing_time = end_time - start_time
            now_str = datetime.now().strftime("%Y年%m月%d日 %H:%M:%S")
            result_placeholder.success(f"**更新完了！** Drive上のExcelファイルが更新されました。(日時: {now_str}, 処理時間: {processing_time:.2f}秒)")

        except Exception as e:
            result_placeholder.error(f"**エラーが発生しました:** {e}")
