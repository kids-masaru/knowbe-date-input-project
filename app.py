# app.py (コンパクト版)

import streamlit as st
import pandas as pd
from gspread_dataframe import set_with_dataframe
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import time

# --- ページ設定 ---
st.set_page_config(
    page_title="出勤管理効率化システム",
    page_icon="📄",
    layout="centered"
)

# --- CSSの読み込み ---
def load_css(file_name):
    """指定されたCSSファイルを読み込み、Streamlitアプリに適用する"""
    with open(file_name, encoding="utf-8") as f:
        st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

load_css('style.css')

# --- Google Sheets API 認証 ---
def get_gspread_client():
    """StreamlitのSecretsから認証情報を読み込み、Google Sheets APIのクライアントを返す"""
    try:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
        )
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Google Sheetsへの認証に失敗しました: {e}")
        st.info("管理者の方へ: .streamlit/secrets.toml の設定が正しいか確認してください。")
        return None

# --- SVGによるカスタムタイトル ---
# ご提示の画像を参考に、SVGコードでタイトルを直接描画します。
svg_title = """
<div class="svg-title-container">
    <svg width="100%" height="60" xmlns="http://www.w3.org/2000/svg">
        <defs>
            <linearGradient id="titleGradient" x1="0%" y1="0%" x2="100%" y2="0%">
                <stop offset="0%" style="stop-color:#3b82f6;stop-opacity:1" />
                <stop offset="100%" style="stop-color:#1e40af;stop-opacity:1" />
            </linearGradient>
        </defs>
        <rect x="0" y="52" width="200" height="4" fill="url(#titleGradient)" rx="2"></rect>
        <circle cx="10" cy="15" r="8" fill="#dbeafe"></circle>
        <rect x="25" y="8" width="16" height="16" fill="#93c5fd" rx="4"></rect>
        <text x="55" y="35" font-family="Noto Sans JP, sans-serif" font-size="28" font-weight="700" fill="url(#titleGradient)">
            出勤管理効率化システム
        </text>
    </svg>
</div>
"""
st.markdown(svg_title, unsafe_allow_html=True)


# --- メインのUI ---
st.markdown("業務データ（ExcelまたはCSV）をアップロードしてください。")

uploaded_files = st.file_uploader(
    "Excelファイルを1つ、またはCSVファイルを2つ選択してください",
    type=['csv', 'xlsx', 'xls'],
    label_visibility="collapsed",
    accept_multiple_files=True
)

if uploaded_files:
    file_names = " | ".join([f.name for f in uploaded_files])
    st.info(f"**選択中のファイル:** {file_names}")

is_pressed = st.button("アップロード開始", use_container_width=True, disabled=(not uploaded_files))

st.markdown("---") # 細い区切り線

result_placeholder = st.empty()
result_placeholder.info("アップロードを開始すると、ここに結果が表示されます。")


# --- ボタンが押された後の処理ロジック ---
if is_pressed:
    start_time = time.time()
    client = get_gspread_client()

    if client:
        try:
            with st.spinner('データを処理し、スプレッドシートに書き込んでいます...'):
                spreadsheet_url = st.secrets["g_spreadsheet_url"]
                spreadsheet = client.open_by_url(spreadsheet_url)

                # ----- ファイル数に応じた処理分岐 -----
                if len(uploaded_files) == 1: # Excelの場合
                    uploaded_file = uploaded_files[0]
                    if not uploaded_file.name.endswith(('.xlsx', '.xls')):
                        result_placeholder.error("エラー: ファイルが1つの場合は、Excel (.xlsx, .xls) ファイルである必要があります。")
                        st.stop()
                    excel_data = pd.read_excel(uploaded_file, engine='openpyxl', sheet_name=None)
                    sheet_names = list(excel_data.keys())
                    if len(sheet_names) < 2:
                        result_placeholder.error("エラー: Excelファイルにシートが2枚以上存在しません。")
                        st.stop()
                    df1, df2 = excel_data[sheet_names[0]], excel_data[sheet_names[1]]

                elif len(uploaded_files) == 2: # CSVの場合
                    file1, file2 = uploaded_files[0], uploaded_files[1]
                    if not (file1.name.endswith('.csv') and file2.name.endswith('.csv')):
                        result_placeholder.error("エラー: ファイルが2つの場合は、両方ともCSVファイルである必要があります。")
                        st.stop()
                    df1, df2 = pd.read_csv(file1), pd.read_csv(file2)
                else:
                    result_placeholder.error("エラー: ファイルは「Excel1つ」または「CSV2つ」のどちらかでアップロードしてください。")
                    st.stop()

                # --- スプレッドシートへの書き込み処理 ---
                worksheet1 = spreadsheet.worksheet("貼り付け用①")
                worksheet1.clear()
                set_with_dataframe(worksheet1, df1)
                
                worksheet2 = spreadsheet.worksheet("貼り付け用②")
                worksheet2.clear()
                set_with_dataframe(worksheet2, df2)

            # --- 正常終了時のメッセージ ---
            end_time = time.time()
            processing_time = end_time - start_time
            now_str = datetime.now().strftime("%Y/%m/%d %H:%M:%S")

            result_placeholder.success(f"**更新完了！** (更新日時: {now_str}, 処理時間: {processing_time:.2f}秒)")
            st.balloons()

        # --- エラー処理 ---
        except gspread.exceptions.WorksheetNotFound:
            result_placeholder.error("**シート名エラー:** Googleスプレッドシートに「貼り付け用①」または「貼り付け用②」という名前のシートが見つかりません。")
        except Exception as e:
            result_placeholder.error(f"**予期せぬエラー:** {e}")