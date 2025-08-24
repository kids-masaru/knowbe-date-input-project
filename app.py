# app.py (エラーハンドリング改善版)

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

# --- 設定の確認 ---
def check_secrets():
    """必要なsecrets設定が存在するかチェック"""
    missing_keys = []
    
    if "gcp_service_account" not in st.secrets:
        missing_keys.append("gcp_service_account")
    
    if "target_excel_file_id" not in st.secrets:
        missing_keys.append("target_excel_file_id")
    
    return missing_keys

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

# --- 設定チェック ---
missing_keys = check_secrets()
if missing_keys:
    st.error(f"""
    **設定エラー:** 以下の設定が不足しています：
    - {', '.join(missing_keys)}
    
    📝 **対応方法:**
    1. `.streamlit/secrets.toml` ファイルを作成してください
    2. 必要な認証情報とファイルIDを追加してください
    
    詳細については、[Streamlit Secrets管理](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/secrets-management)を参照してください。
    """)
    st.stop()

# --- Google Drive ファイルID の設定 ---
st.subheader("📁 更新対象のGoogle DriveファイルID")

# デフォルトファイルIDの取得
default_file_id = ""
try:
    default_file_id = st.secrets.get("target_excel_file_id", "")
except:
    pass

# ファイルIDの入力UI
col1, col2 = st.columns([3, 1])
with col1:
    file_id = st.text_input(
        "Google DriveファイルのIDまたはURLを入力してください",
        value=default_file_id,
        placeholder="例: 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms",
        help="DriveのURL: https://drive.google.com/file/d/【この部分がID】/view"
    )

with col2:
    if st.button("🔗 URLから抽出", help="Drive URLからファイルIDを自動抽出"):
        # URLからファイルIDを抽出する処理は後で実装
        pass

# URLからファイルIDを抽出する関数
def extract_file_id_from_url(url_or_id):
    """URLまたはファイルIDからファイルIDを抽出"""
    if not url_or_id:
        return ""
    
    # すでにファイルIDの形式の場合（英数字とハイフン、アンダースコア）
    if len(url_or_id) > 10 and '/' not in url_or_id:
        return url_or_id.strip()
    
    # URL形式の場合
    import re
    patterns = [
        r'/file/d/([a-zA-Z0-9-_]+)',
        r'id=([a-zA-Z0-9-_]+)',
        r'/folders/([a-zA-Z0-9-_]+)'
    ]
    
    for pattern in patterns:
        match = re.search(pattern, url_or_id)
        if match:
            return match.group(1)
    
    return url_or_id.strip()

# ファイルIDの処理
file_id = extract_file_id_from_url(file_id)

if not file_id:
    st.warning("**ファイルIDが入力されていません。** Google DriveのファイルIDまたはURLを入力してください。")
    st.info("""
    📝 **ファイルIDの取得方法:**
    1. Google Driveで対象のExcelファイルを開く
    2. URLをコピー: `https://drive.google.com/file/d/【この部分】/view`
    3. 上記の【この部分】がファイルIDです
    """)
else:
    st.success(f"**対象ファイルID:** `{file_id}`")
    
    # ファイル情報を表示する機能
    if st.checkbox("📋 ファイル情報を確認"):
        creds = get_google_creds()
        if creds:
            try:
                drive_service = build('drive', 'v3', credentials=creds)
                file_info = drive_service.files().get(
                    fileId=file_id, 
                    fields='name,mimeType,modifiedTime,size'
                ).execute()
                
                st.info(f"""
                **ファイル名:** {file_info.get('name', 'N/A')}  
                **ファイル形式:** {file_info.get('mimeType', 'N/A')}  
                **更新日時:** {file_info.get('modifiedTime', 'N/A')}  
                **サイズ:** {file_info.get('size', 'N/A')} bytes
                """)
            except Exception as e:
                st.error(f"ファイル情報の取得に失敗しました: {e}")
                st.error("ファイルIDが正しいか、またはファイルへのアクセス権限があるか確認してください。")

# --- メインのUI ---
uploaded_file = st.file_uploader(
    "更新元となるファイル（CSVまたはExcel）を選択してください",
    type=['csv', 'xlsx', 'xls'],
    label_visibility="collapsed"
)

if uploaded_file:
    st.info(f"**選択中のファイル:** {uploaded_file.name}")

is_pressed = st.button(
    "Drive上のExcelを更新実行", 
    type="primary", 
    use_container_width=True, 
    disabled=(uploaded_file is None or not file_id)
)

st.markdown("---")
result_placeholder = st.empty()

# --- ボタンが押された後の処理ロジック ---
if is_pressed:
    # 処理開始前の最終チェック
    if uploaded_file is None:
        st.error("エラー: ファイルがアップロードされていません。")
        st.stop()
    
    if not file_id:
        st.error("エラー: Google DriveのファイルIDが入力されていません。")
        st.stop()
    
    start_time = time.time()
    creds = get_google_creds()

    if creds:
        try:
            with st.spinner('処理を実行中... しばらくお待ちください。'):
                # 1. Drive APIサービスを構築
                drive_service = build('drive', 'v3', credentials=creds)

                # 2. アップロードされたファイル（A）からデータをDataFrameとして読み込む
                if uploaded_file is None:
                    st.error("ファイルがアップロードされていません。")
                    st.stop()
                
                # ファイル形式をチェックして適切に読み込み
                file_extension = uploaded_file.name.lower()
                if file_extension.endswith('.csv'):
                    source_df = pd.read_csv(uploaded_file)
                elif file_extension.endswith(('.xlsx', '.xls')):
                    source_df = pd.read_excel(uploaded_file, sheet_name=0)
                else:
                    st.error(f"サポートされていないファイル形式です: {uploaded_file.name}")
                    st.stop()

                # 3. Drive上のExcelファイル（B）をダウンロード
                st.write("ステップ1/3: Drive上の既存ファイルをダウンロード中...")
                try:
                    request = drive_service.files().get_media(fileId=file_id)
                    file_content_bytes = request.execute()
                    fh = io.BytesIO(file_content_bytes)
                except Exception as e:
                    st.error(f"Driveからのファイルダウンロードに失敗しました。ファイルIDが正しいか確認してください: {e}")
                    st.stop()
                
output_buffer)
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
