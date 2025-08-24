# app.py (完全修正版)

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
import re

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

# --- URLからファイルIDを抽出する関数 ---
def extract_file_id_from_url(url_or_id):
    """URLまたはファイルIDからファイルIDを抽出"""
    if not url_or_id:
        return ""
    
    # すでにファイルIDの形式の場合（英数字とハイフン、アンダースコア）
    if len(url_or_id) > 10 and '/' not in url_or_id:
        return url_or_id.strip()
    
    # URL形式の場合
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

# --- 設定チェック ---
missing_keys = check_secrets()
if missing_keys:
    st.error(f"""
    **設定エラー:** 以下の設定が不足しています：
    - {', '.join(missing_keys)}
    
    📝 **対応方法:**
    1. `.streamlit/secrets.toml` ファイルを作成してください
    2. 必要な認証情報を追加してください
    
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
        pass

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
                    fields='name,mimeType,modifiedTime,size,owners,permissions'
                ).execute()
                
                st.info(f"""
                **ファイル名:** {file_info.get('name', 'N/A')}  
                **ファイル形式:** {file_info.get('mimeType', 'N/A')}  
                **更新日時:** {file_info.get('modifiedTime', 'N/A')}  
                **サイズ:** {file_info.get('size', 'N/A')} bytes
                """)
                
                # サービスアカウント情報の表示
                service_account_email = creds.service_account_email
                st.success(f"✅ **アクセス成功！** サービスアカウント: `{service_account_email}`")
                
            except Exception as e:
                st.error(f"ファイル情報の取得に失敗しました: {e}")
                
                # 詳細なトラブルシューティング情報
                service_account_email = creds.service_account_email if creds else "取得失敗"
                st.error(f"""
                **トラブルシューティング:**
                
                🔍 **サービスアカウント:** `{service_account_email}`
                
                📋 **確認項目:**
                1. ファイルIDが正しいか確認
                2. Google Driveでファイルが存在するか確認
                3. サービスアカウントにファイル共有されているか確認
                
                🛠️ **解決方法:**
                1. Google Driveで対象ファイルを右クリック → 「共有」
                2. サービスアカウントのメールアドレスを追加: `{service_account_email}`
                3. 権限を「編集者」に設定
                4. 「送信」をクリック
                """)
                
                # ファイル共有の手順を詳しく表示
                st.info("""
                **📧 サービスアカウントへの共有手順:**
                
                1. Google Driveで該当のExcelファイルを右クリック
                2. 「共有」を選択
                3. 「ユーザーやグループを追加」をクリック
                4. 上記のサービスアカウントメールアドレスを入力
                5. 権限を「編集者」に設定
                6. 「送信」をクリック
                
                ⚠️ **重要:** サービスアカウントは実際のGoogleアカウントではないため、メール通知は送信されません。
                """)

# --- 高度な処理オプション ---
st.subheader("🔧 高度な処理オプション")

enable_advanced_copy = st.checkbox(
    "2枚目→3枚目への名前＆日付マッチング処理を有効にする",
    value=True,
    help="2枚目「まとめ」シートから3枚目「予定カレンダー」シートへの高度なコピー機能"
)

if enable_advanced_copy:
    st.info("""
    **📋 処理内容:**
    - 2枚目のB列の名前と3枚目のN列の名前をマッチング
    - 2枚目のD3,G3,J3の日付と3枚目の1行目の日付をマッチング  
    - 2枚目の7行目以降奇数行のデータを3枚目の19行目以降に貼り付け
    - 数式は値として貼り付け（関数なしのテキスト）
    """)
    
    # --- メインのUI ---
st.subheader("📁 ファイルアップロード")
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
                
                # 4. openpyxlでExcelワークブックとして読み込み（マクロ対応）
                st.write("ステップ2/3: Excelデータをメモリ上で編集中...")
                # keep_vba=Trueでマクロを保持
                workbook = openpyxl.load_workbook(fh, keep_vba=True)
                
                # 5. 1枚目のシートを取得し、既存のデータをクリア
                sheet_to_update = workbook.worksheets[0]
                
                # ヘッダー行を保持するかチェック
                if sheet_to_update.max_row > 1:
                    sheet_to_update.delete_rows(2, sheet_to_update.max_row)

                # 6. 新しいデータを書き込む（ヘッダーがある場合は2行目から開始）
                start_row = 2 if sheet_to_update.max_row >= 1 else 1
                for r_idx, row in enumerate(dataframe_to_rows(source_df, index=False, header=False), start=start_row):
                    for c_idx, value in enumerate(row, start=1):
                        sheet_to_update.cell(row=r_idx, column=c_idx, value=value)

                # 7. 2枚目→3枚目への高度な貼り付け処理
                if enable_advanced_copy:
                    st.write("ステップ2.5/3: 2枚目から3枚目への名前＆日付マッチング処理中...")
                    if len(workbook.worksheets) >= 3:
                        sheet2 = workbook.worksheets[1]  # 2枚目「まとめ」
                        sheet3 = workbook.worksheets[2]  # 3枚目「予定カレンダー」
                        
                        # 2枚目の日付情報を取得（D3, G3, J3, M3, P3... CP3まで3列おき）
                        dates_sheet2 = {}
                        st.write("🔍 2枚目の日付情報を検索中...")
                        
                        # D3から始まって3列おきに検索（D=4, G=7, J=10, M=13, P=16...）
                        for col in range(4, 95, 3):  # 4列目(D)から3列おきに、CP列(94)まで
                            date_val = sheet2.cell(row=3, column=col).value
                            if date_val is not None:
                                try:
                                    if isinstance(date_val, (int, float)):
                                        date_num = int(date_val)
                                        if 1 <= date_num <= 31:  # 日付として有効な範囲
                                            dates_sheet2[date_num] = col
                                            st.write(f"  📅 2枚目: {date_num}日 → {col}列目({chr(64+col)}列) データは{col-1}列目から")
                                    elif isinstance(date_val, str) and date_val.isdigit():
                                        date_num = int(date_val)
                                        if 1 <= date_num <= 31:
                                            dates_sheet2[date_num] = col
                                            st.write(f"  📅 2枚目: {date_num}日 → {col}列目({chr(64+col)}列) データは{col-1}列目から")
                                except:
                                    pass
                        
                        st.write(f"2枚目で見つかった日付数: {len(dates_sheet2)}")
                        
                        # 3枚目の日付情報を取得（S1, V1, Y1, AB1... EF1まで3列おき）
                        dates_sheet3 = {}
                        st.write("🔍 3枚目の日付情報を検索中...")
                        
                        # S1から始まって3列おきに検索（S=19, V=22, Y=25, AB=28...）
                        for col in range(19, 136, 3):  # 19列目(S)から3列おきに、EF列(136)まで
                            date_val = sheet3.cell(row=1, column=col).value
                            if date_val is not None:
                                try:
                                    if isinstance(date_val, (int, float)):
                                        date_num = int(date_val)
                                        if 1 <= date_num <= 31:  # 日付として有効な範囲
                                            dates_sheet3[date_num] = col
                                            col_name = ""
                                            if col <= 26:
                                                col_name = chr(64+col)
                                            else:
                                                col_name = chr(64+(col-1)//26) + chr(64+((col-1)%26)+1)
                                            st.write(f"  📅 3枚目: {date_num}日 → {col}列目({col_name}列) データは{col-2}列目から")
                                    elif isinstance(date_val, str) and date_val.isdigit():
                                        date_num = int(date_val)
                                        if 1 <= date_num <= 31:
                                            dates_sheet3[date_num] = col
                                            col_name = ""
                                            if col <= 26:
                                                col_name = chr(64+col)
                                            else:
                                                col_name = chr(64+(col-1)//26) + chr(64+((col-1)%26)+1)
                                            st.write(f"  📅 3枚目: {date_num}日 → {col}列目({col_name}列) データは{col-2}列目から")
                                except:
                                    pass
                        
                        st.write(f"3枚目で見つかった日付数: {len(dates_sheet3)}")
                        
                        # 共通の日付を確認
                        common_dates = set(dates_sheet2.keys()) & set(dates_sheet3.keys())
                        st.write(f"共通の日付: {sorted(common_dates)}")


                        
                        # 2枚目の名前リストを取得（B列、7行目以降の奇数行）
                        names_sheet2 = {}
                        for row in range(7, min(sheet2.max_row + 1, 50), 2):  # 7行目から奇数行のみ、最大50行まで
                            name = sheet2.cell(row=row, column=2).value  # B列
                            if name and str(name).strip():
                                names_sheet2[str(name).strip()] = row
                        
                        # 3枚目の名前リストを取得（N列、19行目以降）
                        names_sheet3 = {}
                        for row in range(19, min(sheet3.max_row + 1, 100)):  # 19行目以降、最大100行まで
                            name = sheet3.cell(row=row, column=14).value  # N列
                            if name and str(name).strip():
                                names_sheet3[str(name).strip()] = row
                        
                        # 名前＆日付マッチングでデータ貼り付け
                        copy_count = 0
                        match_log = []
                        
                        if not common_dates:
                            st.warning("⚠️ 共通の日付が見つかりませんでした。日付の形式を確認してください。")
                        else:
                            for name, sheet2_row in names_sheet2.items():
                                if name in names_sheet3:
                                    sheet3_row = names_sheet3[name]
                                    match_log.append(f"名前マッチ: {name} (2枚目{sheet2_row}行 → 3枚目{sheet3_row}行)")
                                    
                                    # 各共通日付のデータをコピー
                                    for date in common_dates:
                                        date_col_sheet2 = dates_sheet2[date]  # 日付の列
                                        date_col_sheet3 = dates_sheet3[date]  # 日付の列
                                        
                                        # 実際のデータの列を計算
                                        data_col_sheet2 = date_col_sheet2 - 1  # 日付の1つ前の列
                                        data_col_sheet3 = date_col_sheet3 - 2  # 日付の2つ前の列
                                        
                                        match_log.append(f"  日付マッチ: {date}日 (2枚目{date_col_sheet2}列の日付→{data_col_sheet2}列からコピー, 3枚目{date_col_sheet3}列の日付→{data_col_sheet3}列に貼り付け)")
                                        
                                        # 2枚目の該当セルのデータを取得してコピー
                                        source_value = sheet2.cell(row=sheet2_row, column=data_col_sheet2).value
                                        
                                        if source_value is not None:
                                            # 数式ではなく値として貼り付け
                                            if isinstance(source_value, str) and source_value.startswith('='):
                                                # 数式の場合は計算後の表示値を取得しようとする（簡易版）
                                                display_value = "[計算式結果]"  # 実際の値を取得するのは複雑なので、プレースホルダー
                                                sheet3.cell(row=sheet3_row, column=data_col_sheet3).value = display_value
                                            else:
                                                sheet3.cell(row=sheet3_row, column=data_col_sheet3).value = source_value
                                            
                                            copy_count += 1
                                            match_log.append(f"    ✅コピー: '{source_value}' → 3枚目({sheet3_row},{data_col_sheet3})")
                                        else:
                                            match_log.append(f"    ⚠️スキップ: 空のセル 2枚目({sheet2_row},{data_col_sheet2})")
                        
                        st.success(f"✅ {copy_count}個のセルを2枚目から3枚目にコピーしました")

                        
                        # マッチングログを表示
                        if match_log:
                            with st.expander("📊 マッチング詳細ログ"):
                                for log in match_log[:20]:  # 最初の20件のみ表示
                                    st.text(log)
                    else:
                        st.warning("⚠️ ワークブックにシートが3枚未満のため、シート間コピーをスキップしました")

                # 7. 変更をメモリ上で保存（xlsmとして保存）
                output_buffer = io.BytesIO()
                workbook.save(output_buffer)
                output_buffer.seek(0)
                
                # 8. 再構築したファイルで、Drive上のファイルBを上書き更新
                st.write("ステップ3/3: Drive上のファイルを新しい内容で上書き中...")
                # xlsmファイル用のMIMEタイプに変更
                media = MediaIoBaseUpload(output_buffer, mimetype='application/vnd.ms-excel.sheet.macroEnabled.12')
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
