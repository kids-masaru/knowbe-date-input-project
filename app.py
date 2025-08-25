# app.py (修正版 - 固定範囲コピー対応)

import streamlit as st

# --- デバッグ: secrets.toml の中身を確認 ---
st.write("DEBUG: st.secrets の中身", dict(st.secrets))
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

# --- 列番号から文字に変換する関数 ---
def col_num_to_letter(col_num):
    """列番号を文字に変換 (1=A, 26=Z, 27=AA)"""
    result = ""
    while col_num > 0:
        col_num -= 1
        result = chr(65 + col_num % 26) + result
        col_num //= 26
    return result

# --- 列文字から番号に変換する関数 ---
def col_letter_to_num(col_letter):
    """列文字を番号に変換 (A=1, Z=26, AA=27)"""
    num = 0
    for char in col_letter.upper():
        num = num * 26 + ord(char) - ord('A') + 1
    return num

# --- セル参照をパースする関数 ---
def parse_cell_reference(cell_ref):
    """セル参照（例：Y1）を列番号と行番号に分解"""
    match = re.match(r'^([A-Z]+)(\d+)$', cell_ref.upper())
    if not match:
        return None, None
    col_letter = match.group(1)
    row_num = int(match.group(2))
    col_num = col_letter_to_num(col_letter)
    return col_num, row_num

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
    "2枚目→3枚目への名前マッチング処理を有効にする",
    value=True,
    help="2枚目「まとめ」シートから3枚目「予定カレンダー」シートへの固定範囲コピー機能"
)

if enable_advanced_copy:
    # 貼り付け基準セルの設定
    target_cell = st.text_input(
        "3枚目の基準セル（例：Y1）",
        value="Y1",
        help="指定したセルの2つ前の列から貼り付けます。Y1→W列, Z1→X列"
    )
    
    # セル参照のバリデーション
    target_col, target_row = parse_cell_reference(target_cell)
    if target_col and target_row:
        paste_col = target_col - 2
        paste_col_letter = col_num_to_letter(paste_col)
        st.info(f"""
        **📋 処理設定:**
        - コピー元: 2枚目「まとめ」シートのC7〜CQ7（名前マッチング行）
        - 基準セル: 3枚目「予定カレンダー」の{target_cell}
        - 貼り付け開始列: {paste_col_letter}列（{target_cell}の2つ前）
        - 貼り付け開始行: 19行目（名前マッチング行）
        - 名前マッチング: 2枚目B列 ←→ 3枚目N列
        """)
    else:
        st.error(f"無効なセル参照です: {target_cell}")

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

                # 7. 2枚目→3枚目への名前マッチング＆固定範囲コピー処理
                if enable_advanced_copy and target_col and target_row:
                    st.write("ステップ2.5/3: 2枚目から3枚目への名前マッチング＆固定範囲コピー処理中...")
                    if len(workbook.worksheets) >= 3:
                        sheet2 = workbook.worksheets[1]  # 2枚目「まとめ」
                        sheet3 = workbook.worksheets[2]  # 3枚目「予定カレンダー」
                        
                        # 基本情報表示
                        st.write(f"📊 シート情報:")
                        st.write(f"  2枚目シート名: '{sheet2.title}'")
                        st.write(f"  3枚目シート名: '{sheet3.title}'")
                        
                        # 2枚目の名前リストを取得（B列、全ての行を確認）
                        names_sheet2 = {}
                        st.write("🔍 2枚目の名前を収集中（B列、全行）...")
                        for row in range(1, min(sheet2.max_row + 1, 200)):
                            name = sheet2.cell(row=row, column=2).value  # B列
                            if name and str(name).strip():
                                clean_name = str(name).strip()
                                names_sheet2[clean_name] = row
                                st.write(f"  ✅ B{row}: '{clean_name}'")
                        
                        # 3枚目の名前リストを取得（全列を確認して名前らしきものを探す）
                        names_sheet3 = {}
                        st.write("🔍 3枚目の名前を収集中（全列、全行）...")
                        
                        # 複数の列で名前を探す（N列を重点的に、その他の列も確認）
                        search_columns = [14]  # N列(14)をメインに検索
                        
                        for col in search_columns:
                            col_letter = col_num_to_letter(col)
                            st.write(f"  {col_letter}列を検索中...")
                            
                            for row in range(1, min(sheet3.max_row + 1, 200)):
                                name = sheet3.cell(row=row, column=col).value
                                if name and str(name).strip():
                                    clean_name = str(name).strip()
                                    if clean_name not in names_sheet3:  # 重複を避ける
                                        names_sheet3[clean_name] = row
                                        st.write(f"    ✅ {col_letter}{row}: '{clean_name}'")
                        
                        st.write(f"📊 収集結果:")
                        st.write(f"  2枚目の名前数: {len(names_sheet2)} 個")
                        st.write(f"  3枚目の名前数: {len(names_sheet3)} 個")
                        st.write(f"  2枚目の名前: {list(names_sheet2.keys())}")
                        st.write(f"  3枚目の名前: {list(names_sheet3.keys())}")
                        
                        # 名前のマッチング確認
                        matched_names = set(names_sheet2.keys()) & set(names_sheet3.keys())
                        st.write(f"  名前マッチ数: {len(matched_names)} 個")
                        st.write(f"  マッチした名前: {list(matched_names)}")
                        
                        # 固定範囲コピー処理
                        copy_count = 0
                        copy_log = []
                        
                        # コピー範囲の定義
                        copy_start_col = col_letter_to_num('C')  # C列 = 3
                        copy_end_col = col_letter_to_num('CQ')   # CQ列 = 95
                        
                        paste_start_col = target_col - 2  # 基準セルの2つ前の列
                        
                        if not matched_names:
                            st.warning("⚠️ マッチする名前が見つかりませんでした。")
                            st.info("💡 名前の大文字小文字、スペース、全角半角などを確認してください")
                        else:
                            st.write(f"🚀 コピー開始:")
                            st.write(f"  📄 コピー元: 2枚目 C〜CQ列（{copy_end_col - copy_start_col + 1}列）")
                            st.write(f"  📍 貼り付け先: 3枚目 {col_num_to_letter(paste_start_col)}列から開始")
                            
                            for name in matched_names:
                                sheet2_row = names_sheet2[name]
                                sheet3_row = names_sheet3[name]
                                copy_log.append(f"名前マッチ: '{name}' → 2枚目B{sheet2_row} から 3枚目{col_num_to_letter(paste_start_col)}{sheet3_row}へ")
                                
                                # C列からCQ列まで（固定範囲）をコピー
                                for col_offset in range(copy_end_col - copy_start_col + 1):
                                    source_col = copy_start_col + col_offset  # C列から始まる
                                    target_col_for_paste = paste_start_col + col_offset  # 基準セルの2つ前から始まる
                                    
                                    # 2枚目の該当セルのデータを取得
                                    source_value = sheet2.cell(row=sheet2_row, column=source_col).value
                                    
                                    # 3枚目の該当セルに貼り付け（値として）
                                    if source_value is not None:
                                        # 数式の場合は値として貼り付け（数式は除去）
                                        if isinstance(source_value, str) and source_value.startswith('='):
                                            sheet3.cell(row=sheet3_row, column=target_col_for_paste).value = "[数式]"
                                        else:
                                            sheet3.cell(row=sheet3_row, column=target_col_for_paste).value = source_value
                                        
                                        copy_count += 1
                                    else:
                                        # Noneの場合は空文字を設定
                                        sheet3.cell(row=sheet3_row, column=target_col_for_paste).value = ""
                                    
                                    # 最初の5列のみ詳細ログ
                                    if col_offset < 5:
                                        source_col_letter = col_num_to_letter(source_col)
                                        target_col_letter = col_num_to_letter(target_col_for_paste)
                                        copy_log.append(f"    {source_col_letter}{sheet2_row}('{source_value}') → {target_col_letter}{sheet3_row}")
                            
                            st.success(f"✅ {copy_count}個のセルを2枚目から3枚目にコピーしました")
                            
                            # コピーログを表示
                            if copy_log:
                                with st.expander("📊 コピー詳細ログ"):
                                    for log in copy_log:
                                        st.text(log)
                    else:
                        st.warning("⚠️ ワークブックにシートが3枚未満のため、シート間コピーをスキップしました")

                # 8. 変更をメモリ上で保存（xlsmとして保存）
                output_buffer = io.BytesIO()
                workbook.save(output_buffer)
                output_buffer.seek(0)
                
                # 9. 再構築したファイルで、Drive上のファイルBを上書き更新
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
            import traceback
            st.text(traceback.format_exc())

