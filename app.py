"""
Streamlit Webアプリ - メンバー管理表変換ツール
認証情報を毎回入力する仕様
プログレス表示と完了通知を追加
"""

import streamlit as st
import pandas as pd
import io
import logging
from datetime import datetime, timedelta
from member_processor import MemberListProcessor
from sharepoint_handler import SharePointHandler
import config
import time

# ============ ページ設定 ============
st.set_page_config(
    page_title=config.APP_TITLE,
    page_icon=config.APP_ICON,
    layout=config.PAGE_LAYOUT,
    initial_sidebar_state="expanded"
)

# ============ ロギング設定 ============
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# ============ CSS カスタマイズ ============
st.markdown("""
<style>
    .main {
        padding: 2rem;
    }
    .stTabs [data-baseweb="tab-list"] button {
        font-size: 1.1em;
    }
    .auth-box {
        padding: 1.5rem;
        border-radius: 0.5rem;
        background-color: #f0f2f6;
        border: 2px solid #1f77b4;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
    }
</style>
""", unsafe_allow_html=True)

# ============ セッション状態の初期化 ============
def initialize_session_state():
    """セッション状態を初期化"""
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if "sharepoint_handler" not in st.session_state:
        st.session_state.sharepoint_handler = None
    if "processor" not in st.session_state:
        st.session_state.processor = None
    if "processed" not in st.session_state:
        st.session_state.processed = False
    if "auth_time" not in st.session_state:
        st.session_state.auth_time = None
    if "df_sharepoint" not in st.session_state:
        st.session_state.df_sharepoint = None

initialize_session_state()

# ============ 認証状態チェック ============
def check_session_timeout():
    """セッションタイムアウトをチェック"""
    if st.session_state.auth_time is None:
        return
    
    elapsed = datetime.now() - st.session_state.auth_time
    if elapsed > timedelta(seconds=config.SESSION_TIMEOUT):
        st.session_state.authenticated = False
        st.session_state.sharepoint_handler = None
        st.warning("⏱️ セッションがタイムアウトしました。再度認証してください。")

# ============ 認証画面 ============
def show_authentication_form():
    """認証フォームを表示"""
    st.title(f"{config.APP_ICON} {config.APP_TITLE}")
    
    st.markdown("""
    ### 🔐 SharePoint 認証
    
    職場のメールアドレスとパスワードで認証してください。
    """)
    
    with st.form("auth_form"):
        st.markdown("#### 認証情報を入力")
        
        email = st.text_input(
            "📧 メールアドレス",
            placeholder="example@company.com",
            help="企業のメールアドレスを入力してください"
        )
        
        password = st.text_input(
            "🔑 パスワード",
            type="password",
            placeholder="PCログインパスワード",
            help="PCログイン時と同じパスワードを入力してください"
        )
        
        col1, col2 = st.columns(2)
        
        with col1:
            submit = st.form_submit_button(
                "✅ 認証",
                use_container_width=True,
                type="primary"
            )
        
        with col2:
            st.form_submit_button(
                "🔄 クリア",
                use_container_width=True,
                on_click=lambda: None
            )
    
    if submit:
        if not email or not password:
            st.error("❌ メールアドレスとパスワードを入力してください")
            return
        
        with st.spinner("🔐 認証中..."):
            handler = SharePointHandler(email, password)
            
            if handler.authenticate():
                st.session_state.authenticated = True
                st.session_state.sharepoint_handler = handler
                st.session_state.auth_time = datetime.now()
                st.success("✅ 認証に成功しました！")
                time.sleep(1)
                st.rerun()
            else:
                st.error(f"❌ 認証に失敗しました\n\n{handler.error_message}")
                st.info("""
                #### 🆘 トラブルシューティング
                - メールアドレスが正しいか確認してください
                - パスワードが正しいか確認してください
                - 職場PCネットワークに接続しているか確認してください
                - 勤務時間外の場合はVPN接続を確認してください
                """)

# ============ メイン画面 ============
def show_main_screen():
    """メイン処理画面を表示"""
    st.title(f"{config.APP_ICON} {config.APP_TITLE}")
    
    # ヘッダー
    col1, col2, col3 = st.columns([2, 2, 1])
    
    with col1:
        st.markdown(f"**✅ 認証状態**: ログイン中")
    
    with col2:
        if st.session_state.auth_time:
            elapsed = (datetime.now() - st.session_state.auth_time).total_seconds()
            st.caption(f"認証時刻: {st.session_state.auth_time.strftime('%H:%M:%S')}")
    
    with col3:
        if st.button("🚪 ログアウト"):
            st.session_state.authenticated = False
            st.session_state.sharepoint_handler = None
            st.session_state.processor = None
            st.session_state.processed = False
            st.rerun()
    
    st.divider()
    
    tab1, tab2, tab3 = st.tabs(["📤 処理実行", "📊 結果確認", "ℹ️ 詳細情報"])
    
    with tab1:
        show_process_tab()
    
    with tab2:
        show_result_tab()
    
    with tab3:
        show_info_tab()

def show_process_tab():
    """処理実行タブ"""
    st.subheader("Step 1️⃣ : ファイルを取得")
    
    file_source = st.radio(
        "取得方法",
        ("📁 ファイルアップロード", "🔗 SharePoint から自動取得"),
        horizontal=True,
        key="file_source"
    )
    
    uploaded_file = None
    df_loaded = None
    
    if file_source == "📁 ファイルアップロード":
        uploaded_file = st.file_uploader(
            "Excelファイルを選択",
            type=["xlsx"],
            help="メンバー管理表のExcelファイルをアップロードしてください",
            key="file_upload"
        )
        
        if uploaded_file:
            st.success(f"📄 ファイル: {uploaded_file.name} ({uploaded_file.size / 1024 / 1024:.2f} MB)")
    
    else:  # SharePoint
        st.markdown("#### 🔗 SharePoint から取得")
        
        handler = st.session_state.sharepoint_handler
        
        if handler:
            with st.spinner("📁 ファイル一覧を取得中..."):
                files = handler.list_files()
                
                if files:
                    excel_files = [f for f in files if f.endswith('.xlsx')]
                    
                    if excel_files:
                        selected_file = st.selectbox(
                            "Excelファイルを選択",
                            excel_files,
                            index=0 if config.EXCEL_FILENAME in excel_files else None
                        )
                        
                        if st.button("📥 ダウンロード", use_container_width=True, type="primary"):
                            with st.spinner(f"⏳ ダウンロード中: {selected_file}"):
                                df_loaded, error_msg = handler.read_excel(selected_file)
                                
                                if df_loaded is not None:
                                    st.success(f"✅ ダウンロード完了: {len(df_loaded)}行")
                                    st.session_state.df_sharepoint = df_loaded
                                else:
                                    st.error(f"❌ エラー: {error_msg}")
                    else:
                        st.warning("⚠️ Excelファイルが見つかりません")
                else:
                    st.error(f"❌ ファイル一覧取得エラー: {handler.error_message}")
    
    st.divider()
    
    st.subheader("Step 2️⃣ : 処理を実行")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("▶️ 処理を実行", type="primary", use_container_width=True):
            # プログレス表示エリア
            progress_area = st.container()
            
            try:
                with progress_area:
                    st.info("⏳ 処理を開始しています...")
                    progress_bar = st.progress(0)
                
                processor = MemberListProcessor()
                
                # ステップ1: ファイル読込
                with progress_area:
                    st.info("⏳ **ステップ 1/4: ファイルを読み込み中...**")
                    progress_bar.progress(20)
                
                success = False
                
                if file_source == "📁 ファイルアップロード":
                    if uploaded_file is None:
                        st.error("❌ ファイルをアップロードしてください")
                        progress_area.empty()
                    else:
                        success = processor.load_from_file(uploaded_file)
                else:
                    if "df_sharepoint" not in st.session_state or st.session_state.df_sharepoint is None:
                        st.error("❌ SharePoint からファイルをダウンロードしてください")
                        progress_area.empty()
                        success = False
                    else:
                        success = processor.load_from_dataframe(st.session_state.df_sharepoint)
                
                if success:
                    # ステップ2: バリデーション
                    with progress_area:
                        st.info("⏳ **ステップ 2/4: 列の検証中...**")
                        progress_bar.progress(40)
                    
                    if not processor.validate_columns():
                        st.error("❌ 必要な列が不足しています")
                        with st.expander("📋 エラーログ"):
                            for log in processor.get_logs():
                                st.code(log)
                        progress_area.empty()
                    else:
                        # ステップ3: フィルタリング・変換
                        with progress_area:
                            st.info("⏳ **ステップ 3/4: フィルタリングと変換処理中...**")
                            progress_bar.progress(60)
                        
                        if processor.process():
                            # ステップ4: 完了処理
                            with progress_area:
                                st.info("⏳ **ステップ 4/4: 完了処理中...**")
                                progress_bar.progress(90)
                            
                            st.session_state.processor = processor
                            st.session_state.processed = True
                            
                            progress_bar.progress(100)
                            time.sleep(0.5)
                            progress_area.empty()
                            
                            # 成功通知
                            st.balloons()
                            st.success("""
                            ✅ **処理が完了しました！**
                            
                            📊 **次のステップ**:
                            - 「結果確認」タブで統計情報とプレビューを確認
                            - 「CSVダウンロード」ボタンでファイルをダウンロード
                            """)
                            
                            time.sleep(1)
                            st.rerun()
                        else:
                            progress_area.empty()
                            st.error("❌ 処理中にエラーが発生しました")
                            with st.expander("📋 詳細ログを表示"):
                                for log in processor.get_logs():
                                    if "ERROR" in log or "WARNING" in log:
                                        st.code(log, language="text")
                else:
                    progress_area.empty()
                    st.error("❌ ファイルの読み込みに失敗しました")
                    with st.expander("📋 詳細ログを表示"):
                        for log in processor.get_logs():
                            st.code(log, language="text")
                        
            except Exception as e:
                progress_area.empty()
                st.error(f"❌ 予期しないエラーが発生しました: {str(e)}")
    
    with col2:
        if st.button("🔄 リセット", use_container_width=True):
            st.session_state.processor = None
            st.session_state.processed = False
            st.session_state.df_sharepoint = None
            st.rerun()

def show_result_tab():
    """結果確認タブ"""
    st.subheader("Step 3️⃣ : 結果確認")
    
    if not st.session_state.processed or st.session_state.processor is None:
        st.info("📝 まずは「処理実行」タブで処理を実行してください")
        return
    
    processor = st.session_state.processor
    summary = processor.get_summary()
    
    if not summary:
        st.warning("⚠️ 結果がありません")
        return
    
    # サマリー
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("📊 抽出行数", summary["total_rows"])
    
    with col2:
        st.metric("🏢 部署数", len(summary["departments"]))
    
    with col3:
        st.metric("👔 役職種", len(summary["position_count"]))
    
    st.divider()
    
    # 統計グラフ
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("部署別件数")
        dept_df = pd.DataFrame(
            list(summary["dept_count"].items()),
            columns=["部署", "件数"]
        ).sort_values("件数", ascending=False)
        st.bar_chart(dept_df.set_index("部署"))
    
    with col2:
        st.subheader("役職別件数")
        pos_df = pd.DataFrame(
            list(summary["position_count"].items()),
            columns=["役職", "件数"]
        ).sort_values("件数", ascending=False)
        st.bar_chart(pos_df.set_index("役職"))
    
    st.divider()
    
    # プレビュー
    st.subheader("📋 データプレビュー")
    preview = processor.get_preview(20)
    st.dataframe(preview, use_container_width=True, height=400)
    
    st.divider()
    
    # ダウンロード
    st.subheader("Step 4️⃣ : ダウンロード")
    
    csv_bytes = processor.get_csv_bytes()
    if csv_bytes:
        st.download_button(
            label="⬇️ CSV ダウンロード",
            data=csv_bytes,
            file_name=config.OUTPUT_FILENAME,
            mime="text/csv",
            use_container_width=True,
            type="primary"
        )
        
        st.success("""
        ✅ ダウンロード準備完了
        
        **⚠️ ご注意**: このCSVは個人情報を含むため、慎重に取り扱ってください
        """)
    
    # ログ
    with st.expander("📜 処理ログ"):
        for log in processor.get_logs():
            st.code(log)

def show_info_tab():
    """詳細情報タブ"""
    st.subheader("ℹ️ 詳細情報")
    
    st.markdown("""
    ### 🔍 フィルタリングルール
    
    #### 部署フィルター
    - ✅ **含む**: """ + ", ".join(config.INCLUDE_KEYWORDS) + """
    - ❌ **含まない**: """ + ", ".join(config.EXCLUDE_KEYWORDS) + """
    
    #### 役職変換ルール
    - メール列に `""" + config.MAIL_FILTER_DOMAIN + """` が含まれる場合
    - → 役職を `""" + config.派遣_ROLE + """` に変換
    
    #### 社員コード抽出ルール
    - ニックネーム列から `tg` で始まるデータを除外
    """)
    
    st.divider()
    
    st.subheader("📊 出力データ構造")
    sample_data = {col: "サンプル" for col in config.COLUMN_MAPPING.values()}
    st.dataframe(pd.DataFrame([sample_data]), use_container_width=True)
    
    st.divider()
    
    st.subheader("🆘 トラブルシューティング")
    
    with st.expander("Q: 認証に失敗します"):
        st.write("""
        A: 以下を確認してください：
        - メールアドレスが正しいか
        - パスワードが正しいか
        - 職場PCネットワークに接続しているか
        - VPN接続が必要な場合は接続しているか
        """)
    
    with st.expander("Q: ファイルが見つかりません"):
        st.write(f"""
        A: SharePoint の「{config.SHAREPOINT_LIBRARY}」フォルダに
        「{config.EXCEL_FILENAME}」が存在するか確認してください
        """)
    
    with st.expander("Q: 期待した行数が出力されません"):
        st.write("""
        A: フィルタリング条件を確認してください：
        - 部署に「ＦＣ技」または「出向」が含まれていますか？
        - 部署に「ＴＧテクノ」が含まれていませんか？
        """)

# ============ メイン処理 ============
def main():
    """メインプログラム"""
    check_session_timeout()
    
    if not st.session_state.authenticated:
        show_authentication_form()
    else:
        show_main_screen()

if __name__ == "__main__":
    main()