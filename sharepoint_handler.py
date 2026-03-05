"""
SharePoint アクセスハンドラー
メールアドレスとパスワードでの認証を管理（更新版）
"""

import io
import logging
from typing import Optional, Tuple
import pandas as pd

try:
    from office365.sharepoint.client_context import ClientContext
    from office365.runtime.auth.user_credential import UserCredential
except ImportError:
    raise ImportError("office365ライブラリのインストールが必要です: pip install office365-rest-python-client")

import config

logger = logging.getLogger(__name__)


class SharePointHandler:
    """
    SharePoint ファイルアクセスクラス
    メールアドレスとパスワードで認証
    """

    def __init__(self, email: str, password: str):
        """
        初期化
        
        Args:
            email: SharePoint ユーザーメールアドレス
            password: パスワード
        """
        self.email = email
        self.password = password
        self.ctx: Optional[ClientContext] = None
        self.is_authenticated = False
        self.error_message = ""

    def authenticate(self) -> bool:
        """
        SharePoint に認証
        
        Returns:
            bool: 認証成功時 True
        """
        try:
            logger.info(f"SharePoint 認証開始: {self.email}")
            
            # ユーザー認証情報を作成
            credentials = UserCredential(self.email, self.password)
            
            # クライアントコンテキストを作成（新しいAPI）
            self.ctx = ClientContext(config.SHAREPOINT_SITE_URL).with_credentials(credentials)
            
            # 接続テスト
            web = self.ctx.web
            self.ctx.execute_query()
            
            logger.info(f"SharePoint 認証成功")
            self.is_authenticated = True
            return True
            
        except Exception as e:
            self.error_message = f"認証エラー: {str(e)}"
            logger.error(self.error_message)
            logger.error(f"詳細: {type(e).__name__}")
            return False

    def download_file(self, file_name: str) -> Optional[io.BytesIO]:
        """
        SharePoint からファイルをダウンロード
        
        Args:
            file_name: ダウンロードするファイル名
            
        Returns:
            BytesIO: ファイル内容、失敗時 None
        """
        if not self.is_authenticated:
            self.error_message = "SharePoint に認証されていません"
            logger.error(self.error_message)
            return None

        try:
            logger.info(f"ファイルダウンロード開始: {file_name}")
            
            # ファイルURLを構築
            file_url = f"{config.SHAREPOINT_LIBRARY}/{file_name}"
            
            # ファイルを取得
            file_obj = self.ctx.web.get_file_by_server_relative_url(file_url)
            self.ctx.execute_query()
            
            # バイナリデータを読み込み
            file_content = io.BytesIO(file_obj.read())
            file_content.seek(0)
            
            logger.info(f"ファイルダウンロード完了: {file_name}")
            return file_content
            
        except Exception as e:
            self.error_message = f"ダウンロードエラー: {str(e)}"
            logger.error(self.error_message)
            return None

    def list_files(self) -> Optional[list]:
        """
        SharePoint ライブラリ内のファイル一覧を取得
        
        Returns:
            list: ファイル名のリスト、失敗時 None
        """
        if not self.is_authenticated:
            self.error_message = "SharePoint に認証されていません"
            logger.error(self.error_message)
            return None

        try:
            logger.info("ファイル一覧取得開始")
            
            folder = self.ctx.web.get_folder_by_server_relative_url(
                config.SHAREPOINT_LIBRARY
            )
            files = folder.files
            self.ctx.execute_query()
            
            file_names = [f.name for f in files]
            logger.info(f"ファイル一覧取得完了: {len(file_names)}個")
            return file_names
            
        except Exception as e:
            self.error_message = f"ファイル一覧取得エラー: {str(e)}"
            logger.error(self.error_message)
            return None

    def read_excel(self, file_name: str) -> Optional[Tuple[pd.DataFrame, str]]:
        """
        SharePoint から Excel ファイルを読み込んで DataFrame に変換
        
        Args:
            file_name: ファイル名
            
        Returns:
            Tuple[DataFrame, ""]: (DataFrame, エラーメッセージ)
        """
        file_content = self.download_file(file_name)
        
        if file_content is None:
            return None, self.error_message

        try:
            df = pd.read_excel(file_content, engine='openpyxl')
            logger.info(f"Excel読み込み完了: {len(df)}行")
            return df, ""
            
        except Exception as e:
            error_msg = f"Excel読み込みエラー: {str(e)}"
            logger.error(error_msg)
            return None, error_msg