"""
メンバー管理表処理クラス（Streamlit環境対応）
"""

import logging
import io
from typing import Optional, Tuple, Dict, List
import pandas as pd
from pathlib import Path
import config

logger = logging.getLogger(__name__)


class MemberListProcessor:
    """
    メンバー管理表を処理して名簿リストを生成するクラス
    """

    def __init__(self):
        """初期化"""
        self.df_source: Optional[pd.DataFrame] = None
        self.df_output: Optional[pd.DataFrame] = None
        self.processing_log: List[str] = []

    def _log(self, message: str, level: str = "info") -> None:
        """ログメッセージを記録"""
        self.processing_log.append(f"[{level.upper()}] {message}")
        if level == "error":
            logger.error(message)
        elif level == "warning":
            logger.warning(message)
        else:
            logger.info(message)

    def load_from_file(self, file_obj: io.BytesIO) -> bool:
        """
        ファイルオブジェクトからExcelを読み込み
        
        Args:
            file_obj: ファイルオブジェクト
            
        Returns:
            bool: 成功時True
        """
        try:
            file_obj.seek(0)
            self.df_source = pd.read_excel(file_obj, engine='openpyxl')
            self._log(f"ファイル読み込み成功: {len(self.df_source)}行")
            return True
        except Exception as e:
            self._log(f"ファイル読み込みエラー: {str(e)}", "error")
            return False

    def load_from_dataframe(self, df: pd.DataFrame) -> bool:
        """
        DataFrameから読み込み
        
        Args:
            df: pandas DataFrame
            
        Returns:
            bool: 成功時True
        """
        try:
            self.df_source = df.copy()
            self._log(f"DataFrame読み込み成功: {len(self.df_source)}行")
            return True
        except Exception as e:
            self._log(f"DataFrame読み込みエラー: {str(e)}", "error")
            return False

    def validate_columns(self) -> bool:
        """必要な列が全て存在するか検証"""
        if self.df_source is None:
            self._log("ソースデータが読み込まれていません", "error")
            return False

        required_columns = list(config.SOURCE_COLUMNS.values())
        available_columns = list(self.df_source.columns)
        missing_columns = [col for col in required_columns if col not in available_columns]

        if missing_columns:
            self._log(f"必要な列が不足: {missing_columns}", "error")
            self._log(f"利用可能な列: {available_columns}", "warning")
            return False

        self._log("列検証完了: 全ての必要な列が存在します")
        return True

    def _should_include_member(self, row: pd.Series) -> bool:
        """行がフィルター条件を満たすかチェック"""
        department = str(row[config.SOURCE_COLUMNS["department"]])

        # 除外キーワードチェック
        if any(keyword in department for keyword in config.EXCLUDE_KEYWORDS):
            return False

        # 含むべきキーワードチェック
        if not any(keyword in department for keyword in config.INCLUDE_KEYWORDS):
            return False

        return True

    def _get_position(self, row: pd.Series) -> str:
        """役職を取得（メール条件で派遣・請負社員に変換）"""
        mail = str(row[config.SOURCE_COLUMNS["mail"]])

        if config.MAIL_FILTER_DOMAIN in mail:
            return config.派遣_ROLE

        return str(row[config.SOURCE_COLUMNS["position"]])

    def _get_full_name(self, row: pd.Series) -> str:
        """姓名を結合して取得"""
        last_name = str(row[config.SOURCE_COLUMNS["last_name"]]).strip()
        first_name = str(row[config.SOURCE_COLUMNS["first_name"]]).strip()
        return f"{last_name} {first_name}".strip()

    def _get_employee_code(self, row: pd.Series) -> str:
        """社員コードを取得（ニックネームから"tg"を除外）"""
        nickname = str(row[config.SOURCE_COLUMNS["nickname"]]).strip().lower()

        if nickname.startswith("tg"):
            return ""

        return str(row[config.SOURCE_COLUMNS["nickname"]]).strip()

    def process(self) -> bool:
        """
        メンバー情報を処理して名簿リストを生成
        
        Returns:
            bool: 成功時True
        """
        if not self.validate_columns():
            return False

        self._log(f"処理開始: {len(self.df_source)}行")

        # フィルタリング
        df_filtered = self.df_source[
            self.df_source.apply(self._should_include_member, axis=1)
        ].copy()

        self._log(f"フィルター後: {len(df_filtered)}行")

        if len(df_filtered) == 0:
            self._log("フィルター条件を満たす行がありません", "warning")
            self.df_output = pd.DataFrame(columns=config.COLUMN_MAPPING.values())
            return True

        # 列の抽出と変換
        output_data = {
            config.COLUMN_MAPPING["department"]: df_filtered[
                config.SOURCE_COLUMNS["department"]
            ],
            config.COLUMN_MAPPING["group"]: "",
            config.COLUMN_MAPPING["team"]: "",
            config.COLUMN_MAPPING["position"]: df_filtered.apply(
                self._get_position, axis=1
            ),
            config.COLUMN_MAPPING["name"]: df_filtered.apply(
                self._get_full_name, axis=1
            ),
            config.COLUMN_MAPPING["employee_code"]: df_filtered.apply(
                self._get_employee_code, axis=1
            )
        }

        self.df_output = pd.DataFrame(output_data)
        self._log(f"処理完了: 出力{len(self.df_output)}行")
        return True

    def get_csv_bytes(self) -> Optional[bytes]:
        """CSVをバイト列で取得（ダウンロード用）"""
        if self.df_output is None:
            self._log("処理済みデータがありません", "error")
            return None

        csv_buffer = io.StringIO()
        self.df_output.to_csv(
            csv_buffer,
            index=False,
            encoding='utf-8',
            quoting=1
        )
        return csv_buffer.getvalue().encode(config.OUTPUT_ENCODING)

    def get_summary(self) -> Dict:
        """処理結果のサマリーを取得"""
        if self.df_output is None:
            return {}

        return {
            "total_rows": len(self.df_output),
            "departments": self.df_output[config.COLUMN_MAPPING["department"]].unique().tolist(),
            "dept_count": self.df_output[config.COLUMN_MAPPING["department"]].value_counts().to_dict(),
            "position_count": self.df_output[config.COLUMN_MAPPING["position"]].value_counts().to_dict(),
        }

    def get_preview(self, n: int = 10) -> Optional[pd.DataFrame]:
        """処理結果のプレビューを取得"""
        if self.df_output is None:
            return None
        return self.df_output.head(n)

    def get_logs(self) -> List[str]:
        """処理ログを取得"""
        return self.processing_log