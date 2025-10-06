import datetime
import logging
import platform
import sys
from logging import FileHandler, Formatter, Logger, StreamHandler
from pathlib import Path

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QMessageBox, QWidget


class LogTools:
    """
    * logger.debug()
    * logger.info()
    * logger.warning()
    * logger.error()
    * logger.critical()
    """

    def __init__(self):
        """初期化します"""
        self.file_path_of_log: str = ""
        # create logger
        self.logger: Logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.DEBUG)

    def setup_file_handler(self, file_path: str) -> bool:
        """FileHandlerを設定します"""
        result: bool = False
        try:
            self.file_handler: FileHandler = logging.FileHandler(file_path, mode='w', encoding='utf-8')
            self.file_handler.setLevel(logging.DEBUG)
            self.STR_OF_FILE_FORMATTER: str = "%(message)s - [%(levelname)s] - (%(filename)s) - %(asctime)s"
            self.file_formatter: Formatter = logging.Formatter(self.STR_OF_FILE_FORMATTER)
            self.file_handler.setFormatter(self.file_formatter)
            self.logger.addHandler(self.file_handler)
        except Exception as e:
            print(f"error: \n{repr(e)}")
            raise
        else:
            result = True
        finally:
            return result

    def setup_stream_handler(self) -> bool:
        """StreamHandlerを設定します"""
        result: bool = False
        try:
            self.stream_handler: StreamHandler = logging.StreamHandler(sys.stdout)
            self.stream_handler.setLevel(logging.DEBUG)
            self.STR_OF_STREAM_FORMATTER: str = "%(message)s"
            self.stream_formatter: Formatter = logging.Formatter(self.STR_OF_STREAM_FORMATTER)
            self.stream_handler.setFormatter(self.stream_formatter)
            self.logger.addHandler(self.stream_handler)
        except Exception as e:
            print(f"error: \n{repr(e)}")
            raise
        else:
            result = True
        finally:
            return result


class PlatformTools:
    def __init__(self):
        """初期化します"""
        pass

    def is_wsl(self) -> bool:
        """WSL環境かどうか判定します"""
        result: bool = False
        try:
            if platform.system().lower() == "linux":
                with open("/proc/version", "r") as f:
                    content: str = f.read().lower()
                    if "microsoft" in content or "wsl" in content:
                        result = True
        except Exception as e:
            print(f"error: \n{repr(e)}")
            raise
        else:
            pass
        finally:
            return result


class DatetimeTools:
    def __init__(self):
        """初期化します"""
        self.dt: datetime.datetime = None

    def convert_dt_to_str(self, dt: datetime.datetime | None = None) -> str:
        """datetime型からstr型に変換します"""
        if dt is None:
            self.dt = datetime.datetime.now()
            dt = self.dt
        # datetime型 => str型
        return dt.strftime("%Y-%m-%d_%H:%M:%S")

    def convert_for_file_name(self, dt: datetime.datetime | None = None) -> str:
        """ファイル名用に変換します"""
        if dt is None:
            self.dt = datetime.datetime.now()
            dt = self.dt
        # datetime型 => str型
        return dt.strftime("%Y%m%d_%H%M%S")

    def convert_for_metadata_in_pdf(self, utc: str, dt: datetime.datetime | None = None) -> str:
        """pdfのメタデータ用に変換します"""
        if dt is None:
            self.dt = datetime.datetime.now()
            dt = self.dt
        # datetime型 => str型
        return dt.strftime(f"D\072%Y%m%d%H%M%S{utc}")


class PathTools:
    """
    * `Path()`は、引数に相対パスもしくは絶対パスの文字列を指定して、その環境に合わせた`Path`オブジェクトを作成します。
    """

    def __init__(self):
        """初期化します"""
        self.obj_of_dt2: DatetimeTools = DatetimeTools()

    def get_file_path_of_log(self, base_path: Path) -> Path:
        """ログファイルのパスを取得します"""
        try:
            # 実行するファイルのディレクトリ
            folder_of_exe_p: Path = base_path.parent
            # ログフォルダのパス
            folder_of_log_p: Path = folder_of_exe_p / "__log__"
            # ログフォルダが存在しない場合は作成します
            folder_of_log_p.mkdir(parents=True, exist_ok=True)
            # ログファイル名
            file_name_of_log: str = f"log_{self.obj_of_dt2.convert_for_file_name()}.log"
        except Exception as e:
            print(f"error: \n{repr(e)}")
            raise
        else:
            # ログファイルのパス
            return folder_of_log_p / file_name_of_log
        finally:
            pass


class GUITools:
    def __init__(self, parent: QWidget | None = None):
        """初期化します"""
        self.parent: QWidget = parent

    def show_error(self, msg: str):
        """エラーのウィンドウを表示します"""
        box: QMessageBox = QMessageBox(self.parent)
        box.setIcon(QMessageBox.Icon.Critical)
        box.setWindowTitle("エラー")
        box.setText(msg)
        box.setStyleSheet("""
            QMessageBox { font-size: 12pt; }
        """)
        # 一定時間後に自動終了する
        MILLI_SECONDS: int = 10000
        QTimer.singleShot(MILLI_SECONDS, box.accept)
        box.exec()
