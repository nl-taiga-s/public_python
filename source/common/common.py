import datetime
import logging
import platform
import shutil
import sys
from pathlib import Path

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QMessageBox, QWidget


class LogTools:
    """
    logger.debug()
    logger.info()
    logger.warning()
    logger.error()
    logger.critical()
    """

    def __init__(self):
        """初期化します"""
        self.file_path_of_log = ""
        # create logger
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.DEBUG)

    def setup_file_handler(self, file_path: str):
        """FileHandlerを設定します"""
        self.file_handler = logging.FileHandler(file_path, mode='w', encoding='utf-8')
        self.file_handler.setLevel(logging.DEBUG)
        self.STR_OF_FILE_FORMATTER = "%(message)s - [%(levelname)s] - (%(filename)s) - %(asctime)s"
        self.file_formatter = logging.Formatter(self.STR_OF_FILE_FORMATTER)
        self.file_handler.setFormatter(self.file_formatter)
        self.logger.addHandler(self.file_handler)

    def setup_stream_handler(self):
        """StreamHandlerを設定します"""
        self.stream_handler = logging.StreamHandler(sys.stdout)
        self.stream_handler.setLevel(logging.DEBUG)
        self.STR_OF_STREAM_FORMATTER = "%(message)s"
        self.stream_formatter = logging.Formatter(self.STR_OF_STREAM_FORMATTER)
        self.stream_handler.setFormatter(self.stream_formatter)
        self.logger.addHandler(self.stream_handler)


class PlatformTools:
    def __init__(self):
        """初期化します"""
        pass

    def is_wsl(self) -> bool:
        """WSL環境かどうか判定します"""
        if platform.system().lower() != "linux":
            return False
        try:
            with open("/proc/version", "r") as f:
                content = f.read().lower()
                return "microsoft" in content or "wsl" in content
        except Exception:
            return False


class DatetimeTools:
    def __init__(self):
        """初期化します"""
        self.dt = None

    def convert_dt_to_str(self, dt: datetime.datetime = None) -> str:
        """datetime型からstr型に変換します"""
        if dt is None:
            self.dt = datetime.datetime.now()
            dt = self.dt
        # datetime型 => str型
        return dt.strftime("%Y-%m-%d_%H:%M:%S")

    def convert_for_file_name(self, dt: datetime.datetime = None) -> str:
        """ファイル名用に変換します"""
        if dt is None:
            self.dt = datetime.datetime.now()
            dt = self.dt
        # datetime型 => str型
        return dt.strftime("%Y%m%d_%H%M%S")

    def convert_for_metadata_in_pdf(self, utc: str, dt: datetime.datetime = None) -> str:
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
        self.obj_of_dt2 = DatetimeTools()

    def get_file_path_of_log(self, base_path: Path) -> Path:
        """ログファイルのパスを取得します"""
        try:
            # 実行するファイルのディレクトリ
            folder_of_exe_as_path_type = base_path.parent
            # ログフォルダのパス
            folder_of_log_as_path_type = folder_of_exe_as_path_type / "__log__"
            # ログフォルダが存在しない場合は作成します
            folder_of_log_as_path_type.mkdir(parents=True, exist_ok=True)
            # ログファイル名
            file_name_of_log = f"log_{self.obj_of_dt2.convert_for_file_name()}.log"
            # ログファイルのパス
            return folder_of_log_as_path_type / file_name_of_log
        except Exception as e:
            print(e)


class GUITools:
    def __init__(self, parent: QWidget = None):
        """初期化します"""
        self.parent = parent

    def show_error(self, msg: str):
        """エラーのウィンドウを表示します"""
        box = QMessageBox(self.parent)
        box.setIcon(QMessageBox.Icon.Critical)
        box.setWindowTitle("エラー")
        box.setText(msg)
        # 一定時間後に自動終了する
        MILLI_SECONDS = 10000
        QTimer.singleShot(MILLI_SECONDS, box.accept)
        box.exec()


class FileSystemTools:
    def __init__(self):
        """初期化します"""
        pass

    def clear_folder(self, path: Path):
        """フォルダを空にします"""
        if path.exists():
            for element in path.iterdir():
                if element.is_dir():
                    shutil.rmtree(element)
                else:
                    element.unlink()
