import datetime
import platform
from pathlib import Path


class PlatformTools:
    """プラットフォームのツール"""

    def __init__(self):
        """初期化します"""
        pass

    def is_wsl(self) -> bool:
        """WSL環境かどうか判定します"""
        if platform.system() != "Linux":
            return False
        try:
            with open("/proc/version", "r") as f:
                content = f.read().lower()
                return "microsoft" in content or "wsl" in content
        except Exception:
            return False


class DatetimeTools:
    """日付と時間のツール"""

    def __init__(self):
        """初期化します"""
        self.dt = datetime.datetime.now()

    def convert_dt_to_str(self, dt: datetime.datetime = None) -> str:
        """datetime型からstr型に変換します"""
        dt = dt or self.dt
        # datetime型 => str型
        return dt.strftime("%Y-%m-%d_%H:%M:%S")

    def convert_for_file_name(self, dt: datetime.datetime = None) -> str:
        """ファイル名用に変換します"""
        dt = dt or self.dt
        # datetime型 => str型
        return dt.strftime("%Y%m%d_%H%M%S")


class PathTools:
    """
    パスのツール
    * `Path()`は、引数の文字列をもとに、その環境に合わせたパスオブジェクトを作成します。
    """

    def __init__(self):
        """初期化します"""
        self.obj_of_dt2 = DatetimeTools()

    def get_dir_path(self, p: Path) -> Path:
        """ディレクトリのパスを取得します"""
        if p.is_file():
            return p.parent
        else:
            return None

    def get_absolute_file_path(self, p: Path) -> Path:
        """ファイルの絶対パスを取得します"""
        if p.is_file():
            return p.resolve()
        else:
            return None

    def get_entire_file_name(self, p: Path) -> str:
        """ファイル名の全体を取得します"""
        if p.is_file():
            return p.name
        else:
            return None

    def get_file_name_without_extension(self, p: Path) -> str:
        """拡張子を除いたファイル名を取得します"""
        if p.is_file():
            return p.stem
        else:
            return None

    def get_extension(self, p: Path) -> str:
        """ファイルの拡張子を取得します"""
        if p.is_file():
            return p.suffix.lower()
        else:
            return None

    def get_file_path_of_log(self, file_path_of_exe: Path) -> Path:
        """logファイルのパスを取得します"""
        try:
            # 実行するファイルのディレクトリを取得します
            folder_path_of_exe = self.get_dir_path(file_path_of_exe)
            # logフォルダのパス
            folder_path_of_log = folder_path_of_exe / "__log__"
            # logフォルダが存在しない場合は作成します
            folder_path_of_log.mkdir(parents=True, exist_ok=True)
            # 作成するファイル名
            file_name_of_log = f"log_{self.obj_of_dt2.convert_for_file_name()}.log"
            # 作成するファイルのパス
            return folder_path_of_log / file_name_of_log
        except Exception as e:
            print(e)

    def if_unc_path(self, file_path: str) -> Path:
        """
        UNC(Universal Naming Convention)パスの条件分岐をします
        ex. \\\\ZZ.ZZZ.ZZZ.Z
        """
        fp = Path(file_path)
        if file_path.startswith(r"\\"):
            return fp
        return fp.resolve()
