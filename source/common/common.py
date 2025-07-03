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
    """パスのツール"""

    def __init__(self):
        """初期化します"""
        self.obj_of_dt2 = DatetimeTools()

    def convert_path_to_str(self, p: Path) -> str:
        """Path型からstr型に変換します"""
        return str(p)

    def get_file_path_of_result(self, file_path_of_exe: str) -> Path:
        """resultファイルのパスを取得します"""
        # 実行するファイルのディレクトリを取得します
        folder_path_of_exe = Path(file_path_of_exe).parent
        # resultsフォルダのパス
        folder_path_of_results = folder_path_of_exe / 'results'
        # resultsフォルダが存在しない場合は作成します
        folder_path_of_results.mkdir(parents=True, exist_ok=True)
        # 作成するファイル名
        file_name = f"result_{self.obj_of_dt2.convert_for_file_name()}.txt"
        # 作成するファイルのパス
        return folder_path_of_results / file_name

    # def if_unc_path(self, file_path: str) -> Path:
    #     """
    #     UNC(Universal Naming Convention)パスの条件分岐をします
    #     ex. \\\\ZZ.ZZZ.ZZZ.Z
    #     """
    #     fp = Path(file_path)
    #     if fp.startswith(r"\\"):
    #         return fp
    #     return fp.resolve()
