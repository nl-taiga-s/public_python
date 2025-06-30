import csv
import datetime
import os
import platform


class PlatFormTools:
    """プラットフォームのツール"""

    def __init__(self):
        """初期化します"""
        print(self.__class__.__doc__)

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


class DateTimeTools:
    """日付と時間のツール"""

    def __init__(self):
        """初期化します"""
        print(self.__class__.__doc__)
        self.dt = datetime.datetime.now()

    def get_datetime_now(self) -> str:
        """現在の日時を取得します"""
        # datetime型 => str型
        return self.dt.strftime("%Y-%m-%d_%H:%M:%S")

    def format_for_file_name(self) -> str:
        """ファイル名用に整形した時間を取得します"""
        return self.dt.strftime("%Y%m%d_%H%M%S")


class ListTools:
    """リスト型のツール"""

    def __init__(self):
        """初期化します"""
        print(self.__class__.__doc__)

    def is_list_of_more_than_2nd(self, lst: list) -> bool:
        """2次元以上のリストか判定します"""
        return isinstance(lst, list) and all(isinstance(i, list) for i in lst)


class CsvTools:
    """CSVのツール"""

    def __init__(self):
        """初期化します"""
        print(self.__class__.__doc__)
        self.obj_of_l = ListTools()

    def write_list(self, file_path: str, lst: list):
        """リストを書き込みます"""
        with open(file_path, "w", encoding="utf-8") as f:
            w_of_csv = csv.writer(f)
            if self.obj_of_l.is_list_of_more_than_2nd(lst):
                w_of_csv.writerows(lst)
            else:
                w_of_csv.writerow(lst)


class PathTools:
    """パスのツール"""

    def __init__(self):
        """初期化します"""
        print(self.__class__.__doc__)
        self.obj_of_pf = PlatFormTools()

    def get_file_path_of_log(self, script: str, dt: str) -> str:
        """ログファイルパスを取得します"""
        # スクリプトのあるディレクトリを取得します
        script_dir = os.path.dirname(os.path.abspath(script))
        # resultフォルダのパス
        result_dir = os.path.join(script_dir, 'result')
        # フォルダが存在しない場合は作成します
        os.makedirs(result_dir, exist_ok=True)
        # 作成するファイルのパス
        return os.path.join(result_dir, f'result_{dt}.log')

    def to_path_seaparator_for_os(self, target_path: str) -> str:
        """
        OSに応じて、パスの区切り文字を統一します
        * Windows: "/" => "\\"
        * WSL: "\\" => "/"
        """
        system_name = platform.system()
        try:
            if self.obj_of_pf.is_wsl():
                # WSL
                return os.path.normpath(target_path)
            elif system_name == "Windows":
                # Windows
                return target_path.replace("\\", "/")
            else:
                return target_path
        except Exception as e:
            print(e)

    def if_unc_path(self, target: str) -> str:
        """
        UNC(Universal Naming Convention)パスの条件分岐をします
        ex. \\\\ZZ.ZZZ.ZZZ.Z
        """
        if target.startswith(r"\\"):
            return target
        return os.path.abspath(target)
