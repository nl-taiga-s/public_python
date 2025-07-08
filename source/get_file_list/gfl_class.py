import glob
from pathlib import Path

from source.common.common import DatetimeTools, PathTools


class GetFileList:
    """指定のフォルダ内のファイルのリストを取得します"""

    def __init__(self, folder_path: str, bool_of_r: bool):
        """初期化します"""
        print(self.__class__.__doc__)
        self.obj_of_pt = PathTools()
        self.obj_of_dt2 = DatetimeTools()
        pattern = "**" if bool_of_r else "*"
        folder_of_search_as_path_type = Path(folder_path) / pattern
        search_folder = str(folder_of_search_as_path_type)
        self.list_file_before = []
        for f in glob.glob(
            search_folder,
            recursive=bool_of_r,
        ):
            if Path(f).is_file():
                self.list_file_before.append(f)

    def extract_by_pattern(self, pattern: str):
        """検索パターンで抽出します"""
        self.list_file_after = [f for f in self.list_file_before if pattern in f]

    def write_log(self, file_of_log_as_path_type: Path, lst: list):
        """処理結果をログに書き出す"""
        file_path_of_log = str(file_of_log_as_path_type)
        try:
            with open(file_path_of_log, "w", encoding="utf-8", newline="") as f:
                for element in lst:
                    f.write(f"{element},")
                    f.write(f"{self.obj_of_dt2.convert_dt_to_str()}\n")
        except Exception as e:
            print(f"ログファイルの出力に失敗しました。: \n{e}")
        else:
            print(f"ログファイルの出力に成功しました。: \n{file_path_of_log}")
