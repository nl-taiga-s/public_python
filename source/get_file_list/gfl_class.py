import glob
import os

from source.common.common import DateTimeTools, PathTools


class GetFileList:
    """指定のフォルダ内のファイルのリストを取得します"""

    def __init__(self, folder_path: str, bool_of_r: bool):
        """初期化します"""
        print(self.__class__.__doc__)
        self.obj_of_pt = PathTools()
        self.obj_of_dt2 = DateTimeTools()
        pattern = "**" if bool_of_r else "*"
        self.list_file_before = [
            self.obj_of_pt.if_unc_path(f)
            for f in glob.glob(os.path.join(folder_path, pattern), recursive=bool_of_r)
            if os.path.isfile(f)
        ]
        self.now = self.obj_of_dt2.get_datetime_now()

    def print_list(self, target: list):
        """リストを出力します"""
        print(*target, sep="\n")

    def extract_by_pattern(self, pattern: str):
        """検索パターンで抽出します"""
        self.list_file_after = [f for f in self.list_file_before if pattern in f]
        self.list_of_csv = [[file_name, self.now] for file_name in self.list_file_after]
