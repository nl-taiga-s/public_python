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
        s_f = Path(folder_path) / pattern
        search_folder = str(s_f)
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
