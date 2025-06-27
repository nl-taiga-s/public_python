import glob
import os


class GetFileList:
    """指定のフォルダ内のファイルのリストを再帰的に取得します"""

    def __init__(self, folder_path: str):
        """初期化します"""
        print(self.__class__.__doc__)
        self.list_file_before = [
            os.path.abspath(f)
            for f in glob.glob(os.path.join(folder_path, "**"), recursive=True)
            if os.path.isfile(f)
        ]

    def print_list(self, target: list):
        """出力します"""
        print(*target, sep="\n")

    def extract_by_pattern(self, pattern: str):
        """検索パターンで抽出します"""
        print(self.__class__.extract_by_pattern.__doc__)
        self.list_file_after = [
            os.path.abspath(f) for f in self.list_file_before if pattern in f
        ]
