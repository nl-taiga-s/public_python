import os
import glob

class SearchFile():
    """ファイルを再帰的に検索します"""
    def __init__(self, folder_path: str):
        print(SearchFile.__doc__)
        """初期化する"""
        self.list_file_before = [
            f
            for f in glob.glob(os.path.join(folder_path, "**"), recursive=True)
            if os.path.isfile(f)
        ]

    def extract_by_pattern(self, pattern: str):
        """ファイルのリストから検索パターンで抽出したものを表示します"""
        print(self.extract_by_pattern.__doc__)
        self.list_file_after = [f for f in self.list_file_before if pattern in f]
        print("検索結果: ")
        print(*self.list_file_after, sep="\n")
