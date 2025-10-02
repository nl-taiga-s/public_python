import glob
from logging import Logger
from pathlib import Path


class GetFileList:
    """指定のフォルダ内のファイルのリストを取得します"""

    def __init__(self, logger: Logger):
        """初期化します"""
        self.log: Logger = logger
        self.log.info(self.__class__.__doc__)
        # フォルダパス
        self.folder_path: str = ""
        # 再帰的に検索するかどうか
        self.recursive: bool = False
        # 検索パターン
        self.pattern: str = ""
        # ファイルパスのリスト
        self.lst_file_before: list = []
        self.lst_file_after: list = []
        # ファイルの数
        self.num_of_f_before: int = 0
        self.num_of_f_after: int = 0

    def search_recursively(self) -> bool:
        """再帰的に検索します"""
        try:
            result: bool = False
            RECURSIVE: str = "**" if self.recursive else "*"
            search_folder_p: Path = Path(self.folder_path) / RECURSIVE
            search_folder_s: str = str(search_folder_p)
            self.log.info(f"起点のフォルダパス: {self.folder_path}")
            self.log.info(f"再帰的に検索: {"する" if self.recursive else "しない"}")
            for f in glob.glob(search_folder_s, recursive=self.recursive):
                if Path(f).is_file():
                    self.lst_file_before.append(f)
            if not self.lst_file_before:
                raise Exception("フォルダにファイルがありませんでした。")
            self.num_of_f_before = len(self.lst_file_before)
            self.log.info(f"{self.num_of_f_before}件のファイルがあります。")
            self.log.info("\n".join(self.lst_file_before))
        except Exception as e:
            self.log.error(f"***{self.search_recursively.__doc__} => 失敗しました。***: \n{repr(e)}")
        else:
            result = True
            self.log.info(f"***{self.search_recursively.__doc__} => 成功しました。***")
        finally:
            return result

    def extract_by_pattern(self) -> bool:
        """検索パターンで抽出します"""
        try:
            result: bool = False
            self.log.info(f"検索パターン: {self.pattern if self.pattern else "なし"}")
            self.lst_file_after = [f for f in self.lst_file_before if self.pattern in f]
            if not self.lst_file_after:
                raise Exception("検索パターンによる抽出結果がありませんでした。")
            self.num_of_f_after = len(self.lst_file_after)
            self.log.info(f"{self.num_of_f_after}件のファイルが抽出されました。")
            self.log.info("\n".join(self.lst_file_after))
        except Exception as e:
            self.log.error(f"***{self.extract_by_pattern.__doc__} => 失敗しました。***: \n{repr(e)}")
        else:
            result = True
            self.log.info(f"***{self.extract_by_pattern.__doc__} => 成功しました。***")
        finally:
            return result
