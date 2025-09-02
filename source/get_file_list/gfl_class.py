import glob
from logging import Logger
from pathlib import Path


class GetFileList:
    """指定のフォルダ内のファイルのリストを取得します"""

    def __init__(self, logger: Logger):
        """初期化します"""
        self.log = logger
        self.log.info(self.__class__.__doc__)
        # フォルダパス
        self.folder_path = ""
        # 再帰的に検索するかどうか
        self.bool_of_r = False
        # 検索パターン
        self.pattern = ""

    def search_recursively(self) -> bool:
        """再帰的に検索します"""
        try:
            result = False
            RECURSIVE = "**" if self.bool_of_r else "*"
            search_folder_as_path_type = Path(self.folder_path) / RECURSIVE
            search_folder = str(search_folder_as_path_type)
            self.log.info(f"起点のフォルダパス: {self.folder_path}")
            self.log.info(f"再帰的に検索: {"する" if self.bool_of_r else "しない"}")
            self.list_file_before = []
            for f in glob.glob(search_folder, recursive=self.bool_of_r):
                if Path(f).is_file():
                    self.list_file_before.append(f)
            if not self.list_file_before:
                raise ValueError("***フォルダにファイルがありませんでした。***")
            self.number_of_f_before = len(self.list_file_before)
            self.log.info(f"{self.number_of_f_before}件のファイルがあります。")
            self.log.info("\n".join(self.list_file_before))
        except ValueError as e:
            self.log.error(str(e))
        except Exception:
            self.log.error(f"***{self.search_recursively.__doc__} => 失敗しました。***")
        else:
            result = True
            self.log.info(f"***{self.search_recursively.__doc__} => 成功しました。***")
        finally:
            return result

    def extract_by_pattern(self) -> bool:
        """検索パターンで抽出します"""
        try:
            result = False
            self.log.info(f"検索パターン: {self.pattern if self.pattern else "なし"}")
            self.list_file_after = [f for f in self.list_file_before if self.pattern in f]
            if not self.list_file_after:
                raise ValueError("***検索パターンによる抽出結果がありませんでした。***")
            # ファイルの数
            self.number_of_f_after = len(self.list_file_after)
            self.log.info(f"{self.number_of_f_after}件のファイルが抽出されました。")
            self.log.info("\n".join(self.list_file_after))
        except ValueError as e:
            self.log.error(str(e))
        except Exception:
            self.log.error(f"***{self.extract_by_pattern.__doc__} => 失敗しました。***")
        else:
            result = True
            self.log.info(f"***{self.extract_by_pattern.__doc__} => 成功しました。***")
        finally:
            return result
