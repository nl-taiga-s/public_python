import glob
from pathlib import Path

from source.common.common import DatetimeTools, PathTools


class GetFileList:
    """指定のフォルダ内のファイルのリストを取得します"""

    def __init__(self, folder_path: str, bool_of_r: bool):
        """初期化します"""
        self.log = []
        self.REPEAT_TIMES = 50
        self.log.append(self.__class__.__doc__)
        self.obj_of_pt = PathTools()
        self.obj_of_dt2 = DatetimeTools()
        self.log.append(f"起点のフォルダパス: {folder_path}")
        pattern = "**" if bool_of_r else "*"
        self.log.append(f"再帰的に検索: {"する" if bool_of_r else "しない"}")
        folder_of_search_as_path_type = Path(folder_path) / pattern
        search_folder = str(folder_of_search_as_path_type)
        self.list_file_before = []
        for f in glob.glob(search_folder, recursive=bool_of_r):
            if Path(f).is_file():
                self.list_file_before.append(f)

    def extract_by_pattern(self, pattern: str) -> list:
        """検索パターンで抽出します"""
        try:
            result = False
            local_log = []
            local_log.append(">" * self.REPEAT_TIMES)
            local_log.append(f"検索パターン: {pattern if pattern else "なし"}")
            self.list_file_after = [f for f in self.list_file_before if pattern in f]
            if self.list_file_after:
                for element in self.list_file_after:
                    local_log.append(f"{element}, {self.obj_of_dt2.convert_dt_to_str()}")
            else:
                raise ValueError
        except ValueError:
            local_log.append("***検索結果がありませんでした。***")
        except Exception:
            local_log.append("***検索が失敗しました。***")
        else:
            result = True
            local_log.append("***検索が成功しました。***")
        finally:
            local_log.append("<" * self.REPEAT_TIMES)
            self.log.extend(local_log)
            return [result, local_log]

    def write_log(self, file_of_log_as_path_type: Path) -> list:
        """処理結果をログに書き出す"""
        file_of_log_as_str_type = str(file_of_log_as_path_type)
        try:
            result = False
            with open(file_of_log_as_str_type, "w", encoding="utf-8", newline="") as f:
                f.write("\n".join(self.log))
        except Exception:
            return [result, ""]
        else:
            result = True
            return [result, file_of_log_as_str_type]
