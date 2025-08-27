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
        # フォルダパス
        self.folder_path = folder_path
        # 再帰的に検索するかどうか
        self.bool_of_r = bool_of_r
        pattern = "**" if self.bool_of_r else "*"
        folder_of_search_as_path_type = Path(self.folder_path) / pattern
        search_folder = str(folder_of_search_as_path_type)
        self.list_file_before = []
        for f in glob.glob(search_folder, recursive=self.bool_of_r):
            if Path(f).is_file():
                self.list_file_before.append(f)

    def extract_by_pattern(self, pattern: str) -> list:
        """検索パターンで抽出します"""
        try:
            result = False
            local_log = []
            # 検索パターン
            self.pattern = pattern
            local_log.append(">" * self.REPEAT_TIMES)
            local_log.append(f"起点のフォルダパス: {self.folder_path}")
            local_log.append(f"再帰的に検索: {"する" if self.bool_of_r else "しない"}")
            local_log.append(f"検索パターン: {self.pattern if self.pattern else "なし"}")
            self.list_file_after = [f for f in self.list_file_before if self.pattern in f]
            # ファイルの数
            self.number_of_f = len(self.list_file_after)
            if self.list_file_after:
                local_log.append(f"{self.number_of_f}件のファイルが抽出されました。")
                for element in self.list_file_after:
                    local_log.append(element)
            else:
                raise ValueError
        except ValueError:
            local_log.append("***検索パターンによる抽出結果がありませんでした。***")
        except Exception:
            local_log.append("***検索パターンによる抽出が失敗しました。***")
        else:
            result = True
            local_log.append("***検索パターンによる抽出が成功しました。***")
        finally:
            time_stamp = self.obj_of_dt2.convert_dt_to_str()
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            self.log.extend(local_log)
            return [result, local_log]

    def write_log(self, file_of_log_as_path_type: Path) -> list:
        """処理結果をログに書き出す"""
        file_of_log_as_str_type = str(file_of_log_as_path_type)
        try:
            result = False
            fp = ""
            with open(file_of_log_as_str_type, "w", encoding="utf-8", newline="") as f:
                f.write("\n".join(self.log))
        except Exception:
            pass
        else:
            result = True
            fp = file_of_log_as_str_type
        finally:
            return [result, fp]
