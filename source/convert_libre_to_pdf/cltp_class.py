import glob
import subprocess
from logging import Logger
from pathlib import Path


class ConvertLibreToPDF:
    """
    Excel, Word, PowerPointをPDFに一括変換します
    Excel => .xls, .xlsx
    Word => .doc, .docx
    PowerPoint => .ppt, .pptx
    LibreOfficeが必要です
    """

    def __init__(self, logger: Logger):
        """初期化します"""
        self.log = logger
        self.log.info(self.__class__.__doc__)
        # 変換元のフォルダパス
        self.folder_path_from = ""
        # 変換先のフォルダパス
        self.folder_path_to = ""
        # 対象の変換元のファイルパス
        self.current_of_file_path_from = ""
        # 拡張子を指定する
        self.file_types = {
            "excel": [".xls", ".xlsx"],
            "word": [".doc", ".docx"],
            "powerpoint": [".ppt", ".pptx"],
        }
        # 対象の拡張子の辞書をリストにまとめる
        self.valid_exts = sum(self.file_types.values(), [])
        # ファイルリストのポインタ
        self.p = 0
        # 処理したファイルの数
        self.count = 0
        # 処理が成功したファイルの数
        self.success = 0
        # すべてのファイルを変換できたかどうか
        self.complete = False

    def create_file_list(self) -> bool:
        """ファイルリストを作成します"""
        try:
            result = False
            folder_of_search_as_path_type = Path(self.folder_path_from) / "*"
            search_folder = str(folder_of_search_as_path_type)
            # フィルター前のファイルのリスト
            unfiltered_list_of_f = glob.glob(search_folder)
            # フィルター後のファイルのリスト
            self.filtered_list_of_f = []
            for f in unfiltered_list_of_f:
                file_as_path_type = Path(f)
                if file_as_path_type.suffix.lower() in self.valid_exts:
                    self.filtered_list_of_f.append(f)
            # ファイルの数
            self.number_of_f = len(self.filtered_list_of_f)
            if self.number_of_f == 0:
                raise Exception("変換元のファイルがありません。")
        except Exception as e:
            self.log.error(f"***{self.create_file_list.__doc__} => 失敗しました。***")
            self.log.error(f"error: \n{str(e)}")
        else:
            result = True
            self.set_file_path()
            self.log.info(f"***{self.create_file_list.__doc__} => 成功しました。***")
            self.log.info(f"{self.number_of_f}件のファイルが見つかりました。")
        finally:
            return result

    def set_file_path(self) -> bool:
        """ファイルパスを設定します"""
        try:
            result = False
            self.current_of_file_path_from = self.filtered_list_of_f[self.p]
        except Exception as e:
            self.log.error(f"***{self.set_file_path.__doc__} => 失敗しました。***")
            self.log.error(f"error: \n{str(e)}")
        else:
            result = True
            self.log.info(f"***{self.set_file_path.__doc__} => 成功しました。***")
        finally:
            return result

    def move_to_previous_file(self) -> bool:
        """前のファイルへ"""
        try:
            result = False
            if self.p == 0:
                self.p = self.number_of_f - 1
            else:
                self.p -= 1
            self.set_file_path()
        except Exception as e:
            self.log.error(f"***{self.move_to_previous_file.__doc__} => 失敗しました。***")
            self.log.error(f"error: \n{str(e)}")
        else:
            result = True
            self.log.info(f"***{self.move_to_previous_file.__doc__} => 成功しました。***")
        finally:
            return result

    def move_to_next_file(self) -> bool:
        """次のファイルへ"""
        try:
            result = False
            if self.p == self.number_of_f - 1:
                self.p = 0
            else:
                self.p += 1
            self.set_file_path()
        except Exception as e:
            self.log.error(f"***{self.move_to_next_file.__doc__} => 失敗しました。***")
            self.log.error(f"error: \n{str(e)}")
        else:
            result = True
            self.log.info(f"***{self.move_to_next_file.__doc__} => 成功しました。***")
        finally:
            return result

    def convert_file(self) -> bool:
        """ファイルの種類を判定して、変換を実行します"""
        try:
            result = False
            self.log.info(f"* [{self.count + 1} / {self.number_of_f}] {self.convert_file.__doc__}: ")
            self.log.info(self.current_of_file_path_from)
            file_of_currentfrom_as_path_type = Path(self.current_of_file_path_from)
            ext = file_of_currentfrom_as_path_type.suffix.lower()
            if ext in self.valid_exts:
                result_obj = subprocess.run(
                    ["soffice", "--headless", "--convert-to", "pdf", "--outdir", self.folder_path_to, self.current_of_file_path_from],
                    capture_output=True,
                    text=True,
                )
                result = result_obj.returncode == 0
            self.count += 1
            if result:
                self.success += 1
            if self.count == self.number_of_f:
                if self.success == self.number_of_f:
                    self.complete = True
                    self.log.info("全てのファイルの変換が完了しました。")
                else:
                    raise Exception("一部のファイルの変換が失敗しました。")
        except Exception as e:
            self.log.error(f"error: \n{str(e)}")
        else:
            pass
        finally:
            return result
