import glob
import subprocess
from logging import Logger
from pathlib import Path
from subprocess import CompletedProcess


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
        self.log: Logger = logger
        self.log.info(self.__class__.__doc__)
        # 変換元のフォルダパス
        self.folder_path_from: str = ""
        # 変換先のフォルダパス
        self.folder_path_to: str = ""
        # 拡張子を指定する
        self.file_types: dict = {
            "excel": [".xls", ".xlsx"],
            "word": [".doc", ".docx"],
            "powerpoint": [".ppt", ".pptx"],
        }
        # 対象の拡張子の辞書をリストにまとめる
        self.valid_exts: list = sum(self.file_types.values(), [])
        # 変換元のフォルダのフィルター後のファイルのリスト
        self.filtered_lst_of_f: list = []
        # 変換元のフォルダのファイルの数
        self.number_of_f: int = 0
        # ファイルリストのポインタ
        self.p: int = 0
        # 変換元のファイルパス
        self.current_file_path_from: str = ""
        # 処理したファイルの数
        self.count: int = 0
        # 処理が成功したファイルの数
        self.success: int = 0
        # すべてのファイルを変換できたかどうか
        self.complete: bool = False

    def create_file_lst(self) -> bool:
        """ファイルリストを作成します"""
        result: bool = False
        try:
            search_folder_p: Path = Path(self.folder_path_from) / "*"
            search_folder_s: str = str(search_folder_p)
            # フィルター前のファイルのリスト
            unfiltered_lst_of_f: list = glob.glob(search_folder_s)
            for f in unfiltered_lst_of_f:
                file_p: Path = Path(f)
                if file_p.suffix.lower() in self.valid_exts:
                    self.filtered_lst_of_f.append(f)
            self.number_of_f = len(self.filtered_lst_of_f)
            if self.number_of_f == 0:
                raise Exception("変換元のファイルがありません。")
        except Exception as e:
            self.log.error(f"***{self.create_file_lst.__doc__} => 失敗しました。***: \n{repr(e)}")
            raise
        else:
            result = True
            self.current_file_path_from = self.filtered_lst_of_f[self.p]
            self.log.info(f"***{self.create_file_lst.__doc__} => 成功しました。***")
            self.log.info(f"{self.number_of_f}件のファイルが見つかりました。")
            self.log.info(f"変換先のフォルダ: {self.folder_path_to}")
        finally:
            return result

    def move_to_previous_file(self) -> bool:
        """前のファイルへ"""
        result: bool = False
        try:
            if self.p == 0:
                self.p = self.number_of_f - 1
            else:
                self.p -= 1
            self.current_file_path_from = self.filtered_lst_of_f[self.p]
        except Exception as e:
            self.log.error(f"***{self.move_to_previous_file.__doc__} => 失敗しました。***: \n{repr(e)}")
        else:
            result = True
            self.log.info(f"***{self.move_to_previous_file.__doc__} => 成功しました。***")
        finally:
            return result

    def move_to_next_file(self) -> bool:
        """次のファイルへ"""
        result: bool = False
        try:
            if self.p == self.number_of_f - 1:
                self.p = 0
            else:
                self.p += 1
            self.current_file_path_from = self.filtered_lst_of_f[self.p]
        except Exception as e:
            self.log.error(f"***{self.move_to_next_file.__doc__} => 失敗しました。***: \n{repr(e)}")
        else:
            result = True
            self.log.info(f"***{self.move_to_next_file.__doc__} => 成功しました。***")
        finally:
            return result

    def convert_file(self) -> bool:
        """ファイルの種類を判定して、変換を実行します"""
        result: bool = False
        try:
            self.log.info(f"* [{self.count + 1} / {self.number_of_f}] {self.convert_file.__doc__}: ")
            self.log.info(f"{self.current_file_path_from} => PDF")
            current_file_from_p: Path = Path(self.current_file_path_from)
            ext: str = current_file_from_p.suffix.lower()
            if ext in self.valid_exts:
                convert_obj: CompletedProcess = subprocess.run(
                    ["soffice", "--headless", "--convert-to", "pdf", "--outdir", self.folder_path_to, self.current_file_path_from],
                    capture_output=True,
                    text=True,
                )
                result = not convert_obj.returncode
            self.count += 1
            if result:
                self.success += 1
                self.log.info("***成功しました。***")
            else:
                self.log.error("***失敗しました。***")
            if self.count == self.number_of_f:
                if self.success == self.number_of_f:
                    self.complete = True
                    self.log.info("全てのファイルの変換が完了しました。")
                else:
                    raise Exception("一部のファイルの変換が失敗しました。")
        except Exception as e:
            self.log.error(f"error: \n{repr(e)}")
        else:
            pass
        finally:
            return result
