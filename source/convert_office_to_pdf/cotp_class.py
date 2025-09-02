import glob
from logging import Logger
from pathlib import Path

from comtypes.client import CreateObject


class ConvertOfficeToPDF:
    """
    Excel, Word, PowerPointをPDFに一括変換します
    Excel => .xls, .xlsx
    Word => .doc, .docx
    PowerPoint => .ppt, .pptx
    Windows + Microsoft Office(デスクトップ版)が必要です
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
        # 対象の変換先のファイルパス
        self.current_of_file_path_to = ""
        # 拡張子を指定する
        self.file_types = {
            "excel": [".xls", ".xlsx"],
            "word": [".doc", ".docx"],
            "powerpoint": [".ppt", ".pptx"],
        }
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
            # 対象の拡張子の辞書をリストにまとめる
            valid_exts = sum(self.file_types.values(), [])
            for f in unfiltered_list_of_f:
                file_as_path_type = Path(f)
                if file_as_path_type.suffix.lower() in valid_exts:
                    self.filtered_list_of_f.append(f)
            # ファイルの数
            self.number_of_f = len(self.filtered_list_of_f)
            if self.number_of_f == 0:
                raise ValueError("変換元のファイルがありません。")
        except ValueError as e:
            self.log.error(str(e))
        except Exception:
            self.log.error(f"***{self.create_file_list.__doc__} => 失敗しました。***")
        else:
            result = True
            self.set_file_path()
            self.log.info(f"{self.create_file_list.__doc__} => 成功しました。")
            self.log.info(f"{self.number_of_f}件のファイルが見つかりました。")
        finally:
            return result

    def set_file_path(self) -> bool:
        """ファイルパスを設定します"""
        try:
            result = False
            self.current_of_file_path_from = self.filtered_list_of_f[self.p]
            file_of_currentfrom_as_path_type = Path(self.current_of_file_path_from)
            file_name_no_ext = file_of_currentfrom_as_path_type.stem
            file_of_currentto_as_path_type = Path(self.folder_path_to) / (file_name_no_ext + ".pdf")
            self.current_of_file_path_to = str(file_of_currentto_as_path_type)
        except Exception:
            self.log.error(f"***{self.set_file_path.__doc__} => 失敗しました。***")
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
        except Exception:
            self.log.error(f"***{self.move_to_previous_file.__doc__} => 失敗しました。***")
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
        except Exception:
            self.log.error(f"***{self.move_to_next_file.__doc__} => 失敗しました。***")
        else:
            result = True
            self.log.info(f"***{self.move_to_next_file.__doc__} => 成功しました。***")
        finally:
            return result

    def convert_excel_to_pdf(self) -> bool:
        """ExcelをPDFに変換します"""
        result = False
        self.log.info(f"* [{self.count + 1} / {self.number_of_f}] {self.__class__.convert_excel_to_pdf.__doc__}: ")
        self.log.info(f"{self.current_of_file_path_from} => {self.current_of_file_path_to}")
        PDF_NUMBER_OF_EXCEL = 0
        try:
            obj = CreateObject("Excel.Application")
            f = obj.Workbooks.Open(self.current_of_file_path_from, ReadOnly=False)
            f.ExportAsFixedFormat(Filename=self.current_of_file_path_to, Type=PDF_NUMBER_OF_EXCEL)
        except Exception:
            self.log.error(f"***{self.convert_excel_to_pdf.__doc__} => 失敗しました。***")
        else:
            result = True
            self.log.info(f"***{self.convert_excel_to_pdf.__doc__} => 成功しました。***")
        finally:
            if f:
                f.Close()
            if obj:
                obj.Quit()
            return result

    def convert_word_to_pdf(self) -> bool:
        """WordをPDFに変換します"""
        result = False
        self.log.info(f"* [{self.count + 1} / {self.number_of_f}] {self.__class__.convert_word_to_pdf.__doc__}: ")
        self.log.info(f"{self.current_of_file_path_from} => {self.current_of_file_path_to}")
        PDF_NUMBER_OF_WORD = 17
        try:
            obj = CreateObject("Word.Application")
            f = obj.Documents.Open(self.current_of_file_path_from, ReadOnly=False)
            f.ExportAsFixedFormat(
                OutputFileName=self.current_of_file_path_to,
                ExportFormat=PDF_NUMBER_OF_WORD,
            )
        except Exception:
            self.log.error(f"***{self.convert_word_to_pdf.__doc__} => 失敗しました。***")
        else:
            result = True
            self.log.info(f"***{self.convert_word_to_pdf.__doc__} => 成功しました。***")
        finally:
            if f:
                f.Close()
            if obj:
                obj.Quit()
            return result

    def convert_powerpoint_to_pdf(self) -> bool:
        """PowerPointをPDFに変換します"""
        result = False
        self.log.info(f"* [{self.count + 1} / {self.number_of_f}] {self.__class__.convert_powerpoint_to_pdf.__doc__}: ")
        self.log.info(f"{self.current_of_file_path_from} => {self.current_of_file_path_to}")
        PDF_NUMBER_OF_POWERPOINT = 2
        try:
            obj = CreateObject("PowerPoint.Application")
            f = obj.Presentations.Open(self.current_of_file_path_from, ReadOnly=False)
            f.ExportAsFixedFormat(
                Path=self.current_of_file_path_to,
                FixedFormatType=PDF_NUMBER_OF_POWERPOINT,
            )
        except Exception:
            self.log.error(f"***{self.convert_powerpoint_to_pdf.__doc__} => 失敗しました。***")
        else:
            result = True
            self.log.info(f"***{self.convert_powerpoint_to_pdf.__doc__} => 成功しました。***")
        finally:
            if f:
                f.Close()
            if obj:
                obj.Quit()
            return result

    def handle_file(self) -> bool:
        """ファイルの種類を判定して、各処理を実行します"""
        try:
            file_of_currentfrom_as_path_type = Path(self.current_of_file_path_from)
            ext = file_of_currentfrom_as_path_type.suffix.lower()
            match ext:
                case var if var in self.file_types["excel"]:
                    result = self.convert_excel_to_pdf()
                case var if var in self.file_types["word"]:
                    result = self.convert_word_to_pdf()
                case var if var in self.file_types["powerpoint"]:
                    result = self.convert_powerpoint_to_pdf()
            self.count += 1
            if result:
                self.success += 1
            if self.count == self.number_of_f:
                if self.success == self.number_of_f:
                    self.complete = True
                    self.log.info("***全てのファイルの変換が完了しました。***")
                else:
                    raise Exception("***一部のファイルの変換が失敗しました。***")
        except Exception as e:
            self.log.error(str(e))
        else:
            pass
        finally:
            return result
