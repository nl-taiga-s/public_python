import glob
import os
import platform
from pathlib import Path

from source.common.common import DatetimeTools, PathTools

SPECIFIED_OS = "windows"
if platform.system().lower() == SPECIFIED_OS.lower():
    from comtypes.client import CreateObject
else:
    print(f"このスクリプトは、{SPECIFIED_OS}で実行してください。")
    os._exit(0)


class ConvertOfficeToPDF:
    """
    Excel, Word, PowerPointをPDFに一括変換します
    Excel => .xls, .xlsx
    Word => .doc, .docx
    PowerPoint => .ppt, .pptx
    Windows + Microsoft Office(デスクトップ版)が必要です
    """

    def __init__(self, folder_path_from: str, folder_path_to: str):
        """初期化します"""
        print(self.__class__.__doc__)
        self.obj_of_pt = PathTools()
        self.obj_of_dt2 = DatetimeTools()
        self.folder_path_from = folder_path_from
        self.folder_path_to = folder_path_to
        # 拡張子を指定する
        self.file_types = {
            "excel": [".xls", ".xlsx"],
            "word": [".doc", ".docx"],
            "powerpoint": [".ppt", ".pptx"],
        }
        # 対象の拡張子の辞書をリストにまとめる
        valid_exts = sum(self.file_types.values(), [])
        folder_of_search_as_path_type = Path(self.folder_path_from) / "*"
        search_folder = str(folder_of_search_as_path_type)
        # フィルター前のファイルのリスト
        unfiltered_list_of_f = glob.glob(search_folder)
        # フィルター後のファイルのリスト
        self.filtered_list_of_f = []
        for f in unfiltered_list_of_f:
            file_as_path_type = Path(f)
            if self.obj_of_pt.get_extension(file_as_path_type) in valid_exts:
                self.filtered_list_of_f.append(f)
        # ファイルの数
        self.number_of_f = len(self.filtered_list_of_f)
        if self.number_of_f == 0:
            raise ValueError("変換元のファイルがありません。")
        # ファイルリストのポインタ
        self.p = 0
        self.set_file_path()
        # ログファイルのリスト
        self.convert_log = []

    def set_file_path(self):
        """ファイルパスを設定します"""
        # 対象の変換元のファイルパス
        self.current_of_file_path_from = self.filtered_list_of_f[self.p]
        file_of_currentfrom_as_path_type = Path(self.current_of_file_path_from)
        file_name_no_ext = self.obj_of_pt.get_file_name_without_extension(file_of_currentfrom_as_path_type)
        # 対象の変換先のファイルパス
        file_of_currentto_as_path_type = Path(self.folder_path_to) / (file_name_no_ext + ".pdf")
        self.current_of_file_path_to = str(file_of_currentto_as_path_type)

    def move_to_previous_file(self):
        """前のファイルへ"""
        if self.p == 0:
            self.p = self.number_of_f - 1
        else:
            self.p -= 1
        self.set_file_path()

    def move_to_next_file(self):
        """次のファイルへ"""
        if self.p == self.number_of_f - 1:
            self.p = 0
        else:
            self.p += 1
        self.set_file_path()

    def handle_file(self):
        """ファイルの種類を判定して、各処理を実行します"""
        file_of_currentfrom_as_path_type = Path(self.current_of_file_path_from)
        ext = self.obj_of_pt.get_extension(file_of_currentfrom_as_path_type)
        match ext:
            case var if var in self.file_types["excel"]:
                self.convert_excel_to_pdf()
            case var if var in self.file_types["word"]:
                self.convert_word_to_pdf()
            case var if var in self.file_types["powerpoint"]:
                self.convert_powerpoint_to_pdf()

    def convert_excel_to_pdf(self):
        """ExcelをPDFに変換します"""
        print(f"* {self.__class__.convert_excel_to_pdf.__doc__}: ")
        print(f"{self.current_of_file_path_from} => {self.current_of_file_path_to}")
        PDF_NUMBER_OF_EXCEL = 0
        try:
            obj = CreateObject("Excel.Application")
            f = obj.Workbooks.Open(self.current_of_file_path_from, ReadOnly=False)
            f.ExportAsFixedFormat(Filename=self.current_of_file_path_to, Type=PDF_NUMBER_OF_EXCEL)
        except Exception as e:
            print(f"Convert Error from Excel to PDF: {e}")
        else:
            print("It was Successful!")
            self.log_conversion(self.current_of_file_path_from)
        finally:
            if f:
                f.Close()
            if obj:
                obj.Quit()

    def convert_word_to_pdf(self):
        """WordをPDFに変換します"""
        print(f"* {self.__class__.convert_word_to_pdf.__doc__}: ")
        print(f"{self.current_of_file_path_from} => {self.current_of_file_path_to}")
        PDF_NUMBER_OF_WORD = 17
        try:
            obj = CreateObject("Word.Application")
            f = obj.Documents.Open(self.current_of_file_path_from, ReadOnly=False)
            f.ExportAsFixedFormat(
                OutputFileName=self.current_of_file_path_to,
                ExportFormat=PDF_NUMBER_OF_WORD,
            )
        except Exception as e:
            print(f"Convert Error from Word to PDF: {e}")
        else:
            print("It was Successful!")
            self.log_conversion(self.current_of_file_path_from)
        finally:
            if f:
                f.Close()
            if obj:
                obj.Quit()

    def convert_powerpoint_to_pdf(self):
        """PowerPointをPDFに変換します"""
        print(f"* {self.__class__.convert_powerpoint_to_pdf.__doc__}: ")
        print(f"{self.current_of_file_path_from} => {self.current_of_file_path_to}")
        PDF_NUMBER_OF_POWERPOINT = 2
        try:
            obj = CreateObject("PowerPoint.Application")
            f = obj.Presentations.Open(self.current_of_file_path_from, ReadOnly=False)
            f.ExportAsFixedFormat(
                Path=self.current_of_file_path_to,
                FixedFormatType=PDF_NUMBER_OF_POWERPOINT,
            )
        except Exception as e:
            print(f"Convert Error from PowerPoint to PDF: {e}")
        else:
            print("It was Successful!")
            self.log_conversion(self.current_of_file_path_from)
        finally:
            if f:
                f.Close()
            if obj:
                obj.Quit()

    def convert_all(self):
        """指定のフォルダ内の全てのファイルを変換します"""
        print(self.__class__.convert_all.__doc__)
        for _ in range(self.number_of_f):
            self.handle_file()
            self.move_to_next_file()

    def log_conversion(self, file_path: str):
        """変換の結果を記録する"""
        time_stamp = self.obj_of_dt2.convert_dt_to_str()
        self.convert_log.append(f"{file_path},{time_stamp}")

    def write_log(self, file_of_log_as_path_type: Path):
        """処理結果をログに書き出す"""
        file_path_of_log = str(file_of_log_as_path_type)
        try:
            with open(file_path_of_log, "w", encoding="utf-8", newline="") as f:
                f.write("\n".join(self.convert_log))
        except Exception as e:
            print(f"ログファイルの出力に失敗しました。: \n{e}")
        else:
            print(f"ログファイルの出力に成功しました。: \n{file_path_of_log}")
