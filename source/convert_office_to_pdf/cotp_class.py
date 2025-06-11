import glob
import os
import platform

if platform.system() != "Windows":
    raise EnvironmentError("このスクリプトは、Windowsで実行してください。")
from comtypes.client import CreateObject

class ConvertOfficeToPdf:
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
        # 変換元のフォルダのファイルリストを取得する
        self.list_of_f = glob.glob(os.path.join(self.folder_path_from, "*"))
        # 絶対パスに変換する
        self.list_of_f = [
            os.path.abspath(f)
            for f in self.list_of_f
            if os.path.splitext(f)[1].lower() in valid_exts
        ]
        # ファイルの数
        self.number_of_f = len(self.list_of_f)
        if self.number_of_f == 0:
            raise ValueError("変換元のファイルがありません。")
        # ファイルリストのポインタ
        self.p = 0
        self.__set_file_path()

    def __set_file_path(self):
        """ファイルパスを設定します"""
        # 対象の変換元のファイルパス
        self.current_of_file_path_from = self.list_of_f[self.p]
        file_name_no_ext = os.path.splitext(
            os.path.basename(self.current_of_file_path_from)
        )[0]
        # 対象の変換先のファイルパス
        self.current_of_file_path_to = os.path.join(
            self.folder_path_to, file_name_no_ext + ".pdf"
        )

    def __previous(self):
        """前のファイルへ"""
        if self.p == 0:
            self.p = self.number_of_f - 1
        else:
            self.p -= 1
        self.__set_file_path()

    def __next(self):
        """次のファイルへ"""
        if self.p == self.number_of_f - 1:
            self.p = 0
        else:
            self.p += 1
        self.__set_file_path()

    def handle_file(self):
        """ファイルの種類を判定して、各処理を実行します"""
        ext = os.path.splitext(self.current_of_file_path_from)[1].lower()
        match ext:
            case var if var in self.file_types["excel"]:
                self.convert_excel_to_pdf()
            case var if var in self.file_types["word"]:
                self.convert_word_to_pdf()
            case var if var in self.file_types["powerpoint"]:
                self.convert_powerpoint_to_pdf()

    def convert_excel_to_pdf(self):
        """ExcelをPDFに変換します"""
        print(f"{self.__class__.convert_excel_to_pdf.__doc__}: ")
        print(f"{self.current_of_file_path_from} => {self.current_of_file_path_to}")
        PDF_NUMBER_OF_EXCEL = 0
        try:
            obj = CreateObject("Excel.Application")
            f = obj.Workbooks.Open(self.current_of_file_path_from)
            f.ExportAsFixedFormat(
                Filename=self.current_of_file_path_to, Type=PDF_NUMBER_OF_EXCEL
            )
            print("It was Successful!")
        except Exception as e:
            print(f"Convert Error from Excel to PDF: {e}")
        finally:
            if f:
                f.Close()
            if obj:
                obj.Quit()

    def convert_word_to_pdf(self):
        """WordをPDFに変換します"""
        print(f"{self.__class__.convert_word_to_pdf.__doc__}: ")
        print(f"{self.current_of_file_path_from} => {self.current_of_file_path_to}")
        PDF_NUMBER_OF_WORD = 17
        try:
            obj = CreateObject("Word.Application")
            f = obj.Documents.Open(self.current_of_file_path_from)
            f.ExportAsFixedFormat(
                OutputFileName=self.current_of_file_path_to,
                ExportFormat=PDF_NUMBER_OF_WORD,
            )
            print("It was Successful!")
        except Exception as e:
            print(f"Convert Error from Word to PDF: {e}")
        finally:
            if f:
                f.Close()
            if obj:
                obj.Quit()

    def convert_powerpoint_to_pdf(self):
        """PowerPointをPDFに変換します"""
        print(f"{self.__class__.convert_powerpoint_to_pdf.__doc__}: ")
        print(f"{self.current_of_file_path_from} => {self.current_of_file_path_to}")
        PDF_NUMBER_OF_POWERPOINT = 2
        try:
            obj = CreateObject("PowerPoint.Application")
            f = obj.Presentations.Open(self.current_of_file_path_from)
            f.ExportAsFixedFormat(
                Path=self.current_of_file_path_to,
                FixedFormatType=PDF_NUMBER_OF_POWERPOINT,
            )
            print("It was Successful!")
        except Exception as e:
            print(f"Convert Error from PowerPoint to PDF: {e}")
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
            self.__next()
