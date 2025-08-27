import glob
from pathlib import Path

from comtypes.client import CreateObject

from source.common.common import DatetimeTools, PathTools


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
        self.log = []
        self.REPEAT_TIMES = 50
        self.log.append(self.__class__.__doc__)
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
        # ファイルリストのポインタ
        self.p = 0
        # 処理したファイルの数
        self.count = 0
        # 処理が成功したファイルの数
        self.success = 0
        # すべてのファイルを変換できたかどうか
        self.complete = False
        try:
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
            print(str(e))
        else:
            self.set_file_path()
            self.log.append(f"{self.number_of_f}件のファイルを一括変換します。")

    def set_file_path(self):
        """ファイルパスを設定します"""
        # 対象の変換元のファイルパス
        self.current_of_file_path_from = self.filtered_list_of_f[self.p]
        file_of_currentfrom_as_path_type = Path(self.current_of_file_path_from)
        file_name_no_ext = file_of_currentfrom_as_path_type.stem
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

    def convert_excel_to_pdf(self) -> list:
        """ExcelをPDFに変換します"""
        result = False
        local_log = []
        local_log.append(">" * self.REPEAT_TIMES)
        local_log.append(f"* [{self.count + 1} / {self.number_of_f}] {self.__class__.convert_excel_to_pdf.__doc__}: ")
        local_log.append(f"{self.current_of_file_path_from} => {self.current_of_file_path_to}")
        PDF_NUMBER_OF_EXCEL = 0
        try:
            obj = CreateObject("Excel.Application")
            f = obj.Workbooks.Open(self.current_of_file_path_from, ReadOnly=False)
            f.ExportAsFixedFormat(Filename=self.current_of_file_path_to, Type=PDF_NUMBER_OF_EXCEL)
        except Exception as e:
            local_log.append(f"Convert Error from Excel to PDF: {str(e)}")
        else:
            result = True
            local_log.append("It was Successful!")
        finally:
            if f:
                f.Close()
            if obj:
                obj.Quit()
            time_stamp = self.obj_of_dt2.convert_dt_to_str()
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            return [result, local_log]

    def convert_word_to_pdf(self) -> list:
        """WordをPDFに変換します"""
        result = False
        local_log = []
        local_log.append(">" * self.REPEAT_TIMES)
        local_log.append(f"* [{self.count + 1} / {self.number_of_f}] {self.__class__.convert_word_to_pdf.__doc__}: ")
        local_log.append(f"{self.current_of_file_path_from} => {self.current_of_file_path_to}")
        PDF_NUMBER_OF_WORD = 17
        try:
            obj = CreateObject("Word.Application")
            f = obj.Documents.Open(self.current_of_file_path_from, ReadOnly=False)
            f.ExportAsFixedFormat(
                OutputFileName=self.current_of_file_path_to,
                ExportFormat=PDF_NUMBER_OF_WORD,
            )
        except Exception as e:
            local_log.append(f"Convert Error from Word to PDF: {str(e)}")
        else:
            result = True
            local_log.append("It was Successful!")
        finally:
            if f:
                f.Close()
            if obj:
                obj.Quit()
            time_stamp = self.obj_of_dt2.convert_dt_to_str()
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            return [result, local_log]

    def convert_powerpoint_to_pdf(self) -> list:
        """PowerPointをPDFに変換します"""
        result = False
        local_log = []
        local_log.append(">" * self.REPEAT_TIMES)
        local_log.append(f"* [{self.count + 1} / {self.number_of_f}] {self.__class__.convert_powerpoint_to_pdf.__doc__}: ")
        local_log.append(f"{self.current_of_file_path_from} => {self.current_of_file_path_to}")
        PDF_NUMBER_OF_POWERPOINT = 2
        try:
            obj = CreateObject("PowerPoint.Application")
            f = obj.Presentations.Open(self.current_of_file_path_from, ReadOnly=False)
            f.ExportAsFixedFormat(
                Path=self.current_of_file_path_to,
                FixedFormatType=PDF_NUMBER_OF_POWERPOINT,
            )
        except Exception as e:
            local_log.append(f"Convert Error from PowerPoint to PDF: {str(e)}")
        else:
            result = True
            local_log.append("It was Successful!")
        finally:
            if f:
                f.Close()
            if obj:
                obj.Quit()
            time_stamp = self.obj_of_dt2.convert_dt_to_str()
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            return [result, local_log]

    def handle_file(self) -> list:
        """ファイルの種類を判定して、各処理を実行します"""
        file_of_currentfrom_as_path_type = Path(self.current_of_file_path_from)
        ext = file_of_currentfrom_as_path_type.suffix.lower()
        match ext:
            case var if var in self.file_types["excel"]:
                result, log = self.convert_excel_to_pdf()
            case var if var in self.file_types["word"]:
                result, log = self.convert_word_to_pdf()
            case var if var in self.file_types["powerpoint"]:
                result, log = self.convert_powerpoint_to_pdf()
        self.count += 1
        if result:
            self.success += 1
        if self.count == self.number_of_f:
            if self.success == self.number_of_f:
                self.complete = True
                log.append("全てのファイルの変換が完了しました。")
            else:
                log.append("一部のファイルの変換が失敗しました。")
        self.log.extend(log)
        return [result, log]

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
