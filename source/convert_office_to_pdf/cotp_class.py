import glob
import os

import comtypes.client


class ConvertOfficeToPdf:
    """
    Excel, Word, PowerPointをPDFに一括変換します
    Windows + Microsoft Office(デスクトップ版)が必要です
    """

    def __init__(self):
        """初期化"""
        print(ConvertOfficeToPdf.__doc__)
        self.__input()
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
            print("変換元のファイルがありません。")
            return
        # ファイルリストのポインタ
        self.p = 0
        self.__set_file_path()

    def __input(self):
        """入力する"""
        while True:
            self.folder_path_from = input(
                "ファイルを一括変換するフォルダを指定してください。"
            )
            self.folder_path_to = input(
                "一括変換したファイルを格納するフォルダを指定してください。"
            )
            if self.folder_path_from == "" or self.folder_path_to == "":
                return
            if os.path.exists(self.folder_path_from) and os.path.exists(
                self.folder_path_to
            ):
                # 存在する場合
                if os.path.isdir(self.folder_path_from) and os.path.isdir(
                    self.folder_path_to
                ):
                    # フォルダの場合
                    break

    def __set_file_path(self):
        """ファイルパスを設定する"""
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
                print(f"Excelの処理: {self.current_of_file_path_from}")
                self.convert_excel_to_pdf()
            case var if var in self.file_types["word"]:
                print(f"Wordの処理: {self.current_of_file_path_from}")
                self.convert_word_to_pdf()
            case var if var in self.file_types["powerpoint"]:
                print(f"PowerPointの処理: {self.current_of_file_path_from}")
                self.convert_powerpoint_to_pdf()

    def convert_excel_to_pdf(self):
        """
        ExcelをPDFに変換します
        .xls or .xlsx => .pdf
        """
        try:
            obj = comtypes.client.CreateObject("Excel.Application")
            obj.Visible = False
            f = obj.Workbooks.Open(self.current_of_file_path_from)
            # 0 = pdf
            f.SaveAs(self.current_of_file_path_to, 0)
            f.Close(False)
            obj.Quit()
        except Exception as e:
            print(f"Convert Error from Excel to PDF: {e}")

    def convert_word_to_pdf(self):
        """
        WordをPDFに変換します
        .doc or .docx => .pdf
        """
        try:
            obj = comtypes.client.CreateObject("Word.Application")
            obj.Visible = False
            f = obj.Documents.Open(self.current_of_file_path_from)
            # 17 => pdf
            f.SaveAs(self.current_of_file_path_to, 17)
            f.Close(False)
            obj.Quit()
        except Exception as e:
            print(f"Convert Error from Word to PDF: {e}")

    def convert_powerpoint_to_pdf(self):
        """
        PowerPointをPDFに変換します
        .ppt or .pptx => .pdf
        """
        try:
            obj = comtypes.client.CreateObject("PowerPoint.Application")
            obj.Visible = False
            f = obj.Presentations.Open(self.current_of_file_path_from, WithWindow=False)
            # 32 = pdf
            f.SaveAs(self.current_of_file_path_to, 32)
            f.Close()
            obj.Quit()
        except Exception as e:
            print(f"Convert Error from PowerPoint to PDF: {e}")

    def convert_all(self):
        """フォルダ内のすべてのファイルを処理する"""
        for _ in range(self.number_of_f):
            self.handle_file()
            self.__next()
