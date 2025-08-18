import datetime
from pathlib import Path

from pypdf import PdfReader, PdfWriter

from source.common.common import DatetimeTools, PathTools


class PdfTools:
    """PDFを扱うツール"""

    def __init__(self):
        """初期化します"""
        print(self.__class__.__doc__)
        self.obj_of_dt2 = DatetimeTools()
        self.obj_of_pt = PathTools()
        self.file_path = None
        self.reader = None
        self.num_of_pages = None
        self.writer = None
        self.metadata_of_reader = None
        self.metadata_of_writer = {}
        self.EXTENSION = ".pdf"
        self.UTC_OF_JP = "+09'00'"
        self.fields = [
            ("title", "/Title"),  # タイトル
            ("author", "/Author"),  # 作成者
            ("subject", "/Subject"),  # サブタイトル
            ("creator", "/Creator"),  # アプリケーション
            ("producer", "/Producer"),  # PDF変換
            ("keywords", "/Keywords"),  # キーワード
            ("creation_date", "/CreationDate"),  # 作成日
            ("modification_date", "/ModDate"),  # 更新日
        ]
        self.creation_date = None
        self.modification_date = None
        self.lst_of_text_in_pages = []
        self.log = []
        self.REPEAT_TIMES = 50

    def encrypt(self, file_path: str, password: str) -> bool:
        """暗号化します"""
        try:
            result = False
            local_log = []
            local_log.append(">" * self.REPEAT_TIMES)
            self.file_path = file_path
            self.reader = PdfReader(self.file_path)
            self.writer = PdfWriter(clone_from=self.reader)
            self.writer.encrypt(password, algorithm="AES-256")
            with open(self.file_path, "wb") as f:
                self.writer.write(f)
        except Exception as e:
            local_log.append(f"暗号化に失敗しました。: \n{e}")
        else:
            result = True
            local_log.append(f"暗号化に成功しました。: \n{self.file_path}")
        finally:
            time_stamp = self.obj_of_dt2.convert_dt_to_str(datetime.datetime.now())
            local_log.append(time_stamp)
            local_log.append(f"password: {password}")
            local_log.append("<" * self.REPEAT_TIMES)
            print(local_log)
            self.log.extend(local_log)
            return result

    def decrypt(self, file_path: str, password: str) -> bool:
        """復号化します"""
        try:
            result = False
            local_log = []
            local_log.append(">" * self.REPEAT_TIMES)
            self.file_path = file_path
            self.reader = PdfReader(self.file_path)
            self.reader.decrypt(password)
            self.writer = PdfWriter(clone_from=self.reader)
            with open(self.file_path, "wb") as f:
                self.writer.write(f)
        except Exception as e:
            local_log.append(f"復号化に失敗しました。: \n{e}")
        else:
            result = True
            local_log.append(f"復号化に成功しました。: \n{self.file_path}")
        finally:
            time_stamp = self.obj_of_dt2.convert_dt_to_str(datetime.datetime.now())
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            print(local_log)
            self.log.extend(local_log)
            return result

    def read_file(self, file_path: str) -> bool:
        """ファイルを読み込みます"""
        try:
            result = False
            local_log = []
            local_log.append(">" * self.REPEAT_TIMES)
            self.file_path = file_path
            self.reader = PdfReader(self.file_path)
            self.metadata_of_reader = self.reader.metadata
            self.num_of_pages = len(self.reader.pages)
            self.creation_date = self.metadata_of_reader.get("/CreationDate")
        except Exception as e:
            local_log.append(f"ファイルの読み込みに失敗しました。: \n{e}")
        else:
            result = True
            local_log.append(f"ファイルの読み込みに成功しました。: \n{self.file_path}")
        finally:
            time_stamp = self.obj_of_dt2.convert_dt_to_str(datetime.datetime.now())
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            print(local_log)
            self.log.extend(local_log)
            return result

    def print_metadata(self) -> bool:
        """メタデータを出力します"""
        try:
            result = False
            local_log = []
            local_log.append(">" * self.REPEAT_TIMES)
            if not self.metadata_of_reader:
                print("ファイルを読み込んでください。")
                raise Exception
            else:
                for key, _ in self.fields:
                    value = getattr(self.metadata_of_reader, key, None)
                    print(f"{key.capitalize().replace("_", " ")}: {value or None}")
        except Exception as e:
            local_log.append(f"メタデータの出力に失敗しました。: \n{e}")
        else:
            result = True
            local_log.append("メタデータの出力に成功しました。")
        finally:
            time_stamp = self.obj_of_dt2.convert_dt_to_str(datetime.datetime.now())
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            print(local_log)
            self.log.extend(local_log)
            return result

    def write_metadata(self, file_path: str, metadata_of_writer: dict) -> bool:
        """メタデータを書き込みます"""
        try:
            result = False
            local_log = []
            local_log.append(">" * self.REPEAT_TIMES)
            self.file_path = file_path
            if not self.metadata_of_reader:
                print("ファイルを読み込んでください。")
                raise Exception
            else:
                self.writer = PdfWriter()
                for page in self.reader.pages:
                    self.writer.add_page(page)
                self.writer.add_metadata(metadata_of_writer)
                with open(self.file_path, "wb") as f:
                    self.writer.write(f)
        except Exception as e:
            local_log.append(f"メタデータの書き込みに失敗しました。: \n{e}")
        else:
            result = True
            local_log.append(f"メタデータの書き込みに成功しました。: \n{self.file_path}")
        finally:
            time_stamp = self.obj_of_dt2.convert_dt_to_str(datetime.datetime.now())
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            print(local_log)
            self.log.extend(local_log)
            return result

    def merge(self, pdfs: list) -> bool:
        """マージします"""
        try:
            result = False
            local_log = []
            local_log.append(">" * self.REPEAT_TIMES)
            first_file_of_pdf_as_path_type = Path(pdfs[0])
            folder_of_pdf_as_path_type = first_file_of_pdf_as_path_type.parent
            dt = self.obj_of_dt2.convert_for_file_name()
            file_name_of_pdf_as_str_type = f"merged_file_{dt}.pdf"
            file_of_pdf_as_path_type = folder_of_pdf_as_path_type / file_name_of_pdf_as_str_type
            file_path_of_pdf_as_str_type = str(file_of_pdf_as_path_type)
            self.writer = PdfWriter()
            for pdf in pdfs:
                self.writer.append(pdf)
            with open(file_path_of_pdf_as_str_type, "wb") as f:
                self.writer.write(f)
            self.read_file(file_path_of_pdf_as_str_type)
            self.add_creation_date_in_metadata(file_path_of_pdf_as_str_type)
        except Exception as e:
            local_log.append(f"マージが失敗しました。: \n{e}")
        else:
            result = True
            local_log.append(f"マージが成功しました。\nfrom: \n{"\n".join(pdfs)}\nto: \n{file_path_of_pdf_as_str_type}\n")
        finally:
            self.read_file(self.file_path)
            time_stamp = self.obj_of_dt2.convert_dt_to_str(datetime.datetime.now())
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            print(local_log)
            self.log.extend(local_log)
            return result

    def extract_pages(self, file_path: str, begin_page: int, end_page: int) -> bool:
        """ページを抽出します"""
        try:
            result = False
            local_log = []
            local_log.append(">" * self.REPEAT_TIMES)
            self.file_path = file_path
            self.read_file(self.file_path)
            self.writer = PdfWriter()
            b = begin_page - 1
            e = end_page - 1
            for i in range(self.num_of_pages):
                if b <= i and i <= e:
                    self.writer.add_page(self.reader.pages[i])
            file_of_exe_as_path_type = Path(self.file_path)
            folder_of_pdf_as_path_type = file_of_exe_as_path_type.parent
            dt = self.obj_of_dt2.convert_for_file_name()
            file_name_of_pdf_as_str_type = f"edited_file_{dt}.pdf"
            file_of_pdf_as_path_type = folder_of_pdf_as_path_type / file_name_of_pdf_as_str_type
            file_path_of_pdf_as_str_type = str(file_of_pdf_as_path_type)
            with open(file_path_of_pdf_as_str_type, "wb") as f:
                self.writer.write(f)
            self.read_file(file_path_of_pdf_as_str_type)
            self.add_creation_date_in_metadata(file_path_of_pdf_as_str_type)
        except Exception as e:
            local_log.append(f"ページの抽出に失敗しました。: \n{e}")
        else:
            result = True
            local_log.append(f"ページの抽出に成功しました。\nfrom: \n{self.file_path}\nbegin page: {begin_page}\nend page: {end_page}\nto: \n{file_path_of_pdf_as_str_type}\n")
        finally:
            time_stamp = self.obj_of_dt2.convert_dt_to_str(datetime.datetime.now())
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            print(local_log)
            self.log.extend(local_log)
            return result

    def delete_pages(self, file_path: str, begin_page: int, end_page: int) -> bool:
        """ページを削除します"""
        try:
            result = False
            local_log = []
            local_log.append(">" * self.REPEAT_TIMES)
            self.file_path = file_path
            self.read_file(self.file_path)
            p = end_page - begin_page + 1
            if p == self.num_of_pages:
                raise ValueError
            self.writer = PdfWriter()
            b = begin_page - 1
            e = end_page - 1
            for i in range(self.num_of_pages):
                if b <= i and i <= e:
                    continue
                self.writer.add_page(self.reader.pages[i])
            file_of_exe_as_path_type = Path(self.file_path)
            folder_of_pdf_as_path_type = file_of_exe_as_path_type.parent
            dt = self.obj_of_dt2.convert_for_file_name()
            file_name_of_pdf_as_str_type = f"edited_file_{dt}.pdf"
            file_of_pdf_as_path_type = folder_of_pdf_as_path_type / file_name_of_pdf_as_str_type
            file_path_of_pdf_as_str_type = str(file_of_pdf_as_path_type)
            with open(file_path_of_pdf_as_str_type, "wb") as f:
                self.writer.write(f)
            self.read_file(file_path_of_pdf_as_str_type)
            self.add_creation_date_in_metadata(file_path_of_pdf_as_str_type)
        except ValueError:
            local_log.append("全ページが指定されたため、処理は行われていません。")
        except Exception as e:
            local_log.append(f"ページの削除に失敗しました。: \n{e}")
        else:
            result = True
            local_log.append(f"ページの削除に成功しました。\nfrom: \n{self.file_path}\nbegin page: {begin_page}\nend page: {end_page}\nto: \n{file_path_of_pdf_as_str_type}\n")
        finally:
            time_stamp = self.obj_of_dt2.convert_dt_to_str(datetime.datetime.now())
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            print(local_log)
            self.log.extend(local_log)
            return result

    def extract_text(self, file_path: str, begin_page: int, end_page: int) -> bool:
        """テキストを抽出します"""
        try:
            result = False
            local_log = []
            local_log.append(">" * self.REPEAT_TIMES)
            self.file_path = file_path
            b = begin_page - 1
            e = end_page - 1
            for i in range(self.num_of_pages):
                if b <= i and i <= e:
                    self.lst_of_text_in_pages.append(f"{i + 1}ページ: \n{self.reader.pages[i].extract_text()}")
        except Exception as e:
            local_log.append(f"テキストの抽出に失敗しました。: \n{e}")
        else:
            result = True
            local_log.append(f"テキストの抽出に成功しました。\n{self.file_path}\nbegin page: {begin_page}\nend page: {end_page}\n")
        finally:
            time_stamp = self.obj_of_dt2.convert_dt_to_str(datetime.datetime.now())
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            print(local_log)
            self.log.extend(local_log)
            return result

    def rotate_page_clockwise(self, file_path: str, page: int, degrees: int) -> bool:
        """ページを時計回りで回転します"""
        try:
            result = False
            local_log = []
            local_log.append(">" * self.REPEAT_TIMES)
            self.file_path = file_path
            self.writer = PdfWriter()
            for p in range(self.num_of_pages):
                self.writer.add_page(self.reader.pages[p])
                if p == page - 1:
                    self.writer.pages[p].rotate(degrees)
            with open(self.file_path, "wb") as f:
                self.writer.write(f)
        except Exception as e:
            local_log.append(f"ページの時計回りの回転に失敗しました。: \n{e}")
        else:
            result = True
            local_log.append(f"ページの時計回りの回転に成功しました。\n{self.file_path}\npage: {page}\ndegrees: {degrees}\n")
        finally:
            time_stamp = self.obj_of_dt2.convert_dt_to_str(datetime.datetime.now())
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            print(local_log)
            self.log.extend(local_log)
            return result

    def add_creation_date_in_metadata(self, file_path: str):
        """メタデータの作成日を追加します"""
        self.metadata_of_writer = {}
        for value, key in self.fields:
            if value == "creation_date":
                self.metadata_of_writer[key] = self.obj_of_dt2.convert_for_metadata_in_pdf(self.UTC_OF_JP)
                break
        self.file_path = file_path
        self.write_metadata(self.file_path, self.metadata_of_writer)

    def write_log(self, file_of_log_as_path_type: Path):
        """処理結果をログに書き出す"""
        file_of_log_as_str_type = str(file_of_log_as_path_type)
        try:
            with open(file_of_log_as_str_type, "w", encoding="utf-8", newline="") as f:
                f.write("\n".join(self.log))
        except Exception as e:
            print(f"ログファイルの出力に失敗しました。: \n{e}")
        else:
            print(f"ログファイルの出力に成功しました。: \n{file_of_log_as_str_type}")
