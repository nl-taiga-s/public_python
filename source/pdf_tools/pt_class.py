import datetime
import textwrap
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
        self.reader = None
        self.num_of_pages = None
        self.writer = None
        self.metadata_of_reader = None
        self.metadata_of_writer = None
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
        self.log = []

    def encrypt(self, file_path: str, password: str) -> bool:
        """暗号化します"""
        try:
            b = False
            log_msg = None
            self.reader = PdfReader(file_path)
            self.writer = PdfWriter(clone_from=self.reader)
            self.writer.encrypt(password, algorithm="AES-256")
            with open(file_path, "wb") as f:
                self.writer.write(f)
        except Exception as e:
            log_msg = f"暗号化に失敗しました。: {e}"
        else:
            b = True
            log_msg = f"暗号化に成功しました。: {file_path}"
        finally:
            print(log_msg)
            time_stamp = self.obj_of_dt2.convert_dt_to_str(datetime.datetime.now())
            self.log.append(f"{log_msg},{time_stamp}")
            self.log.append(f"password: {password}")
            return b

    def decrypt(self, file_path: str, password: str) -> bool:
        """復号化します"""
        try:
            b = False
            log_msg = None
            self.reader = PdfReader(file_path)
            self.reader.decrypt(password)
            self.writer = PdfWriter(clone_from=self.reader)
            with open(file_path, "wb") as f:
                self.writer.write(f)
        except Exception as e:
            log_msg = f"復号化に失敗しました。: {e}"
        else:
            b = True
            log_msg = f"復号化に成功しました。: {file_path}"
        finally:
            print(log_msg)
            time_stamp = self.obj_of_dt2.convert_dt_to_str(datetime.datetime.now())
            self.log.append(f"{log_msg},{time_stamp}")
            return b

    def read_metadata(self, file_path: str) -> bool:
        """メタデータを読み込みます"""
        try:
            b = False
            log_msg = None
            self.reader = PdfReader(file_path)
            self.meta_of_reader = self.reader.metadata
        except Exception as e:
            log_msg = f"メタデータの読み込みに失敗しました。: {e}"
        else:
            b = True
            log_msg = f"メタデータの読み込みに成功しました。: {file_path}"
        finally:
            print(log_msg)
            time_stamp = self.obj_of_dt2.convert_dt_to_str(datetime.datetime.now())
            self.log.append(f"{log_msg},{time_stamp}")
            return b

    def print_metadata(self) -> bool:
        """メタデータを出力します"""
        try:
            b = False
            log_msg = None
            for key, _ in self.fields:
                value = getattr(self.meta_of_reader, key, None)
                print(f"{key.capitalize().replace("_", " ")}: {value or None}")
        except Exception as e:
            log_msg = f"メタデータの出力に失敗しました。: {e}"
        else:
            b = True
            log_msg = "メタデータの出力に成功しました。"
        finally:
            print(log_msg)
            time_stamp = self.obj_of_dt2.convert_dt_to_str(datetime.datetime.now())
            self.log.append(f"{log_msg},{time_stamp}")
            return b

    def write_metadata(self, file_path: str) -> bool:
        """メタデータを書き込みます"""
        try:
            b = False
            log_msg = None
            self.writer = PdfWriter()
            for page in self.reader.pages:
                self.writer.add_page(page)
            self.metadata_of_writer = {}
            for key_of_r, key_of_w in self.fields:
                match key_of_r:
                    case "creation_date":
                        pass
                    case "modification_date":
                        time = self.obj_of_dt2.convert_for_metadata_in_pdf(self.UTC_OF_JP)
                        self.metadata_of_writer[f"{key_of_w}"] = time
                    case _:
                        value = input(f"{key_of_r.capitalize().replace("_", " ")}: ")
                        self.metadata_of_writer[f"{key_of_w}"] = value
            self.writer.add_metadata(self.metadata_of_writer)
            with open(file_path, "wb") as f:
                self.writer.write(f)
        except Exception as e:
            log_msg = f"メタデータの書き込みに失敗しました。: {e}"
        else:
            b = True
            log_msg = f"メタデータの書き込みに成功しました。{file_path}"
        finally:
            print(log_msg)
            time_stamp = self.obj_of_dt2.convert_dt_to_str(datetime.datetime.now())
            self.log.append(f"{log_msg},{time_stamp}")
            return b

    def merge(self, pdfs: list) -> bool:
        """マージします"""
        try:
            b = False
            log_msg = None
            first_file_of_pdf_as_path_type = Path(pdfs[0])
            folder_of_pdf_as_path_type = self.obj_of_pt.get_dir_path(first_file_of_pdf_as_path_type)
            dt = self.obj_of_dt2.convert_for_file_name()
            file_name_of_pdf_as_str_type = f"merged_file_{dt}.pdf"
            file_of_pdf_as_path_type = folder_of_pdf_as_path_type / file_name_of_pdf_as_str_type
            file_path_of_pdf_as_str_type = str(file_of_pdf_as_path_type)
            self.writer = PdfWriter()
            for pdf in pdfs:
                self.writer.append(pdf)
            with open(file_path_of_pdf_as_str_type, "wb") as f:
                self.writer.write(f)
        except Exception as e:
            log_msg = f"マージが失敗しました。: {e}"
        else:
            b = True
            log_msg = textwrap.dedent(f"""\
                マージが成功しました。
                from:
                {"\n".join(pdfs)}
                to:
                {file_path_of_pdf_as_str_type}
                """)
        finally:
            print(log_msg)
            time_stamp = self.obj_of_dt2.convert_dt_to_str(datetime.datetime.now())
            self.log.append(f"{log_msg}{time_stamp}")
            return b

    def extract_pages(self, file_path: str, begin_page: int, end_page: int) -> bool:
        """ページを抽出します"""
        try:
            b = False
            log_msg = None
            self.reader = PdfReader(file_path)
            self.writer = PdfWriter()
            for i in range(begin_page - 1, end_page):
                self.writer.add_page(self.reader.pages[i])
            file_of_exe_as_path_type = Path(file_path)
            folder_of_pdf_as_path_type = self.obj_of_pt.get_dir_path(file_of_exe_as_path_type)
            dt = self.obj_of_dt2.convert_for_file_name()
            file_name_of_pdf_as_str_type = f"extracted_file_{dt}.pdf"
            file_of_pdf_as_path_type = folder_of_pdf_as_path_type / file_name_of_pdf_as_str_type
            file_path_of_pdf_as_str_type = str(file_of_pdf_as_path_type)
            with open(file_path_of_pdf_as_str_type, "wb") as f:
                self.writer.write(f)
        except Exception as e:
            log_msg = f"ページの抽出に失敗しました。: {e}"
        else:
            b = True
            log_msg = textwrap.dedent(f"""\
                ページの抽出に成功しました。
                from:
                {file_path}
                begin page: {begin_page}
                end page: {end_page}
                to:
                {file_path_of_pdf_as_str_type}
                """)
        finally:
            print(log_msg)
            time_stamp = self.obj_of_dt2.convert_dt_to_str(datetime.datetime.now())
            self.log.append(f"{log_msg}{time_stamp}")
            return b

    def extract_text(self, file_path: str, begin_page: int, end_page: int) -> bool:
        """テキストを抽出します"""
        try:
            b = False
            log_msg = None
            self.reader = PdfReader(file_path)
            for i in range(begin_page - 1, end_page):
                print(f"{i + 1}ページ: ")
                print(self.reader.pages[i].extract_text())
        except Exception as e:
            log_msg = f"テキストの抽出に失敗しました。: {e}"
        else:
            b = True
            log_msg = textwrap.dedent(f"""\
                テキストの抽出に成功しました。
                {file_path}
                begin page: {begin_page}
                tend page: {end_page}
                """)
        finally:
            print(log_msg)
            time_stamp = self.obj_of_dt2.convert_dt_to_str(datetime.datetime.now())
            self.log.append(f"{log_msg}{time_stamp}")
            return b

    def rotate_page_clockwise(self, file_path: str, page: int, degrees: int) -> bool:
        """ページを時計回りで回転します"""
        try:
            b = False
            log_msg = None
            self.reader = PdfReader(file_path)
            self.writer = PdfWriter()
            for p in range(self.num_of_pages):
                self.writer.add_page(self.reader.pages[p])
                if p == page - 1:
                    self.writer.pages[p].rotate(degrees)
            with open(file_path, "wb") as f:
                self.writer.write(f)
        except Exception as e:
            log_msg = f"ページの時計回りの回転に失敗しました。: {e}"
        else:
            b = True
            log_msg = textwrap.dedent(f"""\
                ページの時計回りの回転に成功しました。
                {file_path}
                page: {page}
                degrees: {degrees}
                """)
        finally:
            print(log_msg)
            time_stamp = self.obj_of_dt2.convert_dt_to_str(datetime.datetime.now())
            self.log.append(f"{log_msg}{time_stamp}")
            return b

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
