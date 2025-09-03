from logging import Logger
from pathlib import Path

from pypdf import PdfReader, PdfWriter

from source.common.common import DatetimeTools


class PdfTools:
    """PDFを扱うツール"""

    def __init__(self, log: Logger):
        """初期化します"""
        self.log = log
        self.log.info(self.__class__.__doc__)
        self.obj_of_dt2 = DatetimeTools()
        self.file_path = ""
        self.reader = None
        self.num_of_pages = ""
        self.writer = None
        self.metadata_of_reader = None
        self.metadata_of_writer = {}
        # 暗号化されているかどうか
        self.encrypted = False
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
        self.creation_date = ""
        self.modification_date = ""
        self.lst_of_text_in_pages = []

    def read_file(self, file_path: str = "") -> bool:
        """ファイルを読み込みます"""
        try:
            result = False
            error = False
            if file_path != "":
                self.file_path = file_path
            self.reader = PdfReader(self.file_path)
            if self.encrypted:
                raise Exception("ファイルが暗号化されています。")
            self.metadata_of_reader = self.reader.metadata
            self.num_of_pages = len(self.reader.pages)
            self.creation_date = self.metadata_of_reader.get("/CreationDate")
        except Exception as e:
            error = True
            self.log.error(f"***{self.read_file.__doc__} => 失敗しました。***")
            self.log.error(str(e))
        else:
            result = True
            self.log.info(f"***{self.read_file.__doc__} => 成功しました。***")
        finally:
            self.log.info(self.file_path)
            if error:
                raise
            return result

    def encrypt(self, password: str) -> bool:
        """暗号化します"""
        try:
            result = False
            error = False
            self.reader = PdfReader(self.file_path)
            self.writer = PdfWriter(clone_from=self.reader)
            self.writer.encrypt(password, algorithm="AES-256")
            with open(self.file_path, "wb") as f:
                self.writer.write(f)
        except Exception as e:
            error = True
            self.log.error(f"***{self.encrypt.__doc__} => 失敗しました。***")
            self.log.error(str(e))
        else:
            result = True
            self.encrypted = True
            self.log.info(f"***{self.encrypt.__doc__} => 成功しました。***")
            self.log.info(f"password: {password}")
        finally:
            self.log.info(self.file_path)
            if error:
                raise
            return result

    def decrypt(self, password: str) -> bool:
        """復号化します"""
        try:
            result = False
            error = False
            self.reader = PdfReader(self.file_path)
            self.reader.decrypt(password)
            self.writer = PdfWriter(clone_from=self.reader)
            with open(self.file_path, "wb") as f:
                self.writer.write(f)
        except Exception as e:
            error = True
            self.log.error(f"***{self.decrypt.__doc__} => 失敗しました。***")
            self.log.error(str(e))
        else:
            result = True
            self.encrypted = False
            self.log.info(f"***{self.decrypt.__doc__} => 成功しました。***")
            self.log.info(f"password: {password}")
        finally:
            self.log.info(self.file_path)
            if error:
                raise
            return result

    def get_metadata(self) -> bool:
        """メタデータを取得します"""
        try:
            result = False
            error = False
            for key, _ in self.fields:
                value = getattr(self.metadata_of_reader, key, None)
                self.log.info(f"{key.capitalize().replace("_", " ")}: {value or None}")
        except Exception as e:
            error = True
            self.log.error(f"***{self.get_metadata.__doc__} => 失敗しました。***")
            self.log.error(str(e))
        else:
            result = True
            self.log.info(f"***{self.get_metadata.__doc__} => 成功しました。***")
        finally:
            self.log.info(self.file_path)
            if error:
                raise
            return result

    def write_metadata(self, metadata_of_writer: dict) -> bool:
        """メタデータを書き込みます"""
        try:
            result = False
            error = False
            self.writer = PdfWriter()
            for page in self.reader.pages:
                self.writer.add_page(page)
            self.writer.add_metadata(metadata_of_writer)
            with open(self.file_path, "wb") as f:
                self.writer.write(f)
        except Exception as e:
            error = True
            self.log.error(f"***{self.write_metadata.__doc__} => 失敗しました。***")
            self.log.error(str(e))
        else:
            result = True
            self.log.info(f"***{self.write_metadata.__doc__} => 成功しました。***")
        finally:
            self.log.info(self.file_path)
            if error:
                raise
            return result

    def merge(self, pdfs: list) -> bool:
        """マージします"""
        try:
            result = False
            error = False
            # ファイルパスを退避させる
            file_path_of_tmp = self.file_path
            # 暗号化されたファイルのリスト
            is_encrypted_list = []
            first_file_of_pdf_as_path_type = Path(pdfs[0])
            folder_of_pdf_as_path_type = first_file_of_pdf_as_path_type.parent
            dt = self.obj_of_dt2.convert_for_file_name()
            file_name_of_pdf_as_str_type = f"merged_file_{dt}.pdf"
            file_of_pdf_as_path_type = folder_of_pdf_as_path_type / file_name_of_pdf_as_str_type
            file_path_of_pdf_as_str_type = str(file_of_pdf_as_path_type)
            self.writer = PdfWriter()
            for pdf in pdfs:
                r = PdfReader(pdf)
                if r.is_encrypted:
                    is_encrypted_list.append(pdf)
                else:
                    self.writer.append(pdf)
            if is_encrypted_list:
                raise Exception("暗号化されたファイルがあります。")
            with open(file_path_of_pdf_as_str_type, "wb") as f:
                self.writer.write(f)
            # マージされたファイルを読み込む
            if not self.read_file(file_path_of_pdf_as_str_type):
                raise Exception
            # マージされたファイルにメタデータの作成日を付与する
            if not self.add_creation_date_in_metadata():
                raise Exception
        except Exception as e:
            error = True
            self.log.error(f"***{self.merge.__doc__} => 失敗しました。***")
            self.log.error(str(e))
            if is_encrypted_list:
                self.log.error("暗号化されたファイルの一覧です。: ")
                self.log.error("\n".join(is_encrypted_list))
        else:
            result = True
            self.log.info(f"***{self.merge.__doc__} => 成功しました。***")
            self.log.info("from: ")
            self.log.info("\n".join(pdfs))
            self.log.info("to: ")
        finally:
            self.log.info(file_path_of_pdf_as_str_type)
            # 退避させたファイルパスを読み込む
            self.read_file(file_path_of_tmp)
            if error:
                raise
            return result

    def extract_pages(self, begin_page: int, end_page: int) -> bool:
        """ページを抽出します"""
        try:
            result = False
            error = False
            # ファイルパスを退避させる
            file_path_of_tmp = self.file_path
            b = begin_page - 1
            e = end_page - 1
            p = end_page - begin_page + 1
            if p == self.num_of_pages:
                raise ValueError("全ページが指定されたため、処理は行われていません。")
            self.writer = PdfWriter()
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
            # マージされたファイルを読み込む
            if not self.read_file(file_path_of_pdf_as_str_type):
                raise Exception
            # ページを抽出したファイルにメタデータの作成日を付与する
            if not self.add_creation_date_in_metadata():
                raise Exception
        except ValueError as e:
            error = True
            self.log.error(f"***{str(e)}***")
        except Exception as e:
            error = True
            self.log.error(f"***{self.extract_pages.__doc__} => 失敗しました。***")
            self.log.error(str(e))
        else:
            result = True
            self.log.info(f"***{self.extract_pages.__doc__} => 成功しました。***")
            self.log.info("from: ")
            self.log.info(self.file_path)
            self.log.info(f"begin page: {begin_page}")
            self.log.info(f"end page: {end_page}")
            self.log.info("to: ")
        finally:
            self.log.info(file_path_of_pdf_as_str_type)
            # 退避させたファイルパスを読み込む
            self.read_file(file_path_of_tmp)
            if error:
                raise
            return result

    def delete_pages(self, begin_page: int, end_page: int) -> bool:
        """ページを削除します"""
        try:
            result = False
            error = False
            # ファイルパスを退避させる
            file_path_of_tmp = self.file_path
            b = begin_page - 1
            e = end_page - 1
            p = end_page - begin_page + 1
            if p == self.num_of_pages:
                raise ValueError("全ページが指定されたため、処理は行われていません。")
            self.writer = PdfWriter()
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
            # マージされたファイルを読み込む
            if not self.read_file(file_path_of_pdf_as_str_type):
                raise Exception
            # ページを削除したファイルにメタデータの作成日を付与する
            if not self.add_creation_date_in_metadata():
                raise Exception
        except ValueError as e:
            error = True
            self.log.error(f"***{str(e)}***")
        except Exception as e:
            error = True
            self.log.error(f"***{self.delete_pages.__doc__} => 失敗しました。***")
            self.log.error(str(e))
        else:
            result = True
            self.log.info(f"***{self.delete_pages.__doc__} => 成功しました。***")
            self.log.info("from: ")
            self.log.info(self.file_path)
            self.log.info(f"begin page: {begin_page}")
            self.log.info(f"end page: {end_page}")
            self.log.info("to: ")
        finally:
            self.log.info(file_path_of_pdf_as_str_type)
            # 退避させたファイルパスを読み込む
            self.read_file(file_path_of_tmp)
            if error:
                raise
            return result

    def extract_text(self, begin_page: int, end_page: int) -> bool:
        """テキストを抽出します"""
        try:
            result = False
            error = False
            lst_of_text_in_pages = []
            b = begin_page - 1
            e = end_page - 1
            for i in range(self.num_of_pages):
                if b <= i and i <= e:
                    lst_of_text_in_pages.append(f"{i + 1}ページ: \n{self.reader.pages[i].extract_text()}")
            self.log.info("\n".join(lst_of_text_in_pages))
        except Exception as e:
            error = True
            self.log.error(f"***{self.extract_text.__doc__} => 失敗しました。***")
            self.log.error(str(e))
        else:
            result = True
            self.log.info(f"***{self.extract_text.__doc__} => 成功しました。***")
            self.log.info(f"begin page: {begin_page}")
            self.log.info(f"end page: {end_page}")
        finally:
            self.log.info(self.file_path)
            if error:
                raise
            return result

    def rotate_page_clockwise(self, page: int, degrees: int) -> bool:
        """ページを時計回りで回転します"""
        try:
            result = False
            error = False
            self.writer = PdfWriter()
            for p in range(self.num_of_pages):
                self.writer.add_page(self.reader.pages[p])
                if p == page - 1:
                    self.writer.pages[p].rotate(degrees)
            with open(self.file_path, "wb") as f:
                self.writer.write(f)
        except Exception as e:
            error = True
            self.log.error(f"***{self.rotate_page_clockwise.__doc__} => 失敗しました。***")
            self.log.error(str(e))
        else:
            result = True
            self.log.info(f"***{self.rotate_page_clockwise.__doc__} =>成功しました。***")
            self.log.info(f"page: {page}")
            self.log.info(f"degrees: {degrees}")
        finally:
            self.log.info(self.file_path)
            if error:
                raise
            return result

    def add_creation_date_in_metadata(self) -> bool:
        """メタデータの作成日を追加します"""
        try:
            result = False
            error = False
            self.metadata_of_writer = {}
            for value, key in self.fields:
                if value == "creation_date":
                    self.metadata_of_writer[key] = self.obj_of_dt2.convert_for_metadata_in_pdf(self.UTC_OF_JP)
                    break
            self.write_metadata(self.metadata_of_writer)
        except Exception as e:
            error = True
            self.log.error(f"***{self.add_creation_date_in_metadata.__doc__} => 失敗しました。***")
            self.log.error(str(e))
        else:
            result = True
            self.log.info(f"***{self.add_creation_date_in_metadata.__doc__} => 成功しました。***")
        finally:
            self.log.info(self.file_path)
            if error:
                raise
            return result
