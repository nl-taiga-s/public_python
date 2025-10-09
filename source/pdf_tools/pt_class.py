from logging import Logger
from pathlib import Path
from typing import Any

from pypdf import DocumentInformation, PdfReader, PdfWriter

from source.common.common import DatetimeTools


class PdfTools:
    """PDFを扱うツール"""

    def __init__(self, logger: Logger):
        """初期化します"""
        self.log: Logger = logger
        self.log.info(self.__class__.__doc__)
        self.obj_of_dt2: DatetimeTools = DatetimeTools()
        # ファイルパス
        self.file_path: str = ""
        self.reader: PdfReader = None
        # ページ数
        self.num_of_pages: int = 0
        self.writer: PdfWriter = None
        self.metadata_of_reader: DocumentInformation = None
        self.metadata_of_writer: dict = {}
        # 暗号化されているかどうか
        self.encrypted: bool = False
        # pdfファイルの拡張子
        self.EXTENSION: str = ".pdf"
        # 日本時間
        self.UTC_OF_JP: str = "+09'00'"
        # メタデータの書き込み用のフィールド
        self.fields: list = [
            ("title", "/Title"),  # タイトル
            ("author", "/Author"),  # 作成者
            ("subject", "/Subject"),  # サブタイトル
            ("creator", "/Creator"),  # アプリケーション
            ("producer", "/Producer"),  # PDF変換
            ("keywords", "/Keywords"),  # キーワード
            ("creation_date", "/CreationDate"),  # 作成日
            ("modification_date", "/ModDate"),  # 更新日
        ]
        # メタデータの書き込み用の作成日
        self.creation_date: str = ""
        # メタデータの書き込み用の更新日
        self.modification_date: str = ""

    def read_file(self, file_path: str = "") -> bool:
        """ファイルを読み込みます"""
        result: bool = False
        try:
            if file_path != "":
                self.file_path = file_path
            self.log.info(f"対象のファイルパス: {self.file_path}")
            self.reader = PdfReader(self.file_path)
            if self.encrypted:
                raise Exception("ファイルが暗号化されています。")
            self.metadata_of_reader = self.reader.metadata
            self.num_of_pages = len(self.reader.pages)
            self.creation_date = self.metadata_of_reader.get("/CreationDate")
        except Exception as e:
            self.log.debug(repr(e))
            self.log.error(f"***{self.read_file.__doc__} => 失敗しました。***: \n{str(e)}")
            raise
        else:
            result = True
            self.log.info(f"***{self.read_file.__doc__} => 成功しました。***")
        finally:
            pass
        return result

    def encrypt(self, password: str) -> bool:
        """暗号化します"""
        result: bool = False
        try:
            self.log.info(f"対象のファイルパス: {self.file_path}")
            self.log.info(f"password: {password}")
            self.reader = PdfReader(self.file_path)
            self.writer = PdfWriter(clone_from=self.reader)
            self.writer.encrypt(password, algorithm="AES-256")
            with open(self.file_path, "wb") as f:
                self.writer.write(f)
        except Exception as e:
            self.log.debug(repr(e))
            self.log.error(f"***{self.encrypt.__doc__} => 失敗しました。***: \n{str(e)}")
            raise
        else:
            result = True
            self.encrypted = True
            self.log.info(f"***{self.encrypt.__doc__} => 成功しました。***")
        finally:
            pass
        return result

    def decrypt(self, password: str) -> bool:
        """復号化します"""
        result: bool = False
        try:
            self.log.info(f"対象のファイルパス: {self.file_path}")
            self.log.info(f"password: {password}")
            self.reader = PdfReader(self.file_path)
            self.reader.decrypt(password)
            self.writer = PdfWriter(clone_from=self.reader)
            with open(self.file_path, "wb") as f:
                self.writer.write(f)
        except Exception as e:
            self.log.debug(repr(e))
            self.log.error(f"***{self.decrypt.__doc__} => 失敗しました。***: \n{str(e)}")
            raise
        else:
            result = True
            self.encrypted = False
            self.log.info(f"***{self.decrypt.__doc__} => 成功しました。***")
        finally:
            pass
        return result

    def get_metadata(self) -> bool:
        """メタデータを取得します"""
        result: bool = False
        try:
            self.log.info(f"対象のファイルパス: {self.file_path}")
            for key, _ in self.fields:
                value: Any = getattr(self.metadata_of_reader, key, None)
                self.log.info(f"{key.capitalize().replace("_", " ")}: {value or None}")
        except Exception as e:
            self.log.debug(repr(e))
            self.log.error(f"***{self.get_metadata.__doc__} => 失敗しました。***: \n{str(e)}")
            raise
        else:
            result = True
            self.log.info(f"***{self.get_metadata.__doc__} => 成功しました。***")
        finally:
            pass
        return result

    def write_metadata(self, metadata_of_writer: dict) -> bool:
        """メタデータを書き込みます"""
        result: bool = False
        try:
            self.log.info(f"対象のファイルパス: {self.file_path}")
            self.writer = PdfWriter()
            for page in self.reader.pages:
                self.writer.add_page(page)
            self.writer.add_metadata(metadata_of_writer)
            with open(self.file_path, "wb") as f:
                self.writer.write(f)
        except Exception as e:
            self.log.debug(repr(e))
            self.log.error(f"***{self.write_metadata.__doc__} => 失敗しました。***: \n{str(e)}")
            raise
        else:
            result = True
            self.log.info(f"***{self.write_metadata.__doc__} => 成功しました。***")
        finally:
            pass
        return result

    def merge(self, pdfs: list) -> bool:
        """マージします"""
        result: bool = False
        try:
            # ファイルパスを退避させる
            file_path_of_tmp: str = self.file_path
            # 暗号化されたファイルのリスト
            is_encrypted_lst: list = []
            first_file_p: Path = Path(pdfs[0])
            folder_p: Path = first_file_p.parent
            dt: str = self.obj_of_dt2.convert_for_file_name()
            file_name_s: str = f"merged_file_{dt}.pdf"
            file_p: Path = folder_p / file_name_s
            file_s: str = str(file_p)
            self.writer = PdfWriter()
            for pdf in pdfs:
                r = PdfReader(pdf)
                if r.is_encrypted:
                    is_encrypted_lst.append(pdf)
                else:
                    self.writer.append(pdf)
            if is_encrypted_lst:
                raise Exception("暗号化されたファイルがあります。")
            with open(file_s, "wb") as f:
                self.writer.write(f)
            # マージされたファイルを読み込む
            self.read_file(file_s)
            # マージされたファイルにメタデータの作成日を付与する
            self.add_creation_date_in_metadata()
        except Exception as e:
            if is_encrypted_lst:
                self.log.error("暗号化されたファイルの一覧です。: ")
                self.log.error("\n".join(is_encrypted_lst))
            self.log.debug(repr(e))
            self.log.error(f"***{self.merge.__doc__} => 失敗しました。***: \n{str(e)}")
            raise
        else:
            result = True
            self.log.info("from: ")
            self.log.info("\n".join(pdfs))
            self.log.info("to: ")
            self.log.info(file_s)
            self.log.info(f"***{self.merge.__doc__} => 成功しました。***")
        finally:
            # 退避させたファイルパスを読み込む
            self.read_file(file_path_of_tmp)
        return result

    def extract_pages(self, begin_page: int, end_page: int) -> bool:
        """ページを抽出します"""
        result: bool = False
        try:
            # ファイルパスを退避させる
            file_path_of_tmp: str = self.file_path
            b: int = begin_page - 1
            e: int = end_page - 1
            p: int = end_page - begin_page + 1
            if p == self.num_of_pages:
                raise Exception("全ページが指定されたため、処理は行われていません。")
            self.writer = PdfWriter()
            for i in range(self.num_of_pages):
                if b <= i and i <= e:
                    self.writer.add_page(self.reader.pages[i])
            file_of_exe_p: Path = Path(self.file_path)
            folder_p: Path = file_of_exe_p.parent
            dt: str = self.obj_of_dt2.convert_for_file_name()
            file_name_s: str = f"edited_file_{dt}.pdf"
            file_p: Path = folder_p / file_name_s
            file_s: str = str(file_p)
            with open(file_s, "wb") as f:
                self.writer.write(f)
            # マージされたファイルを読み込む
            self.read_file(file_s)
            # ページを抽出したファイルにメタデータの作成日を付与する
            self.add_creation_date_in_metadata()
        except Exception as e:
            self.log.debug(repr(e))
            self.log.error(f"***{self.extract_pages.__doc__} => 失敗しました。***: \n{str(e)}")
            raise
        else:
            result = True
            self.log.info("from: ")
            self.log.info(self.file_path)
            self.log.info(f"begin page: {begin_page}")
            self.log.info(f"end page: {end_page}")
            self.log.info("to: ")
            self.log.info(file_s)
            self.log.info(f"***{self.extract_pages.__doc__} => 成功しました。***")
        finally:
            # 退避させたファイルパスを読み込む
            self.read_file(file_path_of_tmp)
        return result

    def delete_pages(self, begin_page: int, end_page: int) -> bool:
        """ページを削除します"""
        result: bool = False
        try:
            # ファイルパスを退避させる
            file_path_of_tmp: str = self.file_path
            b: int = begin_page - 1
            e: int = end_page - 1
            p: int = end_page - begin_page + 1
            if p == self.num_of_pages:
                raise Exception("全ページが指定されたため、処理は行われていません。")
            self.writer = PdfWriter()
            for i in range(self.num_of_pages):
                if b <= i and i <= e:
                    continue
                self.writer.add_page(self.reader.pages[i])
            file_of_exe_p: Path = Path(self.file_path)
            folder_p: Path = file_of_exe_p.parent
            dt: str = self.obj_of_dt2.convert_for_file_name()
            file_name_s: str = f"edited_file_{dt}.pdf"
            file_p: Path = folder_p / file_name_s
            file_s: str = str(file_p)
            with open(file_s, "wb") as f:
                self.writer.write(f)
            # マージされたファイルを読み込む
            self.read_file(file_s)
            # ページを削除したファイルにメタデータの作成日を付与する
            self.add_creation_date_in_metadata()
        except Exception as e:
            self.log.debug(repr(e))
            self.log.error(f"***{self.delete_pages.__doc__} => 失敗しました。***: \n{str(e)}")
            raise
        else:
            result = True
            self.log.info("from: ")
            self.log.info(self.file_path)
            self.log.info(f"begin page: {begin_page}")
            self.log.info(f"end page: {end_page}")
            self.log.info("to: ")
            self.log.info(file_s)
            self.log.info(f"***{self.delete_pages.__doc__} => 成功しました。***")
        finally:
            # 退避させたファイルパスを読み込む
            self.read_file(file_path_of_tmp)
        return result

    def extract_text(self, begin_page: int, end_page: int) -> bool:
        """テキストを抽出します"""
        result: bool = False
        try:
            self.log.info(f"対象のファイルパス: {self.file_path}")
            lst_of_text_in_pages: list = []
            b: int = begin_page - 1
            e: int = end_page - 1
            for i in range(self.num_of_pages):
                if b <= i and i <= e:
                    lst_of_text_in_pages.append(f"{i + 1}ページ: \n{self.reader.pages[i].extract_text()}")
            self.log.info("\n".join(lst_of_text_in_pages))
        except Exception as e:
            self.log.debug(repr(e))
            self.log.error(f"***{self.extract_text.__doc__} => 失敗しました。***: \n{str(e)}")
            raise
        else:
            result = True
            self.log.info(f"begin page: {begin_page}")
            self.log.info(f"end page: {end_page}")
            self.log.info(f"***{self.extract_text.__doc__} => 成功しました。***")
        finally:
            pass
        return result

    def rotate_page_clockwise(self, page: int, degrees: int) -> bool:
        """ページを時計回りで回転します"""
        result: bool = False
        try:
            self.log.info(f"対象のファイルパス: {self.file_path}")
            self.writer = PdfWriter()
            for p in range(self.num_of_pages):
                self.writer.add_page(self.reader.pages[p])
                if p == page - 1:
                    self.writer.pages[p].rotate(degrees)
            with open(self.file_path, "wb") as f:
                self.writer.write(f)
        except Exception as e:
            self.log.debug(repr(e))
            self.log.error(f"***{self.rotate_page_clockwise.__doc__} => 失敗しました。***: \n{str(e)}")
            raise
        else:
            result = True
            self.log.info(f"page: {page}")
            self.log.info(f"degrees: {degrees}")
            self.log.info(f"***{self.rotate_page_clockwise.__doc__} => 成功しました。***")
        finally:
            pass
        return result

    def add_creation_date_in_metadata(self) -> bool:
        """メタデータの作成日を追加します"""
        result: bool = False
        try:
            self.log.info(f"対象のファイルパス: {self.file_path}")
            self.metadata_of_writer = {}
            for value, key in self.fields:
                if value == "creation_date":
                    self.metadata_of_writer[key] = self.obj_of_dt2.convert_for_metadata_in_pdf(self.UTC_OF_JP)
                    break
            self.write_metadata(self.metadata_of_writer)
        except Exception as e:
            self.log.debug(repr(e))
            self.log.error(f"***{self.add_creation_date_in_metadata.__doc__} => 失敗しました。***: \n{str(e)}")
            raise
        else:
            result = True
            self.log.info(f"***{self.add_creation_date_in_metadata.__doc__} => 成功しました。***")
        finally:
            pass
        return result
