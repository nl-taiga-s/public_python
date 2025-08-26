from pathlib import Path

from pypdf import PdfReader, PdfWriter

from source.common.common import DatetimeTools, PathTools


class PdfTools:
    """PDFを扱うツール"""

    def __init__(self):
        """初期化します"""
        self.log = []
        self.log.append(self.__class__.__doc__)
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
        self.REPEAT_TIMES = 50

    def read_file(self, file_path: str) -> list:
        """ファイルを読み込みます"""
        try:
            result = False
            local_log = []
            local_log.append(">" * self.REPEAT_TIMES)
            self.file_path = file_path
            self.reader = PdfReader(self.file_path)
            if self.reader.is_encrypted:
                local_log.append("ファイルが暗号化されています。: ")
                local_log.append(self.file_path)
                raise Exception
            self.metadata_of_reader = self.reader.metadata
            self.num_of_pages = len(self.reader.pages)
            self.creation_date = self.metadata_of_reader.get("/CreationDate")
        except Exception as e:
            local_log.append("***ファイルの読み込みに失敗しました。***")
            local_log.append(str(e))
        else:
            result = True
            local_log.append("***ファイルの読み込みに成功しました。***")
            local_log.append(self.file_path)
        finally:
            time_stamp = self.obj_of_dt2.convert_dt_to_str()
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            self.log.extend(local_log)
            return [result, local_log]

    def encrypt(self, file_path: str, password: str) -> list:
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
            local_log.append("***暗号化に失敗しました。***")
            local_log.append(str(e))
        else:
            result = True
            local_log.append("***暗号化に成功しました。***")
            local_log.append(self.file_path)
            local_log.append(f"password: {password}")
        finally:
            time_stamp = self.obj_of_dt2.convert_dt_to_str()
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            self.log.extend(local_log)
            return [result, local_log]

    def decrypt(self, file_path: str, password: str) -> list:
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
            local_log.append("***復号化に失敗しました。***")
            local_log.append(str(e))
        else:
            result = True
            local_log.append("***復号化に成功しました。***")
            local_log.append(self.file_path)
        finally:
            time_stamp = self.obj_of_dt2.convert_dt_to_str()
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            self.log.extend(local_log)
            return [result, local_log]

    def get_metadata(self, file_path: str) -> list:
        """メタデータを取得します"""
        try:
            result = False
            local_log = []
            local_log.append(">" * self.REPEAT_TIMES)
            for key, _ in self.fields:
                value = getattr(self.metadata_of_reader, key, None)
                local_log.append(f"{key.capitalize().replace("_", " ")}: {value or None}")
        except Exception as e:
            local_log.append("***メタデータの取得に失敗しました。***")
            local_log.append(str(e))
        else:
            result = True
            local_log.append("***メタデータの取得に成功しました。***")
            local_log.append(self.file_path)
        finally:
            time_stamp = self.obj_of_dt2.convert_dt_to_str()
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            self.log.extend(local_log)
            return [result, local_log]

    def write_metadata(self, file_path: str, metadata_of_writer: dict) -> list:
        """メタデータを書き込みます"""
        try:
            result = False
            local_log = []
            local_log.append(">" * self.REPEAT_TIMES)
            self.writer = PdfWriter()
            for page in self.reader.pages:
                self.writer.add_page(page)
            self.writer.add_metadata(metadata_of_writer)
            with open(self.file_path, "wb") as f:
                self.writer.write(f)
        except Exception as e:
            local_log.append("***メタデータの書き込みに失敗しました。***")
            local_log.append(str(e))
        else:
            result = True
            local_log.append("***メタデータの書き込みに成功しました。***")
            local_log.append(self.file_path)
        finally:
            time_stamp = self.obj_of_dt2.convert_dt_to_str()
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            self.log.extend(local_log)
            return [result, local_log]

    def merge(self, pdfs: list) -> list:
        """マージします"""
        try:
            result = False
            local_log = []
            is_encrypted_list = []
            local_log.append(">" * self.REPEAT_TIMES)
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
                raise Exception
            with open(file_path_of_pdf_as_str_type, "wb") as f:
                self.writer.write(f)
            self.read_file(file_path_of_pdf_as_str_type)
            self.add_creation_date_in_metadata(file_path_of_pdf_as_str_type)
        except Exception as e:
            local_log.append("***マージが失敗しました。***")
            local_log.append(str(e))
            if is_encrypted_list:
                local_log.append("暗号化されたファイルの一覧です。: ")
                local_log.extend(is_encrypted_list)
        else:
            result = True
            local_log.append("***マージが成功しました。***")
            local_log.append("from: ")
            local_log.append("\n".join(pdfs))
            local_log.append("to: ")
            local_log.append(file_path_of_pdf_as_str_type)
        finally:
            self.read_file(self.file_path)
            time_stamp = self.obj_of_dt2.convert_dt_to_str()
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            self.log.extend(local_log)
            return [result, local_log]

    def extract_pages(self, file_path: str, begin_page: int, end_page: int) -> list:
        """ページを抽出します"""
        try:
            result = False
            local_log = []
            local_log.append(">" * self.REPEAT_TIMES)
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
            local_log.append("***ページの抽出に失敗しました。***")
            local_log.append(str(e))
        else:
            result = True
            local_log.append("***ページの抽出に成功しました。***")
            local_log.append("from: ")
            local_log.append(self.file_path)
            local_log.append(f"begin page: {begin_page}")
            local_log.append(f"end page: {end_page}")
            local_log.append("to: ")
            local_log.append(file_path_of_pdf_as_str_type)
        finally:
            time_stamp = self.obj_of_dt2.convert_dt_to_str()
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            self.log.extend(local_log)
            return [result, local_log]

    def delete_pages(self, file_path: str, begin_page: int, end_page: int) -> list:
        """ページを削除します"""
        try:
            result = False
            local_log = []
            local_log.append(">" * self.REPEAT_TIMES)
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
            local_log.append("***全ページが指定されたため、処理は行われていません。***")
        except Exception as e:
            local_log.append("***ページの削除に失敗しました。***")
            local_log.append(str(e))
        else:
            result = True
            local_log.append("***ページの削除に成功しました。***")
            local_log.append("from: ")
            local_log.append(self.file_path)
            local_log.append(f"begin page: {begin_page}")
            local_log.append(f"end page: {end_page}")
            local_log.append("to: ")
            local_log.append(file_path_of_pdf_as_str_type)
        finally:
            time_stamp = self.obj_of_dt2.convert_dt_to_str()
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            self.log.extend(local_log)
            return [result, local_log]

    def extract_text(self, file_path: str, begin_page: int, end_page: int) -> list:
        """テキストを抽出します"""
        try:
            result = False
            local_log = []
            lst_of_text_in_pages = []
            local_log.append(">" * self.REPEAT_TIMES)
            b = begin_page - 1
            e = end_page - 1
            for i in range(self.num_of_pages):
                if b <= i and i <= e:
                    lst_of_text_in_pages.append(f"{i + 1}ページ: \n{self.reader.pages[i].extract_text()}")
            local_log.extend(lst_of_text_in_pages)
        except Exception as e:
            local_log.append("***テキストの抽出に失敗しました。***")
            local_log.append(str(e))
        else:
            result = True
            local_log.append("***テキストの抽出に成功しました。***")
            local_log.append(self.file_path)
            local_log.append(f"begin page: {begin_page}")
            local_log.append(f"end page: {end_page}")
        finally:
            time_stamp = self.obj_of_dt2.convert_dt_to_str()
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            self.log.extend(local_log)
            return [result, local_log]

    def rotate_page_clockwise(self, file_path: str, page: int, degrees: int) -> list:
        """ページを時計回りで回転します"""
        try:
            result = False
            local_log = []
            local_log.append(">" * self.REPEAT_TIMES)
            self.writer = PdfWriter()
            for p in range(self.num_of_pages):
                self.writer.add_page(self.reader.pages[p])
                if p == page - 1:
                    self.writer.pages[p].rotate(degrees)
            with open(self.file_path, "wb") as f:
                self.writer.write(f)
        except Exception as e:
            local_log.append("***ページの時計回りの回転に失敗しました。***")
            local_log.append(str(e))
        else:
            result = True
            local_log.append("***ページの時計回りの回転に成功しました。***")
            local_log.append(self.file_path)
            local_log.append(f"page: {page}")
            local_log.append(f"degrees: {degrees}")
        finally:
            time_stamp = self.obj_of_dt2.convert_dt_to_str()
            local_log.append(time_stamp)
            local_log.append("<" * self.REPEAT_TIMES)
            self.log.extend(local_log)
            return [result, local_log]

    def add_creation_date_in_metadata(self, file_path: str):
        """メタデータの作成日を追加します"""
        self.metadata_of_writer = {}
        for value, key in self.fields:
            if value == "creation_date":
                self.metadata_of_writer[key] = self.obj_of_dt2.convert_for_metadata_in_pdf(self.UTC_OF_JP)
                break
        self.file_path = file_path
        self.write_metadata(self.file_path, self.metadata_of_writer)

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
