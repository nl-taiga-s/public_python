import os
import re
from pathlib import Path

from pt_class import PdfTools
from pypdf import PdfReader

from source.common.common import PathTools


class PT:
    def __init__(self):
        """初期化します"""
        self.d_of_bool = {
            "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
            "no": ["いいえ", "0", "No", "no", "N", "n"],
        }
        self.obj_of_pt = PathTools()
        self.obj_of_cls = PdfTools()
        self.result = False

    def input_target_of_pdf(self) -> str:
        """対象のPDFファイルのパスを入力します。"""
        try:
            TILDE = "~"
            file_path_of_pdf_as_str_type = input("対象のPDFファイルのパスを入力してください。: ")
            file_of_pdf_as_path_type = Path(file_path_of_pdf_as_str_type)
            if TILDE in file_path_of_pdf_as_str_type:
                file_of_pdf_as_path_type = self.obj_of_pt.get_expanded_in_home_dir(file_of_pdf_as_path_type)
                file_path_of_pdf_as_str_type = str(file_of_pdf_as_path_type)
            if not file_of_pdf_as_path_type.exists() or self.obj_of_pt.get_extension(file_of_pdf_as_path_type) != self.obj_of_cls.EXTENSION:
                raise FileNotFoundError(f"ファイルが不正です。: {file_path_of_pdf_as_str_type}")
            self.obj_of_cls.reader = PdfReader(file_path_of_pdf_as_str_type)
            self.obj_of_cls.num_of_pages = len(self.obj_of_cls.reader.pages)
            return file_path_of_pdf_as_str_type
        except FileNotFoundError as e:
            print(e)
        except KeyboardInterrupt:
            os._exit(0)

    def input_bool_of_encrypt(self) -> bool:
        """暗号化するかどうかを入力します"""
        try:
            if self.obj_of_cls.reader.is_encrypted:
                print("既に暗号化されています。")
            else:
                while True:
                    str_of_bool = input("暗号化するかどうかを入力してください。: ")
                    match str_of_bool:
                        case var if var in self.d_of_bool["yes"]:
                            return True
                        case var if var in self.d_of_bool["no"]:
                            return False
                        case _:
                            raise ValueError(f"無効な入力です。: {str_of_bool}")
        except ValueError as e:
            print(e)
        except KeyboardInterrupt:
            os._exit(0)

    def input_password_of_encrypt(self) -> str:
        """暗号化のパスワードを入力します"""
        try:
            while True:
                password = input("暗号化のパスワードを入力してください。: ")
                if re.fullmatch(r"[A-Za-z0-9_-]+", password):
                    return password
                else:
                    print("以下の文字で入力してください。\n* 半角英数字\n* アンダーバー\n* ハイフン")
        except KeyboardInterrupt:
            os._exit(0)

    def input_bool_of_decrypt(self) -> bool:
        """復号化するかどうかを入力します"""
        try:
            if not self.obj_of_cls.reader.is_encrypted:
                print("既に復号化されています。")
            else:
                while True:
                    str_of_bool = input("復号化するかどうかを入力してください。: ")
                    match str_of_bool:
                        case var if var in self.d_of_bool["yes"]:
                            return True
                        case var if var in self.d_of_bool["no"]:
                            return False
                        case _:
                            raise ValueError(f"無効な入力です。: {str_of_bool}")
        except ValueError as e:
            print(e)
        except KeyboardInterrupt:
            os._exit(0)

    def input_password_of_decrypt(self) -> str:
        """復号化のパスワードを入力します"""
        try:
            while True:
                password = input("復号化のパスワードを入力してください。: ")
                if re.fullmatch(r"[A-Za-z0-9_-]+", password):
                    return password
                else:
                    print("以下の文字で入力してください。\n* 半角英数字\n* アンダーバー\n* ハイフン")
        except KeyboardInterrupt:
            os._exit(0)

    def input_bool_of_merge(self) -> bool:
        """マージするかどうかを入力します"""
        try:
            while True:
                str_of_bool = input("マージするかどうかを入力してください。: ")
                match str_of_bool:
                    case var if var in self.d_of_bool["yes"]:
                        return True
                    case var if var in self.d_of_bool["no"]:
                        return False
                    case _:
                        raise ValueError(f"無効な入力です。: {str_of_bool}")
        except ValueError as e:
            print(e)
        except KeyboardInterrupt:
            os._exit(0)

    def input_list_of_merge(self) -> list:
        """マージ元の全てのファイルを入力します"""
        try:
            pdfs = []
            while True:
                file_as_str_type = input("マージ元のファイルのパスを順番に入力してください。: ")
                file_as_path_type = Path(file_as_str_type)
                if file_as_path_type.exists() and self.obj_of_pt.get_extension(file_as_path_type) == self.obj_of_cls.EXTENSION:
                    pdfs.append(file_as_str_type)
                else:
                    print("有効なマージ元のファイルのパスを入力してください。")
                str_of_bool = input("対象のファイルは、まだありますか？: ")
                match str_of_bool:
                    case var if var in self.d_of_bool["yes"]:
                        pass
                    case var if var in self.d_of_bool["no"]:
                        break
                    case _:
                        raise ValueError(f"無効な入力です。: {str_of_bool}")
            return pdfs
        except ValueError as e:
            print(e)
        except KeyboardInterrupt:
            os._exit(0)

    def input_bool_of_extracting_pages(self) -> bool:
        """ページを抽出するかどうかを入力します"""
        try:
            while True:
                str_of_bool = input("ページを抽出するかどうかを入力してください。: ")
                match str_of_bool:
                    case var if var in self.d_of_bool["yes"]:
                        return True
                    case var if var in self.d_of_bool["no"]:
                        return False
                    case _:
                        raise ValueError(f"無効な入力です。: {str_of_bool}")
        except ValueError as e:
            print(e)
        except KeyboardInterrupt:
            os._exit(0)

    def input_extracting_pages(self) -> list:
        """抽出するページを入力します"""
        try:
            while True:
                print(self.__class__.input_extracting_pages.__doc__)
                begin_page = input("始めのページを入力してください。: ")
                end_page = input("終わりのページを入力してください。: ")
                if not begin_page.isdigit() and end_page.isdigit():
                    print("数字を入力してください。")
                    continue
                begin_page = int(begin_page)
                end_page = int(end_page)
                if begin_page < 1 or end_page > self.obj_of_cls.num_of_pages or begin_page > end_page:
                    raise ValueError("指定のページ範囲が不正です。")
                else:
                    break
            return begin_page, end_page
        except ValueError as e:
            print(e)
        except KeyboardInterrupt:
            os._exit(0)

    def input_bool_of_extracting_text(self) -> bool:
        """テキストを抽出するかどうかを入力します"""
        try:
            while True:
                str_of_bool = input("テキストを抽出するかどうかを入力してください。: ")
                match str_of_bool:
                    case var if var in self.d_of_bool["yes"]:
                        return True
                    case var if var in self.d_of_bool["no"]:
                        return False
                    case _:
                        raise ValueError(f"無効な入力です。: {str_of_bool}")
        except ValueError as e:
            print(e)
        except KeyboardInterrupt:
            os._exit(0)

    def input_pages_of_extracting_text(self) -> list:
        """テキストを抽出するページを入力します"""
        try:
            while True:
                print(self.__class__.input_pages_of_extracting_text.__doc__)
                begin_page = input("始めのページを入力してください。: ")
                end_page = input("終わりのページを入力してください。: ")
                if not begin_page.isdigit() and end_page.isdigit():
                    print("数字を入力してください。")
                    continue
                begin_page = int(begin_page)
                end_page = int(end_page)
                if begin_page < 1 or end_page > self.obj_of_cls.num_of_pages or begin_page > end_page:
                    raise ValueError("指定のページ範囲が不正です。")
                else:
                    break
            return begin_page, end_page
        except ValueError as e:
            print(e)
        except KeyboardInterrupt:
            os._exit(0)

    def main(self):
        """主要関数"""
        file_path_of_pdf_as_str_type = self.input_target_of_pdf()
        if self.input_bool_of_encrypt():
            password = self.input_password_of_encrypt()
            self.result = self.obj_of_cls.encrypt(file_path_of_pdf_as_str_type, password)
        elif self.input_bool_of_decrypt():
            password = self.input_password_of_decrypt()
            self.result = self.obj_of_cls.decrypt(file_path_of_pdf_as_str_type, password)
        if not self.result:
            self.obj_of_cls.read_metadata(file_path_of_pdf_as_str_type)
            self.obj_of_cls.print_metadata()
            self.obj_of_cls.write_metadata(file_path_of_pdf_as_str_type)
            self.obj_of_cls.read_metadata(file_path_of_pdf_as_str_type)
            self.obj_of_cls.print_metadata()
            if self.input_bool_of_merge():
                pdfs = self.input_list_of_merge()
                self.obj_of_cls.merge(pdfs)
            if self.input_bool_of_extracting_pages():
                begin_page, end_page = self.input_extracting_pages()
                self.obj_of_cls.extract_pages(file_path_of_pdf_as_str_type, begin_page, end_page)
            if self.input_bool_of_extracting_text():
                begin_page, end_of_page = self.input_pages_of_extracting_text()
                self.obj_of_cls.extract_text(file_path_of_pdf_as_str_type, begin_page, end_page)
        file_of_exe_as_path_type = Path(__file__)
        file_of_log_as_path_type = self.obj_of_pt.get_file_path_of_log(file_of_exe_as_path_type)
        self.obj_of_cls.write_log(file_of_log_as_path_type)


if __name__ == "__main__":
    obj_with_cui = PT()
    obj_with_cui.main()
