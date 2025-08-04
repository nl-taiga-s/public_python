import re
import sys
from enum import Enum
from pathlib import Path

from pypdf import PdfReader

from source.common.common import DatetimeTools, PathTools
from source.pdf_tools.pt_class import PdfTools


class PT_With_Cui:
    def __init__(self):
        """初期化します"""
        self.d_of_bool = {
            "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
            "no": ["いいえ", "0", "No", "no", "N", "n"],
        }
        self.obj_of_pt = PathTools()
        self.obj_of_dt2 = DatetimeTools()
        self.MENU = Enum(
            "MENU",
            [
                "ファイルを暗号化します",
                "ファイルを復号化します",
                "メタデータを読み込みます",
                "メタデータを出力します",
                "メタデータを書き込みます",
                "ファイルをマージします",
                "ページを抽出します",
                "テキストを抽出します",
                "ページを時計回りで回転します",
                "終了します",
            ],
        )
        self.DEGREES = Enum("DEGREES", ["90", "180", "270"])

    def select_menu(self) -> str:
        """メニューを選択します"""
        str_of_menu = [f"({m.value}) {m.name}" for m in self.MENU]
        while True:
            try:
                print(*str_of_menu, sep="\n")
                n = input("メニューの番号を入力してください。: ")
                if n == "":
                    return None
                elif n.isdecimal():
                    num = int(n)
                    if 1 <= num <= len(self.MENU):
                        return self.MENU(num)
                print()
            except Exception as e:
                print(e)
            except KeyboardInterrupt:
                sys.exit(0)

    def input_target_of_pdf(self, extension: str) -> str:
        """対象のPDFファイルのパスを入力します。"""
        TILDE = "~"
        while True:
            try:
                file_path_of_pdf_as_str_type = input("対象のPDFファイルのパスを入力してください。: ")
                file_of_pdf_as_path_type = Path(file_path_of_pdf_as_str_type)
                if TILDE in file_path_of_pdf_as_str_type:
                    file_of_pdf_as_path_type = self.obj_of_pt.get_expanded_in_home_dir(file_of_pdf_as_path_type)
                    file_path_of_pdf_as_str_type = str(file_of_pdf_as_path_type)
                if file_of_pdf_as_path_type.exists() and self.obj_of_pt.get_extension(file_of_pdf_as_path_type) == extension:
                    return file_path_of_pdf_as_str_type
                else:
                    print(f"PDFファイルが不正です。: {file_path_of_pdf_as_str_type}")
            except Exception as e:
                print(e)
            except KeyboardInterrupt:
                sys.exit(0)

    def input_password_of_encrypt(self) -> str:
        """暗号化のパスワードを入力します"""
        while True:
            try:
                password = input("暗号化のパスワードを入力してください。: ")
                if re.fullmatch(r"[A-Za-z0-9_-]+", password):
                    return password
                else:
                    print("以下の文字で入力してください。\n* 半角英数字\n* アンダーバー\n* ハイフン")
            except Exception as e:
                print(e)
            except KeyboardInterrupt:
                sys.exit(0)

    def input_password_of_decrypt(self) -> str:
        """復号化のパスワードを入力します"""
        while True:
            try:
                password = input("復号化のパスワードを入力してください。: ")
                if re.fullmatch(r"[A-Za-z0-9_-]+", password):
                    return password
                else:
                    print("以下の文字で入力してください。\n* 半角英数字\n* アンダーバー\n* ハイフン")
            except Exception as e:
                print(e)
            except KeyboardInterrupt:
                sys.exit(0)

    def input_writing_metadata(self, metadata_of_writer: dict, fields: list, creation_date: str, utc: str) -> dict:
        """書き込み用のメタデータを入力します"""
        try:
            for key_of_r, key_of_w in fields:
                match key_of_r:
                    case "creation_date":
                        metadata_of_writer[key_of_w] = creation_date
                    case "modification_date":
                        time = self.obj_of_dt2.convert_for_metadata_in_pdf(utc)
                        metadata_of_writer[key_of_w] = time
                    case _:
                        value = input(f"{key_of_r.capitalize().replace("_", " ")}: ")
                        metadata_of_writer[key_of_w] = value
            return metadata_of_writer
        except Exception as e:
            print(e)
        except KeyboardInterrupt:
            sys.exit(0)

    def input_list_of_merge(self, extension: str) -> list:
        """マージ元の全てのファイルを入力します"""
        TILDE = "~"
        pdfs = []
        while True:
            try:
                file_as_str_type = input("マージ元のファイルのパスを順番に入力してください。: ")
                file_as_path_type = Path(file_as_str_type)
                if TILDE in file_as_str_type:
                    file_as_path_type = self.obj_of_pt.get_expanded_in_home_dir(file_as_path_type)
                    file_as_str_type = str(file_as_path_type)
                if file_as_path_type.exists() and self.obj_of_pt.get_extension(file_as_path_type) == extension:
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
                        print(f"無効な入力です。: {str_of_bool}")
            except Exception as e:
                print(e)
            except KeyboardInterrupt:
                sys.exit(0)
        return pdfs

    def input_extracting_pages(self, num_of_pages: int) -> list:
        """抽出するページを入力します"""
        while True:
            try:
                begin_page = input("始めのページを入力してください。: ")
                end_page = input("終わりのページを入力してください。: ")
                if not begin_page.isdecimal() and not end_page.isdecimal():
                    print("数字を入力してください。")
                    continue
                begin_page = int(begin_page)
                end_page = int(end_page)
                if begin_page < 1 or end_page > num_of_pages or begin_page > end_page:
                    print("指定のページ範囲が不正です。")
                else:
                    break
            except Exception as e:
                print(e)
            except KeyboardInterrupt:
                sys.exit(0)
        return [begin_page, end_page]

    def input_pages_to_extract_text(self, num_of_pages: int) -> list:
        """テキストを抽出するページを入力します"""
        while True:
            try:
                begin_page = input("始めのページを入力してください。: ")
                end_page = input("終わりのページを入力してください。: ")
                if not begin_page.isdecimal() and not end_page.isdecimal():
                    print("数字を入力してください。")
                    continue
                begin_page = int(begin_page)
                end_page = int(end_page)
                if begin_page < 1 or end_page > num_of_pages or begin_page > end_page:
                    print("指定のページ範囲が不正です。")
                else:
                    break
            except Exception as e:
                print(e)
            except KeyboardInterrupt:
                sys.exit(0)
        return [begin_page, end_page]

    def input_rotating_page(self, num_of_pages: int) -> int:
        """回転するページを入力します"""
        while True:
            try:
                page = input("ページを入力してください。: ")
                if not page.isdecimal():
                    print("数字を入力してください。")
                    continue
                page = int(page)
                if page < 1 or page > num_of_pages:
                    print("指定のページ範囲が不正です。")
                else:
                    break
            except Exception as e:
                print(e)
            except KeyboardInterrupt:
                sys.exit(0)
        return page

    def input_degrees(self) -> int:
        """回転する度数を入力します"""
        str_of_degrees = [f"({d.value}) {d.name}" for d in self.DEGREES]
        while True:
            try:
                print(*str_of_degrees, sep="\n")
                n = input("度数の番号を入力してください。: ")
                if n.isdecimal():
                    num = int(n)
                    if 1 <= num <= len(self.DEGREES):
                        return int(self.DEGREES(num).name)
                print()
            except Exception as e:
                print(e)
            except KeyboardInterrupt:
                sys.exit(0)


def main() -> bool:
    """主要関数"""
    obj_with_cui = PT_With_Cui()
    obj_of_cls = PdfTools()
    while True:
        try:
            # メニューを選択します
            option = obj_with_cui.select_menu()
            if option is None:
                return False
            match option:
                case var if var == obj_with_cui.MENU.ファイルを暗号化します:
                    # ファイルを暗号化します
                    file_path_of_pdf_as_str_type = obj_with_cui.input_target_of_pdf(obj_of_cls.EXTENSION)
                    obj_of_cls.read_metadata(file_path_of_pdf_as_str_type)
                    password = obj_with_cui.input_password_of_encrypt()
                    obj_of_cls.encrypt(file_path_of_pdf_as_str_type, password)
                case var if var == obj_with_cui.MENU.ファイルを復号化します:
                    # ファイルを復号化します
                    file_path_of_pdf_as_str_type = obj_with_cui.input_target_of_pdf(obj_of_cls.EXTENSION)
                    obj_of_cls.read_metadata(file_path_of_pdf_as_str_type)
                    password = obj_with_cui.input_password_of_decrypt()
                    obj_of_cls.decrypt(file_path_of_pdf_as_str_type, password)
                case var if var == obj_with_cui.MENU.メタデータを読み込みます:
                    # メタデータを読み込みます
                    file_path_of_pdf_as_str_type = obj_with_cui.input_target_of_pdf(obj_of_cls.EXTENSION)
                    obj_of_cls.read_metadata(file_path_of_pdf_as_str_type)
                case var if var == obj_with_cui.MENU.メタデータを出力します:
                    # メタデータを出力します
                    if obj_of_cls.reader is None:
                        print("メタデータを読み込んでください。")
                    else:
                        obj_of_cls.print_metadata()
                case var if var == obj_with_cui.MENU.メタデータを書き込みます:
                    # メタデータを書き込みます
                    file_path_of_pdf_as_str_type = obj_with_cui.input_target_of_pdf(obj_of_cls.EXTENSION)
                    obj_of_cls.read_metadata(file_path_of_pdf_as_str_type)
                    obj_of_cls.metadata_of_writer = obj_with_cui.input_writing_metadata(
                        obj_of_cls.metadata_of_writer, obj_of_cls.fields, obj_of_cls.creation_date, obj_of_cls.UTC_OF_JP
                    )
                    obj_of_cls.write_metadata(file_path_of_pdf_as_str_type, obj_of_cls.metadata_of_writer)
                case var if var == obj_with_cui.MENU.ファイルをマージします:
                    # ファイルをマージします
                    pdfs = obj_with_cui.input_list_of_merge(obj_of_cls.EXTENSION)
                    for pdf in pdfs:
                        obj_of_cls.reader = PdfReader(pdf)
                    obj_of_cls.merge(pdfs)
                case var if var == obj_with_cui.MENU.ページを抽出します:
                    # ページを抽出します
                    file_path_of_pdf_as_str_type = obj_with_cui.input_target_of_pdf(obj_of_cls.EXTENSION)
                    obj_of_cls.read_metadata(file_path_of_pdf_as_str_type)
                    begin_page, end_page = obj_with_cui.input_extracting_pages(obj_of_cls.num_of_pages)
                    obj_of_cls.extract_pages(file_path_of_pdf_as_str_type, begin_page, end_page)
                case var if var == obj_with_cui.MENU.テキストを抽出します:
                    # テキストを抽出します
                    file_path_of_pdf_as_str_type = obj_with_cui.input_target_of_pdf(obj_of_cls.EXTENSION)
                    obj_of_cls.read_metadata(file_path_of_pdf_as_str_type)
                    begin_page, end_page = obj_with_cui.input_pages_to_extract_text(obj_of_cls.num_of_pages)
                    obj_of_cls.extract_text(file_path_of_pdf_as_str_type, begin_page, end_page)
                    for i in range(begin_page - 1, end_page):
                        print(obj_of_cls.lst_of_text_in_pages[i])
                case var if var == obj_with_cui.MENU.ページを時計回りで回転します:
                    # ページを時計回りで回転します
                    file_path_of_pdf_as_str_type = obj_with_cui.input_target_of_pdf(obj_of_cls.EXTENSION)
                    obj_of_cls.read_metadata(file_path_of_pdf_as_str_type)
                    page = obj_with_cui.input_rotating_page(obj_of_cls.num_of_pages)
                    degrees = obj_with_cui.input_degrees()
                    obj_of_cls.rotate_page_clockwise(file_path_of_pdf_as_str_type, page, degrees)
                case var if var == obj_with_cui.MENU.終了します:
                    # 終了します
                    break
                case _:
                    pass
        except Exception as e:
            print(f"エラー: {e}")
        print()
    obj_of_pt = PathTools()
    file_of_exe_as_path_type = Path(__file__)
    file_of_log_as_path_type = obj_of_pt.get_file_path_of_log(file_of_exe_as_path_type)
    obj_of_cls.write_log(file_of_log_as_path_type)
    return True


if __name__ == "__main__":
    main()
