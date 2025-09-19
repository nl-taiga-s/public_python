import re
from enum import Enum
from pathlib import Path

from source.common.common import DatetimeTools, LogTools, PathTools
from source.pdf_tools.pt_class import PdfTools


class PT_With_Cui:
    def __init__(self):
        """初期化します"""
        self.dct_of_bool = {
            "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
            "no": ["いいえ", "0", "No", "no", "N", "n"],
        }
        self.obj_of_dt2 = DatetimeTools()
        self.MENU = Enum(
            "MENU",
            [
                "ファイルを暗号化します",
                "ファイルを復号化します",
                "メタデータを出力します",
                "メタデータを書き込みます",
                "ファイルをマージします",
                "ページを抽出します",
                "ページを削除します",
                "テキストを抽出します",
                "ページを時計回りで回転します",
            ],
        )
        self.DEGREES = Enum("DEGREES", ["90", "180", "270"])

    def select_menu(self) -> str:
        """メニューを選択します"""
        str_of_menu = [f"({m.value}) {m.name}" for m in self.MENU]
        while True:
            try:
                result = False
                cancel = False
                print(*str_of_menu, sep="\n")
                n = input("メニューの番号を入力してください。: ").strip()
                if n == "":
                    raise Exception("番号が未入力です。")
                if not n.isdecimal():
                    raise Exception("数字を入力してください。")
                num = int(n)
                if 1 > num or num > len(self.MENU):
                    raise Exception("入力した番号が範囲外です。")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                cancel = True
            else:
                result = True
            finally:
                if cancel or result:
                    break
        if cancel:
            raise
        return self.MENU(num)

    def input_file_path(self, extension: str) -> str:
        """ファイルパスを入力します。"""
        while True:
            try:
                result = False
                cancel = False
                file_path_of_pdf_as_str_type = input("ファイルパスを入力してください。: ").strip()
                if file_path_of_pdf_as_str_type == "":
                    raise Exception("未入力です。")
                file_of_pdf_as_path_type = Path(file_path_of_pdf_as_str_type).expanduser()
                file_path_of_pdf_as_str_type = str(file_of_pdf_as_path_type)
                if not file_of_pdf_as_path_type.exists():
                    raise Exception("存在しません。")
                if not file_of_pdf_as_path_type.is_file():
                    raise Exception("ファイル以外は入力しないでください。")
                if file_of_pdf_as_path_type.suffix.lower() != extension:
                    raise Exception(f"{extension}以外は入力しないでください。")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                cancel = True
            else:
                result = True
            finally:
                if cancel or result:
                    break
        if cancel:
            raise
        return file_path_of_pdf_as_str_type

    def input_password(self, keyword: str) -> str:
        """パスワードを入力します"""
        while True:
            try:
                result = False
                cancel = False
                password = input(f"{keyword}のパスワードを入力してください。: ").strip()
                if password == "":
                    raise Exception("パスワードが未入力です。")
                if re.fullmatch(r"[A-Za-z0-9_-]+", password):
                    result = True
                else:
                    raise Exception("以下の文字で入力してください。\n* 半角英数字\n* アンダーバー\n* ハイフン")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                cancel = True
            else:
                pass
            finally:
                if cancel or result:
                    break
        if cancel:
            raise
        return password

    def input_writing_metadata(self, obj_of_cls: PdfTools) -> dict:
        """書き込み用のメタデータを入力します"""
        while True:
            try:
                result = False
                cancel = False
                for key_of_r, key_of_w in obj_of_cls.fields:
                    match key_of_r:
                        case "creation_date":
                            obj_of_cls.metadata_of_writer[key_of_w] = obj_of_cls.creation_date
                        case "modification_date":
                            time = self.obj_of_dt2.convert_for_metadata_in_pdf(obj_of_cls.UTC_OF_JP)
                            obj_of_cls.metadata_of_writer[key_of_w] = time
                        case _:
                            value = input(f"{key_of_r.capitalize().replace("_", " ")}: ").strip()
                            obj_of_cls.metadata_of_writer[key_of_w] = value
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                cancel = True
            else:
                result = True
            finally:
                if cancel or result:
                    break
        if cancel:
            raise
        return obj_of_cls.metadata_of_writer

    def input_list_of_file_path(self, extension: str) -> list:
        """複数のファイルパスを入力します"""
        list_of_fp = []
        try:
            while True:
                cancel = False
                print("ファイルパスを順番に入力してください。")
                file_as_str_type = self.input_file_path(extension)
                list_of_fp.append(file_as_str_type)
                keep = self.input_bool("対象のファイルは、まだありますか？")
                if not keep:
                    if list_of_fp:
                        break
                    else:
                        print("ファイルが何も入力されていません。")
        except Exception as e:
            print(f"error: \n{str(e)}")
        except KeyboardInterrupt:
            cancel = True
        else:
            pass
        finally:
            if cancel:
                raise
            return list_of_fp

    def input_pages_range(self, num_of_pages: int) -> list:
        """ページ範囲を入力します"""
        while True:
            try:
                result = False
                cancel = False
                begin_page = input("始めのページを入力してください。: ").strip()
                end_page = input("終わりのページを入力してください。: ").strip()
                if begin_page == "" or end_page == "":
                    raise Exception("未入力です。")
                if not begin_page.isdecimal() or not end_page.isdecimal():
                    raise Exception("数字を入力してください。")
                begin_page = int(begin_page)
                end_page = int(end_page)
                if begin_page < 1 or end_page > num_of_pages or begin_page > end_page:
                    raise Exception("指定のページ範囲が不正です。")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                cancel = True
            else:
                result = True
            finally:
                if cancel or result:
                    break
        if cancel:
            raise
        return [begin_page, end_page]

    def input_page(self, num_of_pages: int) -> int:
        """ページを入力します"""
        while True:
            try:
                result = False
                cancel = False
                page = input("ページを入力してください。: ").strip()
                if page == "":
                    raise Exception("未入力です。")
                if not page.isdecimal():
                    raise Exception("数字を入力してください。")
                page = int(page)
                if page < 1 or page > num_of_pages:
                    raise Exception("指定のページ範囲が不正です。")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                cancel = True
            else:
                result = True
            finally:
                if cancel or result:
                    break
        if cancel:
            raise
        return page

    def input_degrees(self) -> int:
        """回転する度数を入力します"""
        str_of_degrees = [f"({d.value}) {d.name}" for d in self.DEGREES]
        while True:
            try:
                result = False
                cancel = False
                print(*str_of_degrees, sep="\n")
                n = input("度数の番号を入力してください。: ").strip()
                if n == "":
                    raise Exception("未入力です。")
                if not n.isdecimal():
                    raise Exception("数字を入力してください。")
                num = int(n)
                if 1 > num or num > len(self.DEGREES):
                    raise Exception("入力した番号が範囲外です。")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                cancel = True
            else:
                result = True
            finally:
                if cancel or result:
                    break
        if cancel:
            raise
        return int(self.DEGREES(num).name)

    def input_bool(self, msg: str) -> bool:
        """はいかいいえをを入力します"""
        while True:
            try:
                result = False
                cancel = False
                error = False
                str_of_bool = input(f"{msg}\n(Yes => y or No => n): ").strip()
                match str_of_bool:
                    case var if var in self.dct_of_bool["yes"]:
                        result = True
                    case var if var in self.dct_of_bool["no"]:
                        pass
                    case _:
                        raise Exception("無効な入力です。")
            except Exception as e:
                error = True
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                cancel = True
            else:
                pass
            finally:
                if not error:
                    break
        if cancel:
            raise
        return result


def main() -> bool:
    """主要関数"""
    try:
        result = False
        obj_of_pt = PathTools()
        obj_of_lt = LogTools()
        file_of_exe_as_path_type = Path(__file__)
        file_of_log_as_path_type = obj_of_pt.get_file_path_of_log(file_of_exe_as_path_type)
        obj_of_lt.file_path_of_log = str(file_of_log_as_path_type)
        if not obj_of_lt.setup_file_handler(obj_of_lt.file_path_of_log):
            raise
        if not obj_of_lt.setup_stream_handler():
            raise
    except Exception as e:
        print(f"error: \n{str(e)}")
    else:
        result = True
    finally:
        if not result:
            return result
    while True:
        try:
            result = False
            cancel = False
            obj_with_cui = PT_With_Cui()
            obj_of_cls = PdfTools(obj_of_lt.logger)
            # メニューを選択します
            option = obj_with_cui.select_menu()
            match option:
                case var if var == obj_with_cui.MENU.ファイルを暗号化します:
                    # ファイルを暗号化します
                    obj_of_cls.file_path = obj_with_cui.input_file_path(obj_of_cls.EXTENSION)
                    password = obj_with_cui.input_password("暗号化")
                    if not obj_of_cls.encrypt(password):
                        raise
                case var if var == obj_with_cui.MENU.ファイルを復号化します:
                    # ファイルを復号化します
                    obj_of_cls.file_path = obj_with_cui.input_file_path(obj_of_cls.EXTENSION)
                    password = obj_with_cui.input_password("復号化")
                    if not obj_of_cls.decrypt(password):
                        raise
                case var if var == obj_with_cui.MENU.メタデータを出力します:
                    # メタデータを出力します
                    obj_of_cls.file_path = obj_with_cui.input_file_path(obj_of_cls.EXTENSION)
                    if not obj_of_cls.read_file():
                        raise
                    if not obj_of_cls.get_metadata():
                        raise
                case var if var == obj_with_cui.MENU.メタデータを書き込みます:
                    # メタデータを書き込みます
                    obj_of_cls.file_path = obj_with_cui.input_file_path(obj_of_cls.EXTENSION)
                    if not obj_of_cls.read_file():
                        raise
                    obj_of_cls.metadata_of_writer = obj_with_cui.input_writing_metadata(obj_of_cls)
                    if not obj_of_cls.write_metadata(obj_of_cls.metadata_of_writer):
                        raise
                case var if var == obj_with_cui.MENU.ファイルをマージします:
                    # ファイルをマージします
                    pdfs = obj_with_cui.input_list_of_file_path(obj_of_cls.EXTENSION)
                    if not obj_of_cls.merge(pdfs):
                        raise
                case var if var == obj_with_cui.MENU.ページを抽出します:
                    # ページを抽出します
                    obj_of_cls.file_path = obj_with_cui.input_file_path(obj_of_cls.EXTENSION)
                    if not obj_of_cls.read_file():
                        raise
                    begin_page, end_page = obj_with_cui.input_pages_range(obj_of_cls.num_of_pages)
                    if not obj_of_cls.extract_pages(begin_page, end_page):
                        raise
                case var if var == obj_with_cui.MENU.ページを削除します:
                    # ページを削除します
                    obj_of_cls.file_path = obj_with_cui.input_file_path(obj_of_cls.EXTENSION)
                    if not obj_of_cls.read_file():
                        raise
                    begin_page, end_page = obj_with_cui.input_pages_range(obj_of_cls.num_of_pages)
                    if not obj_of_cls.delete_pages(begin_page, end_page):
                        raise
                case var if var == obj_with_cui.MENU.テキストを抽出します:
                    # テキストを抽出します
                    obj_of_cls.file_path = obj_with_cui.input_file_path(obj_of_cls.EXTENSION)
                    if not obj_of_cls.read_file():
                        raise
                    begin_page, end_page = obj_with_cui.input_pages_range(obj_of_cls.num_of_pages)
                    if not obj_of_cls.extract_text(begin_page, end_page):
                        raise
                case var if var == obj_with_cui.MENU.ページを時計回りで回転します:
                    # ページを時計回りで回転します
                    obj_of_cls.file_path = obj_with_cui.input_file_path(obj_of_cls.EXTENSION)
                    if not obj_of_cls.read_file():
                        raise
                    page = obj_with_cui.input_page(obj_of_cls.num_of_pages)
                    degrees = obj_with_cui.input_degrees()
                    if not obj_of_cls.rotate_page_clockwise(page, degrees):
                        raise
                case _:
                    pass
        except Exception as e:
            print("処理が失敗しました。")
            print(f"error: \n{str(e)}")
        except KeyboardInterrupt:
            cancel = True
        else:
            result = True
            print("処理が成功しました。")
        finally:
            if cancel:
                break
            if obj_with_cui.input_bool("終了しますか？"):
                break
            else:
                continue
    return result


if __name__ == "__main__":
    main()
