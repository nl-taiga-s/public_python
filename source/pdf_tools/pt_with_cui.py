import re
from enum import Enum
from pathlib import Path

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
                print(str(e))
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

    def input_target_of_pdf(self, extension: str) -> str:
        """対象のPDFファイルのパスを入力します。"""
        while True:
            try:
                result = False
                cancel = False
                file_path_of_pdf_as_str_type = input("対象のPDFファイルパスを入力してください。: ").strip()
                if file_path_of_pdf_as_str_type == "":
                    raise Exception("未入力です。")
                file_of_pdf_as_path_type = Path(file_path_of_pdf_as_str_type).expanduser()
                file_path_of_pdf_as_str_type = str(file_of_pdf_as_path_type)
                if not file_of_pdf_as_path_type.exists():
                    raise Exception("存在しません。")
                if not file_of_pdf_as_path_type.is_file():
                    raise Exception("ファイル以外は入力しないでください。")
                if file_of_pdf_as_path_type.suffix.lower() != extension:
                    raise Exception("PDF以外は入力しないでください。")
            except Exception as e:
                print(str(e))
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
                print(str(e))
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

    def input_writing_metadata(self, metadata_of_writer: dict, fields: list, creation_date: str, utc: str) -> dict:
        """書き込み用のメタデータを入力します"""
        while True:
            try:
                result = False
                cancel = False
                for key_of_r, key_of_w in fields:
                    match key_of_r:
                        case "creation_date":
                            metadata_of_writer[key_of_w] = creation_date
                        case "modification_date":
                            time = self.obj_of_dt2.convert_for_metadata_in_pdf(utc)
                            metadata_of_writer[key_of_w] = time
                        case _:
                            value = input(f"{key_of_r.capitalize().replace("_", " ")}: ").strip()
                            metadata_of_writer[key_of_w] = value
            except Exception as e:
                print(str(e))
            except KeyboardInterrupt:
                cancel = True
            else:
                result = True
            finally:
                if cancel or result:
                    break
        if cancel:
            raise
        return metadata_of_writer

    def input_list_of_merge(self, extension: str) -> list:
        """マージ元の全てのファイルを入力します"""
        pdfs = []
        try:
            while True:
                cancel = False
                print("マージ元のPDFファイルパスを順番に入力してください。")
                file_as_str_type = self.input_target_of_pdf(extension)
                pdfs.append(file_as_str_type)
                keep = self.input_bool("対象のファイルは、まだありますか？")
                if not keep:
                    if pdfs:
                        break
                    else:
                        print("ファイルが何も入力されていません。")
        except Exception as e:
            print(str(e))
        except KeyboardInterrupt:
            cancel = True
        else:
            pass
        finally:
            if cancel:
                raise
            return pdfs

    def input_page_range(self, num_of_pages: int) -> list:
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
                print(str(e))
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

    def input_rotating_page(self, num_of_pages: int) -> int:
        """回転するページを入力します"""
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
                print(str(e))
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
                print(str(e))
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
                    case var if var in self.d_of_bool["yes"]:
                        result = True
                    case var if var in self.d_of_bool["no"]:
                        pass
                    case _:
                        raise Exception("無効な入力です。")
            except Exception as e:
                error = True
                print(str(e))
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
    while True:
        try:
            result = False
            cancel = False
            obj_of_pt = PathTools()
            obj_with_cui = PT_With_Cui()
            obj_of_cls = PdfTools()
            print(*obj_of_cls.log, sep="\n")
            # メニューを選択します
            option = obj_with_cui.select_menu()
            match option:
                case var if var == obj_with_cui.MENU.ファイルを暗号化します:
                    # ファイルを暗号化します
                    file_path_of_pdf_as_str_type = obj_with_cui.input_target_of_pdf(obj_of_cls.EXTENSION)
                    password = obj_with_cui.input_password("暗号化")
                    _, log = obj_of_cls.encrypt(file_path_of_pdf_as_str_type, password)
                    print(*log, sep="\n")
                case var if var == obj_with_cui.MENU.ファイルを復号化します:
                    # ファイルを復号化します
                    file_path_of_pdf_as_str_type = obj_with_cui.input_target_of_pdf(obj_of_cls.EXTENSION)
                    password = obj_with_cui.input_password("復号化")
                    _, log = obj_of_cls.decrypt(file_path_of_pdf_as_str_type, password)
                    print(*log, sep="\n")
                case var if var == obj_with_cui.MENU.メタデータを出力します:
                    # メタデータを出力します
                    file_path_of_pdf_as_str_type = obj_with_cui.input_target_of_pdf(obj_of_cls.EXTENSION)
                    result, log = obj_of_cls.read_file(file_path_of_pdf_as_str_type)
                    print(*log, sep="\n")
                    if not result:
                        raise Exception
                    _, log = obj_of_cls.get_metadata(file_path_of_pdf_as_str_type)
                    print(*log, sep="\n")
                case var if var == obj_with_cui.MENU.メタデータを書き込みます:
                    # メタデータを書き込みます
                    file_path_of_pdf_as_str_type = obj_with_cui.input_target_of_pdf(obj_of_cls.EXTENSION)
                    result, log = obj_of_cls.read_file(file_path_of_pdf_as_str_type)
                    print(*log, sep="\n")
                    if not result:
                        raise Exception
                    obj_of_cls.metadata_of_writer = obj_with_cui.input_writing_metadata(
                        obj_of_cls.metadata_of_writer, obj_of_cls.fields, obj_of_cls.creation_date, obj_of_cls.UTC_OF_JP
                    )
                    _, log = obj_of_cls.write_metadata(file_path_of_pdf_as_str_type, obj_of_cls.metadata_of_writer)
                    print(*log, sep="\n")
                case var if var == obj_with_cui.MENU.ファイルをマージします:
                    # ファイルをマージします
                    pdfs = obj_with_cui.input_list_of_merge(obj_of_cls.EXTENSION)
                    _, log = obj_of_cls.merge(pdfs)
                    print(*log, sep="\n")
                case var if var == obj_with_cui.MENU.ページを抽出します:
                    # ページを抽出します
                    file_path_of_pdf_as_str_type = obj_with_cui.input_target_of_pdf(obj_of_cls.EXTENSION)
                    result, log = obj_of_cls.read_file(file_path_of_pdf_as_str_type)
                    print(*log, sep="\n")
                    if not result:
                        raise Exception
                    begin_page, end_page = obj_with_cui.input_page_range(obj_of_cls.num_of_pages)
                    _, log = obj_of_cls.extract_pages(file_path_of_pdf_as_str_type, begin_page, end_page)
                    print(*log, sep="\n")
                case var if var == obj_with_cui.MENU.ページを削除します:
                    # ページを削除します
                    file_path_of_pdf_as_str_type = obj_with_cui.input_target_of_pdf(obj_of_cls.EXTENSION)
                    result, log = obj_of_cls.read_file(file_path_of_pdf_as_str_type)
                    print(*log, sep="\n")
                    if not result:
                        raise Exception
                    begin_page, end_page = obj_with_cui.input_page_range(obj_of_cls.num_of_pages)
                    _, log = obj_of_cls.delete_pages(file_path_of_pdf_as_str_type, begin_page, end_page)
                    print(*log, sep="\n")
                case var if var == obj_with_cui.MENU.テキストを抽出します:
                    # テキストを抽出します
                    file_path_of_pdf_as_str_type = obj_with_cui.input_target_of_pdf(obj_of_cls.EXTENSION)
                    result, log = obj_of_cls.read_file(file_path_of_pdf_as_str_type)
                    print(*log, sep="\n")
                    if not result:
                        raise Exception
                    begin_page, end_page = obj_with_cui.input_page_range(obj_of_cls.num_of_pages)
                    _, log = obj_of_cls.extract_text(file_path_of_pdf_as_str_type, begin_page, end_page)
                    print(*log, sep="\n")
                case var if var == obj_with_cui.MENU.ページを時計回りで回転します:
                    # ページを時計回りで回転します
                    file_path_of_pdf_as_str_type = obj_with_cui.input_target_of_pdf(obj_of_cls.EXTENSION)
                    result, log = obj_of_cls.read_file(file_path_of_pdf_as_str_type)
                    print(*log, sep="\n")
                    if not result:
                        raise Exception
                    page = obj_with_cui.input_rotating_page(obj_of_cls.num_of_pages)
                    degrees = obj_with_cui.input_degrees()
                    _, log = obj_of_cls.rotate_page_clockwise(file_path_of_pdf_as_str_type, page, degrees)
                    print(*log, sep="\n")
                case _:
                    pass
        except Exception as e:
            print(f"処理が失敗しました。: {str(e)}")
        except KeyboardInterrupt:
            cancel = True
        else:
            result = True
            print("処理が成功しました。")
            file_of_exe_as_path_type = Path(__file__)
            file_of_log_as_path_type = obj_of_pt.get_file_path_of_log(file_of_exe_as_path_type)
            _, _ = obj_of_cls.write_log(file_of_log_as_path_type)
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
