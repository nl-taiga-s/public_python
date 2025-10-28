import re
import sys
from enum import Enum
from pathlib import Path

from source.common.common import DatetimeTools, LogTools, PathTools
from source.pdf_tools.pt_class import PdfTools


class PT_With_Cui:
    def __init__(self):
        """初期化します"""
        self.binary_choices: dict = {
            "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
            "no": ["いいえ", "0", "No", "no", "N", "n"],
        }
        self.obj_of_dt2: DatetimeTools = DatetimeTools()
        self.MENU: Enum = Enum(
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
        self.DEGREES: Enum = Enum("DEGREES", ["90", "180", "270"])

    def select_menu(self) -> str:
        """メニューを選択します"""
        menu: list = [f"({m.value}) {m.name}" for m in self.MENU]
        while True:
            try:
                print(*menu, sep="\n")
                choice: str = input("メニューの番号を入力してください。: ").strip()
                if choice == "":
                    raise Exception("番号が未入力です。")
                if not choice.isdecimal():
                    raise Exception("数字を入力してください。")
                num: int = int(choice)
                if 1 > num or num > len(self.MENU):
                    raise Exception("入力した番号が範囲外です。")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                sys.exit(0)
            else:
                break
            finally:
                pass
        return self.MENU(num)

    def input_file_path(self, ext: str) -> str:
        """ファイルパスを入力します。"""
        while True:
            try:
                file_s: str = input("ファイルパスを入力してください。: ").strip()
                if file_s == "":
                    raise Exception("未入力です。")
                file_p: Path = Path(file_s).expanduser()
                file_s = str(file_p)
                if not file_p.exists():
                    raise Exception("存在しません。")
                if not file_p.is_file():
                    raise Exception("ファイル以外は入力しないでください。")
                if file_p.suffix.lower() != ext:
                    raise Exception(f"{ext}ファイル以外は入力しないでください。")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                sys.exit(0)
            else:
                break
            finally:
                pass
        return file_s

    def input_lst_of_file_path(self, ext: str) -> list:
        """複数のファイルパスを入力します"""
        lst: list = []
        try:
            while True:
                print("ファイルパスを順番に入力してください。")
                file_s: str = self.input_file_path(ext)
                lst.append(file_s)
                keep: bool = self.input_bool("対象のファイルは、まだありますか？")
                if not keep:
                    if lst:
                        break
                    else:
                        print("ファイルが何も入力されていません。")
        except Exception as e:
            print(f"error: \n{str(e)}")
        except KeyboardInterrupt:
            sys.exit(0)
        else:
            pass
        finally:
            pass
        return lst

    def input_password(self, keyword: str) -> str:
        """パスワードを入力します"""
        while True:
            try:
                pw: str = input(f"{keyword}のパスワードを入力してください。: ").strip()
                if pw == "":
                    raise Exception("パスワードが未入力です。")
                if not re.fullmatch(r"[A-Za-z0-9_-]+", pw):
                    raise Exception("以下の文字で入力してください。\n* 半角英数字\n* アンダーバー\n* ハイフン")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                sys.exit(0)
            else:
                break
            finally:
                pass
        return pw

    def input_writing_metadata(self, obj_of_cls: PdfTools) -> dict:
        """書き込み用のメタデータを入力します"""
        while True:
            try:
                for key_of_r, key_of_w in obj_of_cls.fields:
                    match key_of_r:
                        case "creation_date":
                            obj_of_cls.metadata_of_writer[key_of_w] = obj_of_cls.creation_date
                        case "modification_date":
                            time: str = self.obj_of_dt2.convert_for_metadata_in_pdf(obj_of_cls.UTC_OF_JP)
                            obj_of_cls.metadata_of_writer[key_of_w] = time
                        case _:
                            value: str = input(f"{key_of_r.capitalize().replace("_", " ")}: ").strip()
                            obj_of_cls.metadata_of_writer[key_of_w] = value
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                sys.exit(0)
            else:
                break
            finally:
                pass
        return obj_of_cls.metadata_of_writer

    def input_pages_range(self, num_of_pages: int) -> list:
        """ページ範囲を入力します"""
        while True:
            try:
                begin_page: str = input("始めのページを入力してください。: ").strip()
                end_page: str = input("終わりのページを入力してください。: ").strip()
                if begin_page == "" or end_page == "":
                    raise Exception("未入力です。")
                if not begin_page.isdecimal() or not end_page.isdecimal():
                    raise Exception("数字を入力してください。")
                bp: int = int(begin_page)
                ep: int = int(end_page)
                if bp < 1 or ep > num_of_pages or bp > ep:
                    raise Exception("指定のページ範囲が不正です。")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                sys.exit(0)
            else:
                break
            finally:
                pass
        return [bp, ep]

    def input_page(self, num_of_pages: int) -> int:
        """ページを入力します"""
        while True:
            try:
                page: str = input("ページを入力してください。: ").strip()
                if page == "":
                    raise Exception("未入力です。")
                if not page.isdecimal():
                    raise Exception("数字を入力してください。")
                p: int = int(page)
                if p < 1 or p > num_of_pages:
                    raise Exception("指定のページ範囲が不正です。")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                sys.exit(0)
            else:
                break
            finally:
                pass
        return p

    def input_degrees(self) -> int:
        """回転する度数を入力します"""
        degrees: list = [f"({d.value}) {d.name}" for d in self.DEGREES]
        while True:
            try:
                print(*degrees, sep="\n")
                num: str = input("度数の番号を入力してください。: ").strip()
                if num == "":
                    raise Exception("未入力です。")
                if not num.isdecimal():
                    raise Exception("数字を入力してください。")
                n = int(num)
                if 1 > n or n > len(self.DEGREES):
                    raise Exception("入力した番号が範囲外です。")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                sys.exit(0)
            else:
                break
            finally:
                pass
        return int(self.DEGREES(n).name)

    def input_bool(self, msg: str) -> bool:
        """はいかいいえをを入力します"""
        result: bool = False
        while True:
            try:
                binary_choice: str = input(f"{msg}\n(Yes => y or No => n): ").strip()
                match binary_choice:
                    case var if var in self.binary_choices["yes"]:
                        result = True
                    case var if var in self.binary_choices["no"]:
                        pass
                    case _:
                        raise Exception("無効な入力です。")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                sys.exit(0)
            else:
                break
            finally:
                pass
        return result


def main() -> bool:
    """主要関数"""
    result: bool = False
    # ログを設定します
    try:
        obj_of_pt: PathTools = PathTools()
        obj_of_lt: LogTools = LogTools()
        file_of_exe_p: Path = Path(__file__)
        file_of_log_p: Path = obj_of_pt.get_file_path_of_log(file_of_exe_p)
        obj_of_lt.file_path_of_log = str(file_of_log_p)
        obj_of_lt.setup_file_handler(obj_of_lt.file_path_of_log)
        obj_of_lt.setup_stream_handler()
    except Exception as e:
        print(f"error: \n{str(e)}")
        return result
    else:
        pass
    finally:
        pass
    # 処理の本体部分
    obj_with_cui = PT_With_Cui()
    obj_of_cls = PdfTools(obj_of_lt.logger)
    while True:
        result: bool = False
        try:
            # メニューを選択します
            option: str = obj_with_cui.select_menu()
            match option:
                case var if var == obj_with_cui.MENU.ファイルを暗号化します:
                    # ファイルを暗号化します
                    obj_of_cls.file_path = obj_with_cui.input_file_path(obj_of_cls.EXTENSION)
                    pw: str = obj_with_cui.input_password("暗号化")
                    obj_of_cls.encrypt(pw)
                case var if var == obj_with_cui.MENU.ファイルを復号化します:
                    # ファイルを復号化します
                    obj_of_cls.file_path = obj_with_cui.input_file_path(obj_of_cls.EXTENSION)
                    pw: str = obj_with_cui.input_password("復号化")
                    obj_of_cls.decrypt(pw)
                case var if var == obj_with_cui.MENU.メタデータを出力します:
                    # メタデータを出力します
                    obj_of_cls.file_path = obj_with_cui.input_file_path(obj_of_cls.EXTENSION)
                    obj_of_cls.read_file()
                    obj_of_cls.get_metadata()
                case var if var == obj_with_cui.MENU.メタデータを書き込みます:
                    # メタデータを書き込みます
                    obj_of_cls.file_path = obj_with_cui.input_file_path(obj_of_cls.EXTENSION)
                    obj_of_cls.read_file()
                    obj_of_cls.metadata_of_writer = obj_with_cui.input_writing_metadata(obj_of_cls)
                    obj_of_cls.write_metadata(obj_of_cls.metadata_of_writer)
                case var if var == obj_with_cui.MENU.ファイルをマージします:
                    # ファイルをマージします
                    pdfs: list = obj_with_cui.input_lst_of_file_path(obj_of_cls.EXTENSION)
                    obj_of_cls.merge(pdfs)
                case var if var == obj_with_cui.MENU.ページを抽出します:
                    # ページを抽出します
                    obj_of_cls.file_path = obj_with_cui.input_file_path(obj_of_cls.EXTENSION)
                    obj_of_cls.read_file()
                    begin_page, end_page = obj_with_cui.input_pages_range(obj_of_cls.num_of_pages)
                    obj_of_cls.extract_pages(begin_page, end_page)
                case var if var == obj_with_cui.MENU.ページを削除します:
                    # ページを削除します
                    obj_of_cls.file_path = obj_with_cui.input_file_path(obj_of_cls.EXTENSION)
                    obj_of_cls.read_file()
                    begin_page, end_page = obj_with_cui.input_pages_range(obj_of_cls.num_of_pages)
                    obj_of_cls.delete_pages(begin_page, end_page)
                case var if var == obj_with_cui.MENU.テキストを抽出します:
                    # テキストを抽出します
                    obj_of_cls.file_path = obj_with_cui.input_file_path(obj_of_cls.EXTENSION)
                    obj_of_cls.read_file()
                    begin_page, end_page = obj_with_cui.input_pages_range(obj_of_cls.num_of_pages)
                    obj_of_cls.extract_text(begin_page, end_page)
                case var if var == obj_with_cui.MENU.ページを時計回りで回転します:
                    # ページを時計回りで回転します
                    obj_of_cls.file_path = obj_with_cui.input_file_path(obj_of_cls.EXTENSION)
                    obj_of_cls.read_file()
                    page: int = obj_with_cui.input_page(obj_of_cls.num_of_pages)
                    degrees: int = obj_with_cui.input_degrees()
                    obj_of_cls.rotate_page_clockwise(page, degrees)
                case _:
                    pass
        except Exception:
            obj_of_lt.logger.critical("***処理が失敗しました。***")
        except KeyboardInterrupt:
            sys.exit(0)
        else:
            result = True
            obj_of_lt.logger.info("***処理が成功しました。***")
        finally:
            pass
        if obj_with_cui.input_bool("終了しますか？"):
            break
    return result


if __name__ == "__main__":
    main()
