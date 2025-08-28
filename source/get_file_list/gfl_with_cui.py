import sys
from pathlib import Path

from source.common.common import DatetimeTools, PathTools
from source.get_file_list.gfl_class import GetFileList


class GFL_With_Cui:
    def __init__(self):
        """初期化します"""
        self.d_of_bool = {
            "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
            "no": ["いいえ", "0", "No", "no", "N", "n"],
        }
        self.obj_of_pt = PathTools()
        self.obj_of_dt2 = DatetimeTools()

    def input_folder_path(self) -> str:
        """フォルダのパスを入力します"""
        while True:
            try:
                result = False
                folder_path_of_pdf_as_str_type = input("ファイルを検索したいフォルダを入力してください。: ").strip()
                if folder_path_of_pdf_as_str_type == "":
                    raise Exception("フォルダが未入力です。")
                folder_as_path_type = Path(folder_path_of_pdf_as_str_type).expanduser()
                folder_path_of_pdf_as_str_type = str(folder_as_path_type)
                if not folder_as_path_type.exists():
                    raise Exception("フォルダが存在しません。")
                if not folder_as_path_type.is_dir():
                    raise Exception("フォルダではありません。")
            except Exception as e:
                print(str(e))
            except KeyboardInterrupt:
                sys.exit(0)
            else:
                result = True
            finally:
                if result:
                    break
        return folder_path_of_pdf_as_str_type

    def input_bool(self, msg: str) -> bool:
        """はいかいいえをを入力します"""
        while True:
            try:
                result = False
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
                sys.exit(0)
            else:
                pass
            finally:
                if not error:
                    break
        return result

    def input_pattern(self) -> str:
        """検索パターンを入力します"""
        pattern = ""
        try:
            pattern = input("ファイルの検索パターンを入力してください。: ").strip()
        except Exception as e:
            print(str(e))
        except KeyboardInterrupt:
            sys.exit(0)
        else:
            pass
        finally:
            return pattern


def main() -> bool:
    """主要関数"""
    while True:
        try:
            result = False
            # pytest
            test_var = input("test_input\n(Yes => y or No => n): ").strip()
            if test_var == "y":
                raise Exception("test_input")
            obj_of_pt = PathTools()
            obj_with_cui = GFL_With_Cui()
            folder_path = obj_with_cui.input_folder_path()
            bool_of_r = obj_with_cui.input_bool("フォルダを再帰的に検索しますか？")
            obj_of_cls = GetFileList(folder_path, bool_of_r)
            print(*obj_of_cls.log, sep="\n")
            pattern = obj_with_cui.input_pattern()
            _, log = obj_of_cls.extract_by_pattern(pattern)
            print(*log, sep="\n")
        except Exception as e:
            print(f"処理が失敗しました。: {str(e)}")
        else:
            result = True
            print("処理が成功しました。")
            file_of_exe_as_path_type = Path(__file__)
            file_of_log_as_path_type = obj_of_pt.get_file_path_of_log(file_of_exe_as_path_type)
            _, _ = obj_of_cls.write_log(file_of_log_as_path_type)
        finally:
            # pytest
            if test_var == "y":
                break
            if obj_with_cui.input_bool("終了しますか？"):
                break
            else:
                continue
    return result


if __name__ == "__main__":
    main()
