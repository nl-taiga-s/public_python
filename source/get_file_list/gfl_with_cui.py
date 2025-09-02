from pathlib import Path

from source.common.common import LogTools, PathTools
from source.get_file_list.gfl_class import GetFileList


class GFL_With_Cui:
    def __init__(self):
        """初期化します"""
        self.d_of_bool = {
            "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
            "no": ["いいえ", "0", "No", "no", "N", "n"],
        }

    def input_folder_path(self) -> str:
        """フォルダのパスを入力します"""
        while True:
            try:
                result = False
                cancel = False
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
                cancel = True
            else:
                result = True
            finally:
                if cancel or result:
                    break
        if cancel:
            raise
        return folder_path_of_pdf_as_str_type

    def input_pattern(self) -> str:
        """検索パターンを入力します"""
        pattern = ""
        try:
            cancel = False
            pattern = input("ファイルの検索パターンを入力してください。: ").strip()
        except Exception as e:
            print(str(e))
        except KeyboardInterrupt:
            cancel = True
        else:
            pass
        finally:
            if cancel:
                raise
            return pattern

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
    obj_of_pt = PathTools()
    obj_of_lt = LogTools()
    file_of_exe_as_path_type = Path(__file__)
    file_of_log_as_path_type = obj_of_pt.get_file_path_of_log(file_of_exe_as_path_type)
    obj_of_lt.file_path_of_log = str(file_of_log_as_path_type)
    obj_of_lt.setup_file_handler(obj_of_lt.file_path_of_log)
    obj_of_lt.setup_stream_handler()
    while True:
        try:
            result = False
            cancel = False
            obj_with_cui = GFL_With_Cui()
            obj_of_cls = GetFileList(obj_of_lt.logger)
            obj_of_cls.folder_path = obj_with_cui.input_folder_path()
            obj_of_cls.bool_of_r = obj_with_cui.input_bool("フォルダを再帰的に検索しますか？")
            if not obj_of_cls.search_recursively():
                raise
            obj_of_cls.pattern = obj_with_cui.input_pattern()
            if not obj_of_cls.extract_by_pattern():
                raise
        except Exception:
            print("処理が失敗しました。")
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
