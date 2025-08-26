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
                folder_path_of_pdf_as_str_type = input("ファイルを検索したいフォルダを入力してください。: ")
                if folder_path_of_pdf_as_str_type == "":
                    return None
                folder_as_path_type = Path(folder_path_of_pdf_as_str_type).expanduser()
                folder_path_of_pdf_as_str_type = str(folder_as_path_type)
                if folder_as_path_type.exists():
                    # 存在する場合
                    if folder_as_path_type.is_dir():
                        # フォルダの場合
                        return folder_path_of_pdf_as_str_type
            except Exception as e:
                print(e)
            except KeyboardInterrupt:
                sys.exit(0)

    def input_bool_of_recursive(self) -> bool:
        """フォルダを再帰的に検索するかどうかを入力します"""
        try:
            while True:
                str_of_bool = input("フォルダを再帰的に検索するかどうかを入力してください。\n(Yes => y or No => n): ")
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
            sys.exit(0)

    def input_pattern(self) -> str:
        """検索パターンを入力します"""
        try:
            pattern = input("ファイルの検索パターンを入力してください。: ")
            return pattern
        except KeyboardInterrupt:
            sys.exit(0)


def main() -> bool:
    """主要関数"""
    obj_of_pt = PathTools()
    obj_with_cui = GFL_With_Cui()
    folder_path = obj_with_cui.input_folder_path()
    if folder_path is None:
        return False
    bool_of_r = obj_with_cui.input_bool_of_recursive()
    obj_of_cls = GetFileList(folder_path, bool_of_r)
    print(*obj_of_cls.log, sep="\n")
    pattern = obj_with_cui.input_pattern()
    _, log = obj_of_cls.extract_by_pattern(pattern)
    print(*log, sep="\n")
    file_of_exe_as_path_type = Path(__file__)
    file_of_log_as_path_type = obj_of_pt.get_file_path_of_log(file_of_exe_as_path_type)
    obj_of_cls.write_log(file_of_log_as_path_type)
    return True


if __name__ == "__main__":
    main()
