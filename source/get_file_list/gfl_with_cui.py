import os

from source.common.common import CsvTools, PathTools
from source.get_file_list.gfl_class import GetFileList


def input_folder_path() -> str:
    """フォルダのパスを入力します"""
    while True:
        folder_path = input("ファイルを検索したいフォルダを入力してください。: ")
        if folder_path == "":
            print("処理を中止します。")
            return
        if os.path.exists(folder_path):
            # 存在する場合
            if os.path.isdir(folder_path):
                # フォルダの場合
                return folder_path


def input_bool_of_recursive() -> bool:
    """フォルダを再帰的に検索するかどうかを入力します"""
    while True:
        d_of_bool = {
            "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
            "no": ["いいえ", "0", "No", "no", "N", "n"],
        }
        str_of_r = input(
            "フォルダを再帰的に検索するかどうかを入力してください。"
            "(Yes => y or No => n): "
        )
        match str_of_r:
            case var if var in d_of_bool["yes"]:
                return True
            case var if var in d_of_bool["no"]:
                return False
            case _:
                raise ValueError("無効な入力です。: {str_or_r}")


def input_pattern() -> str:
    """検索パターンを入力します"""
    pattern = input("ファイルの検索パターンを入力してください。: ")
    return pattern


def main():
    try:
        folder_path = input_folder_path()
        bool_of_r = input_bool_of_recursive()
        obj_of_gfl = GetFileList(folder_path, bool_of_r)
        obj_of_ct = CsvTools()
        obj_of_pt = PathTools()
        pattern = input_pattern()
        obj_of_gfl.extract_by_pattern(pattern)
        obj_of_gfl.print_list(obj_of_gfl.list_file_after)
        file_path_of_log = obj_of_pt.get_file_path_of_log(__file__, obj_of_gfl.now)
        obj_of_ct.write_list(
            file_path_of_log,
            obj_of_gfl.list_of_csv,
        )
    except Exception as e:
        print(e)


if __name__ == "__main__":
    main()
