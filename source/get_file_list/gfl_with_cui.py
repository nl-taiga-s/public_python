from pathlib import Path

from gfl_class import GetFileList

from source.common.common import DatetimeTools, PathTools

d_of_bool = {
    "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
    "no": ["いいえ", "0", "No", "no", "N", "n"],
}


def input_folder_path() -> str:
    """フォルダのパスを入力します"""
    while True:
        folder_path = input("ファイルを検索したいフォルダを入力してください。: ")
        if folder_path == "":
            print("処理を中止します。")
            return
        fp = Path(folder_path)
        if fp.exists():
            # 存在する場合
            if fp.is_dir():
                # フォルダの場合
                return folder_path


def input_bool_of_recursive() -> bool:
    """フォルダを再帰的に検索するかどうかを入力します"""
    while True:
        str_of_r = input(
            "フォルダを再帰的に検索するかどうかを入力してください。"
            "(Yes => y or No => n): "
        )
        d_of_bool
        match str_of_r:
            case var if var in d_of_bool["yes"]:
                return True
            case var if var in d_of_bool["no"]:
                return False
            case _:
                raise ValueError(f"無効な入力です。: {str_of_r}")


def input_bool_of_result() -> bool:
    """resultファイルを出力するかどうかを入力します"""
    while True:
        str_of_l = input(
            "resultファイルを出力するかどうかを入力してください。"
            "(Yes => y or No => n): "
        )
        d_of_bool
        match str_of_l:
            case var if var in d_of_bool["yes"]:
                return True
            case var if var in d_of_bool["no"]:
                return False
            case _:
                raise ValueError(f"無効な入力です。: {str_of_l}")


def input_pattern() -> str:
    """検索パターンを入力します"""
    pattern = input("ファイルの検索パターンを入力してください。: ")
    return pattern


def main():
    try:
        folder_path = input_folder_path()
        bool_of_r = input_bool_of_recursive()
        obj_of_gfl = GetFileList(folder_path, bool_of_r)
        obj_of_pt = PathTools()
        obj_of_dt2 = DatetimeTools()
        pattern = input_pattern()
        obj_of_gfl.extract_by_pattern(pattern)
        print(*obj_of_gfl.list_file_after, sep="\n")
        if input_bool_of_result():
            file_path_of_result = obj_of_pt.get_file_path_of_result(__file__)
            file_path_of_result = obj_of_pt.convert_path_to_str(file_path_of_result)
            try:
                # ファイルに書き出し
                with open(file_path_of_result, "w", encoding="utf-8", newline="") as f:
                    for element in obj_of_gfl.list_file_after:
                        f.write(f"{element},")
                        f.write(f"{obj_of_dt2.convert_dt_to_str()}\n")
            except Exception as e:
                print(e)
            print(f"ログファイルを出力しました。: {file_path_of_result}")
    except Exception as e:
        print(e)


if __name__ == "__main__":
    main()
