import os

from source.get_file_list import gfl_class


def input_folder_path() -> str:
    """フォルダのパスを入力します"""
    while True:
        folder_path = input("ファイルを検索したいフォルダを入力してください。: ")
        if folder_path == "":
            return
        if os.path.exists(folder_path):
            # 存在する場合
            if os.path.isdir(folder_path):
                # フォルダの場合
                return folder_path


def input_pattern() -> str:
    """検索パターンを入力します"""
    pattern = input("ファイルの検索パターンを入力してください。: ")
    return pattern


def main():
    try:
        folder_path = input_folder_path()
        obj = gfl_class.GetFileList(folder_path)
        pattern = input_pattern()
        obj.extract_by_pattern(pattern)
        obj.print(obj.list_file_after)
    except Exception as e:
        print(e)


if __name__ == "__main__":
    main()
