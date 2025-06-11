import os
from source.search_file import sf_class

def input_folder_path() -> str:
    """入力する"""
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
    pattern=input("ファイルの検索パターンを入力してください。: ")
    return pattern

def main():
    try:
        folder_path=input_folder_path()
        obj=sf_class.SearchFile(folder_path)
        pattern=input_pattern()
        obj.extract_by_pattern(pattern)
    except Exception as e:
        print(e)


if __name__ == "__main__":
    main()
