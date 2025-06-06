import glob
import os


def search_file():
    """ファイルを検索します。"""
    while True:
        folder_path = input("ファイルを検索したいフォルダを入力してください。: ")
        if folder_path == "":
            return
        if os.path.exists(folder_path):
            # 存在する場合
            if os.path.isdir(folder_path):
                # フォルダの場合
                break
    # 検索対象のフォルダから再帰的に検索して、ファイルのリストを取得する
    list_file_before = [
        f
        for f in glob.glob(os.path.join(folder_path, "**"), recursive=True)
        if os.path.isfile(f)
    ]
    pattern_of_file_name = input("ファイルの検索パターンを入力してください。: ")
    # ファイルのリストから検索パターンで抽出したリストを取得する
    list_file_after = [f for f in list_file_before if pattern_of_file_name in f]
    print("検索結果: ")
    print(*list_file_after, sep="\n")


def main():
    print(search_file.__doc__)
    search_file()


if __name__ == "__main__":
    main()
