import os
import platform

def input_folder_path() -> list:
    """入力する"""
    while True:
        folder_path_from = input("ファイルを一括変換するフォルダを指定してください。: ")
        folder_path_to = input(
            "一括変換したファイルを格納するフォルダを指定してください。: "
        )
        if folder_path_from == "" or folder_path_to == "":
            return
        if os.path.exists(folder_path_from) and os.path.exists(folder_path_to):
            # 存在する場合
            if os.path.isdir(folder_path_from) and os.path.isdir(folder_path_to):
                # フォルダの場合
                return folder_path_from, folder_path_to


def main():
    try:
        from source.convert_office_to_pdf import cotp_class
        folder_path_from, folder_path_to = input_folder_path()
        obj = cotp_class.ConvertOfficeToPdf(folder_path_from, folder_path_to)
        obj.convert_all()
    except EnvironmentError as e:
        print("このスクリプトは、Windowsで実行してください。")


if __name__ == "__main__":
    main()
