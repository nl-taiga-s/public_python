import sys
from pathlib import Path

from source.common.common import PathTools


def check_os() -> object:
    try:
        from source.convert_office_to_pdf.cotp_class import ConvertOfficeToPDF
    except ImportError as e:
        print(e)
        return None
    else:
        return ConvertOfficeToPDF


class COTP_With_Cui:
    def __init__(self):
        """初期化します"""
        self.cotp = check_os()
        if self.cotp is None:
            return None

    def input_folder_path(self) -> list:
        """フォルダのパスを入力します"""
        try:
            while True:
                folder_path_from = input("ファイルを一括変換するフォルダを指定してください。: ")
                folder_path_to = input("一括変換したファイルを格納するフォルダを指定してください。: ")
                if folder_path_from == "" and folder_path_to == "":
                    return [None, None]
                folder_of_from_as_path_type = Path(folder_path_from)
                folder_of_to_as_path_type = Path(folder_path_to)
                if folder_of_from_as_path_type.exists() and folder_of_to_as_path_type.exists():
                    # 存在する場合
                    if folder_of_from_as_path_type.is_dir() and folder_of_to_as_path_type.is_dir():
                        # フォルダの場合
                        return [folder_path_from, folder_path_to]
        except Exception as e:
            print(e)
        except KeyboardInterrupt:
            sys.exit(0)


def main() -> bool:
    """主要関数"""
    obj_with_cui = COTP_With_Cui()
    if obj_with_cui.cotp is None:
        return False
    folder_path_from, folder_path_to = obj_with_cui.input_folder_path()
    if folder_path_from is None and folder_path_to is None:
        return False
    obj_of_cls = obj_with_cui.cotp(folder_path_from, folder_path_to)
    obj_of_cls.convert_all()
    obj_of_pt = PathTools()
    file_of_exe_as_path_type = Path(__file__)
    file_of_log_as_path_type = obj_of_pt.get_file_path_of_log(file_of_exe_as_path_type)
    obj_of_cls.write_log(file_of_log_as_path_type)
    return True


if __name__ == "__main__":
    main()
