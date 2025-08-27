import sys
from pathlib import Path

from source.common.common import PathTools


class COTP_With_Cui:
    def __init__(self):
        """初期化します"""
        self.d_of_bool = {
            "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
            "no": ["いいえ", "0", "No", "no", "N", "n"],
        }

    def input_folder_path(self) -> list:
        """フォルダのパスを入力します"""
        try:
            while True:
                folder_path_from = input("ファイルを一括変換するフォルダを指定してください。: ")
                folder_path_to = input("一括変換したファイルを格納するフォルダを指定してください。: ")
                if folder_path_from == "" and folder_path_to == "":
                    return [None, None]
                folder_of_from_as_path_type = Path(folder_path_from).expanduser()
                folder_of_to_as_path_type = Path(folder_path_to).expanduser()
                folder_path_from = str(folder_of_from_as_path_type)
                folder_path_to = str(folder_of_to_as_path_type)
                if folder_of_from_as_path_type.exists() and folder_of_to_as_path_type.exists():
                    # 存在する場合
                    if folder_of_from_as_path_type.is_dir() and folder_of_to_as_path_type.is_dir():
                        # フォルダの場合
                        return [folder_path_from, folder_path_to]
        except Exception as e:
            print(str(e))
        except KeyboardInterrupt:
            sys.exit(0)

    def input_continue(self) -> bool:
        """続けるかどうかを入力します"""
        try:
            str_of_bool = input("再度、行いますか？: ")
            match str_of_bool:
                case var if var in self.d_of_bool["yes"]:
                    return True
                case var if var in self.d_of_bool["no"]:
                    return False
        except Exception as e:
            print(str(e))
        except KeyboardInterrupt:
            sys.exit(0)


def main() -> bool:
    """主要関数"""
    try:
        while True:
            result = False
            from source.convert_office_to_pdf.cotp_class import ConvertOfficeToPDF

            obj_of_pt = PathTools()
            obj_with_cui = COTP_With_Cui()
            folder_path_from, folder_path_to = obj_with_cui.input_folder_path()
            if folder_path_from is None and folder_path_to is None:
                return result
            obj_of_cls = ConvertOfficeToPDF(folder_path_from, folder_path_to)
            print(*obj_of_cls.log, sep="\n")
            for _ in range(obj_of_cls.number_of_f):
                _, log = obj_of_cls.handle_file()
                print(*log, sep="\n")
                obj_of_cls.move_to_next_file()
            if obj_with_cui.input_continue():
                continue
            else:
                break
    except ImportError as e:
        print(str(e))
    except Exception as e:
        print(str(e))
    else:
        result = True
    finally:
        file_of_exe_as_path_type = Path(__file__)
        file_of_log_as_path_type = obj_of_pt.get_file_path_of_log(file_of_exe_as_path_type)
        obj_of_cls.write_log(file_of_log_as_path_type)
        return result


if __name__ == "__main__":
    main()
