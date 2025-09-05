from pathlib import Path

from source.common.common import LogTools, PathTools


class COTP_With_Cui:
    def __init__(self):
        """初期化します"""
        self.d_of_bool = {
            "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
            "no": ["いいえ", "0", "No", "no", "N", "n"],
        }

    def input_folder_path(self) -> list:
        """フォルダのパスを入力します"""
        while True:
            try:
                result = False
                cancel = False
                folder_path_from = input("ファイルを一括変換するフォルダを指定してください。: ").strip()
                folder_path_to = input("一括変換したファイルを格納するフォルダを指定してください。: ").strip()
                if folder_path_from == "" or folder_path_to == "":
                    raise Exception("未入力です。")
                folder_of_from_as_path_type = Path(folder_path_from).expanduser()
                folder_of_to_as_path_type = Path(folder_path_to).expanduser()
                folder_path_from = str(folder_of_from_as_path_type)
                folder_path_to = str(folder_of_to_as_path_type)
                if not folder_of_from_as_path_type.exists() or not folder_of_to_as_path_type.exists():
                    raise Exception("存在しません。")
                if not folder_of_from_as_path_type.is_dir() or not folder_of_to_as_path_type.is_dir():
                    raise Exception("フォルダではありません。")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                cancel = True
            else:
                result = True
            finally:
                if cancel or result:
                    break
        if cancel:
            raise
        return [folder_path_from, folder_path_to]

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
                print(f"error: \n{str(e)}")
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
    try:
        result = False
        obj_of_pt = PathTools()
        obj_of_lt = LogTools()
        file_of_exe_as_path_type = Path(__file__)
        file_of_log_as_path_type = obj_of_pt.get_file_path_of_log(file_of_exe_as_path_type)
        obj_of_lt.file_path_of_log = str(file_of_log_as_path_type)
        if not obj_of_lt.setup_file_handler(obj_of_lt.file_path_of_log):
            raise
        if not obj_of_lt.setup_stream_handler():
            raise
    except Exception as e:
        print(f"error: \n{str(e)}")
    else:
        result = True
    finally:
        if not result:
            return result
    while True:
        try:
            result = False
            cancel = False
            from source.convert_office_to_pdf.cotp_class import ConvertOfficeToPDF

            obj_with_cui = COTP_With_Cui()
            obj_of_cls = ConvertOfficeToPDF(obj_of_lt.logger)
            obj_of_cls.folder_path_from, obj_of_cls.folder_path_to = obj_with_cui.input_folder_path()
            if not obj_of_cls.create_file_list():
                raise
            for _ in range(obj_of_cls.number_of_f):
                obj_of_cls.handle_file()
                if obj_of_cls.complete:
                    break
                obj_of_cls.move_to_next_file()
        except ImportError as e:
            cancel = True
            print(f"error: \n{str(e)}")
        except Exception as e:
            print("処理が失敗しました。")
            print(f"error: \n{str(e)}")
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
