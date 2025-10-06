import shutil
from pathlib import Path

from source.common.common import LogTools, PathTools
from source.convert_libre_to_pdf.cltp_class import ConvertLibreToPDF


class CLTP_With_Cui:
    def __init__(self):
        """初期化します"""
        self.binary_choices: dict = {
            "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
            "no": ["いいえ", "0", "No", "no", "N", "n"],
        }

    def input_folder_path(self) -> list:
        """フォルダのパスを入力します"""
        result: bool = False
        cancel: bool = False
        while True:
            try:
                folder_path_from: str = input("ファイルを一括変換するフォルダを指定してください。: ").strip()
                folder_path_to: str = input("一括変換したファイルを格納するフォルダを指定してください。: ").strip()
                if folder_path_from == "" or folder_path_to == "":
                    raise Exception("未入力です。")
                folder_of_from_p: Path = Path(folder_path_from).expanduser()
                folder_of_to_p: Path = Path(folder_path_to).expanduser()
                folder_path_from = str(folder_of_from_p)
                folder_path_to = str(folder_of_to_p)
                if not folder_of_from_p.exists() or not folder_of_to_p.exists():
                    raise Exception("存在しません。")
                if not folder_of_from_p.is_dir() or not folder_of_to_p.is_dir():
                    raise Exception("フォルダではありません。")
            except Exception as e:
                print(f"error: \n{repr(e)}")
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
        result: bool = False
        cancel: bool = False
        while True:
            try:
                binary_choice: str = input(f"{msg}\n(Yes => y or No => n): ").strip()
                match binary_choice:
                    case var if var in self.binary_choices["yes"]:
                        result = True
                    case var if var in self.binary_choices["no"]:
                        continue
                    case _:
                        raise Exception("無効な入力です。")
            except Exception as e:
                print(f"error: \n{repr(e)}")
            except KeyboardInterrupt:
                cancel = True
            else:
                pass
            finally:
                if cancel:
                    break
        if cancel:
            raise
        return result


def main() -> bool:
    """主要関数"""
    # LibreOfficeのコマンドが使用可能か確認する
    result: bool = False
    try:
        LIBRE_COMMAND: str = "soffice"
        if not shutil.which(LIBRE_COMMAND):
            raise ImportError("LibreOfficeをインストールしてください。: \nhttps://ja.libreoffice.org/")
    except ImportError as e:
        print(f"error: \n{repr(e)}")
        return result
    else:
        pass
    finally:
        pass
    # ログを設定する
    try:
        obj_of_pt: PathTools = PathTools()
        obj_of_lt: LogTools = LogTools()
        file_of_exe_p: Path = Path(__file__)
        file_of_log_p: Path = obj_of_pt.get_file_path_of_log(file_of_exe_p)
        obj_of_lt.file_path_of_log = str(file_of_log_p)
        if not obj_of_lt.setup_file_handler(obj_of_lt.file_path_of_log):
            raise
        if not obj_of_lt.setup_stream_handler():
            raise
    except Exception as e:
        print(f"error: \n{repr(e)}")
        return result
    else:
        pass
    finally:
        pass
    # 処理の本体部分
    cancel: bool = False
    while True:
        try:
            obj_with_cui: CLTP_With_Cui = CLTP_With_Cui()
            obj_of_cls: ConvertLibreToPDF = ConvertLibreToPDF(obj_of_lt.logger)
            obj_of_cls.folder_path_from, obj_of_cls.folder_path_to = obj_with_cui.input_folder_path()
            if not obj_of_cls.create_file_lst():
                raise
            for _ in range(obj_of_cls.number_of_f):
                obj_of_cls.convert_file()
                if obj_of_cls.complete:
                    break
                obj_of_cls.move_to_next_file()
        except Exception as e:
            print(f"***処理が失敗しました。***: \n{repr(e)}")
        except KeyboardInterrupt:
            cancel = True
        else:
            result = True
            print("***処理が成功しました。***")
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
