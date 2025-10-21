from pathlib import Path

from source.common.common import LogTools, PathTools
from source.get_file_list.gfl_class import GetFileList


class GFL_With_Cui:
    def __init__(self):
        """初期化します"""
        self.binary_choices: dict = {
            "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
            "no": ["いいえ", "0", "No", "no", "N", "n"],
        }

    def input_folder_path(self) -> str:
        """フォルダのパスを入力します"""
        while True:
            try:
                folder_s: str = input("ファイルを検索したいフォルダを入力してください。: ").strip()
                if folder_s == "":
                    raise Exception("フォルダが未入力です。")
                folder_p: Path = Path(folder_s).expanduser()
                folder_s = str(folder_p)
                if not folder_p.exists():
                    raise Exception("フォルダが存在しません。")
                if not folder_p.is_dir():
                    raise Exception("フォルダではありません。")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                raise
            else:
                break
            finally:
                pass
        return folder_s

    def input_text(self, msg: str) -> str:
        """文字列を入力します"""
        while True:
            try:
                text: str = input(f"{msg}: ").strip()
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                raise
            else:
                break
            finally:
                pass
        return text

    def input_bool(self, msg: str) -> bool:
        """はいかいいえをを入力します"""
        result: bool = False
        while True:
            try:
                binary_choice: str = input(f"{msg}\n(Yes => y or No => n): ").strip()
                match binary_choice:
                    case var if var in self.binary_choices["yes"]:
                        result = True
                    case var if var in self.binary_choices["no"]:
                        pass
                    case _:
                        raise Exception("無効な入力です。")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                raise
            else:
                break
            finally:
                pass
        return result


def main() -> bool:
    """主要関数"""
    result: bool = False
    # ログを設定する
    try:
        obj_of_pt: PathTools = PathTools()
        obj_of_lt: LogTools = LogTools()
        file_of_exe_p: Path = Path(__file__)
        file_of_log_p: Path = obj_of_pt.get_file_path_of_log(file_of_exe_p)
        obj_of_lt.file_path_of_log = str(file_of_log_p)
        obj_of_lt.setup_file_handler(obj_of_lt.file_path_of_log)
        obj_of_lt.setup_stream_handler()
    except Exception as e:
        print(f"error: \n{str(e)}")
        return result
    else:
        pass
    finally:
        pass
    # 処理の本体部分
    obj_with_cui: GFL_With_Cui = GFL_With_Cui()
    obj_of_cls: GetFileList = GetFileList(obj_of_lt.logger)
    while True:
        result = False
        try:
            obj_of_cls.folder_path = obj_with_cui.input_folder_path()
            obj_of_cls.recursive = obj_with_cui.input_bool("フォルダを再帰的に検索しますか？")
            obj_of_cls.search_recursively()
            obj_of_cls.pattern = obj_with_cui.input_text("ファイルの検索パターンを入力してください。")
            obj_of_cls.extract_by_pattern()
        except Exception:
            print("***処理が失敗しました。***")
        except KeyboardInterrupt:
            raise
        else:
            result = True
            print("***処理が成功しました。***")
        finally:
            pass
        if obj_with_cui.input_bool("終了しますか？"):
            break
        else:
            continue
    return result


if __name__ == "__main__":
    main()
