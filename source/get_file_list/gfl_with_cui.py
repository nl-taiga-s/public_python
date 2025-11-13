import sys
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

    def _input_folder_path(self) -> str:
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
            except KeyboardInterrupt:
                raise
            except Exception:
                raise
            else:
                break
            finally:
                pass
        return folder_s

    def _input_bool(self, msg: str) -> bool:
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
            except KeyboardInterrupt:
                raise
            except Exception:
                raise
            else:
                break
            finally:
                pass
        return result


def main() -> bool:
    """主要関数"""
    result: bool = False
    # ログを設定します
    try:
        obj_of_pt: PathTools = PathTools()
        obj_of_lt: LogTools = LogTools()
        file_of_exe_p: Path = Path(__file__)
        file_of_log_p: Path = obj_of_pt._get_file_path_of_log(file_of_exe_p)
        obj_of_lt.file_path_of_log = str(file_of_log_p)
        obj_of_lt._setup_file_handler(obj_of_lt.file_path_of_log)
        obj_of_lt._setup_stream_handler()
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
            obj_of_cls.folder_path = obj_with_cui._input_folder_path()
            obj_of_cls.recursive = obj_with_cui._input_bool("フォルダを再帰的に検索しますか？")
            obj_of_cls.search_directly_under_folder()
            obj_of_cls.pattern = input("ファイルの検索パターンを入力してください。: ")
            obj_of_cls.extract_by_pattern()
        except KeyboardInterrupt:
            sys.exit(0)
        except Exception as e:
            obj_of_lt.logger.critical(f"***処理が失敗しました。***: \n{str(e)}")
        else:
            result = True
            obj_of_lt.logger.info("***処理が成功しました。***")
        finally:
            pass
        if obj_with_cui._input_bool("終了しますか？"):
            break
    return result


if __name__ == "__main__":
    main()
