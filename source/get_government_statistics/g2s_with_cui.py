import asyncio
import sys
from pathlib import Path
from typing import Any

from pandas import DataFrame

from source.common.common import LogTools, PathTools
from source.get_government_statistics.g2s_class import GetGovernmentStatistics

# クレジット表示
# このサービスは、政府統計総合窓口(e-Stat)のAPI機能を使用していますが、サービスの内容は国によって保証されたものではありません。


class GS_With_Cui:
    def __init__(self):
        """初期化します"""
        self.binary_choices: dict = {
            "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
            "no": ["いいえ", "0", "No", "no", "N", "n"],
        }

    def select_element(self, elements: Any) -> list:
        """要素を選択します"""
        lst: list = []
        while True:
            try:
                match elements:
                    case dict():
                        # リストに変換する
                        lst = list(elements.items())
                        for i, (k, v) in enumerate(lst, start=1):
                            print(f"{i}. {k}: {v}")
                    case list():
                        lst = elements
                        for i, s in enumerate(lst, start=1):
                            print(f"{i}. {s}")
                    case _:
                        raise Exception("この変数の型は、対象外です。")
                text: str = input("番号を入力してください。: ").strip()
                if text == "":
                    raise Exception("番号が未入力です。")
                if not text.isdecimal():
                    raise Exception("数字を入力してください。")
                num: int = int(text)
                if num < 1 or num > len(lst):
                    raise Exception("入力した番号が範囲外です。")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                sys.exit(0)
            else:
                break
            finally:
                pass
        return lst[num - 1]

    def input_text(self, msg: str) -> str:
        """文字列を入力します"""
        text: str = ""
        while True:
            try:
                text = input(f"{msg}: ").strip()
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                sys.exit(0)
            else:
                break
            finally:
                pass
        return text

    def input_lst_of_text(self, msg: str) -> list:
        """複数の文字列を入力します"""
        lst: list = []
        try:
            while True:
                text: str = self.input_text(msg)
                lst.append(text)
                keep: bool = self.input_bool("入力する文字列は、まだありますか？")
                if not keep:
                    if lst:
                        break
                    else:
                        print("文字列が何も入力されていません。")
        except Exception as e:
            print(f"error: \n{str(e)}")
        except KeyboardInterrupt:
            sys.exit(0)
        else:
            pass
        finally:
            pass
        return lst

    def input_stats_data_id(self) -> str:
        """統計表IDを入力します"""
        text: str = ""
        # 桁
        DIGIT: int = 10
        while True:
            try:
                text = input("統計表IDを入力してください。: ").strip()
                if text == "":
                    raise Exception("統計表IDが未入力です。")
                if not text.isdecimal():
                    raise Exception("数字を入力してください。")
                if len(text) != DIGIT:
                    raise Exception(f"{DIGIT}桁で入力してください。")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                sys.exit(0)
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
                text: str = input(f"{msg}\n(Yes => y or No => n): ").strip()
                match text:
                    case var if var in self.binary_choices["yes"]:
                        result = True
                    case var if var in self.binary_choices["no"]:
                        pass
                    case _:
                        raise Exception("無効な入力です。")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                sys.exit(0)
            else:
                break
            finally:
                pass
        return result


async def main() -> bool:
    """主要関数"""
    # ログを設定します
    result: bool = False
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
    obj_with_cui: GS_With_Cui = GS_With_Cui()
    obj_of_cls: GetGovernmentStatistics = GetGovernmentStatistics(obj_of_lt.logger)
    while True:
        try:
            obj_of_cls.lst_of_data_type = obj_with_cui.select_element(obj_of_cls.dct_of_data_type)
            if obj_with_cui.input_bool(f"{obj_of_cls.write_stats_data_ids_to_file.__doc__} => 行いますか？"):
                obj_of_lt.logger.info(f"{obj_of_cls.write_stats_data_ids_to_file.__doc__} => 開始しました。")
                # 統計表IDをテキストファイルに書き出す
                await obj_of_cls.write_stats_data_ids_to_file(
                    page_generator=obj_of_cls.get_stats_data_ids(),
                    data_type_key=obj_of_cls.lst_of_data_type[obj_of_cls.KEY],
                    chunk_size=100,
                )
                obj_of_lt.logger.info(f"{obj_of_cls.write_stats_data_ids_to_file.__doc__} => 終了しました。")
            obj_of_cls.STATS_DATA_ID = obj_with_cui.input_stats_data_id()
            df: DataFrame = obj_of_cls.get_data_from_api()
            obj_of_cls.lst_of_match = obj_with_cui.select_element(obj_of_cls.dct_of_match)
            if obj_of_cls.lst_of_match[obj_of_cls.KEY] != "何もしない":
                obj_of_cls.lst_of_keyword = obj_with_cui.input_lst_of_text("抽出するキーワードを入力してください。")
                if len(obj_of_cls.lst_of_keyword) > 1:
                    obj_of_cls.lst_of_logic = obj_with_cui.select_element(obj_of_cls.dct_of_logic)
                df: DataFrame = obj_of_cls.filter_data(df)
            obj_of_cls.show_table(df)
        except Exception:
            obj_of_lt.logger.critical("***処理が失敗しました。***")
        except KeyboardInterrupt:
            sys.exit(0)
        else:
            result = True
            obj_of_lt.logger.info("***処理が成功しました。***")
        finally:
            pass
        if obj_with_cui.input_bool("終了しますか？"):
            break
    return result


if __name__ == "__main__":
    asyncio.run(main())
