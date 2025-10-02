import asyncio
from pathlib import Path
from typing import Any

from pandas import DataFrame

from source.common.common import LogTools, PathTools
from source.get_government_statistics.g2s_class import GetGovernmentStatistics


class GS_With_Cui:
    def __init__(self):
        """初期化します"""
        self.dct_of_bool: dict = {
            "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
            "no": ["いいえ", "0", "No", "no", "N", "n"],
        }
        # 取得するデータ形式
        self.dct_of_data_type: dict = {
            "xml": "タグ構造のデータ",
            "json": "キーと値のペアのデータ",
            "csv": "カンマ区切りのデータ",
        }
        # 検索方式
        self.dct_of_match: dict = {
            "部分一致": "フィールドの値にキーワードが含まれている",
            "完全一致": "フィールドの値がキーワードと完全に一致している",
        }
        # 表示順
        self.lst_of_order: list = ["先頭", "末尾"]
        # 抽出方式
        self.dct_of_logic: dict = {"OR抽出": "複数のキーワードのいずれかが含まれている", "AND抽出": "複数のキーワードの全てが含まれている"}

    def select_element(self, elements: Any) -> list:
        """要素を選択します"""
        flag: bool = False
        cancel: bool = False
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
                print(f"error: \n{repr(e)}")
            except KeyboardInterrupt:
                cancel = True
            else:
                flag = True
            finally:
                if cancel or flag:
                    break
        if cancel:
            raise
        return lst[num - 1]

    def input_text(self, msg: str) -> str:
        """文字列を入力します"""
        flag: bool = False
        cancel: bool = False
        while True:
            try:
                text: str = input(f"{msg}: ").strip()
            except Exception as e:
                print(f"error: \n{repr(e)}")
            except KeyboardInterrupt:
                cancel = True
            else:
                flag = True
            finally:
                if cancel or flag:
                    break
        if cancel:
            raise
        return text

    def input_list_of_text(self, msg: str) -> list:
        """複数の文字列を入力します"""
        cancel: bool = False
        list_of_txt: list = []
        try:
            while True:
                text: str = self.input_text(msg)
                list_of_txt.append(text)
                keep: bool = self.input_bool("入力する文字列は、まだありますか？")
                if not keep:
                    if list_of_txt:
                        break
                    else:
                        print("文字列が何も入力されていません。")
        except Exception as e:
            print(f"error: \n{repr(e)}")
        except KeyboardInterrupt:
            cancel = True
        else:
            pass
        finally:
            if cancel:
                raise
            return list_of_txt

    def input_stats_data_id(self) -> str:
        """統計表IDを入力します"""
        flag: bool = False
        cancel: bool = False
        # 桁
        DIGIT: int = 10
        while True:
            try:
                text: str = input("統計表IDを入力してください。: ").strip()
                if text == "":
                    raise Exception("統計表IDが未入力です。")
                if not text.isdecimal():
                    raise Exception("数字を入力してください。")
                if len(text) != DIGIT:
                    raise Exception(f"{DIGIT}桁で入力してください。")
            except Exception as e:
                print(f"error: \n{repr(e)}")
            except KeyboardInterrupt:
                cancel = True
            else:
                flag = True
            finally:
                if cancel or flag:
                    break
        if cancel:
            raise
        return text

    def input_bool(self, msg: str) -> bool:
        """はいかいいえをを入力します"""
        flag: bool = False
        cancel: bool = False
        while True:
            try:
                text: str = input(f"{msg}\n(Yes => y or No => n): ").strip()
                match text:
                    case var if var in self.dct_of_bool["yes"]:
                        flag = True
                    case var if var in self.dct_of_bool["no"]:
                        continue
                    case _:
                        raise Exception("無効な入力です。")
            except Exception as e:
                print(f"error: \n{repr(e)}")
            except KeyboardInterrupt:
                cancel = True
            else:
                flag = True
            finally:
                if flag or cancel:
                    break
        if cancel:
            raise
        return flag


async def main() -> bool:
    """主要関数"""
    flag: bool = False
    try:
        obj_of_pt: PathTools = PathTools()
        obj_of_lt: LogTools = LogTools()
        file_of_exe_as_path_type: Path = Path(__file__)
        file_of_log_as_path_type = obj_of_pt.get_file_path_of_log(file_of_exe_as_path_type)
        obj_of_lt.file_path_of_log = str(file_of_log_as_path_type)
        if not obj_of_lt.setup_file_handler(obj_of_lt.file_path_of_log):
            raise
        if not obj_of_lt.setup_stream_handler():
            raise
    except Exception as e:
        print(f"error: \n{repr(e)}")
    else:
        flag = True
    finally:
        if not flag:
            return flag
    flag = False
    cancel: bool = False
    while True:
        try:
            obj_with_cui: GS_With_Cui = GS_With_Cui()
            obj_of_cls: GetGovernmentStatistics = GetGovernmentStatistics(obj_of_lt.logger)
            obj_of_cls.lst_of_data_type = obj_with_cui.select_element(obj_with_cui.dct_of_data_type)
            async for page in obj_of_cls.get_stats_data_ids():
                for stat_id, info in page.items():
                    match obj_of_cls.lst_of_data_type[obj_of_cls.KEY]:
                        case "xml":
                            stat_name: str = info.get("stat_name", "")
                            title: str = info.get("title", "")
                            print(f"{stat_id}, {stat_name}, {title}")
                        case "json":
                            statistics_name: str = info.get("statistics_name", "")
                            title: str = info.get("title", "")
                            print(f"{stat_id}, {statistics_name}, {title}")
                        case "csv":
                            stat_name: str = info.get("stat_name", "")
                            category: str = info.get("category", "")
                            print(f"{stat_id}, {stat_name}, {category}")
                        case _:
                            raise Exception("データタイプが対応していません。")
            obj_of_cls.STATS_DATA_ID = obj_with_cui.input_stats_data_id()
            df: DataFrame = obj_of_cls.get_data_from_api()
            obj_of_cls.lst_of_match = obj_with_cui.select_element(obj_with_cui.dct_of_match)
            obj_of_cls.lst_of_keyword = obj_with_cui.input_list_of_keyword("抽出するキーワードを入力してください。")
            if len(obj_of_cls.lst_of_keyword) > 1:
                obj_of_cls.lst_of_logic = obj_with_cui.select_element(obj_with_cui.dct_of_logic)
            filtered_df: DataFrame = obj_of_cls.filter_data(df)
            obj_of_cls.order = str(obj_with_cui.select_element(obj_with_cui.lst_of_order))
            if not obj_of_cls.show_data(filtered_df):
                raise
        except Exception as e:
            print(f"***処理が失敗しました。\n{repr(e)}***")
        except KeyboardInterrupt:
            cancel = True
        else:
            flag = True
            print("***処理が成功しました。***")
        finally:
            if cancel:
                break
            if obj_with_cui.input_bool("終了しますか？"):
                break
            else:
                continue
    return flag


if __name__ == "__main__":
    asyncio.run(main())
