import asyncio
from pathlib import Path
from typing import Any

from source.common.common import LogTools, PathTools
from source.get_government_statistics.g2s_class import GetGovernmentStatistics


class GS_With_Cui:
    def __init__(self):
        """初期化します"""
        self.dct_of_bool = {
            "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
            "no": ["いいえ", "0", "No", "no", "N", "n"],
        }
        # 取得するデータ形式
        self.dct_of_data_type = {
            "xml": "タグ構造のデータ",
            "json": "キーと値のペアのデータ",
            "csv": "カンマ区切りのデータ",
        }
        # 検索方式
        self.dct_of_match = {"部分一致": "フィールドの値にキーワードが含まれている", "完全一致": "フィールドの値がキーワードと完全に一致している"}
        # 表示順
        self.lst_of_order = ["先頭", "末尾"]
        # 抽出方式
        self.dct_of_logic = {"OR抽出": "複数のキーワードのいずれかが含まれている", "AND抽出": "複数のキーワードの全てが含まれている"}

    def select_target(self, target: Any) -> list:
        """対象を選択します"""
        while True:
            try:
                result = False
                cancel = False
                match target:
                    case dict():
                        # リストに変換する
                        lst = list(target.items())
                        for i, (k, v) in enumerate(lst, start=1):
                            print(f"{i}. {k}: {v}")
                    case list():
                        lst = target
                        for i, s in enumerate(lst, start=1):
                            print(f"{i}. {s}")
                    case _:
                        raise Exception("この変数の型は、対象外です。")
                num_of_target = input("番号を入力してください。: ").strip()
                if num_of_target == "":
                    raise Exception("番号が未入力です。")
                if not num_of_target.isdecimal():
                    raise Exception("数字を入力してください。")
                num_of_target = int(num_of_target)
                if num_of_target < 1 or num_of_target > len(lst):
                    raise Exception("入力した番号が範囲外です。")
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
        return lst[num_of_target - 1]

    def input_keyword(self, msg: str) -> str:
        """キーワードを入力します"""
        while True:
            try:
                result = False
                cancel = False
                pattern = input(f"{msg}: ").strip()
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
        return pattern

    def input_list_of_keyword(self, msg: str) -> list:
        """複数のキーワードを入力します"""
        list_of_kw = []
        try:
            while True:
                cancel = False
                keyword = self.input_keyword(msg)
                list_of_kw.append(keyword)
                keep = self.input_bool("キーワードは、まだありますか？")
                if not keep:
                    if list_of_kw:
                        break
                    else:
                        print("キーワードが何も入力されていません。")
        except Exception as e:
            print(f"error: \n{str(e)}")
        except KeyboardInterrupt:
            cancel = True
        else:
            pass
        finally:
            if cancel:
                raise
            return list_of_kw

    def input_stats_data_id(self) -> str:
        """統計表IDを入力します"""
        while True:
            try:
                result = False
                cancel = False
                id = input("統計表IDを入力してください。: ").strip()
                if id == "":
                    raise Exception("統計表IDが未入力です。")
                if not id.isdecimal():
                    raise Exception("数字を入力してください。")
                # 桁
                DIGIT = 10
                if len(id) != DIGIT:
                    raise Exception(f"{DIGIT}桁で入力してください。")
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
        return id

    def input_bool(self, msg: str) -> bool:
        """はいかいいえをを入力します"""
        while True:
            try:
                result = False
                cancel = False
                error = False
                str_of_bool = input(f"{msg}\n(Yes => y or No => n): ").strip()
                match str_of_bool:
                    case var if var in self.dct_of_bool["yes"]:
                        result = True
                    case var if var in self.dct_of_bool["no"]:
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


async def main() -> bool:
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
            obj_with_cui = GS_With_Cui()
            obj_of_cls = GetGovernmentStatistics(obj_of_lt.logger)
            obj_of_cls.lst_of_data_type = obj_with_cui.select_target(obj_with_cui.dct_of_data_type)
            async for page in obj_of_cls.get_stats_data_ids():
                for stat_id, info in page.items():
                    match obj_of_cls.lst_of_data_type[obj_of_cls.KEY]:
                        case "xml":
                            stat_name = info.get("stat_name", "")
                            title = info.get("title", "")
                            obj_of_cls.log.info(f"{stat_id}, {stat_name}, {title}")
                        case "json":
                            statistics_name = info.get("statistics_name", "")
                            title = info.get("title", "")
                            obj_of_cls.log.info(f"{stat_id}, {statistics_name}, {title}")
                        case "csv":
                            stat_name = info.get("stat_name", "")
                            category = info.get("category", "")
                            obj_of_cls.log.info(f"{stat_id}, {stat_name}, {category}")
                        case _:
                            raise Exception("データタイプが対応していません。")
            df = None
            filtered_df = None
            obj_of_cls.STATS_DATA_ID = obj_with_cui.input_stats_data_id()
            df = obj_of_cls.get_data_from_api()
            obj_of_cls.lst_of_match = obj_with_cui.select_target(obj_with_cui.dct_of_match)
            obj_of_cls.lst_of_keyword = obj_with_cui.input_list_of_keyword("抽出するキーワードを入力してください。")
            if len(obj_of_cls.lst_of_keyword) > 1:
                obj_of_cls.lst_of_logic = obj_with_cui.select_target(obj_with_cui.dct_of_logic)
            filtered_df = obj_of_cls.filter_data(df)
            obj_of_cls.order = str(obj_with_cui.select_target(obj_with_cui.lst_of_order))
            if not obj_of_cls.show_data(filtered_df):
                raise
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
    asyncio.run(main())
