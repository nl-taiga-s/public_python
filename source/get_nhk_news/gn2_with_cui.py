from pathlib import Path

from source.common.common import LogTools, PathTools
from source.get_nhk_news.gn2_class import GetNHKNews


class GN2_With_Cui:
    def __init__(self):
        """初期化します"""
        self.dct_of_bool = {
            "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
            "no": ["いいえ", "0", "No", "no", "N", "n"],
        }

    def select_genre(self, d: dict) -> list:
        """ジャンルを選択します"""
        while True:
            try:
                result = False
                cancel = False
                print("ジャンルの一覧")
                genres = list(d.keys())
                for i, genre in enumerate(genres, 1):
                    print(f"{i}: {genre}")
                choice = input("番号を入力してください。: ").strip()
                if choice == "":
                    raise Exception("番号が未入力です。")
                if not choice.isdecimal():
                    raise Exception("数字を入力してください。")
                choice = int(choice)
                if 1 > choice or choice > len(genres):
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
        return [choice, genres[choice - 1]]

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
            obj_with_cui = GN2_With_Cui()
            obj_of_cls = GetNHKNews(obj_of_lt.logger)
            num_of_genre, key_of_genre = obj_with_cui.select_genre(obj_of_cls.rss_feeds)
            if not obj_of_cls.parse_rss(num_of_genre, key_of_genre):
                raise
            if not obj_of_cls.get_standard_time_and_today(obj_of_cls.TIMEZONE_OF_JAPAN):
                raise
            if not obj_of_cls.extract_news_of_today_from_standard_time():
                raise
            if not obj_of_cls.get_news(obj_of_cls.NUM_OF_NEWS_TO_SHOW):
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
    main()
