import sys
from pathlib import Path

from source.common.common import PathTools
from source.get_nhk_news.gn2_class import GetNHKNews


class GN2_With_Cui:
    def __init__(self):
        """初期化します"""
        self.d_of_bool = {
            "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
            "no": ["いいえ", "0", "No", "no", "N", "n"],
        }

    def select_genre(self, d: dict) -> list:
        """ジャンルを選択します"""
        while True:
            try:
                result = False
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
                print(str(e))
            except KeyboardInterrupt:
                sys.exit(0)
            else:
                result = True
            finally:
                if result:
                    break
        return [choice, genres[choice - 1]]

    def input_bool(self, msg: str) -> bool:
        """はいかいいえをを入力します"""
        while True:
            try:
                result = False
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
                print(str(e))
            except KeyboardInterrupt:
                sys.exit(0)
            else:
                pass
            finally:
                if not error:
                    break
        return result


def main() -> bool:
    """主要関数"""
    while True:
        try:
            result = False
            # pytest
            test_var = input("test_input\n(Yes => y or No => n): ").strip()
            if test_var == "y":
                raise Exception("test_input")
            obj_of_pt = PathTools()
            obj_with_cui = GN2_With_Cui()
            obj_of_cls = GetNHKNews()
            print(*obj_of_cls.log, sep="\n")
            num_of_genre, key_of_genre = obj_with_cui.select_genre(obj_of_cls.rss_feeds)
            obj_of_cls.parse_rss(num_of_genre, key_of_genre)
            obj_of_cls.get_standard_time_and_today(obj_of_cls.TIMEZONE_OF_JAPAN)
            obj_of_cls.extract_news_of_today_from_standard_time()
            _, log = obj_of_cls.get_news(obj_of_cls.NUM_OF_NEWS_TO_SHOW)
            print(*log, sep="\n")
        except Exception as e:
            print(f"処理が失敗しました。: {str(e)}")
        else:
            result = True
            print("処理が成功しました。")
            file_of_exe_as_path_type = Path(__file__)
            file_of_log_as_path_type = obj_of_pt.get_file_path_of_log(file_of_exe_as_path_type)
            _, _ = obj_of_cls.write_log(file_of_log_as_path_type)
        finally:
            # pytest
            if test_var == "y":
                break
            if obj_with_cui.input_bool("終了しますか？"):
                break
            else:
                continue
    return result


if __name__ == "__main__":
    main()
