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
        try:
            while True:
                print("ジャンルを選択してください。")
                genres = list(d.keys())
                for i, genre in enumerate(genres, 1):
                    print(f"{i}: {genre}")
                choice = input("番号を入力してください。: ")
                if choice == "":
                    return [None, None]
                elif choice.isdecimal():
                    choice = int(choice)
                    if 1 <= choice and choice <= len(genres):
                        return [choice, genres[choice - 1]]
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
            obj_of_pt = PathTools()
            obj_of_cls = GetNHKNews()
            obj_with_cui = GN2_With_Cui()
            # ジャンルを選択します
            num_of_genre, key_of_genre = obj_with_cui.select_genre(obj_of_cls.rss_feeds)
            if num_of_genre is None and key_of_genre is None:
                return result
            print(*obj_of_cls.log, sep="\n")
            obj_of_cls.parse_rss(num_of_genre, key_of_genre)
            obj_of_cls.get_standard_time_and_today(obj_of_cls.TIMEZONE_OF_JAPAN)
            obj_of_cls.extract_news_of_today_from_standard_time()
            _, log = obj_of_cls.get_news(obj_of_cls.NUM_OF_NEWS_TO_SHOW)
            print(*log, sep="\n")
            if obj_with_cui.input_continue():
                continue
            else:
                break
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
