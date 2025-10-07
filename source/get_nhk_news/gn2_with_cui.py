from pathlib import Path

from source.common.common import LogTools, PathTools
from source.get_nhk_news.gn2_class import GetNHKNews


class GN2_With_Cui:
    def __init__(self):
        """初期化します"""
        self.binary_choices: dict = {
            "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
            "no": ["いいえ", "0", "No", "no", "N", "n"],
        }

    def select_genre(self, dct: dict) -> list:
        """ジャンルを選択します"""
        while True:
            try:
                genres: list = list(dct.keys())
                for i, genre in enumerate(genres, 1):
                    print(f"{i}: {genre}")
                choice: str = input("番号を入力してください。: ").strip()
                if choice == "":
                    raise Exception("番号が未入力です。")
                if not choice.isdecimal():
                    raise Exception("数字を入力してください。")
                num: int = int(choice)
                if 1 > num or num > len(genres):
                    raise Exception("入力した番号が範囲外です。")
            except Exception as e:
                print(f"error: \n{str(e)}")
            except KeyboardInterrupt:
                raise
            else:
                break
            finally:
                pass
        return [num, genres[num - 1]]

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
        result = obj_of_lt.setup_file_handler(obj_of_lt.file_path_of_log)
        result = obj_of_lt.setup_stream_handler()
    except Exception as e:
        print(f"error: \n{str(e)}")
        return result
    else:
        pass
    finally:
        pass
    # 処理の本体部分
    while True:
        result = False
        try:
            obj_with_cui: GN2_With_Cui = GN2_With_Cui()
            obj_of_cls: GetNHKNews = GetNHKNews(obj_of_lt.logger)
            obj_of_cls.num_of_genre, obj_of_cls.key_of_genre = obj_with_cui.select_genre(obj_of_cls.rss_feeds)
            result = obj_of_cls.parse_rss()
            result = obj_of_cls.get_standard_time_and_today()
            result = obj_of_cls.extract_news_of_today_from_standard_time()
            result = obj_of_cls.get_news()
        except Exception:
            print("***処理が失敗しました。***")
        except KeyboardInterrupt:
            raise
        else:
            print("***処理が成功しました。***")
        finally:
            if obj_with_cui.input_bool("終了しますか？"):
                break
            else:
                continue
    return result


if __name__ == "__main__":
    main()
