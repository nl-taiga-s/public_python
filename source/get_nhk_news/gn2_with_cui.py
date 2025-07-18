import os
from pathlib import Path

from source.common.common import PathTools
from source.get_nhk_news.gn2_class import GetNHKNews


class GN2:
    def __init__(self):
        """初期化します"""
        self.obj_of_pt = PathTools()

    def select_genre(self, d: dict) -> list:
        """ジャンルを選択します"""
        try:
            while True:
                print("ジャンルを選択してください。")
                genres = list(d.keys())
                for i, genre in enumerate(genres):
                    print(f"{i}: {genre}")
                choice = int(input("番号を入力してください。: "))
                return choice, genres[choice]
        except (ValueError, IndexError):
            print("無効な入力です。")
        except KeyboardInterrupt:
            os._exit(0)

    def main(self):
        """主要関数"""
        self.obj_of_cls = GetNHKNews()
        # ジャンルを選択します
        num_of_genre, key_of_genre = self.select_genre(self.obj_of_cls.rss_feeds)
        self.obj_of_cls.parse_rss(num_of_genre, key_of_genre)
        self.obj_of_cls.get_standard_time_and_today(self.obj_of_cls.TIMEZONE_OF_JAPAN)
        self.obj_of_cls.extract_news_of_today_from_standard_time()
        self.obj_of_cls.print_specified_number_of_news(self.obj_of_cls.NUM_OF_NEWS_TO_SHOW)
        file_of_exe_as_path_type = Path(__file__)
        file_of_log_as_path_type = self.obj_of_pt.get_file_path_of_log(file_of_exe_as_path_type)
        self.obj_of_cls.write_log(file_of_log_as_path_type)


if __name__ == "__main__":
    obj_with_cui = GN2()
    obj_with_cui.main()
