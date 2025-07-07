from pathlib import Path

from source.common.common import PathTools
from source.get_nhk_news.gn2_class import GetNHKNews


class GN2:
    def __init__(self):
        """初期化します"""
        self.d_of_bool = {
            "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
            "no": ["いいえ", "0", "No", "no", "N", "n"],
        }
        self.obj_of_pt = PathTools()

    def select_genre(self, d: dict) -> list:
        """ジャンルを選択します"""
        while True:
            print("ジャンルを選択してください。")
            genres = list(d.keys())
            for i, genre in enumerate(genres):
                print(f"{i}: {genre}")
            try:
                choice = int(input("番号を入力してください。: "))
            except (ValueError, IndexError):
                print("無効な入力です。")
            else:
                return choice, genres[choice]

    def input_bool_of_log(self) -> bool:
        """logファイルを出力するかどうかを入力します"""
        while True:
            str_of_l = input(
                "logファイルを出力するかどうかを入力してください。"
                "(Yes => y or No => n): "
            )
            match str_of_l:
                case var if var in self.d_of_bool["yes"]:
                    return True
                case var if var in self.d_of_bool["no"]:
                    return False
                case _:
                    raise ValueError("無効な入力です。: {str_or_l}")

    def output_log_file(self):
        """logファイルを出力します"""
        if self.input_bool_of_log():
            fp_e = Path(__file__)
            fp_l = self.obj_of_pt.get_file_path_of_log(fp_e)
            file_path_of_log = str(fp_l)
            try:
                with open(file_path_of_log, "w", encoding="utf-8", newline="") as f:
                    f.write(f"<<<日付: {self.obj_of_cls.today}>>>\n")
                    f.write(f"<<<ジャンル: {self.obj_of_cls.key_of_genre}>>>\n\n")
                    for i, news in enumerate(
                        self.obj_of_cls.today_news[
                            : self.obj_of_cls.NUM_OF_NEWS_TO_SHOW
                        ],
                        start=1,
                    ):
                        f.write(f"{i}. {news.title}\n")
                        f.write(f"{news.link}\n\n")
            except Exception as e:
                print(e)
            else:
                print(f"logファイルを出力しました。: {file_path_of_log}")

    def main(self):
        """メイン関数"""
        try:
            self.obj_of_cls = GetNHKNews()
            # ジャンルを選択します
            num_of_genre, key_of_genre = self.select_genre(self.obj_of_cls.rss_feeds)
            self.obj_of_cls.parse_rss(num_of_genre, key_of_genre)
            self.obj_of_cls.get_standard_time_and_today(
                self.obj_of_cls.TIMEZONE_OF_JAPAN
            )
            self.obj_of_cls.extract_news_of_today_from_standard_time()
            self.obj_of_cls.print_specified_number_of_news(
                self.obj_of_cls.NUM_OF_NEWS_TO_SHOW
            )
            self.output_log_file()
        except Exception as e:
            print(e)


if __name__ == "__main__":
    obj_with_cui = GN2()
    obj_with_cui.main()
