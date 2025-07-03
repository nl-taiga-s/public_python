from gn2_class import GetNHKNews

from source.common.common import PathTools

d_of_bool = {
    "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
    "no": ["いいえ", "0", "No", "no", "N", "n"],
}


def select_genre(d: dict) -> list:
    """ジャンルを選択します"""
    while True:
        print("ジャンルを選択してください。")
        genres = list(d.keys())
        for i, genre in enumerate(genres):
            print(f"{i}: {genre}")
        try:
            choice = int(input("番号を入力してください。: "))
            return choice, genres[choice]
        except (ValueError, IndexError):
            print("無効な入力です。")


def input_bool_of_result() -> bool:
    """resultファイルを出力するかどうかを入力します"""
    while True:
        str_of_l = input(
            "resultファイルを出力するかどうかを入力してください。"
            "(Yes => y or No => n): "
        )
        match str_of_l:
            case var if var in d_of_bool["yes"]:
                return True
            case var if var in d_of_bool["no"]:
                return False
            case _:
                raise ValueError("無効な入力です。: {str_or_l}")


def main():
    try:
        obj_of_gn2 = GetNHKNews()
        obj_of_pt = PathTools()
        # ジャンルを選択します
        num_of_genre, key_of_genre = select_genre(obj_of_gn2.rss_feeds)
        obj_of_gn2.parse_rss(num_of_genre, key_of_genre)
        obj_of_gn2.get_standard_time_and_today(obj_of_gn2.TIMEZONE_OF_JAPAN)
        obj_of_gn2.extract_news_of_today_from_standard_time()
        obj_of_gn2.print_specified_number_of_news(obj_of_gn2.NUM_OF_NEWS_TO_SHOW)
        if input_bool_of_result():
            file_path_of_result = obj_of_pt.get_file_path_of_result(__file__)
            file_path_of_result = obj_of_pt.convert_path_to_str(file_path_of_result)
            try:
                with open(file_path_of_result, "w", encoding="utf-8", newline="") as f:
                    f.write(f"<<<日付: {obj_of_gn2.today}>>>\n")
                    f.write(f"<<<ジャンル: {key_of_genre}>>>\n\n")
                    for i, news in enumerate(
                        obj_of_gn2.today_news[: obj_of_gn2.NUM_OF_NEWS_TO_SHOW],
                        start=1,
                    ):
                        f.write(f"{i}. {news.title}\n")
                        f.write(f"{news.link}\n\n")
            except Exception as e:
                print(e)
            print(f"resultファイルを出力しました。: {file_path_of_result}")
    except Exception as e:
        print(e)


if __name__ == "__main__":
    main()
