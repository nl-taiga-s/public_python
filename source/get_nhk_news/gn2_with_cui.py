from gn2_class import GetNHKNews


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


def main():
    # 日本のタイムゾーン
    TIMEZONE_OF_JAPAN = 9
    # 上位3件
    SPECIFIED_NUMBER = 3
    try:
        obj_of_gn2 = GetNHKNews()
        num_of_genre, key_of_genre = select_genre(obj_of_gn2.rss_feeds)
        # 指定のジャンル
        obj_of_gn2.parse_rss(num_of_genre, key_of_genre)
        # 日本の標準時の今日の日付
        obj_of_gn2.get_standard_time_and_today(TIMEZONE_OF_JAPAN)
        obj_of_gn2.extract_news_of_today_from_standard_time()
        # 上位3件
        obj_of_gn2.print_specified_number_of_news(SPECIFIED_NUMBER)
    except Exception as e:
        print(e)


if __name__ == "__main__":
    main()
