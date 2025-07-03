from datetime import datetime, timedelta, timezone

import feedparser


class GetNHKNews:
    """NHKニュースを取得します。"""

    def __init__(self):
        """初期化します"""
        print(self.__class__.__doc__)
        # NHKニュースのジャンルごとのRSS URL
        self.rss_feeds = {
            "主要": "https://www.nhk.or.jp/rss/news/cat0.xml",
            "社会": "https://www.nhk.or.jp/rss/news/cat1.xml",
            "科学・医療": "https://www.nhk.or.jp/rss/news/cat3.xml",
            "政治": "https://www.nhk.or.jp/rss/news/cat4.xml",
            "経済": "https://www.nhk.or.jp/rss/news/cat5.xml",
            "国際": "https://www.nhk.or.jp/rss/news/cat6.xml",
            "スポーツ": "https://www.nhk.or.jp/rss/news/cat7.xml",
            "文化・エンタメ": "https://www.nhk.or.jp/rss/news/cat2.xml",
            "LIVE": "https://www.nhk.or.jp/rss/news/cat-live.xml",
        }
        # 日本のタイムゾーン
        self.TIMEZONE_OF_JAPAN = 9
        # 表示するニュースの数
        self.NUM_OF_NEWS_TO_SHOW = 10

    def parse_rss(self, num_of_genre: int, key_of_genre: str):
        """RSSを解析します"""
        self.num_of_genre = num_of_genre
        self.key_of_genre = key_of_genre
        self.genre = list(self.rss_feeds.keys())[self.num_of_genre]
        self.url_of_rss = self.rss_feeds[self.key_of_genre]
        self.feed = feedparser.parse(self.url_of_rss)

    def get_standard_time_and_today(self, time_difference: float):
        """指定の標準時と今日の日付を取得します"""
        # 指定の標準時
        self.standard_time = timezone(timedelta(hours=time_difference))
        # 今日の日付
        self.today = datetime.now(self.standard_time).date()

    def extract_news_of_today_from_standard_time(self):
        """指定の標準時の今日のニュースを抽出します"""
        self.today_news = []
        for entry in self.feed.entries:
            if hasattr(entry, "published_parsed"):
                pub_date = (
                    datetime(*entry.published_parsed[:6], tzinfo=timezone.utc)
                    .astimezone(self.standard_time)
                    .date()
                )
                if pub_date == self.today:
                    self.today_news.append(entry)

    def print_specified_number_of_news(self, num_of_news: int):
        """上位指定の件数のニュースを出力します"""
        print(f"<<<日付: {self.today}>>>")
        print(f"<<<ジャンル: {self.key_of_genre}>>>")
        if not self.today_news:
            print("***ニュースは、まだありません。***")
        else:
            for i, news in enumerate(self.today_news[:num_of_news], start=1):
                print(f"{i}. {news.title}: ")
                print(f"\t{news.link}")
