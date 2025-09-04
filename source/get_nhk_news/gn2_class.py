from datetime import datetime, timedelta, timezone
from logging import Logger

import feedparser


class GetNHKNews:
    """NHKニュースを取得します"""

    def __init__(self, logger: Logger):
        """初期化します"""
        self.log = logger
        self.log.info(self.__class__.__doc__)
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
        self.TIMEZONE_OF_JAPAN = 9.0
        # 表示するニュースの数
        self.NUM_OF_NEWS_TO_SHOW = 10

    def parse_rss(self, num_of_genre: int, key_of_genre: str) -> bool:
        """RSSを解析します"""
        try:
            result = False
            self.num_of_genre = num_of_genre
            self.key_of_genre = key_of_genre
            self.genre = list(self.rss_feeds.keys())[self.num_of_genre]
            self.url_of_rss = self.rss_feeds[self.key_of_genre]
            self.feed = feedparser.parse(self.url_of_rss)
        except Exception as e:
            self.log.error(f"***{self.parse_rss.__doc__} => 失敗しました。***")
            self.log.error(f"error: \n{str(e)}")
        else:
            result = True
            self.log.info(f"***{self.parse_rss.__doc__} => 成功しました。***")
        finally:
            return result

    def get_standard_time_and_today(self, time_difference: float) -> bool:
        """指定の標準時と今日の日付を取得します"""
        try:
            result = False
            # 指定の標準時
            self.standard_time = timezone(timedelta(hours=time_difference))
            # 今日の日付
            self.today = datetime.now(self.standard_time).date()
        except Exception as e:
            self.log.error(f"***{self.get_standard_time_and_today.__doc__} => 失敗しました。***")
            self.log.error(f"error: \n{str(e)}")
        else:
            result = True
            self.log.info(f"***{self.get_standard_time_and_today.__doc__} => 成功しました。***")
        finally:
            return result

    def extract_news_of_today_from_standard_time(self) -> bool:
        """指定の標準時の今日のニュースを抽出します"""
        try:
            result = False
            self.today_news = []
            for entry in self.feed.entries:
                if hasattr(entry, "published_parsed"):
                    pub_date = datetime(*entry.published_parsed[:6], tzinfo=timezone.utc).astimezone(self.standard_time).date()
                    if pub_date == self.today:
                        self.today_news.append(entry)
        except Exception as e:
            self.log.error(f"***{self.extract_news_of_today_from_standard_time.__doc__} => 失敗しました。***")
            self.log.error(f"error: \n{str(e)}")
        else:
            result = True
            self.log.info(f"***{self.extract_news_of_today_from_standard_time.__doc__} => 成功しました。***")
        finally:
            return result

    def get_news(self, num_of_news: int) -> bool:
        """ニュースを取得します"""
        try:
            result = False
            self.log.info(f"* 日付: {self.today}")
            self.log.info(f"* ジャンル: {self.key_of_genre}")
            if not self.today_news:
                raise Exception("ニュースは、まだありません。")
            else:
                for i, news in enumerate(self.today_news[:num_of_news], start=1):
                    self.log.info(f"{i}. {news.title}: ")
                    self.log.info(f"\t{news.link}")
        except Exception as e:
            self.log.error(f"***{self.get_news.__doc__} => 失敗しました。***")
            self.log.error(f"error: \n{str(e)}")
        else:
            result = True
            self.log.info(f"***{self.get_news.__doc__} => 成功しました。***")
        finally:
            return result
