from datetime import datetime, timedelta, timezone
from pathlib import Path

import feedparser


class GetNHKNews:
    """NHKニュースを取得します"""

    def __init__(self):
        """初期化します"""
        self.log = []
        self.REPEAT_TIMES = 50
        self.log.append(self.__class__.__doc__)
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
                pub_date = datetime(*entry.published_parsed[:6], tzinfo=timezone.utc).astimezone(self.standard_time).date()
                if pub_date == self.today:
                    self.today_news.append(entry)

    def get_news(self, num_of_news: int) -> list:
        """ニュースを取得します"""
        try:
            result = False
            local_log = []
            local_log.append(">" * self.REPEAT_TIMES)
            local_log.append(f"日付: {self.today}")
            local_log.append(f"ジャンル: {self.key_of_genre}")
            if not self.today_news:
                raise ValueError
            else:
                for i, news in enumerate(self.today_news[:num_of_news], start=1):
                    local_log.append(f"{i}. {news.title}: ")
                    local_log.append(f"\t{news.link}")
        except ValueError:
            local_log.append("***ニュースは、まだありません。***")
        except Exception:
            local_log.append("***ニュースの出力に失敗しました。***")
        else:
            result = True
            local_log.append("***ニュースの出力に成功しました。***")
        finally:
            local_log.append("<" * self.REPEAT_TIMES)
            self.log.extend(local_log)
            return [result, local_log]

    def write_log(self, file_of_log_as_path_type: Path) -> list:
        """処理結果をログに書き出す"""
        file_of_log_as_str_type = str(file_of_log_as_path_type)
        try:
            result = False
            fp = ""
            with open(file_of_log_as_str_type, "w", encoding="utf-8", newline="") as f:
                f.write("\n".join(self.log))
        except Exception:
            pass
        else:
            result = True
            fp = file_of_log_as_str_type
        finally:
            return [result, fp]
