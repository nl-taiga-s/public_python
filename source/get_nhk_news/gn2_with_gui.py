import platform
import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QFontDatabase
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from source.common.common import DatetimeTools, PathTools, PlatformTools
from source.get_nhk_news.gn2_class import GetNHKNews


class NHKNewsApp(QWidget):
    def __init__(self):
        super().__init__()
        self.obj_of_pft = PlatformTools()
        self.obj_of_dt2 = DatetimeTools()
        self.obj_of_pt = PathTools()
        self.news_obj = GetNHKNews()
        # WSL-Ubuntuでフォント設定
        if self.obj_of_pft.is_wsl():
            font_path = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"
            font_id = QFontDatabase.addApplicationFont(font_path)
            font_family = QFontDatabase.applicationFontFamilies(font_id)[0]
            font = QFont(font_family)
            self.setFont(font)
        self.setWindowTitle("NHKニュース取得アプリ")

        self.setup_ui()

    def setup_ui(self):
        # ウィジェット作成
        layout = QVBoxLayout()
        genre_layout = QHBoxLayout()
        genre_label = QLabel("ジャンル:")
        self.genre_combo = QComboBox()
        self.fetch_button = QPushButton("ニュース取得")
        self.save_button = QPushButton("テキストに保存")
        self.news_list = QListWidget()

        # レイアウト
        self.genre_combo.addItems(self.news_obj.rss_feeds.keys())
        genre_layout.addWidget(genre_label)
        genre_layout.addWidget(self.genre_combo)
        layout.addLayout(genre_layout)
        layout.addWidget(self.fetch_button)
        layout.addWidget(self.save_button)
        layout.addWidget(self.news_list)
        self.setLayout(layout)

        # シグナル接続
        self.fetch_button.clicked.connect(self.fetch_news)
        self.save_button.clicked.connect(self.save_news_to_file)
        self.news_list.itemClicked.connect(self.open_news_link)

    def fetch_news(self):
        self.news_list.clear()
        genre_key = self.genre_combo.currentText()

        try:
            genre_index = list(self.news_obj.rss_feeds.keys()).index(genre_key)
            self.news_obj.parse_rss(genre_index, genre_key)
            self.news_obj.get_standard_time_and_today(self.news_obj.TIMEZONE_OF_JAPAN)
            self.news_obj.extract_news_of_today_from_standard_time()

            if not self.news_obj.today_news:
                QMessageBox.information(
                    self, "情報", "今日のニュースはまだありません。"
                )
                return
            for news in self.news_obj.today_news[: self.news_obj.NUM_OF_NEWS_TO_SHOW]:
                title = news.title
                summary = (
                    (news.summary or "").splitlines()[0]
                    if hasattr(news, "summary")
                    else ""
                )

                # QListWidgetItem + カスタムWidgetのセット
                item = QListWidgetItem()
                widget = NewsItemWidget(title, summary)

                # アイテムにURL情報をセット
                item.setData(Qt.UserRole, news.link)
                item.setSizeHint(widget.sizeHint())

                self.news_list.addItem(item)
                self.news_list.setItemWidget(item, widget)

        except Exception as e:
            QMessageBox.critical(self, "エラー", f"ニュースの取得に失敗しました:\n{e}")

    def save_news_to_file(self):
        if not self.news_obj.today_news:
            QMessageBox.information(self, "情報", "保存するニュースがありません。")
            return

        # exe 実行時とスクリプト実行時に対応した保存先
        if getattr(sys, 'frozen', False):
            fp_e = Path(sys.executable)
        else:
            fp_e = Path(__file__)

        fp_l = self.obj_of_pt.get_file_path_of_log(fp_e)
        file_path_of_log = str(fp_l)

        try:
            with open(file_path_of_log, "w", encoding="utf-8", newline="") as f:
                f.write(f"<<<日付: {self.news_obj.today}>>>\n")
                f.write(f"<<<ジャンル: {self.genre_combo.currentText()}>>>\n\n")
                for i, news in enumerate(
                    self.news_obj.today_news[: self.news_obj.NUM_OF_NEWS_TO_SHOW],
                    start=1,
                ):
                    f.write(f"{i}. {news.title}\n")
                    f.write(f"{news.link}\n\n")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"ファイルの保存に失敗しました:\n{e}")
        else:
            QMessageBox.information(
                self, "成功", f"ニュースを保存しました:\n{file_path_of_log}"
            )

    def open_news_link(self, item: QListWidgetItem):
        url = item.data(Qt.UserRole)
        if not url:
            return

        try:
            system_name = platform.system()
            if system_name == "Windows" or self.obj_of_pft.is_wsl():
                subprocess.run(['powershell.exe', 'Start-Process', url], check=True)
        except Exception as e:
            QMessageBox.warning(self, "警告", f"ブラウザを開くのに失敗しました:\n{e}")


class NewsItemWidget(QWidget):
    def __init__(self, title: str, summary: str):
        super().__init__()
        layout = QVBoxLayout()
        self.title_label = QLabel(title)
        self.title_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        self.summary_label = QLabel(summary)
        self.summary_label.setStyleSheet("color: gray; font-size: 12px;")
        self.summary_label.setWordWrap(True)

        layout.addWidget(self.title_label)
        layout.addWidget(self.summary_label)
        layout.setContentsMargins(5, 5, 5, 5)
        self.setLayout(layout)


def main():
    app = QApplication(sys.argv)
    window = NHKNewsApp()
    window.resize(600, 400)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
