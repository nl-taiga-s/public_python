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


class MainApp_Of_GN2(QWidget):
    def __init__(self):
        """初期化します"""
        super().__init__()
        self.obj_of_pft = PlatformTools()
        self.obj_of_dt2 = DatetimeTools()
        self.obj_of_pt = PathTools()
        self.obj_of_cls = GetNHKNews()
        self.setup_ui()

    def setup_ui(self):
        """User Interfaceを設定します"""
        # WSL-Ubuntuでフォント設定
        if self.obj_of_pft.is_wsl():
            font_path = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"
            font_id = QFontDatabase.addApplicationFont(font_path)
            font_family = QFontDatabase.applicationFontFamilies(font_id)[0]
            font = QFont(font_family)
            self.setFont(font)
        # タイトル
        self.setWindowTitle("NHKニュース取得アプリ")
        # ウィジェット
        layout = QVBoxLayout()
        genre_layout = QHBoxLayout()
        genre_label = QLabel("ジャンル:")
        self.genre_combo = QComboBox()
        self.fetch_button = QPushButton("ニュース取得")
        self.save_button = QPushButton("テキストに保存")
        self.news_list = QListWidget()
        # レイアウト
        self.genre_combo.addItems(self.obj_of_cls.rss_feeds.keys())
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
        """ニュースを取りに行きます"""
        self.news_list.clear()
        genre_key = self.genre_combo.currentText()
        try:
            genre_index = list(self.obj_of_cls.rss_feeds.keys()).index(genre_key)
            self.obj_of_cls.parse_rss(genre_index, genre_key)
            self.obj_of_cls.get_standard_time_and_today(self.obj_of_cls.TIMEZONE_OF_JAPAN)
            self.obj_of_cls.extract_news_of_today_from_standard_time()
            if not self.obj_of_cls.today_news:
                QMessageBox.information(self, "情報", "今日のニュースはまだありません。")
                return
            for news in self.obj_of_cls.today_news[: self.obj_of_cls.NUM_OF_NEWS_TO_SHOW]:
                title = news.title
                summary = (news.summary or "").splitlines()[0] if hasattr(news, "summary") else ""
                # QListWidgetItem + カスタムWidgetのセット
                item = QListWidgetItem()
                widget = NewsItem_Of_GN2(title, summary)
                # アイテムにURL情報をセット
                item.setData(Qt.UserRole, news.link)
                item.setSizeHint(widget.sizeHint())
                self.news_list.addItem(item)
                self.news_list.setItemWidget(item, widget)
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"ニュースの取得に失敗しました。: \n{e}")

    def save_news_to_file(self):
        """ニュースを保存します"""
        if not self.obj_of_cls.today_news:
            QMessageBox.information(self, "情報", "保存するニュースがありません。")
            return
        try:
            file_of_exe_as_path_type = Path(__file__)
            file_of_log_as_path_type = self.obj_of_pt.get_file_path_of_log(file_of_exe_as_path_type)
            self.obj_of_cls.write_log(file_of_log_as_path_type)
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"ニュースの保存に失敗しました。: \n{e}")
        else:
            QMessageBox.information(self, "成功", f"ニュースの保存に成功しました。: \n{str(file_of_log_as_path_type)}")

    def open_news_link(self, item: QListWidgetItem):
        """ニュースのリンクを開きます"""
        url = item.data(Qt.UserRole)
        if not url:
            return
        try:
            if platform.system().lower() == "windows" or self.obj_of_pft.is_wsl():
                subprocess.run(['powershell.exe', 'Start-Process', url], check=True)
        except Exception as e:
            QMessageBox.warning(self, "警告", f"ブラウザを開くのに失敗しました。: \n{e}")


class NewsItem_Of_GN2(QWidget):
    def __init__(self, title: str, summary: str):
        """初期化します"""
        super().__init__()
        self.setup_ui(title, summary)

    def setup_ui(self, title: str, summary: str):
        """User Interfaceを設定します"""
        # ウィジェット
        self.title_label = QLabel(title)
        self.title_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        self.summary_label = QLabel(summary)
        self.summary_label.setStyleSheet("color: gray; font-size: 12px;")
        self.summary_label.setWordWrap(True)
        # レイアウト
        layout = QVBoxLayout()
        layout.addWidget(self.title_label)
        layout.addWidget(self.summary_label)
        layout.setContentsMargins(5, 5, 5, 5)
        self.setLayout(layout)


def main():
    """主要関数"""
    app = QApplication(sys.argv)
    window = MainApp_Of_GN2()
    window.resize(600, 400)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
