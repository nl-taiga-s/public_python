import os
import platform
import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import Qt
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

    def closeEvent(self, event):
        """終了します"""
        self.write_log()
        super().closeEvent(event)

    def setup_ui(self):
        """User Interfaceを設定します"""
        # タイトル
        self.setWindowTitle("NHKニュース表示アプリ")
        # ウィジェット
        layout = QVBoxLayout()
        genre_layout = QHBoxLayout()
        genre_label = QLabel("ジャンル:")
        self.genre_combo = QComboBox()
        self.fetch_button = QPushButton("ニュース取得")
        self.news_list = QListWidget()
        # レイアウト
        self.genre_combo.addItems(self.obj_of_cls.rss_feeds.keys())
        genre_layout.addWidget(genre_label)
        genre_layout.addWidget(self.genre_combo)
        layout.addLayout(genre_layout)
        layout.addWidget(self.fetch_button)
        layout.addWidget(self.news_list)
        self.setLayout(layout)
        # シグナル接続
        self.fetch_button.clicked.connect(self.fetch_news)
        self.news_list.itemClicked.connect(self.open_news_link)

    def fetch_news(self) -> bool:
        """ニュースを取りに行く"""
        self.news_list.clear()
        genre_key = self.genre_combo.currentText()
        try:
            result = False
            genre_index = list(self.obj_of_cls.rss_feeds.keys()).index(genre_key)
            self.obj_of_cls.parse_rss(genre_index, genre_key)
            self.obj_of_cls.get_standard_time_and_today(self.obj_of_cls.TIMEZONE_OF_JAPAN)
            self.obj_of_cls.extract_news_of_today_from_standard_time()
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
            _, _ = self.obj_of_cls.get_news(self.obj_of_cls.NUM_OF_NEWS_TO_SHOW)
            if not self.obj_of_cls.today_news:
                raise Exception("今日のニュースはまだありません。")
        except Exception as e:
            self.show_error(str(e))
        else:
            result = True
        finally:
            return result

    def open_news_link(self, item: QListWidgetItem) -> bool:
        """ニュースのリンクを開きます"""
        POWERSHELL_OF_WSL = "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe"
        POWERSHELL_OF_WINDOWS = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe"
        try:
            result = False
            url = item.data(Qt.UserRole)
            if not url:
                raise Exception
            if platform.system().lower() == "windows":
                subprocess.run([POWERSHELL_OF_WINDOWS, "Start-Process", url], check=True)
            elif self.obj_of_pft.is_wsl():
                subprocess.run([POWERSHELL_OF_WSL, "Start-Process", url], check=True)
        except Exception as e:
            self.show_error(str(e))
        else:
            result = True
        finally:
            return result

    def write_log(self):
        """ログを書き出す"""
        # exe化されている場合とそれ以外を切り分ける
        if getattr(sys, "frozen", False):
            exe_path = Path(sys.executable)
        else:
            exe_path = Path(__file__)
        file_of_log_as_path_type = self.obj_of_pt.get_file_path_of_log(exe_path)
        result, path = self.obj_of_cls.write_log(file_of_log_as_path_type)
        self.show_result(f"ログファイル: \n{path}\nの出力", result)

    def show_result(self, label: str, success: bool):
        """結果を表示します"""
        QMessageBox.information(self, "結果", f"{label}に{'成功' if success else '失敗'}しました。")

    def show_error(self, msg: str):
        """エラーを表示します"""
        QMessageBox.information(self, "エラー", msg)


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
    if platform.system().lower() == "windows":
        # アプリ全体のスケール
        os.environ["QT_SCALE_FACTOR"] = "0.7"
    app = QApplication(sys.argv)
    window = MainApp_Of_GN2()
    window.resize(600, 400)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
