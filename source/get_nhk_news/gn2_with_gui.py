import platform
import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
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

from source.common.common import GUITools, LogTools, PathTools, PlatformTools
from source.get_nhk_news.gn2_class import GetNHKNews


class MainApp_Of_GN2(QWidget):
    def __init__(self):
        """初期化します"""
        super().__init__()
        self.obj_of_lt = LogTools()
        self.obj_of_cls = GetNHKNews(self.obj_of_lt.logger)
        self.setup_ui()
        self.obj_of_pft = PlatformTools()
        self.obj_of_pt = PathTools()
        self.setup_log()

    def closeEvent(self, event):
        """終了します"""
        if self.obj_of_lt:
            self.show_info(f"ログファイルは、\n{self.obj_of_lt.file_path_of_log}\nに出力されました。")
        super().closeEvent(event)

    def show_info(self, msg: str):
        """情報を表示します"""
        QMessageBox.information(self, "情報", msg)

    def show_result(self, label: str, success: bool):
        """結果を表示します"""
        QMessageBox.information(self, "結果", f"{label}に{'成功' if success else '失敗'}しました。")

    def show_error(self, msg: str):
        """エラーを表示します"""
        QMessageBox.critical(self, "エラー", msg)

    def setup_log(self) -> bool:
        """ログを設定する"""
        try:
            result = False
            # exe化されている場合とそれ以外を切り分ける
            if getattr(sys, "frozen", False):
                exe_path = Path(sys.executable)
            else:
                exe_path = Path(__file__)
            file_of_log_as_path_type = self.obj_of_pt.get_file_path_of_log(exe_path)
            self.obj_of_lt.file_path_of_log = str(file_of_log_as_path_type)
            if not self.obj_of_lt.setup_file_handler(self.obj_of_lt.file_path_of_log):
                raise
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            return result

    def setup_ui(self) -> bool:
        """User Interfaceを設定します"""
        try:
            result = False
            # タイトル
            self.setWindowTitle("NHKニュース表示アプリ")
            # ウィジェット
            layout = QVBoxLayout()
            genre_layout = QHBoxLayout()
            genre_label = QLabel("ジャンル:")
            self.genre_combo = QComboBox()
            self.fetch_button = QPushButton("ニュースを取得")
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
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            return result

    def fetch_news(self) -> bool:
        """ニュースを取りに行く"""
        self.news_list.clear()
        genre_key = self.genre_combo.currentText()
        try:
            result = False
            genre_index = list(self.obj_of_cls.rss_feeds.keys()).index(genre_key)
            if not self.obj_of_cls.parse_rss(genre_index, genre_key):
                raise
            if not self.obj_of_cls.get_standard_time_and_today(self.obj_of_cls.TIMEZONE_OF_JAPAN):
                raise
            if not self.obj_of_cls.extract_news_of_today_from_standard_time():
                raise
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
            if not self.obj_of_cls.get_news(self.obj_of_cls.NUM_OF_NEWS_TO_SHOW):
                raise
            if not self.obj_of_cls.today_news:
                raise Exception("今日のニュースはまだありません。")
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
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
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            return result


class NewsItem_Of_GN2(QWidget):
    def __init__(self, title: str, summary: str):
        """初期化します"""
        super().__init__()
        self.setup_ui(title, summary)

    def setup_ui(self, title: str, summary: str) -> bool:
        """User Interfaceを設定します"""
        try:
            result = False
            # ウィジェット
            self.title_label = QLabel(title)
            self.title_label.setStyleSheet("font-weight: bold; font-size: 18px;")
            self.summary_label = QLabel(summary)
            self.summary_label.setStyleSheet("color: gray; font-size: 16px;")
            self.summary_label.setWordWrap(True)
            # レイアウト
            layout = QVBoxLayout()
            layout.addWidget(self.title_label)
            layout.addWidget(self.summary_label)
            layout.setContentsMargins(5, 5, 5, 5)
            self.setLayout(layout)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            return result


def main() -> bool:
    """主要関数"""
    try:
        result = False
        obj_of_gt = GUITools()
        app = QApplication(sys.argv)
        # アプリ単位でフォントを設定する
        font = QFont()
        font.setPointSize(12)
        app.setFont(font)
        window = MainApp_Of_GN2()
        window.resize(1000, 800)
        # 最大化して、表示させる
        window.showMaximized()
        sys.exit(app.exec())
    except Exception as e:
        obj_of_gt.show_error(f"error: \n{str(e)}")
    else:
        result = True
    finally:
        return result


if __name__ == "__main__":
    main()
