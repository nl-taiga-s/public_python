import sys
import os
import platform
from PySide6.QtGui import QFont, QFontDatabase
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLineEdit, QListWidget, QFileDialog, QLabel
)
from source.search_file.sf_class import GetFileList


class FileSearchApp(QWidget):
    def __init__(self):
        super().__init__()
        # WSL-Ubuntuでフォント設定
        if platform.system() == "Linux":
            # install ipafont-gothic
            font_path = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"
            font_id = QFontDatabase.addApplicationFont(font_path)
            font_family = QFontDatabase.applicationFontFamilies(font_id)[0]
            font = QFont(font_family)
            self.setFont(font)
        self.setWindowTitle("ファイル検索ツール")

        self.file_list_obj = None

        # ウィジェット作成
        self.folder_label = QLabel("フォルダ未選択")
        self.select_folder_btn = QPushButton("フォルダを選択")
        self.pattern_input = QLineEdit()
        self.pattern_input.setPlaceholderText("検索パターンを入力...")
        self.search_btn = QPushButton("検索実行")
        self.result_list = QListWidget()

        # レイアウト
        layout = QVBoxLayout()
        layout.addWidget(self.folder_label)
        layout.addWidget(self.select_folder_btn)

        layout.addWidget(QLabel("検索パターン:"))
        layout.addWidget(self.pattern_input)
        layout.addWidget(self.search_btn)

        layout.addWidget(QLabel("検索結果:"))
        layout.addWidget(self.result_list)

        self.setLayout(layout)

        # シグナル接続
        self.select_folder_btn.clicked.connect(self.select_folder)
        self.search_btn.clicked.connect(self.search_files)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "フォルダを選択")
        if folder:
            self.folder_label.setText(folder)
            self.file_list_obj = GetFileList(folder)

    def search_files(self):
        if not self.file_list_obj:
            self.result_list.clear()
            self.result_list.addItem("フォルダが未選択です。")
            return

        pattern = self.pattern_input.text().strip()
        if not pattern:
            self.result_list.clear()
            self.result_list.addItem("検索パターンが空です。")
            return

        self.file_list_obj.extract_by_pattern(pattern)
        self.result_list.clear()
        if self.file_list_obj.list_file_after:
            self.result_list.addItems(self.file_list_obj.list_file_after)
        else:
            self.result_list.addItem("一致するファイルが見つかりませんでした。")

def main():
    app = QApplication(sys.argv)
    window = FileSearchApp()
    window.resize(600, 400)
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
