import glob
import os
import sys

from PySide6.QtGui import QFont, QFontDatabase
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


class FileSearchApp(QWidget):
    def __init__(self):
        super().__init__()
        # WSL-Ubuntuでフォント設定
        if sys.platform == 'linux':
            font_path = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"
            font_id = QFontDatabase.addApplicationFont(font_path)
            font_family = QFontDatabase.applicationFontFamilies(font_id)[0]
            font = QFont(font_family)
            self.setFont(font)
        self.setWindowTitle("ファイル検索アプリ")
        self.resize(600, 400)

        layout = QVBoxLayout()

        self.folder_label = QLabel("選択されたフォルダ: なし")
        layout.addWidget(self.folder_label)

        self.select_button = QPushButton("フォルダを選択")
        self.select_button.clicked.connect(self.select_folder)
        layout.addWidget(self.select_button)

        self.pattern_input = QLineEdit()
        self.pattern_input.setPlaceholderText("検索パターンを入力（例: .txt）")
        layout.addWidget(self.pattern_input)

        self.search_button = QPushButton("検索")
        self.search_button.clicked.connect(self.search_files)
        layout.addWidget(self.search_button)

        self.result_list = QListWidget()
        layout.addWidget(self.result_list)

        self.setLayout(layout)
        self.folder_path = ""

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "フォルダを選択")
        if folder:
            self.folder_path = folder
            self.folder_label.setText(f"選択されたフォルダ: {folder}")

    def search_files(self):
        if not self.folder_path or not os.path.exists(self.folder_path):
            QMessageBox.warning(self, "エラー", "有効なフォルダを選択してください。")
            return

        pattern = self.pattern_input.text().strip()
        if not pattern:
            QMessageBox.warning(self, "エラー", "検索パターンを入力してください。")
            return

        all_files = [
            f
            for f in glob.glob(os.path.join(self.folder_path, "**"), recursive=True)
            if os.path.isfile(f)
        ]
        matched_files = [f for f in all_files if pattern in os.path.basename(f)]

        self.result_list.clear()
        if matched_files:
            self.result_list.addItems(matched_files)
        else:
            self.result_list.addItem("該当ファイルなし")


def main():
    app = QApplication(sys.argv)
    window = FileSearchApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
