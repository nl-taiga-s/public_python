import os
import platform
import subprocess
import sys
from datetime import datetime

from gfl_class import GetFileList
from PySide6.QtGui import QFont, QFontDatabase
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


def is_wsl() -> bool:
    """WSL(Windows Subsystem Linux)かどうかを判定します"""
    if platform.system() != "Linux":
        return False
    try:
        with open("/proc/version", "r") as f:
            content = f.read().lower()
            return "microsoft" in content or "wsl" in content
    except Exception:
        return False


class FileSearchApp(QWidget):
    def __init__(self):
        super().__init__()
        # WSL-Ubuntuでフォント設定
        if is_wsl():
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
        self.open_folder_btn = QPushButton("フォルダを開く")
        self.search_btn = QPushButton("検索実行")
        self.export_btn = QPushButton("検索結果を出力")
        self.result_list = QListWidget()

        # レイアウト
        layout = QVBoxLayout()
        layout.addWidget(self.folder_label)
        layout.addWidget(self.select_folder_btn)
        layout.addWidget(QLabel("検索パターン:"))
        layout.addWidget(self.pattern_input)
        layout.addWidget(self.open_folder_btn)
        layout.addWidget(self.search_btn)
        layout.addWidget(self.export_btn)
        layout.addWidget(QLabel("検索結果:"))
        layout.addWidget(self.result_list)

        self.setLayout(layout)

        # シグナル接続
        self.select_folder_btn.clicked.connect(self.select_folder)
        self.open_folder_btn.clicked.connect(self.open_folder)
        self.search_btn.clicked.connect(self.search_files)
        self.export_btn.clicked.connect(self.export_results)

    def select_folder(self):
        self.folder = QFileDialog.getExistingDirectory(self, "フォルダを選択")
        if self.folder:
            self.folder_label.setText(self.folder)
            self.file_list_obj = GetFileList(self.folder)

    def open_folder(self):
        if hasattr(self, 'folder') and self.folder:
            try:
                system_name = platform.system()
                if system_name == "Windows":
                    os.startfile(self.folder)
                elif is_wsl():
                    # Windowsのパスに変換（/mnt/c/... 形式）
                    wsl_path = (
                        subprocess.check_output(["wslpath", "-w", self.folder])
                        .decode("utf-8")
                        .strip()
                    )
                    subprocess.run(["explorer.exe", wsl_path])
            except Exception as e:
                print(f"エクスプローラー起動エラー: {e}")
        else:
            self.result_list.addItem("フォルダが未選択のため開けません。")

    def search_files(self):
        if not self.file_list_obj:
            self.result_list.clear()
            self.result_list.addItem("フォルダが未選択です。")
            return

        pattern = self.pattern_input.text().strip()
        self.file_list_obj.extract_by_pattern(pattern)
        self.result_list.clear()

        if self.file_list_obj.list_file_after:
            self.result_list.addItems(self.file_list_obj.list_file_after)
        else:
            self.result_list.addItem("一致するファイルが見つかりませんでした。")

    def export_results(self):
        if not self.file_list_obj:
            self.result_list.addItem("検索対象のフォルダが未選択です。")
            return

        if not self.file_list_obj.list_file_after:
            self.result_list.addItem("出力する検索結果がありません。")
            return

        try:
            # exe 実行時とスクリプト実行時に対応した保存先
            if getattr(sys, 'frozen', False):
                exe_dir = os.path.dirname(sys.executable)
            else:
                exe_dir = os.path.dirname(os.path.abspath(__file__))

            output_dir = os.path.join(exe_dir, "output")
            os.makedirs(output_dir, exist_ok=True)

            now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(output_dir, f"search_result_{now_str}.txt")

            with open(output_file, "w", encoding="utf-8") as f:
                f.write("\n".join(self.file_list_obj.list_file_after))

            self.result_list.addItem(f"結果を出力しました: {output_file}")

        except Exception as e:
            self.result_list.addItem(f"出力時にエラーが発生しました: {e}")


def main():
    app = QApplication(sys.argv)
    window = FileSearchApp()
    window.resize(600, 400)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
