import os
import platform
import subprocess
import sys
from pathlib import Path

from PySide6.QtWidgets import QApplication, QCheckBox, QFileDialog, QLabel, QLineEdit, QMessageBox, QPushButton, QTextEdit, QVBoxLayout, QWidget

from source.common.common import (
    DatetimeTools,
    PathTools,
    PlatformTools,
)
from source.get_file_list.gfl_class import GetFileList


class MainApp_Of_GFL(QWidget):
    def __init__(self):
        """初期化します"""
        super().__init__()
        self.obj_of_pft = PlatformTools()
        self.obj_of_dt2 = DatetimeTools()
        self.obj_of_pt = PathTools()
        self.setup_ui()

    def closeEvent(self, event):
        """終了します"""
        self.write_log()
        super().closeEvent(event)

    def setup_ui(self):
        """User Interfaceを設定します"""
        # WSL-Ubuntuでフォント設定
        # if self.obj_of_pft.is_wsl():
        #     font_path = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"
        #     font_id = QFontDatabase.addApplicationFont(font_path)
        #     font_family = QFontDatabase.applicationFontFamilies(font_id)[0]
        #     font = QFont(font_family)
        #     self.setFont(font)
        # タイトル
        self.setWindowTitle("ファイル検索ツール")
        # ウィジェット
        self.folder_label = QLabel("フォルダ未選択")
        select_folder_btn = QPushButton("フォルダを選択")
        self.recursive_checkbox = QCheckBox("サブフォルダも含めて検索（再帰）")
        self.pattern_input = QLineEdit()
        self.pattern_input.setPlaceholderText("検索パターンを入力...")
        open_folder_btn = QPushButton("フォルダを開く")
        search_btn = QPushButton("検索実行")
        self.log_area = QTextEdit()
        # レイアウト
        layout = QVBoxLayout()
        layout.addWidget(self.folder_label)
        layout.addWidget(select_folder_btn)
        layout.addWidget(self.recursive_checkbox)
        layout.addWidget(QLabel("検索パターン:"))
        layout.addWidget(self.pattern_input)
        layout.addWidget(open_folder_btn)
        layout.addWidget(search_btn)
        layout.addWidget(QLabel("検索結果:"))
        layout.addWidget(self.log_area)
        self.setLayout(layout)
        # シグナル接続
        select_folder_btn.clicked.connect(self.select_folder)
        open_folder_btn.clicked.connect(lambda: self.open_explorer(self.folder))
        search_btn.clicked.connect(self.search_files)

    def select_folder(self):
        """フォルダを選択します"""
        self.folder = QFileDialog.getExistingDirectory(self, "フォルダを選択")
        folder_as_path_type = Path(self.folder).expanduser()
        self.folder = str(folder_as_path_type)
        if self.folder:
            self.folder_label.setText(self.folder)
            recursive = self.recursive_checkbox.isChecked()
            self.obj_of_cls = GetFileList(self.folder, recursive)

    def open_explorer(self, folder: str):
        """エクスプローラーを開きます"""
        EXPLORER = "/mnt/c/Windows/explorer.exe"
        if folder:
            try:
                if platform.system().lower() == "windows":
                    os.startfile(folder)
                elif self.obj_of_pft.is_wsl():
                    # Windowsのパスに変換（/mnt/c/... 形式）
                    wsl_path = subprocess.check_output(["wslpath", "-w", folder]).decode("utf-8").strip()
                    subprocess.run([EXPLORER, wsl_path])
            except Exception as e:
                self.show_error(str(e))
        else:
            self.output_log("フォルダが未選択のため開けません。")

    def search_files(self):
        """ファイルを検索します"""
        if not self.obj_of_cls:
            self.output_log("フォルダが未選択です。")
            return
        pattern = self.pattern_input.text().strip()
        recursive = self.recursive_checkbox.isChecked()
        self.obj_of_cls = GetFileList(self.folder, recursive)
        self.obj_of_cls.extract_by_pattern(pattern)
        if self.obj_of_cls.list_file_after:
            self.output_log("\n".join(self.obj_of_cls.list_file_after))
        else:
            self.output_log("一致するファイルが見つかりませんでした。")

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
        QMessageBox.information(self, f"{label}の結果", f"{label}に{'成功' if success else '失敗'}しました。")

    def show_error(self, msg: str):
        """エラーを表示します"""
        QMessageBox.information(self, "エラー", msg)

    def output_log(self, message: str):
        """メッセージを表示します"""
        self.log_area.append(message)


def main():
    """主要関数"""
    if platform.system().lower() == "windows":
        # アプリ全体のスケール
        os.environ["QT_SCALE_FACTOR"] = "0.7"
    app = QApplication(sys.argv)
    window = MainApp_Of_GFL()
    window.resize(600, 400)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
