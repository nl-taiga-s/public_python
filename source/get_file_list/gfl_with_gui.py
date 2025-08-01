import os
import platform
import subprocess
import sys
from pathlib import Path

from PySide6.QtGui import QFont, QFontDatabase
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

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
        self.setWindowTitle("ファイル検索ツール")
        # ウィジェット
        self.folder_label = QLabel("フォルダ未選択")
        self.select_folder_btn = QPushButton("フォルダを選択")
        self.recursive_checkbox = QCheckBox("サブフォルダも含めて検索（再帰）")
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
        layout.addWidget(self.recursive_checkbox)
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
        self.open_folder_btn.clicked.connect(lambda: self.open_explorer(self.folder))
        self.search_btn.clicked.connect(self.search_files)
        self.export_btn.clicked.connect(self.export_results)

    def select_folder(self):
        """フォルダを選択します"""
        self.folder = QFileDialog.getExistingDirectory(self, "フォルダを選択")
        folder_as_path_type = Path(self.folder)
        self.folder = str(folder_as_path_type)
        if self.folder:
            self.folder_label.setText(self.folder)
            recursive = self.recursive_checkbox.isChecked()
            self.obj_of_cls = GetFileList(self.folder, recursive)

    def open_explorer(self, folder: str):
        """エクスプローラーを開きます"""
        if folder:
            try:
                if platform.system().lower() == "windows":
                    os.startfile(folder)
                elif self.obj_of_pft.is_wsl():
                    # Windowsのパスに変換（/mnt/c/... 形式）
                    wsl_path = subprocess.check_output(["wslpath", "-w", folder]).decode("utf-8").strip()
                    subprocess.run(["explorer.exe", wsl_path])
            except Exception as e:
                print(f"エクスプローラー起動エラー: {e}")
        else:
            self.result_list.addItem("フォルダが未選択のため開けません。")

    def search_files(self):
        """ファイルを検索します"""
        if not self.obj_of_cls:
            self.result_list.clear()
            self.result_list.addItem("フォルダが未選択です。")
            return
        pattern = self.pattern_input.text().strip()
        recursive = self.recursive_checkbox.isChecked()
        self.obj_of_cls = GetFileList(self.folder, recursive)
        self.obj_of_cls.extract_by_pattern(pattern)
        self.result_list.clear()
        if self.obj_of_cls.list_file_after:
            self.result_list.addItems(self.obj_of_cls.list_file_after)
        else:
            self.result_list.addItem("一致するファイルが見つかりませんでした。")

    def export_results(self):
        """処理結果を出力します"""
        if not self.obj_of_cls:
            self.result_list.addItem("検索対象のフォルダが未選択です。")
            return
        if not self.obj_of_cls.list_file_after:
            self.log_list.addItem("出力する検索結果がありません。")
            return
        try:
            file_of_exe_as_path_type = Path(__file__)
            file_of_log_as_path_type = self.obj_of_pt.get_file_path_of_log(file_of_exe_as_path_type)
            self.obj_of_cls.write_log(file_of_log_as_path_type, self.obj_of_cls.list_file_after)
        except Exception as e:
            self.result_list.addItem(f"ログファイルの出力に失敗しました。: \n{e}")
        else:
            self.result_list.addItem(f"ログファイルの出力に成功しました。: \n{str(file_of_log_as_path_type)}")


def main():
    """主要関数"""
    app = QApplication(sys.argv)
    window = MainApp_Of_GFL()
    window.resize(600, 400)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
