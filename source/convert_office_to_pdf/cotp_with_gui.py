import os
import platform
import sys
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QLabel,
    QListWidget,
    QProgressBar,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from source.common.common import PathTools, PlatformTools


def check_os() -> object:
    try:
        from source.convert_office_to_pdf.cotp_class import ConvertOfficeToPDF
    except SystemExit:
        return None
    else:
        return ConvertOfficeToPDF


class ConvertToPdfApp(QWidget):
    def __init__(self):
        """初期化します"""
        self.cotp = check_os()
        if self.cotp is None:
            sys.exit(0)
        super().__init__()
        self.obj_of_pft = PlatformTools()
        self.obj_of_pt = PathTools()
        self.folder_path_from = ""
        self.folder_path_to = ""
        self.setup_ui()

    def setup_ui(self):
        """User Interfaceを設定します"""
        # タイトル
        self.setWindowTitle("Officeファイル → PDF 一括変換")
        # ウィジェット
        self.label_from = QLabel("変換元フォルダ: 未選択")
        self.btn_select_from = QPushButton("変換元フォルダを選択")
        self.btn_open_from = QPushButton("変換元フォルダを開く")
        self.label_to = QLabel("変換先フォルダ: 未選択")
        self.btn_select_to = QPushButton("変換先フォルダを選択")
        self.btn_open_to = QPushButton("変換先フォルダを開く")
        self.file_list_widget = QListWidget()
        self.progress_bar = QProgressBar()
        self.btn_convert = QPushButton("一括変換 実行")
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        # レイアウト
        layout = QVBoxLayout()
        layout.addWidget(self.label_from)
        layout.addWidget(self.btn_select_from)
        layout.addWidget(self.btn_open_from)
        layout.addWidget(self.label_to)
        layout.addWidget(self.btn_select_to)
        layout.addWidget(self.btn_open_to)
        layout.addWidget(QLabel("変換対象ファイル一覧:"))
        layout.addWidget(self.file_list_widget)
        layout.addWidget(QLabel("進行状況:"))
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.btn_convert)
        layout.addWidget(QLabel("ログ:"))
        layout.addWidget(self.log_output)
        self.setLayout(layout)
        # シグナル接続
        self.btn_select_from.clicked.connect(self.select_from_folder)
        self.btn_open_from.clicked.connect(lambda: self.open_explorer(self.folder_path_from))
        self.btn_select_to.clicked.connect(self.select_to_folder)
        self.btn_open_to.clicked.connect(lambda: self.open_explorer(self.folder_path_to))
        self.btn_convert.clicked.connect(self.convert_file)

    def log(self, message: str):
        """メッセージを表示します"""
        self.log_output.append(message)

    def open_explorer(self, folder: str):
        """エクスプローラーを開きます"""
        if folder:
            try:
                if platform.system().lower() == "windows":
                    os.startfile(folder)
            except Exception as e:
                print(f"エクスプローラー起動エラー: {e}")
        else:
            self.log("フォルダが未選択のため開けません。")

    def select_from_folder(self):
        """変換元のフォルダを選択します"""
        folder = QFileDialog.getExistingDirectory(self, "変換元のフォルダを選択")
        folder_as_path_type = Path(folder)
        folder = str(folder_as_path_type)
        if folder:
            self.folder_path_from = folder
            self.label_from.setText(f"変換元フォルダ: {folder}")
            self.try_load_files()

    def select_to_folder(self):
        """変換先のフォルダを選択します"""
        folder = QFileDialog.getExistingDirectory(self, "変換先のフォルダを選択")
        folder_as_path_type = Path(folder)
        folder = str(folder_as_path_type)
        if folder:
            self.folder_path_to = folder
            self.label_to.setText(f"変換先フォルダ: {folder}")
            self.try_load_files()

    def try_load_files(self):
        """両フォルダ選択済みならファイル一覧を表示します"""
        if not self.folder_path_from or not self.folder_path_to:
            return
        try:
            self.pdf_converter = self.cotp(self.folder_path_from, self.folder_path_to)
            self.file_list_widget.clear()
            for f in self.pdf_converter.filtered_list_of_f:
                file_as_path_type = Path(f)
                file_path = self.obj_of_pt.get_entire_file_name(file_as_path_type)
                self.file_list_widget.addItem(file_path)
            self.progress_bar.setValue(0)
        except ValueError as e:
            self.file_list_widget.clear()
            self.log(f"⚠️ {e}")
        else:
            self.log(f"✅ {self.pdf_converter.number_of_f} 件のファイルが見つかりました。")

    def convert_file(self):
        """変換します"""
        self.log_output.clear()
        if not self.pdf_converter:
            self.log("⚠️ ファイルリストが初期化されていません。")
            return
        total = self.pdf_converter.number_of_f
        self.progress_bar.setRange(0, total)
        self.log("📄 一括変換を開始します...")
        for i in range(total):
            try:
                file_of_currentfrom_as_path_type = Path(self.pdf_converter.current_of_file_path_from)
                file_name = self.obj_of_pt.get_entire_file_name(file_of_currentfrom_as_path_type)
                self.pdf_converter.handle_file()
            except Exception as e:
                self.log(f"❌ {file_name} → エラー: {e}")
            else:
                self.log(f"✅ {file_name} → 完了")
                self.progress_bar.setValue(i + 1)
                self.pdf_converter.move_to_next_file()
        self.log("🎉 すべてのファイルの変換が完了しました！")
        try:
            file_of_exe_as_path_type = Path(__file__)
            file_of_log_as_path_type = self.obj_of_pt.get_file_path_of_log(file_of_exe_as_path_type)
            self.pdf_converter.write_log(file_of_log_as_path_type)
        except Exception as e:
            self.log(f"📄 ログファイルの出力に失敗しました。: \n{e}")
        else:
            self.log(f"📄 ログファイルの出力に成功しました。: \n{str(file_of_log_as_path_type)}")


def main():
    """主要関数"""
    app = QApplication(sys.argv)
    window = ConvertToPdfApp()
    window.resize(700, 600)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
