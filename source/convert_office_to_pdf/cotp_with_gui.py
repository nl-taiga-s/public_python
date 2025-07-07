import os
import platform
import subprocess
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
from source.convert_office_to_pdf.cotp_class import ConvertOfficeToPdf


class ConvertToPdfApp(QWidget):
    def __init__(self):
        if platform.system() != "Windows":
            raise EnvironmentError("このアプリはWindows専用です。")
        super().__init__()
        self.obj_of_pft = PlatformTools()
        self.obj_of_pt = PathTools()
        self.setWindowTitle("Officeファイル → PDF 一括変換")

        # ウィジェット作成
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

        # 状態変数
        self.folder_path_from = ""
        self.folder_path_to = ""
        self.pdf_converter = None

        # シグナル接続
        self.btn_select_from.clicked.connect(self.select_from_folder)
        self.btn_open_from.clicked.connect(
            lambda: self.open_explorer(self.folder_path_from)
        )
        self.btn_select_to.clicked.connect(self.select_to_folder)
        self.btn_open_to.clicked.connect(
            lambda: self.open_explorer(self.folder_path_to)
        )
        self.btn_convert.clicked.connect(self.start_conversion)

    def log(self, message: str):
        self.log_output.append(message)
        print(message)

    def open_explorer(self, folder: str):
        if folder:
            try:
                system_name = platform.system()
                if system_name == "Windows":
                    os.startfile(folder)
                elif self.obj_of_pft.is_wsl():
                    # Windowsのパスに変換（/mnt/c/... 形式）
                    wsl_path = (
                        subprocess.check_output(["wslpath", "-w", folder])
                        .decode("utf-8")
                        .strip()
                    )
                    subprocess.run(["explorer.exe", wsl_path])
            except Exception as e:
                print(f"エクスプローラー起動エラー: {e}")
        else:
            self.log("フォルダが未選択のため開けません。")

    def select_from_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "変換元フォルダを選択")
        if folder:
            self.folder_path_from = folder
            self.label_from.setText(f"変換元フォルダ: {folder}")
            self.try_load_files()

    def select_to_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "変換先フォルダを選択")
        if folder:
            self.folder_path_to = folder
            self.label_to.setText(f"変換先フォルダ: {folder}")
            self.try_load_files()

    def try_load_files(self):
        """両フォルダ選択済みならファイル一覧を表示"""
        if not self.folder_path_from or not self.folder_path_to:
            return

        try:
            self.pdf_converter = ConvertOfficeToPdf(
                self.folder_path_from, self.folder_path_to
            )
            self.file_list_widget.clear()
            for f in self.pdf_converter.list_of_f:
                fp = Path(f)
                file_path = self.obj_of_pt.get_entire_file_name(fp)
                self.file_list_widget.addItem(file_path)
            self.progress_bar.setValue(0)
        except ValueError as e:
            self.file_list_widget.clear()
            self.log(f"⚠️ {e}")
        else:
            self.log(
                f"✅ {self.pdf_converter.number_of_f} 件のファイルが見つかりました。"
            )

    def start_conversion(self):
        self.log_output.clear()

        if not self.pdf_converter:
            self.log("⚠️ ファイルリストが初期化されていません。")
            return

        total = self.pdf_converter.number_of_f
        self.progress_bar.setRange(0, total)

        self.log("📄 一括変換を開始します...")
        for i in range(total):
            try:
                c_fp_f = Path(self.pdf_converter.current_of_file_path_from)
                file_name = self.obj_of_pt.get_entire_file_name(c_fp_f)
                self.pdf_converter.handle_file()
                self.progress_bar.setValue(i + 1)
                self.pdf_converter.move_to_next_file()
            except Exception as e:
                self.log(f"❌ {file_name} → エラー: {e}")
            else:
                self.log(f"✅ {file_name} → 完了")
        self.log("🎉 すべてのファイルの変換が完了しました！")
        fp_e = Path(__file__)
        self.pdf_converter.write_log(fp_e)
        self.log("📄 変換ログを出力しました。")


def main():
    app = QApplication(sys.argv)
    window = ConvertToPdfApp()
    window.resize(700, 600)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
