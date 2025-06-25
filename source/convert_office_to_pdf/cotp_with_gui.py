import os
import platform
import subprocess
import sys

from cotp_class import ConvertOfficeToPdf
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


class ConvertToPdfApp(QWidget):
    def __init__(self):
        if platform.system() != "Windows":
            raise EnvironmentError("このアプリはWindows専用です。")
        super().__init__()
        self.setWindowTitle("Officeファイル → PDF 一括変換")

        # --- UI要素 ---
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

        # --- レイアウト構築 ---
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

        # シグナル
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
                elif is_wsl():
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
                self.file_list_widget.addItem(os.path.basename(f))
            self.progress_bar.setValue(0)
            self.log(
                f"✅ {self.pdf_converter.number_of_f} 件のファイルが見つかりました。"
            )
        except ValueError as e:
            self.file_list_widget.clear()
            self.log(f"⚠️ {e}")

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
                self.pdf_converter.handle_file()
                self.log(
                    f"✅"
                    f" {os.path.basename(self.pdf_converter.current_of_file_path_from)}"
                    f" → 完了"
                )
            except Exception as e:
                self.log(
                    f"❌"
                    f" {os.path.basename(self.pdf_converter.current_of_file_path_from)}"
                    f" → エラー: {e}"
                )
            self.progress_bar.setValue(i + 1)
            self.pdf_converter._ConvertOfficeToPdf__next()

        self.log("🎉 すべてのファイルの変換が完了しました！")


def main():
    app = QApplication(sys.argv)
    window = ConvertToPdfApp()
    window.resize(700, 600)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
