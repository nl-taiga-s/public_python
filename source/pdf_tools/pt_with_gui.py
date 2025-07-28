import sys
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from source.common.common import GUITools, PathTools
from source.pdf_tools.pt_class import PdfTools


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDFツール - GUI完全版★")
        self.resize(500, 600)

        self.pt = PathTools()
        self.pdf = PdfTools()
        self.gui = GUITools()

        self.init_ui()

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # ファイル選択
        self.file_input = QLineEdit()
        layout.addWidget(QLabel("PDFファイルパス"))
        layout.addWidget(self.file_input)
        browse_btn = QPushButton("参照")
        browse_btn.clicked.connect(self.select_pdf)
        layout.addWidget(browse_btn)

        # パスワード入力
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("パスワード（英数字/アンダーバー/ハイフン）")
        layout.addWidget(QLabel("パスワード"))
        layout.addWidget(self.password_input)

        # 暗号化・復号化
        encrypt_btn = QPushButton("暗号化")
        encrypt_btn.clicked.connect(self.encrypt_pdf)
        layout.addWidget(encrypt_btn)

        decrypt_btn = QPushButton("復号化")
        decrypt_btn.clicked.connect(self.decrypt_pdf)
        layout.addWidget(decrypt_btn)

        # メタデータ
        meta_btn = QPushButton("メタデータ読み込み＆表示")
        meta_btn.clicked.connect(self.read_and_show_metadata)
        layout.addWidget(meta_btn)

        write_meta_btn = QPushButton("メタデータ書き込み")
        write_meta_btn.clicked.connect(self.write_metadata)
        layout.addWidget(write_meta_btn)

        # マージ
        merge_btn = QPushButton("複数PDFをマージ")
        merge_btn.clicked.connect(self.merge_pdfs)
        layout.addWidget(merge_btn)

        # ページ抽出
        self.begin_spin = QSpinBox()
        self.end_spin = QSpinBox()
        page_layout = QHBoxLayout()
        page_layout.addWidget(QLabel("抽出開始"))
        page_layout.addWidget(self.begin_spin)
        page_layout.addWidget(QLabel("抽出終了"))
        page_layout.addWidget(self.end_spin)
        layout.addLayout(page_layout)

        extract_btn = QPushButton("ページ抽出")
        extract_btn.clicked.connect(self.extract_pages)
        layout.addWidget(extract_btn)

        rotate_btn = QPushButton("ページを時計回りに回転（90度）")
        rotate_btn.clicked.connect(self.rotate_page)
        layout.addWidget(rotate_btn)

        # 終了時にログ出力
        self.destroyed.connect(self.write_log)

    def select_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "PDFファイルを選択", "", "PDF Files (*.pdf)")
        if file_path:
            self.file_input.setText(file_path)

    def encrypt_pdf(self):
        path, pw = self.file_input.text(), self.password_input.text()
        if not path or not pw:
            self.show_error("ファイルパスとパスワードを入力してください。")
            return
        success = self.pdf.encrypt(path, pw)
        self.show_result("暗号化", success)

    def decrypt_pdf(self):
        path, pw = self.file_input.text(), self.password_input.text()
        if not path or not pw:
            self.show_error("ファイルパスとパスワードを入力してください。")
            return
        success = self.pdf.decrypt(path, pw)
        self.show_result("復号化", success)

    def read_and_show_metadata(self):
        path = self.file_input.text()
        if not path:
            self.show_error("PDFファイルパスを入力してください。")
            return
        if self.pdf.read_metadata(path):
            self.pdf.print_metadata()
            self.show_result("メタデータの読み込み", True)
        else:
            self.show_result("メタデータの読み込み", False)

    def write_metadata(self):
        path = self.file_input.text()
        if not path:
            self.show_error("PDFファイルパスを入力してください。")
            return
        self.pdf.write_metadata(path)

    def merge_pdfs(self):
        files, _ = QFileDialog.getOpenFileNames(self, "マージするPDFを選択", "", "PDF Files (*.pdf)")
        if files:
            success = self.pdf.merge(files)
            self.show_result("マージ", success)

    def extract_pages(self):
        path = self.file_input.text()
        begin, end = self.begin_spin.value(), self.end_spin.value()
        if not path or begin < 1 or end < begin:
            self.show_error("ページ範囲またはパスが不正です。")
            return
        self.pdf.reader = self.pdf.reader or PdfTools().reader
        self.pdf.num_of_pages = len(PdfTools().reader.pages)
        self.pdf.extract_pages(path, begin, end)

    def rotate_page(self):
        path = self.file_input.text()
        page = self.begin_spin.value()
        if not path or page < 1:
            self.show_error("ページ番号が不正です。")
            return
        self.pdf.num_of_pages = len(PdfTools().reader.pages)
        self.pdf.rotate_page_clockwise(path, page, 90)

    def show_result(self, label: str, success: bool):
        QMessageBox.information(self, f"{label}結果", f"{label}に{'成功' if success else '失敗'}しました。")

    def show_error(self, msg: str):
        self.gui.show_error(msg)

    def write_log(self):
        exe_path = Path(__file__)
        log_path = self.pt.get_file_path_of_log(exe_path)
        self.pdf.write_log(log_path)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main = MainWindow()
    main.show()
    sys.exit(app.exec())
