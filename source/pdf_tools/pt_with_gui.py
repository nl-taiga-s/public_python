import sys
from pathlib import Path

import pypdfium2
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from source.common.common import DatetimeTools, PathTools
from source.pdf_tools.pt_class import PdfTools


class MainApp_Of_PT(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDFツール")
        self.resize(700, 700)

        self.obj_of_pt = PathTools()
        self.obj_of_dt2 = DatetimeTools()
        self.obj_of_cls = PdfTools()

        self.first_init_ui()

    def closeEvent(self, event):
        self.write_log()
        super().closeEvent(event)

    def first_init_ui(self):
        self.central = QWidget()
        self.setCentralWidget(self.central)
        # 主要
        self.main_layout = QHBoxLayout(self.central)
        # 左側
        self.left_layout = QVBoxLayout()
        self.main_layout.addLayout(self.left_layout)
        # 右側
        self.right_scroll_area = QScrollArea()
        self.right_scroll_area.setWidgetResizable(True)
        self.main_layout.addWidget(self.right_scroll_area)
        # 仮想コンテナのウィジェットとレイアウト
        self.right_widget = QWidget()
        self.right_layout = QGridLayout(self.right_widget)
        self.right_scroll_area.setWidget(self.right_widget)

        # ファイル選択
        self.file_input = QLineEdit()
        self.left_layout.addWidget(QLabel("PDFファイルパス"))
        self.left_layout.addWidget(self.file_input)
        browse_btn = QPushButton("参照")
        browse_btn.clicked.connect(self.select_pdf)
        self.left_layout.addWidget(browse_btn)

        # パスワード入力
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("パスワード（英数字/アンダーバー/ハイフン）")
        self.left_layout.addWidget(QLabel("パスワード"))
        self.left_layout.addWidget(self.password_input)

        # 暗号化
        encrypt_btn = QPushButton("暗号化")
        encrypt_btn.clicked.connect(self.encrypt_pdf)
        self.left_layout.addWidget(encrypt_btn)

        # 復号化
        decrypt_btn = QPushButton("復号化")
        decrypt_btn.clicked.connect(self.decrypt_pdf)
        self.left_layout.addWidget(decrypt_btn)

        # メタデータの表示
        meta_btn = QPushButton("メタデータの表示")
        meta_btn.clicked.connect(self.show_metadata)
        self.left_layout.addWidget(meta_btn)

        # メタデータの入力
        self.left_layout.addWidget(QLabel("メタデータの入力"))
        self.widget_of_metadata = {}
        self.line_edits_of_metadata = {}
        for value, key in self.obj_of_cls.fields:
            if value in ["creation_date", "modification_date"]:
                continue
            self.line_edits_of_metadata[key] = QLineEdit()
            self.left_layout.addWidget(QLabel(value.capitalize().replace("_", " ")))
            self.left_layout.addWidget(self.line_edits_of_metadata[key])

        # メタデータの書き込み
        write_meta_btn = QPushButton("メタデータの書き込み")
        write_meta_btn.clicked.connect(self.write_metadata)
        self.left_layout.addWidget(write_meta_btn)

        # マージ
        merge_btn = QPushButton("複数PDFをマージ")
        merge_btn.clicked.connect(self.merge_pdfs)
        self.left_layout.addWidget(merge_btn)

        # ページの抽出
        self.begin_spin_of_ep = QSpinBox()
        self.end_spin_of_ep = QSpinBox()
        page_layout_of_ep = QHBoxLayout()
        page_layout_of_ep.addWidget(QLabel("ページの抽出開始"))
        page_layout_of_ep.addWidget(self.begin_spin_of_ep)
        page_layout_of_ep.addWidget(QLabel("ページの抽出終了"))
        page_layout_of_ep.addWidget(self.end_spin_of_ep)
        self.left_layout.addLayout(page_layout_of_ep)

        extract_page_btn = QPushButton("ページの抽出")
        extract_page_btn.clicked.connect(self.extract_pages)
        self.left_layout.addWidget(extract_page_btn)

        # テキストの抽出
        self.begin_spin_of_et = QSpinBox()
        self.end_spin_of_et = QSpinBox()
        page_layout_of_et = QHBoxLayout()
        page_layout_of_et.addWidget(QLabel("テキストの抽出開始"))
        page_layout_of_et.addWidget(self.begin_spin_of_et)
        page_layout_of_et.addWidget(QLabel("テキストの抽出終了"))
        page_layout_of_et.addWidget(self.end_spin_of_et)
        self.left_layout.addLayout(page_layout_of_et)

        extract_text_btn = QPushButton("テキストの抽出")
        extract_text_btn.clicked.connect(self.extract_text)
        self.left_layout.addWidget(extract_text_btn)

        # ページの回転
        self.spin_of_rp = QSpinBox()
        page_layout_of_rp = QHBoxLayout()
        page_layout_of_rp.addWidget(QLabel("回転するページ"))
        page_layout_of_rp.addWidget(self.spin_of_rp)
        rotate_btn = QPushButton("ページを時計回りに回転（90度）")
        rotate_btn.clicked.connect(self.rotate_page)
        page_layout_of_rp.addWidget(rotate_btn)
        self.left_layout.addLayout(page_layout_of_rp)

        # ログの出力
        self.output = QTextEdit()
        self.output.setReadOnly(True)
        self.left_layout.addWidget(QLabel("出力"))
        self.left_layout.addWidget(self.output)

    def second_init_ui(self, images: list):
        # 既存のレイアウトをクリア（再表示に対応）
        for i in reversed(range(self.right_layout.count())):
            widget = self.right_layout.itemAt(i).widget()
            if widget is not None:
                widget.setParent(None)
        # 各ページを表示する
        for i, element in enumerate(images):
            # 垂直レイアウトを用意する
            page_layout = QVBoxLayout()
            page_widget = QWidget()
            page_widget.setLayout(page_layout)
            # ページ番号のラベル
            page_num_label = QLabel(f"page: {i + 1}")
            page_num_label.setAlignment(Qt.AlignCenter)
            page_layout.addWidget(page_num_label)
            # 画像のラベル
            image_label = QLabel()
            pixmap = QPixmap(element).scaledToWidth(300, Qt.SmoothTransformation)
            image_label.setPixmap(pixmap)
            image_label.setScaledContents(True)
            image_label.setFixedSize(pixmap.size())
            page_layout.addWidget(image_label)
            # グリッドにページごとに追加する
            self.right_layout.addWidget(page_widget, i // 2, i % 2)

    def select_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "PDFファイルを選択", "", "PDF Files (*.pdf)")
        if file_path.strip():
            self.file_input.setText(file_path)
            self.obj_of_cls.read_metadata(file_path)
            images = self.get_image(file_path)
            self.second_init_ui(images)

    def get_image(self, file_path: str) -> list:
        file_as_path_type = Path(file_path)
        file_name_as_str_type = file_as_path_type.stem
        self.output_dir = Path(__file__).parent / "__images__"
        self.output_dir.mkdir(parents=True, exist_ok=True)
        pdf = pypdfium2.PdfDocument(file_path)
        output_files = []
        for i, page in enumerate(pdf):
            pil_image = page.render(scale=1, rotation=0, crop=(0, 0, 0, 0)).to_pil()
            output_file = self.output_dir / f"{file_name_as_str_type}_{i + 1}.png"
            pil_image.save(output_file)
            output_files.append(str(output_file))
            page.close()
        return output_files

    def encrypt_pdf(self):
        path, pw = self.file_input.text(), self.password_input.text()
        if not path.strip() or not pw.strip():
            self.show_error("ファイルパスとパスワードを入力してください。")
            return None
        success = self.obj_of_cls.encrypt(path, pw)
        self.show_result("暗号化", success)
        self.output_log()

    def decrypt_pdf(self):
        path, pw = self.file_input.text(), self.password_input.text()
        if not path.strip() or not pw.strip():
            self.show_error("ファイルパスとパスワードを入力してください。")
            return None
        success = self.obj_of_cls.decrypt(path, pw)
        self.show_result("復号化", success)
        self.output_log()

    def show_metadata(self):
        path = self.file_input.text()
        if not path.strip():
            self.show_error("PDFファイルパスを入力してください。")
            return None
        try:
            lines = []
            for k, v in self.obj_of_cls.metadata_of_reader.items():
                lines.append(f"{k}: {v}")
            self.obj_of_cls.log += lines
        except Exception:
            b = False
        else:
            b = True
        finally:
            self.show_result("メタデータの表示", b)
            self.output_log()

    def write_metadata(self):
        path = self.file_input.text()
        if not path.strip():
            self.show_error("PDFファイルパスを入力してください。")
            return None
        for value, key in self.obj_of_cls.fields:
            match value:
                case "creation_date":
                    self.widget_of_metadata[key] = self.obj_of_cls.creation_date
                case "modification_date":
                    time = self.obj_of_dt2.convert_for_metadata_in_pdf(self.obj_of_cls.UTC_OF_JP)
                    self.widget_of_metadata[key] = time
                case _:
                    self.widget_of_metadata[key] = self.line_edits_of_metadata[key].text()
        success = self.obj_of_cls.write_metadata(path, self.widget_of_metadata)
        self.show_result("メタデータの書き込み", success)
        self.output_log()
        self.obj_of_cls.read_metadata(path)

    def merge_pdfs(self):
        files, _ = QFileDialog.getOpenFileNames(self, "マージするPDFを選択", "", "PDF Files (*.pdf)")
        if files:
            success = self.obj_of_cls.merge(files)
            self.show_result("マージ", success)
        self.output_log()

    def extract_pages(self):
        path = self.file_input.text()
        begin, end = self.begin_spin_of_ep.value(), self.end_spin_of_ep.value()
        if begin == 0 or end == 0:
            self.show_error("ページ範囲を指定してください。")
            return None
        if not path.strip() or begin < 1 or end < begin:
            self.show_error("ページ範囲またはパスが不正です。")
            return None
        success = self.obj_of_cls.extract_pages(path, begin, end)
        self.show_result("ページの抽出", success)
        self.output_log()

    def extract_text(self):
        path = self.file_input.text()
        begin, end = self.begin_spin_of_et.value(), self.end_spin_of_et.value()
        if begin == 0 or end == 0:
            self.show_error("ページ範囲を指定してください。")
            return None
        if not path.strip() or begin < 1 or end < begin:
            self.show_error("ページ範囲またはパスが不正です。")
            return None
        success = self.obj_of_cls.extract_text(path, begin, end)
        self.obj_of_cls.log += self.obj_of_cls.lst_of_text_in_pages
        self.show_result("テキストの抽出", success)
        self.output_log()

    def rotate_page(self):
        path = self.file_input.text()
        page = self.spin_of_rp.value()
        if not path.strip() or page < 1:
            self.show_error("ページ番号が不正です。")
            return None
        success = self.obj_of_cls.rotate_page_clockwise(path, page, 90)
        self.show_result("ページの回転", success)
        self.output_log()

    def show_result(self, label: str, success: bool):
        QMessageBox.information(self, f"{label}の結果", f"{label}に{'成功' if success else '失敗'}しました。")

    def show_error(self, msg: str):
        QMessageBox.information(self, "エラー", msg)

    def output_log(self):
        log = "\n".join(self.obj_of_cls.log)
        self.output.setPlainText(log)

    def write_log(self):
        exe_path = Path(__file__)
        log_path = self.obj_of_pt.get_file_path_of_log(exe_path)
        self.obj_of_cls.write_log(log_path)


def main():
    app = QApplication(sys.argv)
    main = MainApp_Of_PT()
    main.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
