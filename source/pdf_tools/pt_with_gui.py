import os
import platform
import sys
from pathlib import Path

import pypdfium2
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
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

from source.common.common import DatetimeTools, FileSystemTools, PathTools
from source.pdf_tools.pt_class import PdfTools


class MainApp_Of_PT(QMainWindow):
    def __init__(self):
        """初期化します"""
        super().__init__()

        self.obj_of_pt = PathTools()
        self.obj_of_dt2 = DatetimeTools()
        self.obj_of_fst = FileSystemTools()
        self.obj_of_cls = PdfTools()

        self.file_path = ""

        self.first_init_ui()

    def closeEvent(self, event):
        """終了します"""
        image_dir = Path(__file__).parent / "__images__"
        self.obj_of_fst.clear_folder(image_dir)
        self.write_log()
        super().closeEvent(event)

    def first_init_ui(self):
        """1回目のUser Interfaceを設定します"""
        # タイトル
        self.setWindowTitle("PDF編集アプリ")
        central = QWidget()
        self.setCentralWidget(central)
        # 主要
        main_layout = QHBoxLayout(central)
        # 左側
        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("機能"))
        left_func_area = QVBoxLayout()
        left_layout.addLayout(left_func_area)
        main_layout.addLayout(left_layout)
        # 中央
        center_layout = QVBoxLayout()
        center_layout.addWidget(QLabel("ビューワー"))
        center_scroll_area = QScrollArea()
        center_layout.addWidget(center_scroll_area)
        center_scroll_area.setWidgetResizable(True)
        main_layout.addLayout(center_layout)
        # 仮想コンテナのウィジェットとレイアウト
        center_widget = QWidget()
        self.center_viewer = QVBoxLayout()
        center_widget.setLayout(self.center_viewer)
        center_scroll_area.setWidget(center_widget)
        # 右側
        right_layout = QVBoxLayout()
        main_layout.addLayout(right_layout)
        right_layout.addWidget(QLabel("ログ"))
        self.output = QTextEdit()
        self.output.setReadOnly(True)
        right_layout.addWidget(self.output)

        # ファイル選択と再読み込み
        self.file_input = QLineEdit()
        left_func_area.addWidget(QLabel("PDFファイルパス"))
        left_func_area.addWidget(self.file_input)
        select_and_reload_area = QHBoxLayout()
        browse_btn = QPushButton("参照")
        select_and_reload_area.addWidget(browse_btn)
        browse_btn.clicked.connect(lambda: self.select_pdf(False))
        reload_btn = QPushButton("再読み込み")
        select_and_reload_area.addWidget(reload_btn)
        reload_btn.clicked.connect(lambda: self.select_pdf(True))
        left_func_area.addLayout(select_and_reload_area)

        # パスワード入力
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("パスワード（英数字/アンダーバー/ハイフン）")
        left_func_area.addWidget(QLabel("パスワード"))
        left_func_area.addWidget(self.password_input)

        # 暗号化と復号化
        encrypt_and_decrypt_area = QHBoxLayout()
        encrypt_btn = QPushButton("暗号化")
        encrypt_and_decrypt_area.addWidget(encrypt_btn)
        encrypt_btn.clicked.connect(self.encrypt_pdf)
        decrypt_btn = QPushButton("復号化")
        encrypt_and_decrypt_area.addWidget(decrypt_btn)
        decrypt_btn.clicked.connect(self.decrypt_pdf)
        left_func_area.addLayout(encrypt_and_decrypt_area)

        # メタデータの表示
        meta_btn = QPushButton("メタデータの表示")
        meta_btn.clicked.connect(self.show_metadata)
        left_func_area.addWidget(meta_btn)

        # メタデータの入力
        left_func_area.addWidget(QLabel("メタデータの入力"))
        self.widget_of_metadata = {}
        self.line_edits_of_metadata = {}
        for value, key in self.obj_of_cls.fields:
            if value in ["creation_date", "modification_date"]:
                continue
            self.line_edits_of_metadata[key] = QLineEdit()
            left_func_area.addWidget(QLabel(value.capitalize().replace("_", " ")))
            left_func_area.addWidget(self.line_edits_of_metadata[key])

        # メタデータの書き込み
        write_meta_btn = QPushButton("メタデータの書き込み")
        write_meta_btn.clicked.connect(self.write_metadata)
        left_func_area.addWidget(write_meta_btn)

        # マージ
        merge_btn = QPushButton("複数PDFをマージ")
        merge_btn.clicked.connect(self.merge_pdfs)
        left_func_area.addWidget(merge_btn)

        # ページの抽出
        begin_spin_of_ep = QSpinBox()
        end_spin_of_ep = QSpinBox()
        page_layout_of_ep = QHBoxLayout()
        page_layout_of_ep.addWidget(QLabel("ページの抽出開始"))
        page_layout_of_ep.addWidget(begin_spin_of_ep)
        page_layout_of_ep.addWidget(QLabel("ページの抽出終了"))
        page_layout_of_ep.addWidget(end_spin_of_ep)
        left_func_area.addLayout(page_layout_of_ep)

        extract_page_btn = QPushButton("ページの抽出")
        extract_page_btn.clicked.connect(lambda: self.extract_pages(begin_spin_of_ep, end_spin_of_ep))
        left_func_area.addWidget(extract_page_btn)

        # ページの削除
        begin_spin_of_dp = QSpinBox()
        end_spin_of_dp = QSpinBox()
        page_layout_of_dp = QHBoxLayout()
        page_layout_of_dp.addWidget(QLabel("ページの削除開始"))
        page_layout_of_dp.addWidget(begin_spin_of_dp)
        page_layout_of_dp.addWidget(QLabel("ページの削除終了"))
        page_layout_of_dp.addWidget(end_spin_of_dp)
        left_func_area.addLayout(page_layout_of_dp)

        delete_page_btn = QPushButton("ページの削除")
        delete_page_btn.clicked.connect(lambda: self.delete_pages(begin_spin_of_dp, end_spin_of_dp))
        left_func_area.addWidget(delete_page_btn)

        # テキストの抽出
        begin_spin_of_et = QSpinBox()
        end_spin_of_et = QSpinBox()
        page_layout_of_et = QHBoxLayout()
        page_layout_of_et.addWidget(QLabel("テキストの抽出開始"))
        page_layout_of_et.addWidget(begin_spin_of_et)
        page_layout_of_et.addWidget(QLabel("テキストの抽出終了"))
        page_layout_of_et.addWidget(end_spin_of_et)
        left_func_area.addLayout(page_layout_of_et)

        extract_text_btn = QPushButton("テキストの抽出")
        extract_text_btn.clicked.connect(lambda: self.extract_text(begin_spin_of_et, end_spin_of_et))
        left_func_area.addWidget(extract_text_btn)

        # ページの回転
        spin_of_rp = QSpinBox()
        page_layout_of_rp = QHBoxLayout()
        page_layout_of_rp.addWidget(QLabel("回転するページ"))
        page_layout_of_rp.addWidget(spin_of_rp)
        rotate_btn = QPushButton("ページを時計回りに回転（90度）")
        rotate_btn.clicked.connect(lambda: self.rotate_page(spin_of_rp))
        page_layout_of_rp.addWidget(rotate_btn)
        left_func_area.addLayout(page_layout_of_rp)

    def second_init_ui(self, images: list):
        """2回目のUser Interfaceを設定します"""
        # 既存のレイアウトをクリア（再表示に対応）
        for i in reversed(range(self.center_viewer.count())):
            widget = self.center_viewer.itemAt(i).widget()
            if widget is not None:
                widget.setParent(None)
        if images is None:
            return None
        # 各ページを表示する
        for i, element in enumerate(images):
            # 垂直レイアウトを用意する
            page_layout = QVBoxLayout()
            page_layout.setAlignment(Qt.AlignCenter)
            page_widget = QWidget()
            page_widget.setLayout(page_layout)
            # ページ番号のラベル
            page_num_label = QLabel(f"page: {i + 1}\n")
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
            self.center_viewer.addWidget(page_widget)

    def select_pdf(self, reload: bool):
        """選択します"""
        if not reload:
            self.file_path, _ = QFileDialog.getOpenFileName(self, "PDFファイルを選択", "", "PDF Files (*.pdf)")
        if self.file_path.strip():
            # 各OSに応じたパス区切りに変換する
            file_as_path_type = Path(self.file_path).expanduser()
            self.file_path = str(file_as_path_type)
            self.file_input.setText(self.file_path)
            self.obj_of_cls.read_file(self.file_path)
            if self.obj_of_cls.reader.is_encrypted:
                self.show_error("ファイルが暗号化されています。")
            else:
                _, images = self.get_images(self.file_path)
                self.second_init_ui(images)
        else:
            self.show_error("ファイルを選択してください。")
        self.output_log()

    def get_images(self, file_path: str) -> list:
        """ビューワーに表示するファイルの各ページの画像を取得します"""
        try:
            result = False
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
        except Exception as e:
            self.show_error(str(e))
        else:
            result = True
        finally:
            return [result, output_files]

    def encrypt_pdf(self):
        """暗号化します"""
        self.file_path, pw = self.file_input.text(), self.password_input.text()
        if not self.file_path.strip() or not pw.strip():
            self.show_error("ファイルパスとパスワードを入力してください。")
            return None
        result, log = self.obj_of_cls.encrypt(self.file_path, pw)
        self.show_result("暗号化", result)
        self.output_log()

    def decrypt_pdf(self):
        """復号化します"""
        self.file_path, pw = self.file_input.text(), self.password_input.text()
        if not self.file_path.strip() or not pw.strip():
            self.show_error("ファイルパスとパスワードを入力してください。")
            return None
        result, log = self.obj_of_cls.decrypt(self.file_path, pw)
        self.show_result("復号化", result)
        self.output_log()

    def show_metadata(self):
        """メタデータを表示します"""
        self.file_path = self.file_input.text()
        if not self.file_path.strip():
            self.show_error("PDFファイルパスを入力してください。")
        result, log = self.obj_of_cls.get_metadata(self.file_path)
        self.show_result("メタデータの表示", result)
        self.output_log()

    def write_metadata(self):
        """メタデータを書き込みます"""
        self.file_path = self.file_input.text()
        if not self.file_path.strip():
            self.show_error("PDFファイルパスを入力してください。")
            return None
        for value, key in self.obj_of_cls.fields:
            match value:
                case "creation_date":
                    self.widget_of_metadata[key] = self.obj_of_cls.creation_date
                case "modification_date":
                    self.widget_of_metadata[key] = self.obj_of_dt2.convert_for_metadata_in_pdf(self.obj_of_cls.UTC_OF_JP)
                case _:
                    self.widget_of_metadata[key] = self.line_edits_of_metadata[key].text()
        result, log = self.obj_of_cls.write_metadata(self.file_path, self.widget_of_metadata)
        self.show_result("メタデータの書き込み", result)
        self.output_log()

    def merge_pdfs(self):
        """マージします"""
        files, _ = QFileDialog.getOpenFileNames(self, "マージするPDFを選択", "", "PDF Files (*.pdf)")
        if files:
            # 各OSに応じたパス区切りに変換する
            for i, element in enumerate(files):
                files[i] = str(Path(element))
            result, log = self.obj_of_cls.merge(files)
            self.show_result("マージ", result)
        self.output_log()

    def extract_pages(self, b_spin: QSpinBox, e_spin: QSpinBox):
        """ページを抽出します"""
        self.file_path = self.file_input.text()
        begin, end = b_spin.value(), e_spin.value()
        if begin == 0 or end == 0:
            self.show_error("ページ範囲を指定してください。")
            return None
        if not self.file_path.strip() or begin < 1 or end < begin:
            self.show_error("ページ範囲またはパスが不正です。")
            return None
        result, log = self.obj_of_cls.extract_pages(self.file_path, begin, end)
        self.show_result("ページの抽出", result)
        self.output_log()

    def delete_pages(self, b_spin: QSpinBox, e_spin: QSpinBox):
        """ページを削除します"""
        self.file_path = self.file_input.text()
        begin, end = b_spin.value(), e_spin.value()
        if begin == 0 or end == 0:
            self.show_error("ページ範囲を指定してください。")
            return None
        if not self.file_path.strip() or begin < 1 or end < begin:
            self.show_error("ページ範囲またはパスが不正です。")
            return None
        result, log = self.obj_of_cls.delete_pages(self.file_path, begin, end)
        self.show_result("ページの削除", result)
        self.output_log()

    def extract_text(self, b_spin: QSpinBox, e_spin: QSpinBox):
        """テキストを抽出します"""
        self.file_path = self.file_input.text()
        begin, end = b_spin.value(), e_spin.value()
        if begin == 0 or end == 0:
            self.show_error("ページ範囲を指定してください。")
            return None
        if not self.file_path.strip() or begin < 1 or end < begin:
            self.show_error("ページ範囲またはパスが不正です。")
            return None
        result, log = self.obj_of_cls.extract_text(self.file_path, begin, end)
        self.show_result("テキストの抽出", result)
        self.output_log()

    def rotate_page(self, spin: QSpinBox):
        """ページを回転します"""
        self.file_path = self.file_input.text()
        page = spin.value()
        if not self.file_path.strip() or page < 1:
            self.show_error("ページ番号が不正です。")
            return None
        result, log = self.obj_of_cls.rotate_page_clockwise(self.file_path, page, 90)
        self.show_result("ページの回転", result)
        self.output_log()
        _, images = self.get_images(self.file_path)
        self.second_init_ui(images)

    def write_log(self):
        """ログを書き出す"""
        # exe化されている場合とそれ以外を切り分ける
        if getattr(sys, "frozen", False):
            exe_path = Path(sys.executable)
        else:
            exe_path = Path(__file__)
        log_path = self.obj_of_pt.get_file_path_of_log(exe_path)
        result, path = self.obj_of_cls.write_log(log_path)
        self.show_result(f"ログファイル({path})の出力", result)

    def show_result(self, label: str, success: bool):
        """結果を表示します"""
        QMessageBox.information(self, "結果", f"{label}に{'成功' if success else '失敗'}しました。")

    def show_error(self, msg: str):
        """エラーを表示します"""
        QMessageBox.information(self, "エラー", msg)

    def output_log(self):
        """ログを表示します"""
        formatted_log = "\n".join(self.obj_of_cls.log)
        self.output.setPlainText(formatted_log)


def main():
    """主要関数"""
    if platform.system().lower() == "windows":
        # アプリ全体のスケール
        os.environ["QT_SCALE_FACTOR"] = "0.7"
    app = QApplication(sys.argv)
    window = MainApp_Of_PT()
    window.resize(1600, 800)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
