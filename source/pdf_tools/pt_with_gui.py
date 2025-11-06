import logging
import shutil
import sys
from pathlib import Path
from typing import Any

import pypdfium2
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QPixmap
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

from source.common.common import DatetimeTools, GUITools, LogTools, PathTools
from source.pdf_tools.pt_class import PdfTools


# QTextEdit にログを流すためのハンドラ
class QTextEditHandler(logging.Handler):
    def __init__(self, widget: QTextEdit):
        super().__init__()
        self.widget: QTextEdit = widget

    def emit(self, record: logging.LogRecord):
        msg: str = self.format(record)
        self.widget.append(msg)


class MainApp_Of_PT(QMainWindow):
    def __init__(self):
        """初期化します"""
        super().__init__()
        self.obj_of_lt: LogTools = LogTools()
        self.obj_of_cls: PdfTools = PdfTools(self.obj_of_lt.logger)
        self._setup_first_ui()
        self.obj_of_pt: PathTools = PathTools()
        self.obj_of_dt2: DatetimeTools = DatetimeTools()
        self._setup_log()

    def closeEvent(self, event):
        """終了します"""
        image_dir: Path = Path(__file__).parent / "__images__"
        # フォルダを空にする
        if image_dir.exists():
            for element in image_dir.iterdir():
                if element.is_dir():
                    shutil.rmtree(element)
                else:
                    element.unlink()
        if self.obj_of_lt:
            self.show_info(f"ログファイルは、\n{self.obj_of_lt.file_path_of_log}\nに出力されました。")
        super().closeEvent(event)

    def show_info(self, msg: str):
        """情報を表示します"""
        QMessageBox.information(self, "情報", msg)
        self.obj_of_lt.logger.info(msg)

    def show_result(self, label: str, success: bool):
        """結果を表示します"""
        QMessageBox.information(self, "結果", f"{label} => {'成功' if success else '失敗'}しました。")
        if success:
            self.obj_of_lt.logger.info(f"{label} => 成功しました。")
        else:
            self.obj_of_lt.logger.error(f"{label} => 失敗しました。")

    def show_error(self, msg: str):
        """エラーを表示します"""
        QMessageBox.warning(self, "エラー", msg)
        self.obj_of_lt.logger.warning(msg)

    def _setup_log(self) -> bool:
        """ログを設定します"""
        result: bool = False
        try:
            # exe化されている場合とそれ以外を切り分ける
            exe_path: Path = Path(sys.executable) if getattr(sys, "frozen", False) else Path(__file__)
            file_of_log_p: Path = self.obj_of_pt._get_file_path_of_log(exe_path)
            self.obj_of_lt.file_path_of_log = str(file_of_log_p)
            self.obj_of_lt._setup_file_handler(self.obj_of_lt.file_path_of_log)
            text_handler: QTextEditHandler = QTextEditHandler(self.log_area)
            text_handler.setFormatter(self.obj_of_lt.file_formatter)
            self.obj_of_lt.logger.addHandler(text_handler)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def _setup_first_ui(self) -> bool:
        """1番目のUser Interfaceを設定します"""
        result: bool = False
        try:
            # タイトル
            self.setWindowTitle("PDF編集アプリ")
            central: QWidget = QWidget()
            self.setCentralWidget(central)
            base_layout: QVBoxLayout = QVBoxLayout(central)
            # 各要素のタイトル
            element_title_area: QHBoxLayout = QHBoxLayout()
            base_layout.addLayout(element_title_area)
            element_title_area.addWidget(QLabel("機能"))
            element_title_area.addWidget(QLabel("ビューワー"))
            element_title_area.addWidget(QLabel("ログ"))
            # 主要
            main_layout: QHBoxLayout = QHBoxLayout()
            base_layout.addLayout(main_layout)
            # 左側
            left_layout: QVBoxLayout = QVBoxLayout()
            left_func_area: QVBoxLayout = QVBoxLayout()
            left_layout.addLayout(left_func_area)
            main_layout.addLayout(left_layout)
            # 中央
            center_layout: QVBoxLayout = QVBoxLayout()
            center_scroll_area: QScrollArea = QScrollArea()
            center_scroll_area.setWidgetResizable(True)
            center_layout.addWidget(center_scroll_area)
            main_layout.addLayout(center_layout)
            # 仮想コンテナのウィジェットとレイアウト
            center_widget: QWidget = QWidget()
            self.center_viewer: QVBoxLayout = QVBoxLayout()
            center_widget.setLayout(self.center_viewer)
            center_scroll_area.setWidget(center_widget)
            # 右側
            right_layout: QVBoxLayout = QVBoxLayout()
            main_layout.addLayout(right_layout)
            self.log_area: QTextEdit = QTextEdit()
            self.log_area.setReadOnly(True)
            right_layout.addWidget(self.log_area)
            # ファイル選択と再読み込み
            self.file_input: QLineEdit = QLineEdit()
            left_func_area.addWidget(QLabel("PDFファイルパス"))
            left_func_area.addWidget(self.file_input)
            select_and_reload_area: QHBoxLayout = QHBoxLayout()
            browse_btn: QPushButton = QPushButton("参照")
            select_and_reload_area.addWidget(browse_btn)
            reload_btn: QPushButton = QPushButton("再読み込み")
            select_and_reload_area.addWidget(reload_btn)
            left_func_area.addLayout(select_and_reload_area)
            browse_btn.clicked.connect(lambda: self.select_pdf(False))
            reload_btn.clicked.connect(lambda: self.select_pdf(True))
            # パスワード入力
            self.password_input: QLineEdit = QLineEdit()
            self.password_input.setPlaceholderText("パスワード（英数字/アンダーバー/ハイフン）")
            left_func_area.addWidget(QLabel("パスワード"))
            left_func_area.addWidget(self.password_input)
            # 暗号化と復号化
            encrypt_and_decrypt_area: QHBoxLayout = QHBoxLayout()
            encrypt_btn: QPushButton = QPushButton("暗号化")
            encrypt_and_decrypt_area.addWidget(encrypt_btn)
            decrypt_btn: QPushButton = QPushButton("復号化")
            encrypt_and_decrypt_area.addWidget(decrypt_btn)
            left_func_area.addLayout(encrypt_and_decrypt_area)
            encrypt_btn.clicked.connect(self.encrypt_pdf)
            decrypt_btn.clicked.connect(self.decrypt_pdf)
            # メタデータの表示
            meta_btn: QPushButton = QPushButton("メタデータの表示")
            left_func_area.addWidget(meta_btn)
            meta_btn.clicked.connect(self.show_metadata)
            # メタデータの入力
            left_func_area.addWidget(QLabel("メタデータの入力"))
            self.widget_of_metadata: dict = {}
            self.line_edits_of_metadata: dict = {}
            for value, key in self.obj_of_cls.fields:
                if value in ["creation_date", "modification_date"]:
                    continue
                self.line_edits_of_metadata[key] = QLineEdit()
                left_func_area.addWidget(QLabel(value.capitalize().replace("_", " ")))
                left_func_area.addWidget(self.line_edits_of_metadata[key])
            # メタデータの書き込み
            write_meta_btn: QPushButton = QPushButton("メタデータの書き込み")
            left_func_area.addWidget(write_meta_btn)
            write_meta_btn.clicked.connect(self.write_metadata)
            # マージ
            merge_btn: QPushButton = QPushButton("複数PDFをマージ")
            left_func_area.addWidget(merge_btn)
            merge_btn.clicked.connect(self.merge_pdfs)
            # ページの抽出
            begin_spin_of_ep: QSpinBox = QSpinBox()
            end_spin_of_ep: QSpinBox = QSpinBox()
            page_layout_of_ep: QHBoxLayout = QHBoxLayout()
            page_layout_of_ep.addWidget(QLabel("ページの抽出開始"))
            page_layout_of_ep.addWidget(begin_spin_of_ep)
            page_layout_of_ep.addWidget(QLabel("ページの抽出終了"))
            page_layout_of_ep.addWidget(end_spin_of_ep)
            left_func_area.addLayout(page_layout_of_ep)
            extract_page_btn: QPushButton = QPushButton("ページの抽出")
            left_func_area.addWidget(extract_page_btn)
            extract_page_btn.clicked.connect(lambda: self.extract_pages(begin_spin_of_ep, end_spin_of_ep))
            # ページの削除
            begin_spin_of_dp: QSpinBox = QSpinBox()
            end_spin_of_dp: QSpinBox = QSpinBox()
            page_layout_of_dp: QHBoxLayout = QHBoxLayout()
            page_layout_of_dp.addWidget(QLabel("ページの削除開始"))
            page_layout_of_dp.addWidget(begin_spin_of_dp)
            page_layout_of_dp.addWidget(QLabel("ページの削除終了"))
            page_layout_of_dp.addWidget(end_spin_of_dp)
            left_func_area.addLayout(page_layout_of_dp)
            delete_page_btn: QPushButton = QPushButton("ページの削除")
            left_func_area.addWidget(delete_page_btn)
            delete_page_btn.clicked.connect(lambda: self.delete_pages(begin_spin_of_dp, end_spin_of_dp))
            # テキストの抽出
            begin_spin_of_et: QSpinBox = QSpinBox()
            end_spin_of_et: QSpinBox = QSpinBox()
            page_layout_of_et: QHBoxLayout = QHBoxLayout()
            page_layout_of_et.addWidget(QLabel("テキストの抽出開始"))
            page_layout_of_et.addWidget(begin_spin_of_et)
            page_layout_of_et.addWidget(QLabel("テキストの抽出終了"))
            page_layout_of_et.addWidget(end_spin_of_et)
            left_func_area.addLayout(page_layout_of_et)
            extract_text_btn: QPushButton = QPushButton("テキストの抽出")
            left_func_area.addWidget(extract_text_btn)
            extract_text_btn.clicked.connect(lambda: self.extract_text(begin_spin_of_et, end_spin_of_et))
            # ページの回転
            spin_of_rp: QSpinBox = QSpinBox()
            page_layout_of_rp: QHBoxLayout = QHBoxLayout()
            page_layout_of_rp.addWidget(QLabel("回転するページ"))
            page_layout_of_rp.addWidget(spin_of_rp)
            rotate_btn: QPushButton = QPushButton("ページを時計回りに回転（90度）")
            page_layout_of_rp.addWidget(rotate_btn)
            left_func_area.addLayout(page_layout_of_rp)
            rotate_btn.clicked.connect(lambda: self.rotate_page(spin_of_rp))
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def _get_images(self, file_path: str) -> list:
        """ビューワーに表示するファイルの各ページの画像を取得します"""
        output_files: list = []
        try:
            file_p: Path = Path(file_path)
            file_name_s = file_p.stem
            self.output_dir = Path(__file__).parent / "__images__"
            self.output_dir.mkdir(parents=True, exist_ok=True)
            pdf = pypdfium2.PdfDocument(file_path)
            for i, page in enumerate(pdf):
                pil_image: Any = page.render(scale=1, rotation=0, crop=(0, 0, 0, 0)).to_pil()
                output_file = self.output_dir / f"{file_name_s}_{i + 1}.png"
                pil_image.save(output_file)
                output_files.append(str(output_file))
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            pass
        finally:
            pass
        return output_files

    def _setup_second_ui(self, images: list) -> bool:
        """2番目のUser Interfaceを設定します"""
        result: bool = False
        try:
            # 既存のレイアウトをクリア（再表示に対応）
            for i in reversed(range(self.center_viewer.count())):
                widget: QWidget = self.center_viewer.itemAt(i).widget()
                if widget is not None:
                    widget.setParent(None)
            if not images:
                raise Exception("PDFファイルの各ページの画像がありません。")
            # 各ページを表示する
            for i, element in enumerate(images):
                # 垂直レイアウトを用意する
                page_layout: QVBoxLayout = QVBoxLayout()
                page_layout.setAlignment(Qt.AlignCenter)
                page_widget: QWidget = QWidget()
                page_widget.setLayout(page_layout)
                # ページ番号のラベル
                page_num_label: QLabel = QLabel(f"page: {i + 1}\n")
                page_num_label.setAlignment(Qt.AlignCenter)
                page_layout.addWidget(page_num_label)
                # 画像のラベル
                image_label: QLabel = QLabel()
                pixmap: QPixmap = QPixmap(element).scaledToWidth(300, Qt.SmoothTransformation)
                image_label.setPixmap(pixmap)
                image_label.setScaledContents(True)
                image_label.setFixedSize(pixmap.size())
                page_layout.addWidget(image_label)
                # グリッドにページごとに追加する
                self.center_viewer.addWidget(page_widget)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def select_pdf(self, reload: bool) -> bool:
        """選択します"""
        result: bool = False
        try:
            if reload:
                if self.obj_of_cls.file_path == "":
                    raise Exception("PDFファイルを選択してください。")
            else:
                self.obj_of_cls.file_path, _ = QFileDialog.getOpenFileName(self, "PDFファイルを選択", "", "PDF Files (*.pdf)")
            # 各OSに応じたパス区切りに変換する
            file_p: Path = Path(self.obj_of_cls.file_path).expanduser()
            self.obj_of_cls.file_path = str(file_p)
            self.file_input.setText(self.obj_of_cls.file_path)
            # 読み込む
            self.obj_of_cls.read_file()
            images: list = self._get_images(self.obj_of_cls.file_path)
            self._setup_second_ui(images)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def encrypt_pdf(self) -> bool:
        """暗号化します"""
        result: bool = False
        try:
            self.obj_of_cls.file_path, pw = self.file_input.text(), self.password_input.text()
            if self.obj_of_cls.file_path == "" or pw == "":
                # 未入力の場合
                raise Exception("PDFファイルを選択し、パスワードを入力してください。")
            self.obj_of_cls.encrypt(pw)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            self.show_result(self.encrypt_pdf.__doc__, result)
        return result

    def decrypt_pdf(self) -> bool:
        """復号化します"""
        result: bool = False
        try:
            self.obj_of_cls.file_path, pw = self.file_input.text(), self.password_input.text()
            if self.obj_of_cls.file_path == "" or pw == "":
                # 未入力の場合
                raise Exception("PDFファイルを選択し、パスワードを入力してください。")
            self.obj_of_cls.decrypt(pw)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            self.show_result(self.decrypt_pdf.__doc__, result)
        return result

    def show_metadata(self) -> bool:
        """メタデータを表示します"""
        result: bool = False
        try:
            self.obj_of_cls.file_path = self.file_input.text()
            if self.obj_of_cls.file_path == "":
                raise Exception("PDFファイルを選択してください。")
            # 再読み込み
            self.select_pdf(True)
            self.obj_of_cls.get_metadata()
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def write_metadata(self) -> bool:
        """メタデータを書き込みます"""
        result: bool = False
        try:
            self.obj_of_cls.file_path = self.file_input.text()
            if self.obj_of_cls.file_path == "":
                raise Exception("PDFファイルを選択してください。")
            # 再読み込み
            self.select_pdf(True)
            for value, key in self.obj_of_cls.fields:
                match value:
                    case "creation_date":
                        self.widget_of_metadata[key] = self.obj_of_cls.creation_date
                    case "modification_date":
                        self.widget_of_metadata[key] = self.obj_of_dt2._convert_for_metadata_in_pdf(self.obj_of_cls.UTC_OF_JP)
                    case _:
                        self.widget_of_metadata[key] = self.line_edits_of_metadata[key].text()
            self.obj_of_cls.write_metadata(self.widget_of_metadata)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def merge_pdfs(self) -> bool:
        """マージします"""
        result: bool = False
        try:
            files, _ = QFileDialog.getOpenFileNames(self, "マージするPDFを選択", "", "PDF Files (*.pdf)")
            if not files:
                raise Exception("PDFファイルを選択してください。")
            # 各OSに応じたパス区切りに変換する
            for i, element in enumerate(files):
                files[i] = str(Path(element))
            self.obj_of_cls.merge(files)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def extract_pages(self, b_spin: QSpinBox, e_spin: QSpinBox) -> bool:
        """ページを抽出します"""
        result: bool = False
        try:
            self.obj_of_cls.file_path = self.file_input.text()
            if self.obj_of_cls.file_path == "":
                raise Exception("PDFファイルを選択してください。")
            # 再読み込み
            self.select_pdf(True)
            begin, end = b_spin.value(), e_spin.value()
            if begin == 0 or end == 0:
                raise Exception("ページ範囲を指定してください。")
            if begin < 1 or end > self.obj_of_cls.num_of_pages or end < begin:
                raise Exception("ページ範囲またはパスが不正です。")
            self.obj_of_cls.extract_pages(begin, end)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def delete_pages(self, b_spin: QSpinBox, e_spin: QSpinBox) -> bool:
        """ページを削除します"""
        result: bool = False
        try:
            self.obj_of_cls.file_path = self.file_input.text()
            if self.obj_of_cls.file_path == "":
                raise Exception("PDFファイルを選択してください。")
            # 再読み込み
            self.select_pdf(True)
            begin, end = b_spin.value(), e_spin.value()
            if begin == 0 or end == 0:
                raise Exception("ページ範囲を指定してください。")
            if begin < 1 or end > self.obj_of_cls.num_of_pages or end < begin:
                raise Exception("ページ範囲またはパスが不正です。")
            self.obj_of_cls.delete_pages(begin, end)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def extract_text(self, b_spin: QSpinBox, e_spin: QSpinBox) -> bool:
        """テキストを抽出します"""
        result: bool = False
        try:
            self.obj_of_cls.file_path = self.file_input.text()
            if self.obj_of_cls.file_path == "":
                raise Exception("PDFファイルを選択してください。")
            # 再読み込み
            self.select_pdf(True)
            begin, end = b_spin.value(), e_spin.value()
            if begin == 0 or end == 0:
                raise Exception("ページ範囲を指定してください。")
            if begin < 1 or end > self.obj_of_cls.num_of_pages or end < begin:
                raise Exception("ページ範囲またはパスが不正です。")
            self.obj_of_cls.extract_text(begin, end)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def rotate_page(self, spin: QSpinBox) -> bool:
        """ページを回転します"""
        result: bool = False
        try:
            self.obj_of_cls.file_path = self.file_input.text()
            if self.obj_of_cls.file_path == "":
                raise Exception("PDFファイルを選択してください。")
            # 再読み込み
            self.select_pdf(True)
            page: int = spin.value()
            if page == 0:
                raise Exception("ページ範囲を指定してください。")
            if page < 1 or page > self.obj_of_cls.num_of_pages:
                raise Exception("ページ番号が不正です。")
            self.obj_of_cls.rotate_page_clockwise(page, 90)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
            images: list = self._get_images(self.obj_of_cls.file_path)
            self._setup_second_ui(images)
        finally:
            pass
        return result


def main() -> bool:
    """主要関数"""
    result: bool = False
    try:
        obj_of_gt: GUITools = GUITools()
        app: QApplication = QApplication(sys.argv)
        # アプリ単位でフォントを設定する
        font: QFont = QFont()
        font.setPointSize(12)
        app.setFont(font)
        window: MainApp_Of_PT = MainApp_Of_PT()
        window.resize(1000, 800)
        # 最大化して、表示させる
        window.showMaximized()
        sys.exit(app.exec())
    except Exception as e:
        obj_of_gt._show_start_up_error(f"error: \n{str(e)}")
    else:
        result = True
    finally:
        pass
    return result


if __name__ == "__main__":
    main()
