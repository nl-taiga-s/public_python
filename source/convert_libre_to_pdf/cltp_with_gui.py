import logging
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path

from PySide6.QtGui import QFont
from PySide6.QtWidgets import QApplication, QFileDialog, QLabel, QListWidget, QMessageBox, QProgressBar, QPushButton, QTextEdit, QVBoxLayout, QWidget

from source.common.common import GUITools, LogTools, PathTools, PlatformTools
from source.convert_libre_to_pdf.cltp_class import ConvertLibreToPDF


# QTextEdit にログを流すためのハンドラ
class QTextEditHandler(logging.Handler):
    def __init__(self, widget: QTextEdit):
        super().__init__()
        self.widget = widget

    def emit(self, record: logging.LogRecord):
        msg = self.format(record)
        self.widget.append(msg)


class MainApp_Of_CLTP(QWidget):
    def __init__(self):
        """初期化します"""
        super().__init__()
        self.obj_of_lt = LogTools()
        self.obj_of_cls = ConvertLibreToPDF(self.obj_of_lt.logger)
        self.setup_ui()
        self.obj_of_pft = PlatformTools()
        self.obj_of_pt = PathTools()
        self.setup_log()

    def closeEvent(self, event):
        """終了します"""
        if self.obj_of_lt:
            self.show_info(f"ログファイルは、\n{self.obj_of_lt.file_path_of_log}\nに出力されました。")
        super().closeEvent(event)

    def show_info(self, msg: str):
        """情報を表示します"""
        QMessageBox.information(self, "情報", msg)

    def show_result(self, label: str, success: bool):
        """結果を表示します"""
        QMessageBox.information(self, "結果", f"{label}に{'成功' if success else '失敗'}しました。")

    def show_error(self, msg: str):
        """エラーを表示します"""
        QMessageBox.critical(self, "エラー", msg)

    def setup_log(self) -> bool:
        """ログを設定する"""
        try:
            result = False
            # exe化されている場合とそれ以外を切り分ける
            if getattr(sys, "frozen", False):
                exe_path = Path(sys.executable)
            else:
                exe_path = Path(__file__)
            file_of_log_as_path_type = self.obj_of_pt.get_file_path_of_log(exe_path)
            self.obj_of_lt.file_path_of_log = str(file_of_log_as_path_type)
            if not self.obj_of_lt.setup_file_handler(self.obj_of_lt.file_path_of_log):
                raise
            text_handler = QTextEditHandler(self.log_area)
            text_handler.setFormatter(self.obj_of_lt.file_formatter)
            self.obj_of_lt.logger.addHandler(text_handler)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            return result

    def setup_ui(self) -> bool:
        """User Interfaceを設定します"""
        try:
            result = False
            # タイトル
            self.setWindowTitle("Officeファイル => PDF 一括変換アプリ")
            # ウィジェット
            self.label_from = QLabel("変換元フォルダ: 未選択")
            btn_select_from = QPushButton("変換元フォルダを選択")
            btn_open_from = QPushButton("変換元フォルダを開く")
            self.label_to = QLabel("変換先フォルダ: 未選択")
            btn_select_to = QPushButton("変換先フォルダを選択")
            btn_open_to = QPushButton("変換先フォルダを開く")
            self.file_list_widget = QListWidget()
            self.progress_bar = QProgressBar()
            btn_convert = QPushButton("一括変換を実行します")
            self.log_area = QTextEdit()
            self.log_area.setReadOnly(True)
            # レイアウト
            layout = QVBoxLayout()
            layout.addWidget(self.label_from)
            layout.addWidget(btn_select_from)
            layout.addWidget(btn_open_from)
            layout.addWidget(self.label_to)
            layout.addWidget(btn_select_to)
            layout.addWidget(btn_open_to)
            layout.addWidget(QLabel("変換対象ファイル一覧: "))
            layout.addWidget(self.file_list_widget)
            layout.addWidget(QLabel("進行状況: "))
            layout.addWidget(self.progress_bar)
            layout.addWidget(btn_convert)
            layout.addWidget(QLabel("ログ: "))
            layout.addWidget(self.log_area)
            self.setLayout(layout)
            # シグナル接続
            btn_select_from.clicked.connect(self.select_folder_from)
            btn_open_from.clicked.connect(lambda: self.open_explorer(self.obj_of_cls.folder_path_from))
            btn_select_to.clicked.connect(self.select_folder_to)
            btn_open_to.clicked.connect(lambda: self.open_explorer(self.obj_of_cls.folder_path_to))
            btn_convert.clicked.connect(self.convert_file)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            return result

    def select_folder_from(self) -> bool:
        """変換元のフォルダを選択します"""
        try:
            result = False
            self.obj_of_cls.folder_path_from = QFileDialog.getExistingDirectory(self, "変換元のフォルダを選択")
            folder_as_path_type = Path(self.obj_of_cls.folder_path_from).expanduser()
            self.obj_of_cls.folder_path_from = str(folder_as_path_type)
            if self.obj_of_cls.folder_path_from:
                self.label_from.setText(f"変換元フォルダ: {self.obj_of_cls.folder_path_from}")
                if self.obj_of_cls.folder_path_to:
                    self.show_file_list()
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            return result

    def select_folder_to(self) -> bool:
        """変換先のフォルダを選択します"""
        try:
            result = False
            self.obj_of_cls.folder_path_to = QFileDialog.getExistingDirectory(self, "変換先のフォルダを選択")
            folder_as_path_type = Path(self.obj_of_cls.folder_path_to).expanduser()
            self.obj_of_cls.folder_path_to = str(folder_as_path_type)
            if self.obj_of_cls.folder_path_to:
                self.label_to.setText(f"変換先フォルダ: {self.obj_of_cls.folder_path_to}")
                if self.obj_of_cls.folder_path_from:
                    self.show_file_list()
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            return result

    def open_explorer(self, folder_path: str) -> bool:
        """エクスプローラーを開きます"""
        EXPLORER = "/mnt/c/Windows/explorer.exe"
        try:
            result = False
            if not folder_path:
                raise Exception("フォルダを選択してください。")
            if platform.system().lower() == "windows":
                os.startfile(folder_path)
            elif self.obj_of_pft.is_wsl():
                # Windowsのパスに変換（/mnt/c/... 形式）
                wsl_path = subprocess.check_output(["wslpath", "-w", folder_path]).decode("utf-8").strip()
                subprocess.run([EXPLORER, wsl_path])
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            return result

    def show_file_list(self) -> bool:
        """ファイル一覧を表示します"""
        try:
            result = False
            if not self.obj_of_cls.folder_path_from or not self.obj_of_cls.folder_path_to:
                raise Exception("変換元と変換先のフォルダを選択してください。")
            self.obj_of_cls.create_file_list()
            self.file_list_widget.clear()
            for f in self.obj_of_cls.filtered_list_of_f:
                file_as_path_type = Path(f)
                file_path = file_as_path_type.name
                self.file_list_widget.addItem(file_path)
            self.progress_bar.setValue(0)
        except Exception as e:
            self.file_list_widget.clear()
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            return result

    def convert_file(self) -> bool:
        """変換します"""
        try:
            result = False
            if not self.obj_of_cls.filtered_list_of_f:
                raise Exception("ファイルリストが初期化されていません。")
            self.progress_bar.setRange(0, self.obj_of_cls.number_of_f)
            for i in range(self.obj_of_cls.number_of_f):
                self.obj_of_cls.convert_file()
                self.progress_bar.setValue(i + 1)
                if self.obj_of_cls.complete:
                    break
                self.obj_of_cls.move_to_next_file()
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            return result


def main() -> bool:
    """主要関数"""
    try:
        result = False
        obj_of_gt = GUITools()
        app = QApplication(sys.argv)
        # LibreOfficeのコマンド
        LIBRE_COMMAND = "soffice"
        if not shutil.which(LIBRE_COMMAND):
            raise ImportError("各OSのLibreOfficeをインストールしてください。\nhttps://ja.libreoffice.org/")
        # アプリ単位でフォントを設定する
        font = QFont()
        font.setPointSize(12)
        app.setFont(font)
        window = MainApp_Of_CLTP()
        window.resize(1000, 800)
        # 最大化して、表示させる
        window.showMaximized()
        sys.exit(app.exec())
    except ImportError as e:
        obj_of_gt.show_error(f"error: \n{str(e)}")
    except Exception as e:
        obj_of_gt.show_error(f"error: \n{str(e)}")
    else:
        result = True
    finally:
        return result


if __name__ == "__main__":
    main()
