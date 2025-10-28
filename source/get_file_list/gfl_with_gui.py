import logging
import platform
import subprocess
import sys
from pathlib import Path

from PySide6.QtGui import QFont
from PySide6.QtWidgets import QApplication, QCheckBox, QFileDialog, QLabel, QLineEdit, QMessageBox, QPushButton, QTextEdit, QVBoxLayout, QWidget

from source.common.common import GUITools, LogTools, PathTools, PlatformTools
from source.get_file_list.gfl_class import GetFileList


# QTextEdit にログを流すためのハンドラ
class QTextEditHandler(logging.Handler):
    def __init__(self, widget: QTextEdit):
        super().__init__()
        self.widget: QTextEdit = widget

    def emit(self, record: logging.LogRecord):
        msg: str = self.format(record)
        self.widget.append(msg)


class MainApp_Of_GFL(QWidget):
    def __init__(self):
        """初期化します"""
        super().__init__()
        self.obj_of_lt: LogTools = LogTools()
        self.obj_of_cls: GetFileList = GetFileList(self.obj_of_lt.logger)
        self.setup_ui()
        self.obj_of_pt: PathTools = PathTools()
        self.obj_of_pft: PlatformTools = PlatformTools()
        self.setup_log()

    def closeEvent(self, event):
        """終了します"""
        if self.obj_of_lt:
            self.show_info(f"ログファイルは、\n{self.obj_of_lt.file_path_of_log}\nに出力されました。")
        super().closeEvent(event)

    def show_info(self, msg: str):
        """情報を表示します"""
        QMessageBox.information(self, "情報", msg)
        self.obj_of_lt.logger.info(msg)

    def show_result(self, label: str, success: bool):
        """結果を表示します"""
        QMessageBox.information(self, "結果", f"{label}に{'成功' if success else '失敗'}しました。")
        if success:
            self.obj_of_lt.logger.info(f"{label}に成功しました。")
        else:
            self.obj_of_lt.logger.error(f"{label}に失敗しました。")

    def show_error(self, msg: str):
        """エラーを表示します"""
        QMessageBox.critical(self, "エラー", msg)
        self.obj_of_lt.logger.warning(msg)

    def setup_log(self) -> bool:
        """ログを設定します"""
        result: bool = False
        try:
            # exe化されている場合とそれ以外を切り分ける
            exe_path: Path = Path(sys.executable) if getattr(sys, "frozen", False) else Path(__file__)
            file_of_log_p: Path = self.obj_of_pt.get_file_path_of_log(exe_path)
            self.obj_of_lt.file_path_of_log = str(file_of_log_p)
            self.obj_of_lt.setup_file_handler(self.obj_of_lt.file_path_of_log)
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

    def setup_ui(self) -> bool:
        """User Interfaceを設定します"""
        result: bool = False
        try:
            # タイトル
            self.setWindowTitle("ファイル検索アプリ")
            # ウィジェット
            self.folder_label: QLabel = QLabel("フォルダが未選択")
            select_folder_btn: QPushButton = QPushButton("フォルダを選択")
            self.recursive_checkbox: QCheckBox = QCheckBox("サブフォルダも含めて検索（再帰）")
            self.pattern_input: QLineEdit = QLineEdit()
            self.pattern_input.setPlaceholderText("検索パターンを入力...")
            open_folder_btn: QPushButton = QPushButton("フォルダを開く")
            search_btn: QPushButton = QPushButton("検索実行")
            self.log_area: QTextEdit = QTextEdit()
            # レイアウト
            layout: QVBoxLayout = QVBoxLayout()
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
            select_folder_btn.clicked.connect(lambda: self.select_folder(False))
            open_folder_btn.clicked.connect(self.open_explorer)
            search_btn.clicked.connect(self.search_files)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def select_folder(self, reload: bool) -> bool:
        """フォルダを選択します"""
        result: bool = False
        try:
            if reload:
                if self.obj_of_cls.folder_path == "":
                    raise Exception("フォルダを選択してください。")
            else:
                self.obj_of_cls.folder_path = QFileDialog.getExistingDirectory(self, "フォルダを選択")
            folder_p: Path = Path(self.obj_of_cls.folder_path).expanduser()
            self.obj_of_cls.folder_path = str(folder_p)
            if self.obj_of_cls.folder_path:
                self.folder_label.setText(self.obj_of_cls.folder_path)
                self.obj_of_cls.recursive = self.recursive_checkbox.isChecked()
                self.obj_of_cls.search_recursively()
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def open_explorer(self) -> bool:
        """エクスプローラーを開きます"""
        result: bool = False
        EXPLORER_OF_WSL: str = "/mnt/c/Windows/explorer.exe"
        try:
            if self.obj_of_cls.folder_path == "":
                raise Exception("フォルダを選択してください。")
            if platform.system().lower() == "windows":
                subprocess.run(["explorer", self.obj_of_cls.folder_path], shell=False)
            elif self.obj_of_pft.is_wsl():
                # Windowsのパスに変換（/mnt/c/... 形式）
                wsl_path: str = subprocess.check_output(["wslpath", "-w", self.obj_of_cls.folder_path]).decode("utf-8").strip()
                subprocess.run([EXPLORER_OF_WSL, wsl_path])
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def search_files(self) -> bool:
        """ファイルを検索します"""
        result: bool = False
        try:
            if self.obj_of_cls.folder_path == "":
                raise Exception("フォルダを選択してください。")
            # 再読み込み
            self.select_folder(True)
            self.obj_of_cls.pattern = self.pattern_input.text().strip()
            self.obj_of_cls.extract_by_pattern()
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
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
        window: MainApp_Of_GFL = MainApp_Of_GFL()
        window.resize(1000, 800)
        # 最大化して、表示させる
        window.showMaximized()
        sys.exit(app.exec())
    except Exception as e:
        obj_of_gt.show_error(f"error: \n{str(e)}")
    else:
        result = True
    finally:
        pass
    return result


if __name__ == "__main__":
    main()
