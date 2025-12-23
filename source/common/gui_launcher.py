import sys

from PySide6.QtGui import QFont
from PySide6.QtWidgets import QApplication, QFormLayout, QMainWindow, QMessageBox, QPushButton, QScrollArea, QVBoxLayout, QWidget

from source.common.common import GUITools


class MainApp_Of_Gui_Launcher(QMainWindow):
    def __init__(self):
        """初期化します"""
        super().__init__()
        self._setup_ui()

    def closeEvent(self, event):
        """終了します"""
        super().closeEvent(event)

    def _show_info(self, msg: str):
        """情報を表示します"""
        QMessageBox.information(self, "情報", msg)
        self.obj_of_lt.logger.info(msg)

    def _show_result(self, label: str | None, success: bool):
        """結果を表示します"""
        QMessageBox.information(self, "結果", f"{label} => {'成功' if success else '失敗'}しました。")
        if success:
            self.obj_of_lt.logger.info(f"{label} => 成功しました。")
        else:
            self.obj_of_lt.logger.error(f"{label} => 失敗しました。")

    def _show_error(self, msg: str):
        """エラーを表示します"""
        QMessageBox.warning(self, "エラー", msg)
        self.obj_of_lt.logger.warning(msg)

    def _setup_ui(self) -> bool:
        """User Interfaceを設定します"""
        result: bool = False
        try:
            # タイトル
            self.setWindowTitle("Useful Tools - GUI Launcher-")
            central: QWidget = QWidget()
            self.setCentralWidget(central)
            base_layout: QVBoxLayout = QVBoxLayout(central)
            # 主要
            main_scroll_area: QScrollArea = QScrollArea()
            main_scroll_area.setWidgetResizable(True)
            base_layout.addWidget(main_scroll_area)
            main_container: QWidget = QWidget()
            main_container_layout: QFormLayout = QFormLayout(main_container)
            main_scroll_area.setWidget(main_container)
            # 選択
            buttons: dict = {
                "source/convert_libre_to_pdf": self.launch_cltp,
                "source/convert_office_to_pdf": self.launch_cotp,
                "source/convert_to_md": self.launch_ctm,
                "source/get_file_list": self.launch_gfl,
                "source/get_government_statistics": self.launch_g2s,
                "source/pdf_tools": self.launch_pt,
            }
            for text, func in buttons.items():
                btn: QPushButton = QPushButton(text)
                btn.clicked.connect(func)
                main_container_layout.addRow(btn)
        except Exception as e:
            self._show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def launch_cltp(self) -> bool:
        """CLTP"""
        result: bool = False
        try:
            from source.convert_libre_to_pdf.cltp_with_gui import main

            main()
        except Exception:
            raise
        else:
            result = True
        finally:
            pass
        return result

    def launch_cotp(self) -> bool:
        """COTP"""
        result: bool = False
        try:
            from source.convert_office_to_pdf.cotp_with_gui import main

            main()
        except Exception:
            raise
        else:
            result = True
        finally:
            pass
        return result

    def launch_ctm(self) -> bool:
        """CTM"""
        result: bool = False
        try:
            from source.convert_to_md.ctm_with_gui import main

            main()
        except Exception:
            raise
        else:
            result = True
        finally:
            pass
        return result

    def launch_gfl(self) -> bool:
        """GFL"""
        result: bool = False
        try:
            from source.get_file_list.gfl_with_gui import main

            main()
        except Exception:
            raise
        else:
            result = True
        finally:
            pass
        return result

    def launch_g2s(self) -> bool:
        """G2S"""
        result: bool = False
        try:
            from source.get_government_statistics.g2s_with_gui import main

            main()
        except Exception:
            raise
        else:
            result = True
        finally:
            pass
        return result

    def launch_pt(self) -> bool:
        """PT"""
        result: bool = False
        try:
            from source.pdf_tools.pt_with_gui import main

            main()
        except Exception:
            raise
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
        window: MainApp_Of_Gui_Launcher = MainApp_Of_Gui_Launcher()
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
