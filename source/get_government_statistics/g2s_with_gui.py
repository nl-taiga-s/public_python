import asyncio
import csv
import logging
import os
import sys
import threading
from pathlib import Path

from pandas import DataFrame
from PySide6.QtCore import QModelIndex, QObject, Signal
from PySide6.QtGui import QFont, QStandardItem, QStandardItemModel
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QTableView,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from source.common.common import GUITools, LogTools, PathTools, QtSafeLogger
from source.get_government_statistics.g2s_class import GetGovernmentStatistics


class WorkerSignals(QObject):
    finished: Signal = Signal(bool, str)
    info: Signal = Signal(str)
    error: Signal = Signal(str)


# QTextEdit にログを流すためのハンドラ
class QTextEditHandler(logging.Handler):
    def __init__(self, widget: QTextEdit):
        super().__init__()
        self.widget: QTextEdit = widget

    def emit(self, record: logging.LogRecord):
        msg: str = self.format(record)
        self.widget.append(msg)


class MainApp_Of_G2S(QMainWindow):
    def __init__(self):
        """初期化します"""
        super().__init__()
        self.obj_of_lt: LogTools = LogTools()
        self.qt_logger: QtSafeLogger = QtSafeLogger(self.obj_of_lt.logger)
        self.obj_of_cls: GetGovernmentStatistics = GetGovernmentStatistics(self.qt_logger)
        self.setup_ui()
        self.obj_of_pt: PathTools = PathTools()
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
        QMessageBox.information(self, "結果", f"{label} => {'成功' if success else '失敗'}しました。")
        if success:
            self.obj_of_lt.logger.info(f"{label} => 成功しました。")
        else:
            self.obj_of_lt.logger.error(f"{label} => 失敗しました。")

    def show_error(self, msg: str):
        """エラーを表示します"""
        QMessageBox.critical(self, "エラー", msg)
        self.obj_of_lt.logger.critical(msg)

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
            self.setWindowTitle("政府統計表示アプリ")
            central: QWidget = QWidget()
            self.setCentralWidget(central)
            # 主要
            self.main_layout: QVBoxLayout = QVBoxLayout(central)
            self.main_layout.addWidget(QLabel("統計表IDを選択してください。"))
            # 上
            self.top_layout: QHBoxLayout = QHBoxLayout()
            self.main_layout.addLayout(self.top_layout)
            # 左上
            self.top_left_layout: QVBoxLayout = QVBoxLayout()
            self.top_layout.addLayout(self.top_left_layout)
            self.top_left_layout.addWidget(QLabel("統計表ID"))
            self.show_list_of_ids(False)
            # 右上
            self.top_right_layout: QVBoxLayout = QVBoxLayout()
            self.top_layout.addLayout(self.top_right_layout)
            self.top_right_layout.addWidget(QLabel("ログ"))
            self.log_area: QTextEdit = QTextEdit()
            self.log_area.setReadOnly(True)
            self.top_right_layout.addWidget(self.log_area)
            # 下
            self.bottom_layout: QVBoxLayout = QVBoxLayout()
            self.main_layout.addLayout(self.bottom_layout)
            self.bottom_layout.addWidget(QLabel("表"))
            # 仮想コンテナ
            bottom_container: QWidget = QWidget()
            self.bottom_container_layout: QHBoxLayout = QHBoxLayout()
            bottom_container.setLayout(self.bottom_container_layout)
            bottom_scroll_area: QScrollArea = QScrollArea()
            self.bottom_layout.addWidget(bottom_scroll_area)
            bottom_scroll_area.setWidgetResizable(True)
            bottom_scroll_area.setWidget(bottom_container)
            # 統計表
            self.table_area: QVBoxLayout = QVBoxLayout()
            self.bottom_layout.addLayout(self.table_area)
            # 関数
            first_func_area: QHBoxLayout = QHBoxLayout()
            self.bottom_layout.addLayout(first_func_area)
            # アプリケーションID
            app_id_lbl: QLabel = QLabel("アプリケーションID: ")
            self.app_id_text: QLineEdit = QLineEdit()
            first_func_area.addWidget(app_id_lbl)
            first_func_area.addWidget(self.app_id_text)
            # データタイプ
            data_type_lbl: QLabel = QLabel("データタイプ: ")
            self.data_type_combo: QComboBox = QComboBox()
            for key, desc in self.obj_of_cls.dct_of_data_type.items():
                self.data_type_combo.addItem(f"{key}: {desc}", userData=key)
            self.data_type_combo.setCurrentIndex(0)
            first_func_area.addWidget(data_type_lbl)
            first_func_area.addWidget(self.data_type_combo)
            second_func_area: QHBoxLayout = QHBoxLayout()
            self.bottom_layout.addLayout(second_func_area)
            # 取得方法
            self.get_type_combo: QComboBox = QComboBox()
            for key, desc in self.obj_of_cls.dct_of_get_type.items():
                self.get_type_combo.addItem(f"{key}: {desc}", userData=key)
            self.get_type_combo.setCurrentIndex(0)
            second_func_area.addWidget(self.get_type_combo)
            # 統計表IDの一覧を取得して表示する
            get_btn: QPushButton = QPushButton("統計表IDの一覧を取得して表示する")
            second_func_area.addWidget(get_btn)
            get_btn.clicked.connect(lambda: self.show_list_of_ids(True))
            # 指定の統計表を表示する
            show_btn: QPushButton = QPushButton("統計表を表示する")
            second_func_area.addWidget(show_btn)
            show_btn.clicked.connect(self.show_statistical_table)
            # クレジット
            credit_area: QVBoxLayout = QVBoxLayout()
            self.bottom_layout.addLayout(credit_area)
            credit_notation: QLabel = QLabel(
                "クレジット表示\nこのサービスは、政府統計総合窓口(e-Stat)のAPI機能を使用していますが、サービスの内容は国によって保証されたものではありません。"
            )
            credit_area.addWidget(credit_notation)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def get_app_id_from_lineedit(self) -> bool:
        """ラインエディットからアプリケーションIDを取得します"""
        result: bool = False
        try:
            if "first_appid_of_estat" in os.environ:
                self.obj_of_cls.APP_ID = os.environ.get("first_appid_of_estat")
            else:
                if self.app_id_text.text() == "":
                    raise Exception("政府統計のAPIのアプリケーションIDを取得して、入力してください。https://www.e-stat.go.jp/")
                self.obj_of_cls.APP_ID = self.app_id_text.text()
                os.environ["first_appid_of_estat"] = self.obj_of_cls.APP_ID
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def get_data_type_from_combo(self) -> bool:
        """コンボボックスからデータタイプを取得します"""
        result: bool = False
        try:
            index: int = self.data_type_combo.currentIndex()
            key: str = self.data_type_combo.itemData(index)
            desc: str = self.obj_of_cls.dct_of_data_type[key]
            self.obj_of_cls.lst_of_data_type = [key, desc]
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def get_getting_type_from_combo(self) -> bool:
        """コンボボックスから取得方法を取得します"""
        result: bool = False
        try:
            index: int = self.get_type_combo.currentIndex()
            key: str = self.get_type_combo.itemData(index)
            desc: str = self.obj_of_cls.dct_of_get_type[key]
            self.obj_of_cls.lst_of_get_type = [key, desc]
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def clear_layout(self, layout) -> bool:
        """レイアウト内の全ウィジェットを削除します"""
        result: bool = False
        try:
            if layout is not None:
                while layout.count():
                    item = layout.takeAt(0)
                    widget = item.widget()
                    if widget is not None:
                        widget.deleteLater()
                    else:
                        sub_layout = item.layout()
                        if sub_layout is not None:
                            self.clear_layout(sub_layout)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def show_list_of_ids(self, get: bool) -> bool:
        """統計表IDの一覧を表示します"""

        self.worker_signals = WorkerSignals()
        self.worker_signals.finished.connect(lambda success, label: self.show_result(label, success))
        self.worker_signals.error.connect(lambda msg: self.show_error(msg))
        self.worker_signals.info.connect(lambda msg: self.show_info(msg))

        def _run_getting_ids_with_async() -> bool:
            """バックグラウンドスレッドで非同期の処理をします"""
            result: bool = False
            try:
                asyncio.run(self.obj_of_cls.write_stats_data_ids_to_file())
            except Exception as e:
                self.worker_signals.error.emit(str(e))
            else:
                result = True
            finally:
                self.worker_signals.finished.emit(result, _run_getting_ids_with_async.__doc__)
            return result

        result: bool = False
        try:
            self.clear_layout(self.top_left_layout)
            if get:
                self.get_app_id_from_lineedit()
                self.get_data_type_from_combo()
                self.get_getting_type_from_combo()
                match self.obj_of_cls.lst_of_get_type[self.obj_of_cls.KEY]:
                    case "非同期":
                        # asyncioをGUIイベントループと安全に共存させる
                        threading.Thread(target=_run_getting_ids_with_async, daemon=True).start()
                    case "同期":
                        self.obj_of_cls.write_stats_data_ids_to_file()
                    case _:
                        raise Exception("そのような取得方法は、ありません。")
            # 仮想コンテナ
            top_left_container: QWidget = QWidget()
            self.top_left_container_layout: QHBoxLayout = QHBoxLayout()
            top_left_container.setLayout(self.top_left_container_layout)
            top_left_scroll_area: QScrollArea = QScrollArea()
            self.top_left_layout.addWidget(top_left_scroll_area)
            self.top_left_scroll_layout: QVBoxLayout = QVBoxLayout()
            top_left_scroll_area.setWidgetResizable(True)
            top_left_scroll_area.setWidget(top_left_container)
            lst_of_ids: QTableView = QTableView()
            self.top_left_container_layout.addWidget(lst_of_ids)
            self.model: QStandardItemModel = QStandardItemModel()
            # ヘッダーを追加する
            self.model.setHorizontalHeaderLabels(self.obj_of_cls.header_of_ids_l)
            # 検索パターン
            PATTERN: str = "*.csv"
            for csv_file in self.obj_of_cls.folder_p_of_ids.glob(PATTERN):
                with open(csv_file, newline="", encoding="utf-8") as f:
                    reader = csv.reader(f)
                    # ヘッダー行をスキップする
                    next(reader, None)
                    for row in reader:
                        items: list = [QStandardItem(str(cell)) for cell in row]
                        self.model.appendRow(items)
            lst_of_ids.setModel(self.model)
            lst_of_ids.resizeColumnsToContents()
            lst_of_ids.clicked.connect(self.get_id_from_list)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def get_id_from_list(self, index: QModelIndex) -> bool:
        """一覧から統計表IDを取得します"""
        result: bool = False
        try:
            if index is None:
                raise Exception("指定した統計表IDを取得できませんでした。")
            # 行番号
            r: int = index.row()
            # 統計表IDの列番号
            c_of_id: int = 0
            item_of_id: QStandardItem = self.model.item(r, c_of_id)
            self.obj_of_cls.STATS_DATA_ID = item_of_id.text()
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def show_statistical_table(self) -> bool:
        """統計表を表示させます"""
        result: bool = False
        try:
            self.clear_layout(self.bottom_container_layout)
            self.get_app_id_from_lineedit()
            if self.obj_of_cls.STATS_DATA_ID == "":
                raise Exception("統計表IDが選択されていません。")
            self.get_data_type_from_combo()
            df: DataFrame = self.obj_of_cls.get_data_from_api()
            # 統計表IDごとに仮想コンテナでまとめる
            element: QWidget = QWidget()
            element_layout: QVBoxLayout = QVBoxLayout()
            element.setLayout(element_layout)
            stats_id: QLabel = QLabel(f"統計表ID: {self.obj_of_cls.STATS_DATA_ID}")
            stats_table: QTableView = QTableView(self)
            element_layout.addWidget(stats_id)
            element_layout.addWidget(stats_table)
            self.bottom_container_layout.addWidget(element)
            model: QStandardItemModel = QStandardItemModel()
            # ヘッダーを追加する
            model.setHorizontalHeaderLabels(df.columns.tolist())
            for r in df.itertuples(index=False):
                items = [QStandardItem(str(v)) for v in r]
                model.appendRow(items)
            stats_table.setModel(model)
            stats_table.resizeColumnsToContents()
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            pass
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
        window: MainApp_Of_G2S = MainApp_Of_G2S()
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
