import asyncio
import csv
import logging
import sys
from pathlib import Path

import pandas
from pandas import DataFrame
from PySide6.QtCore import QAbstractTableModel, QModelIndex, QObject, Qt, QThread, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableView,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from source.common.common import GUITools, LogTools, PathTools
from source.get_government_statistics.g2s_class import GetGovernmentStatistics


# QTextEdit にログを流すためのハンドラ
class QTextEditHandler(logging.Handler):
    def __init__(self, widget: QTextEdit):
        super().__init__()
        self.widget: QTextEdit = widget

    def emit(self, record: logging.LogRecord):
        msg: str = self.format(record)
        self.widget.append(msg)


class DataFrameModel(QAbstractTableModel):
    def __init__(self, df: DataFrame):
        super().__init__()
        self._df = df

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return len(self._df)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return self._df.shape[1]

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
        if not index.isValid():
            return None
        if role == Qt.DisplayRole:
            return str(self._df.iat[index.row(), index.column()])
        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section])
        return str(section + 1)


class WorkerGetIds(QObject):
    finished = Signal(bool)
    error = Signal(str)

    def __init__(self, obj_of_cls_for_work: GetGovernmentStatistics, data_type_key: str):
        super().__init__()
        self.obj_of_cls_for_work = obj_of_cls_for_work
        self.data_type_key = data_type_key

    def run(self):
        try:
            loop: asyncio.AbstractEventLoop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            loop.run_until_complete(
                self.obj_of_cls_for_work.write_stats_data_ids_to_file(
                    page_generator=self.obj_of_cls_for_work.get_stats_data_ids(), data_type_key=self.data_type_key, chunk_size=100
                )
            )
            loop.close()
        except Exception as e:
            self.error.emit(str(e))
        else:
            self.finished.emit(True)


class MainApp_Of_G2S(QMainWindow):
    def __init__(self):
        """初期化します"""
        super().__init__()
        self.obj_of_lt: LogTools = LogTools()
        self.obj_of_cls: GetGovernmentStatistics = GetGovernmentStatistics(self.obj_of_lt.logger)
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
            self.setWindowTitle('政府統計表示アプリ')
            central = QWidget()
            self.setCentralWidget(central)
            main_layout = QVBoxLayout(central)
            # データタイプ選択
            self.data_type_combo = QComboBox()
            for key, desc in self.obj_of_cls.dct_of_data_type.items():
                self.data_type_combo.addItem(f"{key}: {desc}", key)
            self.data_type_combo.setCurrentIndex(0)
            main_layout.addWidget(QLabel('統計表IDを選択してください。'))
            main_layout.addWidget(self.data_type_combo)
            # ボタン
            button_layout = QHBoxLayout()
            self.get_ids_btn = QPushButton('統計表IDの一覧を取得します')
            self.get_ids_btn.clicked.connect(self.start_getting_ids)
            self.show_table_btn = QPushButton('統計表を表示する')
            self.show_table_btn.clicked.connect(self.show_statistical_table)
            button_layout.addWidget(self.get_ids_btn)
            button_layout.addWidget(self.show_table_btn)
            main_layout.addLayout(button_layout)
            # ログ
            main_layout.addWidget(QLabel('ログ'))
            self.log_area = QTextEdit()
            self.log_area.setReadOnly(True)
            main_layout.addWidget(self.log_area)
            # 統計表表示エリア
            self.table_view = QTableView()
            main_layout.addWidget(self.table_view)
            # 統計表IDリスト表示用
            self.ids_table_view = QTableView()
            main_layout.addWidget(QLabel('統計表ID一覧'))
            main_layout.addWidget(self.ids_table_view)
            # クレジット
            credit_notation: QLabel = QLabel(
                "クレジット表示\nこのサービスは、政府統計総合窓口(e-Stat)のAPI機能を使用していますが、サービスの内容は国によって保証されたものではありません。"
            )
            main_layout.addWidget(credit_notation)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def start_getting_ids(self):
        """統計表IDの一覧の取得を開始します"""
        try:
            key = self.data_type_combo.currentData()
            self.thread: QThread = QThread()
            self.worker = WorkerGetIds(self.obj_of_cls, key)
            self.worker.moveToThread(self.thread)
            self.thread.started.connect(self.worker.run)
            self.worker.finished.connect(lambda _: self.display_ids())
            self.worker.finished.connect(self.thread.quit)
            self.worker.finished.connect(self.worker.deleteLater)
            self.thread.finished.connect(self.thread.deleteLater)
            self.worker.error.connect(lambda msg: self.show_error(msg))
            self.thread.start()
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def display_ids(self) -> bool:
        """統計表IDの一覧を表示します"""
        result: bool = False
        try:
            PATTERN: str = "*.csv"
            model: DataFrameModel = DataFrameModel(DataFrame())
            for csv_file in self.obj_of_cls.folder_p_of_ids.glob(PATTERN):
                with open(csv_file, newline='', encoding='utf-8') as f:
                    df = DataFrame(list(csv.reader(f)))
                    if df.shape[0] > 1:
                        # ヘッダーを除く
                        df = df.iloc[1:, :]
                    model._df = pandas.concat([model._df, df], ignore_index=True) if not model._df.empty else df
            self.ids_table_view.setModel(model)
            self.ids_table_view.resizeColumnsToContents()
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
            index = self.ids_table_view.currentIndex()
            if not index.isValid():
                raise Exception("統計表IDが選択されていません。")
            stats_id = self.ids_table_view.model()._df.iat[index.row(), 0]
            self.obj_of_cls.STATS_DATA_ID = stats_id
            self.obj_of_cls.lst_of_data_type = [self.data_type_combo.currentData(), '']
            df: DataFrame = self.obj_of_cls.get_data_from_api()
            model = DataFrameModel(df)
            self.table_view.setModel(model)
            self.table_view.resizeColumnsToContents()
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
