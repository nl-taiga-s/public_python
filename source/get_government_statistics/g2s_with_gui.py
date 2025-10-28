import csv
import logging
import sys
from pathlib import Path
from typing import Optional

from pandas import DataFrame
from PySide6.QtCore import QModelIndex
from PySide6.QtGui import QFont, QStandardItem, QStandardItemModel
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
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

    def show_result(self, label: str, success: bool):
        """結果を表示します"""
        QMessageBox.information(self, "結果", f"{label}に{'成功' if success else '失敗'}しました。")

    def show_error(self, msg: str):
        """エラーを表示します"""
        QMessageBox.critical(self, "エラー", msg)

    def setup_log(self) -> bool:
        """ログを設定します"""
        result: bool = False
        exe_path: Optional[Path] = None
        try:
            # exe化されている場合とそれ以外を切り分ける
            if getattr(sys, "frozen", False):
                exe_path = Path(sys.executable)
            else:
                exe_path = Path(__file__)
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
            # 仮想コンテナ
            top_left_container: QWidget = QWidget()
            self.top_left_container_layout: QHBoxLayout = QHBoxLayout()
            top_left_container.setLayout(self.top_left_container_layout)
            top_left_scroll_area: QScrollArea = QScrollArea()
            self.top_left_layout.addWidget(top_left_scroll_area)
            self.top_left_scroll_layout: QVBoxLayout = QVBoxLayout()
            top_left_scroll_area.setWidgetResizable(True)
            top_left_scroll_area.setWidget(top_left_container)
            self.show_list_of_ids()
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
            # ボタン
            button_area: QHBoxLayout = QHBoxLayout()
            self.bottom_layout.addLayout(button_area)
            self.data_type_combo: QComboBox = QComboBox()
            for key, desc in self.obj_of_cls.dct_of_data_type.items():
                self.data_type_combo.addItem(f"{key}: {desc}", userData=key)
            # 初期値
            self.data_type_combo.setCurrentIndex(0)
            button_area.addWidget(self.data_type_combo)
            show_btn: QPushButton = QPushButton("統計表を表示する")
            button_area.addWidget(show_btn)
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

    def show_list_of_ids(self) -> bool:
        """統計表IDの一覧を表示します"""
        result: bool = False
        try:
            lst_of_ids: QTableView = QTableView()
            self.top_left_container_layout.addWidget(lst_of_ids)
            self.model: QStandardItemModel = QStandardItemModel()
            # ヘッダーを追加する
            self.model.setHorizontalHeaderLabels(self.obj_of_cls.header_of_ids)
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

        def get_data_type_from_combo() -> bool:
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

        result: bool = False
        try:
            if self.obj_of_cls.STATS_DATA_ID == "":
                raise Exception("統計表IDが選択されていません。")
            get_data_type_from_combo()
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
