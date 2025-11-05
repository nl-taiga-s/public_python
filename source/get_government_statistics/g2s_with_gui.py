import asyncio
import csv
import logging
import os
import sys
import threading
from pathlib import Path

from PySide6.QtCore import QModelIndex
from PySide6.QtGui import QFont, QStandardItem, QStandardItemModel
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QTableView,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from source.common.common import GUITools, LogTools, PathTools, QtSafeLogger
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
        QMessageBox.warning(self, "エラー", msg)
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
            self.setWindowTitle("政府統計表示アプリ")
            central: QWidget = QWidget()
            self.setCentralWidget(central)
            # 主要
            self.main_layout: QVBoxLayout = QVBoxLayout(central)
            # 上
            self.first_element_title_area: QHBoxLayout = QHBoxLayout()
            self.main_layout.addLayout(self.first_element_title_area)
            self.first_element_title_area.addWidget(QLabel("統計表ID"))
            self.first_element_title_area.addWidget(QLabel("ログ"))
            self.top_layout: QHBoxLayout = QHBoxLayout()
            self.main_layout.addLayout(self.top_layout)
            # 左上
            self.top_left_layout: QVBoxLayout = QVBoxLayout()
            self.top_layout.addLayout(self.top_left_layout)
            self.show_lst_of_ids(False)
            # 右上
            self.top_right_layout: QVBoxLayout = QVBoxLayout()
            self.top_layout.addLayout(self.top_right_layout)
            self.log_area: QTextEdit = QTextEdit()
            self.log_area.setReadOnly(True)
            self.top_right_layout.addWidget(self.log_area)
            # 下
            self.second_element_title_area: QHBoxLayout = QHBoxLayout()
            self.main_layout.addLayout(self.second_element_title_area)
            self.second_element_title_area.addWidget(QLabel("統計表"))
            self.second_element_title_area.addWidget(QLabel("機能"))
            self.bottom_layout: QHBoxLayout = QHBoxLayout()
            self.main_layout.addLayout(self.bottom_layout)
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
            func_area: QFormLayout = QFormLayout()
            self.bottom_layout.addLayout(func_area)
            # アプリケーションID
            self.app_id_text: QLineEdit = QLineEdit()
            func_area.addRow(QLabel("アプリケーションID: "), self.app_id_text)
            # データタイプ
            self.data_type_combo: QComboBox = QComboBox()
            for key, desc in self.obj_of_cls.dct_of_data_type.items():
                self.data_type_combo.addItem(f"{key}: {desc}", userData=key)
            self.data_type_combo.setCurrentIndex(0)
            func_area.addRow(QLabel("データタイプ: "), self.data_type_combo)
            # 取得方法
            self.get_type_combo: QComboBox = QComboBox()
            for key, desc in self.obj_of_cls.dct_of_get_type.items():
                self.get_type_combo.addItem(f"{key}: {desc}", userData=key)
            self.get_type_combo.setCurrentIndex(0)
            func_area.addRow(QLabel("取得方法: "), self.get_type_combo)
            # 統計表IDの一覧を取得して表示する
            get_btn: QPushButton = QPushButton("統計表IDの一覧を取得して表示する")
            func_area.addRow(get_btn)
            get_btn.clicked.connect(lambda: self.show_lst_of_ids(True))
            # 検索方法
            self.match_type_combo: QComboBox = QComboBox()
            for key, desc in self.obj_of_cls.dct_of_match_type.items():
                self.match_type_combo.addItem(f"{key}: {desc}", userData=key)
            self.match_type_combo.setCurrentIndex(0)
            func_area.addRow(QLabel("検索方法: "), self.match_type_combo)
            # キーワード
            self.keyword_text: QPlainTextEdit = QPlainTextEdit()
            func_area.addRow(QLabel("キーワード(1行につき、1つのキーワード): "), self.keyword_text)
            # 抽出方法
            self.logic_type_combo: QComboBox = QComboBox()
            for key, desc in self.obj_of_cls.dct_of_logic_type.items():
                self.logic_type_combo.addItem(f"{key}: {desc}", userData=key)
            self.logic_type_combo.setCurrentIndex(0)
            func_area.addRow(QLabel("抽出方法: "), self.logic_type_combo)
            # 指定の統計表を表示する
            show_btn: QPushButton = QPushButton("統計表を表示する")
            func_area.addRow(show_btn)
            show_btn.clicked.connect(self.show_table)
            # 指定の統計表をCSVファイルに出力する
            output_btn: QPushButton = QPushButton("統計表を出力する")
            func_area.addRow(output_btn)
            output_btn.clicked.connect(self.output_table_to_csv)
            # クレジット
            credit_area: QVBoxLayout = QVBoxLayout()
            self.main_layout.addLayout(credit_area)
            credit_notation: QLabel = QLabel(self.obj_of_cls.credit_text)
            credit_area.addWidget(credit_notation)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def get_app_id(self) -> bool:
        """アプリケーションIDを取得します"""
        result: bool = False
        try:
            if self.obj_of_cls.ENV_NAME_OF_APP_ID in os.environ:
                self.obj_of_cls.APP_ID = os.environ.get(self.obj_of_cls.ENV_NAME_OF_APP_ID)
            else:
                if self.app_id_text.text() == "":
                    raise Exception("政府統計のAPIのアプリケーションIDを取得して、入力してください。https://www.e-stat.go.jp/")
                self.obj_of_cls.APP_ID = self.app_id_text.text()
                os.environ[self.obj_of_cls.ENV_NAME_OF_APP_ID] = self.obj_of_cls.APP_ID
        except Exception:
            raise
        else:
            result = True
        finally:
            pass
        return result

    def get_lst(self, combo: QComboBox, dct: dict) -> list:
        """リストを取得します"""
        try:
            index: int = combo.currentIndex()
            key: str = combo.itemData(index)
            desc: str = dct[key]
        except Exception:
            raise
        else:
            pass
        finally:
            pass
        return [key, desc]

    def get_keyword(self, ptxtedit: QPlainTextEdit) -> list:
        """キーワードを取得します"""
        return [line.strip() for line in ptxtedit.toPlainText().splitlines() if line.strip()]

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
        except Exception:
            raise
        else:
            result = True
        finally:
            pass
        return result

    def show_lst_of_ids(self, get: bool) -> bool:
        """統計表IDの一覧を表示します"""

        def _run_getting_ids_with_async() -> bool:
            """バックグラウンドスレッドで非同期の処理をします"""
            result: bool = False
            try:
                asyncio.run(self.obj_of_cls.write_stats_data_ids_to_file())
            except Exception:
                raise
            else:
                result = True
            finally:
                pass
            return result

        result: bool = False
        try:
            self.clear_layout(self.top_left_layout)
            if get:
                self.get_app_id()
                self.obj_of_cls.lst_of_data_type = self.get_lst(self.data_type_combo, self.obj_of_cls.dct_of_data_type)
                self.obj_of_cls.lst_of_get_type = self.get_lst(self.get_type_combo, self.obj_of_cls.dct_of_get_type)
                try:
                    match self.obj_of_cls.lst_of_get_type[self.obj_of_cls.KEY]:
                        case "非同期":
                            # asyncioをGUIイベントループと安全に共存させる
                            threading.Thread(target=_run_getting_ids_with_async, daemon=True).start()
                        case "同期":
                            self.obj_of_cls.write_stats_data_ids_to_file()
                        case _:
                            raise Exception("そのような取得方法は、ありません。")
                except Exception as e:
                    self.show_error(f"error: \n{str(e)}")
                else:
                    pass
                finally:
                    pass
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
            lst_of_ids.clicked.connect(self.get_id_from_lst)
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            self.show_result(self.show_lst_of_ids.__doc__, result)
        return result

    def get_id_from_lst(self, index: QModelIndex) -> bool:
        """一覧から統計表IDを取得します"""
        result: bool = False
        try:
            if index is None:
                raise Exception("指定した統計表IDを取得できませんでした。")
            # 行番号
            r: int = index.row()
            # 列番号
            # 統計表ID
            c_of_id: int = 0
            # 統計名
            c_of_stat_name: int = 1
            # 表題
            c_of_title: int = 2
            self.obj_of_cls.STATS_DATA_ID = self.model.item(r, c_of_id).text()
            self.obj_of_cls.STAT_NAME = self.model.item(r, c_of_stat_name).text()
            self.obj_of_cls.TITLE = self.model.item(r, c_of_title).text()
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def show_table(self) -> bool:
        """指定の統計表を表示します"""
        result: bool = False
        try:
            self.clear_layout(self.bottom_container_layout)
            if self.obj_of_cls.STATS_DATA_ID == "":
                raise Exception("統計表IDが選択されていません。")
            self.get_app_id()
            self.obj_of_cls.lst_of_data_type = self.get_lst(self.data_type_combo, self.obj_of_cls.dct_of_data_type)
            self.obj_of_cls.get_df_from_api()
            self.obj_of_cls.lst_of_match_type = self.get_lst(self.match_type_combo, self.obj_of_cls.dct_of_match_type)
            if self.obj_of_cls.lst_of_match_type[self.obj_of_cls.KEY] != "検索しない":
                self.obj_of_cls.lst_of_keyword = self.get_keyword(self.keyword_text)
                if not self.obj_of_cls.lst_of_keyword:
                    raise Exception("キーワードが入力されていません。")
                if len(self.obj_of_cls.lst_of_keyword) > 1:
                    self.obj_of_cls.lst_of_logic_type = self.get_lst(self.logic_type_combo, self.obj_of_cls.dct_of_logic_type)
                self.obj_of_cls.filter_df()
            # 統計表IDごとに仮想コンテナでまとめる
            element: QWidget = QWidget()
            element_layout: QVBoxLayout = QVBoxLayout()
            element.setLayout(element_layout)
            stats_table: QTableView = QTableView(self)
            element_layout.addWidget(QLabel(f"統計表ID: {self.obj_of_cls.STATS_DATA_ID}"))
            element_layout.addWidget(QLabel(f"統計名: {self.obj_of_cls.STAT_NAME}"))
            element_layout.addWidget(QLabel(f"表題: {self.obj_of_cls.TITLE}"))
            element_layout.addWidget(stats_table)
            self.bottom_container_layout.addWidget(element)
            model: QStandardItemModel = QStandardItemModel()
            # ヘッダーを追加する
            model.setHorizontalHeaderLabels(self.obj_of_cls.df.columns.tolist())
            for r in self.obj_of_cls.df.itertuples(index=False):
                items = [QStandardItem(str(v)) for v in r]
                model.appendRow(items)
            stats_table.setModel(model)
            stats_table.resizeColumnsToContents()
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            self.obj_of_cls.show_df()
        finally:
            pass
        return result

    def output_table_to_csv(self) -> bool:
        """指定の統計表をCSVファイルに出力します"""
        result: bool = False
        try:
            if self.obj_of_cls.df is None:
                raise Exception("統計表を表示してください。")
            self.obj_of_cls.output_df_to_csv()
        except Exception as e:
            self.show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            self.show_result(self.output_df_to_csv.__doc__, result)
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
