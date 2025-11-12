import asyncio
import csv
import logging
import os
import sys
import threading
from pathlib import Path

import pandas
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
        self._setup_first_ui()
        self.obj_of_pt: PathTools = PathTools()
        self._setup_log()

    def closeEvent(self, event):
        """終了します"""
        if self.obj_of_lt:
            self._show_info(f"ログファイルは、\n{self.obj_of_lt.file_path_of_log}\nに出力されました。")
        super().closeEvent(event)

    def _show_info(self, msg: str):
        """情報を表示します"""
        QMessageBox.information(self, "情報", msg)
        self.obj_of_lt.logger.info(msg)

    def _show_result(self, label: str, success: bool):
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
            self._show_error(f"error: \n{str(e)}")
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
            self.top_left_scroll_area: QScrollArea = QScrollArea()
            self.top_left_scroll_area.setWidgetResizable(True)
            self.top_layout.addWidget(self.top_left_scroll_area)
            self._setup_second_ui()
            # 右上
            top_right_scroll_area: QScrollArea = QScrollArea()
            top_right_scroll_area.setWidgetResizable(True)
            self.top_layout.addWidget(top_right_scroll_area)
            top_right_container: QWidget = QWidget()
            self.top_right_container_layout: QVBoxLayout = QVBoxLayout(top_right_container)
            top_right_scroll_area.setWidget(top_right_container)
            self.log_area: QTextEdit = QTextEdit()
            self.log_area.setReadOnly(True)
            self.top_right_container_layout.addWidget(self.log_area)
            # 下
            self.second_element_title_area: QHBoxLayout = QHBoxLayout()
            self.main_layout.addLayout(self.second_element_title_area)
            self.second_element_title_area.addWidget(QLabel("統計表"))
            self.second_element_title_area.addWidget(QLabel("機能"))
            self.bottom_layout: QHBoxLayout = QHBoxLayout()
            self.main_layout.addLayout(self.bottom_layout)
            # 統計表
            self.table_scroll_area: QScrollArea = QScrollArea()
            self.table_scroll_area.setWidgetResizable(True)
            self.bottom_layout.addWidget(self.table_scroll_area)
            # 関数
            func_scroll_area: QScrollArea = QScrollArea()
            func_scroll_area.setWidgetResizable(True)
            self.bottom_layout.addWidget(func_scroll_area)
            func_container: QWidget = QWidget()
            func_container_layout: QVBoxLayout = QVBoxLayout(func_container)
            func_scroll_area.setWidget(func_container)
            func_area: QFormLayout = QFormLayout()
            func_container_layout.addLayout(func_area)
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
            # 統計表IDの一覧を取得する
            get_ids_btn: QPushButton = QPushButton("統計表IDの一覧を取得する")
            get_ids_btn.clicked.connect(self.get_lst_of_ids)
            # 統計表IDの一覧の取得をキャンセルする
            cancel_getting_ids_btn: QPushButton = QPushButton("統計表IDの一覧の取得をキャンセルする")
            cancel_getting_ids_btn.clicked.connect(self.cancel_getting_lst_of_ids)
            func_area.addRow(get_ids_btn, cancel_getting_ids_btn)
            # 統計表IDの一覧を表示する
            show_ids_btn: QPushButton = QPushButton("統計表IDの一覧を表示する")
            func_area.addRow(show_ids_btn)
            show_ids_btn.clicked.connect(self.show_lst_of_ids)
            # 統計表IDの一覧をフィルターにかける
            filter_ids_btn: QPushButton = QPushButton("統計表IDの一覧をフィルターにかける")
            func_area.addRow(filter_ids_btn)
            filter_ids_btn.clicked.connect(self.filter_lst_of_ids)
            # フィルターのキーワード
            self.keyword_text: QPlainTextEdit = QPlainTextEdit()
            func_area.addRow(QLabel("フィルターのキーワード\n(1行につき、1つのキーワード): "), self.keyword_text)
            # 検索方法
            self.match_type_combo: QComboBox = QComboBox()
            for key, desc in self.obj_of_cls.dct_of_match_type.items():
                self.match_type_combo.addItem(f"{key}: {desc}", userData=key)
            self.match_type_combo.setCurrentIndex(0)
            func_area.addRow(QLabel("検索方法: "), self.match_type_combo)
            # 抽出方法
            self.logic_type_combo: QComboBox = QComboBox()
            for key, desc in self.obj_of_cls.dct_of_logic_type.items():
                self.logic_type_combo.addItem(f"{key}: {desc}", userData=key)
            self.logic_type_combo.setCurrentIndex(0)
            func_area.addRow(QLabel("抽出方法: "), self.logic_type_combo)
            # 指定の統計表を表示する
            show_table_btn: QPushButton = QPushButton("統計表を表示する")
            func_area.addRow(show_table_btn)
            show_table_btn.clicked.connect(self.show_table)
            # 指定の統計表をフィルターにかける
            filter_table_btn: QPushButton = QPushButton("統計表をフィルターにかける")
            func_area.addRow(filter_table_btn)
            filter_table_btn.clicked.connect(self.filter_table)
            # 指定の統計表をCSVファイルに出力する
            output_btn: QPushButton = QPushButton("統計表を出力する")
            func_area.addRow(output_btn)
            output_btn.clicked.connect(self.output_table)
            # クレジット
            credit_area: QVBoxLayout = QVBoxLayout()
            self.main_layout.addLayout(credit_area)
            credit_notation: QLabel = QLabel(self.obj_of_cls.credit_text)
            credit_area.addWidget(credit_notation)
        except Exception as e:
            self._show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def _get_id_from_lst(self, index: QModelIndex) -> bool:
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
            self._show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            self._show_info(
                f"選択された統計表ID: {self.obj_of_cls.STATS_DATA_ID}\n統計名: {self.obj_of_cls.STAT_NAME}\n表題: {self.obj_of_cls.TITLE}"
            )
        return result

    def _setup_second_ui(self) -> bool:
        """2番目のUser Interfaceを設定します"""
        result: bool = False
        try:
            self.top_left_container: QWidget = QWidget()
            self.top_left_container_layout: QVBoxLayout = QVBoxLayout(self.top_left_container)
            self.top_left_scroll_area.setWidget(self.top_left_container)
            self.lst_of_ids: QTableView = QTableView()
            self.top_left_container_layout.addWidget(self.lst_of_ids)
            self.model: QStandardItemModel = QStandardItemModel()
            # ヘッダーを追加する
            self.model.setHorizontalHeaderLabels(self.obj_of_cls.header_of_ids_l)
            self.lst_of_ids.setModel(self.model)
            self.lst_of_ids.clicked.connect(self._get_id_from_lst)
        except Exception:
            raise
        else:
            result = True
        finally:
            pass
        return result

    def _setup_third_ui(self) -> bool:
        """3番目のUser Interfaceを設定します"""
        result: bool = False
        try:
            self.table_container: QWidget = QWidget()
            self.table_container_layout: QVBoxLayout = QVBoxLayout(self.table_container)
            self.table_scroll_area.setWidget(self.table_container)
            self.stats_table: QTableView = QTableView(self)
            self.table_container_layout.addWidget(QLabel(f"統計表ID: {self.obj_of_cls.STATS_DATA_ID}"))
            self.table_container_layout.addWidget(QLabel(f"統計名: {self.obj_of_cls.STAT_NAME}"))
            self.table_container_layout.addWidget(QLabel(f"表題: {self.obj_of_cls.TITLE}"))
            self.table_container_layout.addWidget(self.stats_table)
            model: QStandardItemModel = QStandardItemModel()
            # ヘッダーを追加する
            model.setHorizontalHeaderLabels(self.obj_of_cls.df.columns.tolist())
            for r in self.obj_of_cls.df.itertuples(index=False):
                items = [QStandardItem(str(v)) for v in r]
                model.appendRow(items)
            self.stats_table.setModel(model)
            self.stats_table.resizeColumnsToContents()
        except Exception:
            raise
        else:
            result = True
        finally:
            pass
        return result

    def _check_first_form(self) -> bool:
        """1番目のフォームの入力を確認します"""
        result: bool = False
        try:
            self._get_app_id()
            self.obj_of_cls.lst_of_data_type = self._get_lst(self.data_type_combo, self.obj_of_cls.dct_of_data_type)
            if not self.obj_of_cls.lst_of_data_type:
                raise Exception("データ形式を選択してください。")
            self.obj_of_cls.lst_of_get_type = self._get_lst(self.get_type_combo, self.obj_of_cls.dct_of_get_type)
        except Exception:
            raise
        else:
            result = True
        finally:
            pass
        return result

    def _check_second_form(self) -> bool:
        """2番目のフォームの入力を確認します"""
        result: bool = False
        try:
            self.obj_of_cls.lst_of_match_type = self._get_lst(self.match_type_combo, self.obj_of_cls.dct_of_match_type)
            if not self.obj_of_cls.lst_of_match_type:
                raise Exception("検索方法を選択してください。")
            if self.obj_of_cls.lst_of_match_type[self.obj_of_cls.KEY] != "検索しない":
                self.obj_of_cls.lst_of_keyword = self._get_keyword(self.keyword_text)
                if not self.obj_of_cls.lst_of_keyword:
                    raise Exception("キーワードを入力してください。")
                if len(self.obj_of_cls.lst_of_keyword) > 1:
                    self.obj_of_cls.lst_of_logic_type = self._get_lst(self.logic_type_combo, self.obj_of_cls.dct_of_logic_type)
                    if not self.obj_of_cls.lst_of_logic_type:
                        raise Exception("抽出方法を選択してください。")
        except Exception:
            raise
        else:
            result = True
        finally:
            pass
        return result

    def _get_app_id(self) -> bool:
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

    def _get_lst(self, combo: QComboBox, dct: dict) -> list:
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

    def _get_keyword(self, p_txt_edit: QPlainTextEdit) -> list:
        """キーワードを取得します"""
        return [line.strip() for line in p_txt_edit.toPlainText().splitlines() if line.strip()]

    def _clear_widget(self, widget: QWidget) -> bool:
        """ウィジェットやQScrollAreaの中身を安全に削除します"""
        result: bool = False
        try:
            if widget is None:
                raise Exception("ウィジェットが存在しません。")
            # QScrollArea の場合は中身を削除
            elif isinstance(widget, QScrollArea):
                inner_widget = widget.widget()
                if inner_widget is not None:
                    self._clear_widget(inner_widget)
                    inner_widget.deleteLater()
                    widget.takeWidget()
            else:
                # 通常の QWidget の場合
                layout = widget.layout()
                if layout is not None:
                    while layout.count():
                        item = layout.takeAt(0)
                        child_widget = item.widget()
                        child_layout = item.layout()
                        if child_widget is not None:
                            self._clear_widget(child_widget)
                            child_widget.deleteLater()
                        elif child_layout is not None:
                            # 子レイアウトを再帰的にクリア
                            self._clear_widget(QWidget().setLayout(child_layout))
        except Exception:
            raise
        else:
            result = True
        finally:
            return result

    def get_lst_of_ids(self) -> bool:
        """統計表IDの一覧を取得します"""

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
            # 取得方法は非同期のみ
            self.get_type_combo.setCurrentIndex(0)
            self._check_first_form()
            threading.Thread(target=_run_getting_ids_with_async, daemon=True).start()
        except Exception as e:
            self._show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            pass
        return result

    def cancel_getting_lst_of_ids(self) -> None:
        """統計表IDの一覧の取得をキャンセルします"""
        self.obj_of_cls.cancel = True

    def show_lst_of_ids(self) -> bool:
        """統計表IDの一覧を表示します"""
        result: bool = False
        try:
            self._clear_widget(self.top_left_scroll_area)
            self._setup_second_ui()
            # 検索パターン
            PATTERN: str = "*.csv"
            csv_files = self.obj_of_cls.folder_p_of_ids.glob(PATTERN)
            if not any(csv_files):
                raise Exception("統計表IDの一覧を取得してください。")
            for csv_file in csv_files:
                with open(str(csv_file), newline="", encoding="utf-8") as f:
                    reader = csv.reader(f)
                    # ヘッダー行をスキップする
                    next(reader, None)
                    for row in reader:
                        items: list = [QStandardItem(str(cell)) for cell in row]
                        self.model.appendRow(items)
            self.lst_of_ids.resizeColumnsToContents()
        except Exception as e:
            self._show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            self._show_result(self.show_lst_of_ids.__doc__, result)
        return result

    def filter_lst_of_ids(self) -> bool:
        """統計表IDの一覧をフィルターにかけます"""
        result: bool = False
        try:
            self._check_second_form()
            self._clear_widget(self.top_left_scroll_area)
            self._setup_second_ui()
            # 検索パターン
            PATTERN: str = "*.csv"
            csv_files = self.obj_of_cls.folder_p_of_ids.glob(PATTERN)
            if not any(csv_files):
                raise Exception("統計表IDの一覧を取得してください。")
            for csv_file in csv_files:
                reader = pandas.read_csv(str(csv_file), chunksize=1, dtype=str)
                # ヘッダー行をスキップする
                next(reader, None)
                for chunk in reader:
                    df = self.obj_of_cls.filter_df(chunk)
                    for _, row in df.iterrows():
                        items: list = [QStandardItem(str(v)) for v in row]
                        self.model.appendRow(items)
            self.lst_of_ids.resizeColumnsToContents()
        except Exception as e:
            self._show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            self._show_result(self.filter_lst_of_ids.__doc__, result)
        return result

    def show_table(self) -> bool:
        """指定の統計表を表示します"""
        result: bool = False
        try:
            if self.obj_of_cls.STATS_DATA_ID == "":
                raise Exception("統計表IDを選択してください。")
            # 取得方法は同期のみ
            self.get_type_combo.setCurrentIndex(1)
            self._check_first_form()
            self._clear_widget(self.table_scroll_area)
            self.obj_of_cls.get_table_from_api()
            self._setup_third_ui()
        except Exception as e:
            self._show_error(f"error: \n{str(e)}")
        else:
            result = True
            self.obj_of_cls.show_table()
        finally:
            self._show_result(self.show_table.__doc__, result)
        return result

    def filter_table(self) -> bool:
        """指定の統計表をフィルターにかけます"""
        result: bool = False
        try:
            if self.obj_of_cls.df is None:
                raise Exception("統計表を表示してください。")
            self._check_second_form()
            self._clear_widget(self.table_scroll_area)
            self.obj_of_cls.df = self.obj_of_cls.filter_df(self.obj_of_cls.df)
            self._setup_third_ui()
        except Exception as e:
            self._show_error(f"error: \n{str(e)}")
        else:
            result = True
            self.obj_of_cls.show_table()
        finally:
            self._show_result(self.filter_table.__doc__, result)
        return result

    def output_table(self) -> bool:
        """指定の統計表をファイルに出力します"""
        result: bool = False
        try:
            if self.obj_of_cls.df is None:
                raise Exception("統計表を表示してください。")
            self.obj_of_cls.output_table_to_csv()
        except Exception as e:
            self._show_error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            self._show_result(self.output_table.__doc__, result)
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
        obj_of_gt._show_start_up_error(f"error: \n{str(e)}")
    else:
        result = True
    finally:
        pass
    return result


if __name__ == "__main__":
    main()
