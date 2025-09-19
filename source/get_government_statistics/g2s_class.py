import json
import os
import xml.etree.ElementTree as et
from io import StringIO
from logging import Logger

import pandas as pd
import pyperclip
import requests
from pandas import DataFrame
from requests import Response
from tabulate import tabulate


class GetGovernmentStatistics:
    """政府の統計データを取得します"""

    def __init__(self, logger: Logger):
        """初期化します"""
        self.log = logger
        self.log.info(self.__class__.__doc__)
        # dict変数のキー番号
        self.KEY = 0
        # dict変数の説明番号
        self.DESCRIPTION = 1
        # 取得するデータ形式
        self.lst_of_data_type = []
        # APIのURL
        self.URL = ""
        # APIのバージョン
        self.VERSION = 3.0
        # APIのURLのパラメータ
        self.params = {}
        # アプリケーションIDを取得して、環境変数で指定しておく
        self.APP_ID = os.environ.get("first_appid_of_estat")
        # 統計表ID
        self.STATS_DATA_ID = ""
        # 統計表IDの一覧
        self.STATS_DATA_IDS = [
            "0003443838",
            "0003443839",
            "0003443840",
            "0003443841",
            "0003448228",
            "0003448229",
            "0003448230",
            "0003448231",
            "0003448232",
            "0003448233",
            "0003448234",
            "0003448235",
            "0003448236",
            "0003448237",
        ]
        # 部分一致か完全一致か
        self.lst_of_match = []
        # 先頭か末尾か
        self.order = ""
        # 抽出するキーワード
        self.lst_of_keyword = []
        # OR抽出かAND抽出か
        self.lst_of_logic = []
        # 表示するデータの件数
        self.DATA_COUNT = 12

    def get_base_url(self) -> bool:
        """APIのURLの基礎部分を取得します"""
        try:
            result = False
            match self.lst_of_data_type[self.KEY]:
                case "xml":
                    self.URL = f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/getStatsData"
                case "json":
                    self.URL = f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/json/getStatsData"
                case "csv":
                    self.URL = f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/getSimpleStatsData"
                case _:
                    raise Exception("ファイル形式が対応していません。")
        except Exception as e:
            self.log.error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            if not result:
                raise
            return result

    def get_params_of_url(self):
        """APIのURLのパラメータを取得します"""
        self.params = {
            "appId": self.APP_ID,  # アプリケーションID
            "statsDataId": self.STATS_DATA_ID,  # 統計表ID
            "lang": "J",  # 言語
            "metaGetFlg": "Y",  # メタ情報の取得フラグ
            "cntGetFlg": "N",  # 件数の取得フラグ
            "explanationGetFlg": "N",  # 解説情報の有無フラグ
            "annotationGetFlg": "N",  # 注釈情報の有無フラグ
            "sectionHeaderFlg": 1,  # 見出し行の有無フラグ
            "replaceSpChars": 0,  # 特殊文字のエスケープフラグ
        }

    def get_data_from_api(self) -> DataFrame:
        """APIからデータを取得します"""

        def with_xml(self: GetGovernmentStatistics, response: Response) -> DataFrame:
            """XMLでデータを取得します"""
            try:
                result = False
                df = None
                # デバッグ用(加工前のデータをクリップボードにコピーする)
                pyperclip.copy(response.text)
                root = et.fromstring(response.text)
                # --- 1. CLASS_INFを辞書化 ---
                mapping = {}
                for obj in root.findall(".//CLASS_OBJ"):
                    obj_id = obj.attrib["id"]
                    code_map = {}
                    for cls in obj.findall("CLASS"):
                        code_map[cls.attrib["code"]] = cls.attrib.get("name", cls.attrib["code"])
                    mapping[obj_id] = code_map
                # --- 2. VALUE要素を取り込む ---
                rows = []
                for v in root.findall(".//VALUE"):
                    row = {}
                    for key, value in v.attrib.items():
                        if key in mapping:
                            row[key] = mapping[key].get(value, value)
                        else:
                            row[key] = value
                    row["値"] = v.text.strip()
                    rows.append(row)
                df = pd.DataFrame(rows)
                # --- 3. 列名をCLASS_OBJのname属性に置換 ---
                id2name = {}
                for obj in root.findall(".//CLASS_OBJ"):
                    obj_id = obj.attrib["id"]
                    obj_name = obj.attrib.get("name", obj_id)
                    id2name[obj_id] = obj_name
                id2name["unit"] = "単位"
                df.rename(columns=id2name, inplace=True)
                # --- 値列を数値型に変換 ---
                if "値" in df.columns:
                    df["値"] = pd.to_numeric(df["値"], errors="coerce")
            except Exception as e:
                self.log.error(f"error: \n{str(e)}")
            else:
                result = True
            finally:
                if not result:
                    raise
                return df

        def with_json(self: GetGovernmentStatistics, response: Response) -> DataFrame:
            """JSONでデータを取得します"""
            try:
                result = False
                df = None
                # デバッグ用(加工前のデータをクリップボードにコピーする)
                pyperclip.copy(json.dumps(response.json(), indent=4, ensure_ascii=False))
                data = response.json()
                # --- CLASS_INF と VALUE を取得 ---
                class_inf = data["GET_STATS_DATA"]["STATISTICAL_DATA"]["CLASS_INF"]["CLASS_OBJ"]
                values = data["GET_STATS_DATA"]["STATISTICAL_DATA"]["DATA_INF"]["VALUE"]
                # --- 列名用のコード(@xxx) → 日本語(@name) の辞書 ---
                col_name_map = {obj["@id"]: obj["@name"] for obj in class_inf}
                # "@unit" を "時間軸" の後に追加
                new_col_map = {}
                for k, v in col_name_map.items():
                    new_col_map[k] = v
                    if k == "time":
                        # "@time"（時間軸）の後
                        new_col_map["unit"] = "単位"
                col_name_map = new_col_map
                # --- コード→日本語辞書を作成 ---
                code_to_name = {}
                for obj in class_inf:
                    cid = obj["@id"]
                    clss = obj["CLASS"]
                    if isinstance(clss, list):
                        code_to_name[cid] = {c["@code"]: c["@name"] for c in clss}
                    else:
                        code_to_name[cid] = {clss["@code"]: clss["@name"]}
                # --- VALUE 内のコードを日本語に置換 ---
                translated_rows = []
                for v in values:
                    row = {}
                    for k, val in v.items():
                        if k.startswith("@") and k[1:] in code_to_name:
                            # 列名を日本語に、値もコードを日本語に置換
                            jp_col = col_name_map.get(k[1:], k[1:])
                            row[jp_col] = code_to_name[k[1:]].get(val, val)
                        elif k == "@unit":
                            # "@unit" は単位列として追加
                            jp_col = col_name_map.get("unit", "unit")
                            row[jp_col] = val
                        elif k == "$":
                            row["値"] = val
                        else:
                            row[k] = val
                    translated_rows.append(row)
                df = pd.DataFrame(translated_rows)
                # --- 値列を数値型に変換 ---
                if "値" in df.columns:
                    df["値"] = pd.to_numeric(df["値"], errors="coerce")
            except Exception as e:
                self.log.error(f"error: \n{str(e)}")
            else:
                result = True
            finally:
                if not result:
                    raise
                return df

        def with_csv(self: GetGovernmentStatistics, response: Response) -> DataFrame:
            """CSVでデータを取得します"""
            try:
                result = False
                df = None
                # デバッグ用(加工前のデータをクリップボードにコピーする)
                pyperclip.copy(response.text)
                lines = response.text.splitlines()
                # VALUE行の位置を探す
                value_idx = None
                for i, line in enumerate(lines):
                    if line.strip().replace('"', '') == "VALUE":
                        value_idx = i
                        break
                if value_idx is None:
                    raise Exception("CSVに 'VALUE' 行が見つかりませんでした。")
                # --- ヘッダー行 ---
                header_cols = [h.strip('"') for h in lines[value_idx + 1].split(',')]
                # --- データ本体をそのまま読み込む ---
                csv_body = "\n".join(lines[value_idx + 2 :])
                df = pd.read_csv(StringIO(csv_body), header=None)
                df.columns = header_cols  # 英語+日本語ペア含めて13列のまま
                # --- 列名置換 ---
                rename_map = {}
                drop_cols = []
                i = 0
                while i < len(header_cols):
                    eng = header_cols[i]
                    if eng.endswith("_code") and i + 1 < len(header_cols):  # コード列
                        drop_cols.append(eng)  # 英語コード列は削除
                        i += 2
                        continue
                    # 単独列の処理
                    elif eng == "unit":
                        rename_map[eng] = "単位"
                    elif eng == "value":
                        rename_map[eng] = "値"
                    else:
                        rename_map[eng] = eng
                    i += 1
                # 列名をリネーム
                df = df.rename(columns=rename_map)
                # 英語コード列を削除
                df = df.drop(columns=drop_cols)
                # --- 値列を数値型に変換 ---
                if "値" in df.columns:
                    df["値"] = pd.to_numeric(df["値"], errors="coerce")
            except Exception as e:
                self.log.error(f"error: \n{str(e)}")
            else:
                result = True
            finally:
                if not result:
                    raise
                return df

        def handle_data_type(self: GetGovernmentStatistics, response: Response) -> DataFrame:
            """データタイプで条件分岐させます"""
            try:
                result = False
                df = None
                match self.lst_of_data_type[self.KEY]:
                    case "xml":
                        df = with_xml(self, response)
                    case "json":
                        df = with_json(self, response)
                    case "csv":
                        df = with_csv(self, response)
                    case _:
                        raise Exception("ファイル形式が対応していません。")
            except Exception as e:
                self.log.error(f"error: \n{str(e)}")
            else:
                result = True
            finally:
                if not result:
                    raise
                return df

        try:
            result = False
            df = None
            self.log.info(self.get_data_from_api.__doc__)
            self.get_base_url()
            self.get_params_of_url()
            # リクエストを送信する
            response = requests.get(self.URL, params=self.params)
            df = handle_data_type(self, response)
        except Exception as e:
            self.log.error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            if not result:
                raise
            return df

    def filter_data(self, df: DataFrame) -> DataFrame:
        """データをフィルターにかけます"""
        try:
            result = False
            filtered_df = None
            self.log.info(self.filter_data.__doc__)
            match self.lst_of_match[self.KEY]:
                case "部分一致":
                    # 全列で部分一致検索する
                    if len(self.lst_of_keyword) == 1:
                        # 単一キーワード
                        kw = str(self.lst_of_keyword[0])
                        filtered_df = df[df.apply(lambda row: row.astype(str).str.contains(kw, case=False, na=False).any(), axis=1)]
                    else:
                        # 複数キーワード
                        match self.lst_of_logic[self.KEY]:
                            case "OR抽出":
                                pattern = "|".join(map(str, self.lst_of_keyword))
                                filtered_df = df[df.apply(lambda row: row.astype(str).str.contains(pattern, case=False, na=False).any(), axis=1)]
                            case "AND抽出":
                                filtered_df = df[
                                    df.apply(
                                        lambda row: all(row.astype(str).str.contains(k, case=False, na=False).any() for k in self.lst_of_keyword),
                                        axis=1,
                                    )
                                ]
                            case _:
                                raise Exception("その抽出方式はありません。")
                case "完全一致":
                    # 全列で完全一致検索する
                    if len(self.lst_of_keyword) == 1:
                        # 単一キーワード
                        kw = str(self.lst_of_keyword[0])
                        filtered_df = df[df.apply(lambda row: row.astype(str).eq(kw).any(), axis=1)]
                    else:
                        # 複数キーワード
                        match self.lst_of_logic[self.KEY]:
                            case "OR抽出":
                                filtered_df = df[df.apply(lambda row: row.astype(str).isin(self.lst_of_keyword).any(), axis=1)]
                            case "AND抽出":
                                filtered_df = df[df.apply(lambda row: all(row.astype(str).eq(k).any() for k in self.lst_of_keyword), axis=1)]
                            case _:
                                raise Exception("その抽出方式はありません。")
                case _:
                    raise Exception("その検索方式はありません。")
        except Exception as e:
            self.log.error(f"error: \n{str(e)}")
        else:
            result = True
        finally:
            if not result:
                raise
            return filtered_df

    def show_data(self, df: DataFrame) -> bool:
        """データを表示させます"""
        try:
            result = False
            self.log.info(self.show_data.__doc__)
            match self.order:
                case "先頭":
                    self.log.info(tabulate(df.head(self.DATA_COUNT), headers="keys", tablefmt="pipe", showindex=False))
                case "末尾":
                    self.log.info(tabulate(df.tail(self.DATA_COUNT), headers="keys", tablefmt="pipe", showindex=False))
                case _:
                    raise Exception("その表示順はありません。")
        except Exception as e:
            self.log.error(f"error: \n{str(e)}")
        else:
            result = True
            self.log.info(f"データの取得形式: {self.lst_of_data_type[self.KEY]} => {self.lst_of_data_type[self.DESCRIPTION]}")
            self.log.info(f"検索方式: {self.lst_of_match[self.KEY]} => {self.lst_of_match[self.DESCRIPTION]}")
            self.log.info(f"抽出するキーワード: {", ".join(map(str, self.lst_of_keyword))}")
            if self.lst_of_logic:
                self.log.info(f"抽出方式: {self.lst_of_logic[self.KEY]} => {self.lst_of_logic[self.DESCRIPTION]}")
            self.log.info(f"表示順: {self.order}")
            self.log.info(f"表示件数: {self.DATA_COUNT}")
        finally:
            return result
