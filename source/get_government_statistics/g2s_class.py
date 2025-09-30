import asyncio
import json
import os
import traceback
import xml.etree.ElementTree as et
from csv import DictReader
from io import StringIO
from logging import Logger
from typing import Any, Callable, Dict, Optional, Tuple

import httpx
import pandas as pd
import pyperclip
import requests
from httpx import AsyncClient, Response
from pandas import DataFrame
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

    async def get_stats_data_ids(self):
        """統計表IDの一覧を取得します"""

        async def fetch_page(
            client: AsyncClient, url: str, params: dict[str, Any], parser: Callable[[Response], Tuple[Dict[str, Dict[str, str]], int]]
        ) -> Tuple[Dict[str, Dict[str, str]], int]:
            """1ページ分を取得します"""
            page_dict: Dict[str, Dict[str, str]] = {}
            count = 0
            error: Optional[Exception] = None
            try:
                res = await client.get(url, params=params)
                res.raise_for_status()
                page_dict, count = parser(res)
            except Exception as e:
                error = e
                tb = traceback.format_exc()
                self.log.error(f"***{fetch_page.__doc__} => 失敗しました。: \n{repr(e)}\n{tb}***")
            else:
                pass
            finally:
                if error is not None:
                    raise error
                return page_dict, count

        async def fetch_all_pages(
            url: str,
            parser: Callable[[Response], Tuple[Dict[str, Dict[str, str]], int]],
            page_limit: int = 100,  # 1回のリクエストで取得する件数
            concurrency: int = 5,  # 非同期で処理する並列数
        ):
            """共通処理でページごとに取得します"""
            error: Optional[Exception] = None
            # タイムアウト時間を設定する
            async with httpx.AsyncClient(timeout=httpx.Timeout(15.0)) as client:
                start = 1
                while True:
                    stop = True
                    try:
                        tasks = [
                            fetch_page(
                                client,
                                url,
                                {"appId": self.APP_ID, "lang": "J", "limit": page_limit, "startPosition": start + i * page_limit},
                                parser,
                            )
                            for i in range(concurrency)
                        ]
                        results = await asyncio.gather(*tasks, return_exceptions=True)
                        for res in results:
                            if isinstance(res, Exception):
                                tb = traceback.format_exc()
                                self.log.error(f"タスクで例外発生: \n{repr(res)}\n{tb}")
                                continue
                            page_dict, count = res
                            if count > 0:
                                yield page_dict
                                stop = False
                        if stop:
                            break
                    except Exception as e:
                        error = e
                        tb = traceback.format_exc()
                        self.log.error(f"***{fetch_all_pages.__doc__} => 失敗しました。: \n{repr(e)}\n{tb}***")
                    else:
                        pass
                    finally:
                        if error is not None:
                            raise error
                        start += page_limit * concurrency

        def parser_xml(res: Response) -> Tuple[Dict[str, Dict[str, str]], int]:
            """XMLのデータを解析します"""
            page_dict: Dict[str, Dict[str, str]] = {}
            table_list: list = []
            error: Optional[Exception] = None
            try:
                root = et.fromstring(res.text)
                table_list = root.findall(".//TABLE_INF")
                for t in table_list:
                    stat_id = t.attrib.get("id", "")
                    element_of_stat_name = t.find("STAT_NAME")
                    stat_name = element_of_stat_name.text
                    stat_code = element_of_stat_name.attrib.get("code")
                    title = t.find("TITLE").text
                    page_dict[stat_id] = {"stat_name": stat_name, "stat_code": stat_code, "title": title}
            except Exception as e:
                error = e
                self.log.error(f"***{parser_xml.__doc__} => 失敗しました。: \n{str(e)}***")
            else:
                self.log.info(f"***{parser_xml.__doc__} => 成功しました。***")
            finally:
                # デバッグ用(加工前のデータをクリップボードにコピーする)
                pyperclip.copy(res.text)
                if error is not None:
                    raise error
                return page_dict, len(table_list)

        def parser_json(res: Response) -> Tuple[Dict[str, Dict[str, str]], int]:
            """JSONのデータを解析します"""
            page_dict: Dict[str, Dict[str, str]] = {}
            table_list: list = []
            error: Optional[Exception] = None
            try:
                data = res.json()
                table_list = data["GET_STATS_LIST"]["DATALIST_INF"]["TABLE_INF"]
                if isinstance(table_list, dict):
                    table_list = [table_list]
                for t in table_list:
                    stat_id = t.get("@id", "")
                    element_of_stat_name = t.get("STAT_NAME", {})
                    stat_name = element_of_stat_name.get("$", "")
                    stat_code = element_of_stat_name.get("@code", "")
                    statistics_name = t.get("STATISTICS_NAME", {})
                    title = t.get("TITLE", {})
                    page_dict[stat_id] = {
                        "stat_name": stat_name,
                        "stat_code": stat_code,
                        "statistics_name": statistics_name,
                        "title": title,
                    }
            except Exception as e:
                error = e
                self.log.error(f"***{parser_json.__doc__} => 失敗しました。: \n{str(e)}***")
            else:
                self.log.info(f"***{parser_json.__doc__} => 成功しました。***")
            finally:
                # デバッグ用(加工前のデータをクリップボードにコピーする)
                pyperclip.copy(json.dumps(data, indent=4, ensure_ascii=False))
                if error is not None:
                    raise error
                return page_dict, len(table_list)

        def parser_csv(res: Response) -> Tuple[Dict[str, Dict[str, str]], int]:
            """CSVのデータを解析します"""
            page_dict: Dict[str, Dict[str, str]] = {}
            row_count = 0
            error: Optional[Exception] = None
            try:
                res.encoding = "utf-8"
                reader = DictReader(StringIO(res.text))
                for row in reader:
                    row_count += 1
                    stat_id = row.get("TABLE_INF", "")
                    stat_name = row.get("STAT_NAME", "")
                    stat_code = row.get("STAT_CODE", "")
                    category = row.get("TABULATION_SUB_CATEGORY3", "")
                    page_dict[stat_id] = {"stat_name": stat_name, "stat_code": stat_code, "category": category}
            except Exception as e:
                error = e
                self.log.error(f"***{parser_csv.__doc__} => 失敗しました。: \n{str(e)}***")
            else:
                self.log.info(f"***{parser_csv.__doc__} => 成功しました。***")
            finally:
                # デバッグ用(加工前のデータをクリップボードにコピーする)
                pyperclip.copy(res.text)
                if error is not None:
                    raise error
                return page_dict, row_count

        URL = ""
        # データタイプに応じてジェネレータを返す
        match self.lst_of_data_type[self.KEY]:
            case "xml":
                URL = f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/getStatsList"
                async for d in fetch_all_pages(URL, parser_xml):
                    yield d
            case "json":
                URL = f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/json/getStatsList"
                async for d in fetch_all_pages(URL, parser_json):
                    yield d
            case "csv":
                URL = f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/getSimpleStatsList"
                async for d in fetch_all_pages(URL, parser_csv):
                    yield d
            case _:
                raise Exception("データタイプが対応していません。")

    def get_data_from_api(self) -> DataFrame:
        """APIからデータを取得します"""

        def get_params_of_url(self: GetGovernmentStatistics):
            """APIのURLのパラメータを取得します"""
            self.params = {
                "appId": self.APP_ID,  # アプリケーションID
                "statsDataId": self.STATS_DATA_ID,  # 統計表ID
                "lang": "J",  # 言語
                "limit": 100,
                "metaGetFlg": "Y",  # メタ情報の取得フラグ
                "cntGetFlg": "N",  # 件数の取得フラグ
                "explanationGetFlg": "N",  # 解説情報の有無フラグ
                "annotationGetFlg": "N",  # 注釈情報の有無フラグ
                "sectionHeaderFlg": 1,  # 見出し行の有無フラグ
                "replaceSpChars": 0,  # 特殊文字のエスケープフラグ
            }

        def with_xml(self: GetGovernmentStatistics) -> DataFrame:
            """XMLでデータを取得します"""
            try:
                result = False
                df = None
                self.URL = f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/getStatsData"
                get_params_of_url(self)
                # リクエストを送信する
                response = requests.get(self.URL, params=self.params)
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
                # デバッグ用(加工前のデータをクリップボードにコピーする)
                pyperclip.copy(response.text)
                if not result:
                    raise
                return df

        def with_json(self: GetGovernmentStatistics) -> DataFrame:
            """JSONでデータを取得します"""
            try:
                result = False
                df = None
                self.URL = f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/json/getStatsData"
                get_params_of_url(self)
                # リクエストを送信する
                response = requests.get(self.URL, params=self.params)
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
                # デバッグ用(加工前のデータをクリップボードにコピーする)
                pyperclip.copy(json.dumps(response.json(), indent=4, ensure_ascii=False))
                if not result:
                    raise
                return df

        def with_csv(self: GetGovernmentStatistics) -> DataFrame:
            """CSVでデータを取得します"""
            try:
                result = False
                df = None
                self.URL = f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/getSimpleStatsData"
                get_params_of_url(self)
                # リクエストを送信する
                response = requests.get(self.URL, params=self.params)
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
                # デバッグ用(加工前のデータをクリップボードにコピーする)
                pyperclip.copy(response.text)
                if not result:
                    raise
                return df

        def handle_data_type(self: GetGovernmentStatistics) -> DataFrame:
            """データタイプで条件分岐させます"""
            try:
                result = False
                df = None
                match self.lst_of_data_type[self.KEY]:
                    case "xml":
                        df = with_xml(self)
                    case "json":
                        df = with_json(self)
                    case "csv":
                        df = with_csv(self)
                    case _:
                        raise Exception("データタイプが対応していません。")
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
            df = handle_data_type(self)
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
