import asyncio
import json
import os
import traceback
import xml.etree.ElementTree as et
from csv import DictReader
from io import StringIO
from logging import Logger
from pathlib import Path
from typing import Any, AsyncGenerator, Callable, Coroutine, Dict, List, Optional, Tuple, Union, cast
from xml.etree.ElementTree import Element

import clipboard
import httpx
import pandas as pd
import pyperclip
from httpx import AsyncClient, Client, Response, Timeout
from pandas import DataFrame
from tabulate import tabulate

from source.common.common import DatetimeTools


class GetGovernmentStatistics:
    """政府の統計データを取得します"""

    def __init__(self, logger: Logger):
        """初期化します"""
        self.log: Logger = logger
        self.log.info(self.__class__.__doc__)
        self.obj_of_dt2: DatetimeTools = DatetimeTools()
        # dict変数のキー番号
        self.KEY: int = 0
        # dict変数の説明番号
        self.DESCRIPTION: int = 1
        # 取得するデータ形式
        self.lst_of_data_type: list = []
        # APIのURL
        self.URL: str = ""
        # APIのバージョン
        self.VERSION: float = 3.0
        # APIのURLのパラメータ
        self.params: dict = {}
        # アプリケーションIDを取得して、環境変数で指定しておく
        self.APP_ID: str = os.environ.get("first_appid_of_estat")
        # 統計表ID
        self.STATS_DATA_ID: str = ""
        # 部分一致か完全一致か
        self.lst_of_match: list = []
        # 先頭か末尾か
        self.order: str = ""
        # 抽出するキーワード
        self.lst_of_keyword: list = []
        # OR抽出かAND抽出か
        self.lst_of_logic: list = []
        # 表示するデータの件数
        self.DATA_COUNT: int = 12
        # セッション管理のタイムアウト
        self.TIMEOUT: Timeout = Timeout(60.0)

    async def get_stats_data_ids(self) -> AsyncGenerator[dict]:
        """統計表IDの一覧を取得します"""

        async def fetch_page(client: AsyncClient, url: str, params: dict, parser: Callable[[Response], tuple[dict, int]]) -> tuple[dict, int]:
            """1ページ分を取得します"""
            try:
                res: Response = await client.get(url, params=params)
                res.encoding = "utf-8"
                res.raise_for_status()
                page_dct, count = parser(res)
            except Exception as e:
                tb: str = traceback.format_exc()
                self.log.error(f"***{fetch_page.__doc__} => 失敗しました。***: \n{str(e)}\n{tb}")
                raise
            else:
                return page_dct, count
            finally:
                pass

        async def fetch_all_pages(
            url: str,
            parser: Callable[[Response], tuple[dict, int]],
            page_limit: int = 100,  # 1回のリクエストで取得する件数
            concurrency: int = 5,  # 非同期で処理する並列数
        ) -> AsyncGenerator[dict]:
            """共通処理でページごとに取得します"""
            # タイムアウト時間を設定する
            async with httpx.AsyncClient(timeout=self.TIMEOUT) as client:
                start: int = 1
                while True:
                    try:
                        tasks: List[Coroutine[Any, Any, Tuple[Dict[Any, Any], int]]] = [
                            fetch_page(
                                client,
                                url,
                                {"appId": self.APP_ID, "lang": "J", "limit": page_limit, "startPosition": start + i * page_limit},
                                parser,
                            )
                            for i in range(concurrency)
                        ]
                        results = cast(List[Union[Tuple[Dict[Any, Any], int], BaseException]], await asyncio.gather(*tasks, return_exceptions=True))
                        stop: bool = True
                        for res in results:
                            if isinstance(res, BaseException):
                                tb: str = traceback.format_exc()
                                self.log.error(f"タスクで例外発生: \n{str(res)}\n{tb}")
                                continue
                            page_dct, count = res
                            if count > 0:
                                yield page_dct
                                stop = False
                        if stop:
                            break
                    except Exception as e:
                        self.log.error(f"***{fetch_all_pages.__doc__} => 失敗しました。***: \n{str(e)}")
                    else:
                        pass
                    finally:
                        start += page_limit * concurrency

        def parser_xml(res: Response) -> tuple[dict, int]:
            """XMLのデータを解析します"""
            page_dct: dict = {}
            try:
                root: Element[str] = et.fromstring(res.text)
                table_lst: list[Element[str]] = root.findall(".//TABLE_INF")
                for t in table_lst:
                    stat_id: str = (t.attrib.get("id", "") or "") if t is not None else ""
                    element_of_stat_name: Optional[Element] = t.find("STAT_NAME")
                    stat_name: str = (element_of_stat_name.text or "") if element_of_stat_name is not None else ""
                    stat_code: str = (element_of_stat_name.attrib.get("code") or "") if element_of_stat_name is not None else ""
                    element_of_title: Optional[Element] = t.find("TITLE")
                    title: str = (element_of_title.text or "") if element_of_title is not None else ""
                    page_dct[stat_id] = {"stat_name": stat_name, "stat_code": stat_code, "title": title}
            except Exception as e:
                self.log.error(f"***{parser_xml.__doc__} => 失敗しました。***: \n{str(e)}")
                raise
            else:
                return page_dct, len(table_lst)
            finally:
                # デバッグ用(加工前のデータをクリップボードにコピーする)
                # clipboard.copy(res.text)
                pass

        def parser_json(res: Response) -> tuple[dict, int]:
            """JSONのデータを解析します"""
            page_dct: dict = {}
            try:
                data: Any = res.json()
                table_data: Any = data["GET_STATS_LIST"]["DATALIST_INF"]["TABLE_INF"]
                table_lst = [table_data] if isinstance(table_data, dict) else table_data
                for t in table_lst:
                    stat_id: str = t.get("@id", "")
                    element_of_stat_name: dict = t.get("STAT_NAME", {})
                    stat_name: str = element_of_stat_name.get("$", "")
                    stat_code: str = element_of_stat_name.get("@code", "")
                    statistics_name: str = t.get("STATISTICS_NAME", {})
                    title: str = t.get("TITLE", {})
                    page_dct[stat_id] = {
                        "stat_name": stat_name,
                        "stat_code": stat_code,
                        "statistics_name": statistics_name,
                        "title": title,
                    }
            except Exception as e:
                self.log.error(f"***{parser_json.__doc__} => 失敗しました。***: \n{str(e)}")
                raise
            else:
                return page_dct, len(table_lst)
            finally:
                # デバッグ用(加工前のデータをクリップボードにコピーする)
                clipboard.copy(json.dumps(data, indent=4, ensure_ascii=False))

        def parser_csv(res: Response) -> tuple[dict, int]:
            """CSVのデータを解析します"""
            page_dct: dict = {}
            row_count: int = 0
            try:
                reader: DictReader[str] = DictReader(StringIO(res.text))
                for row in reader:
                    row_count += 1
                    stat_id: str = row.get("TABLE_INF", "")
                    stat_name: str = row.get("STAT_NAME", "")
                    stat_code: str = row.get("STAT_CODE", "")
                    category: str = row.get("TABULATION_SUB_CATEGORY3", "")
                    page_dct[stat_id] = {"stat_name": stat_name, "stat_code": stat_code, "category": category}
            except Exception as e:
                self.log.error(f"***{parser_csv.__doc__} => 失敗しました。***: \n{str(e)}")
                raise
            else:
                return page_dct, row_count
            finally:
                # デバッグ用(加工前のデータをクリップボードにコピーする)
                clipboard.copy(res.text)

        URL: str = ""
        try:
            # データタイプに応じてジェネレータを返す
            match self.lst_of_data_type[self.KEY]:
                case "xml":
                    URL = f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/getStatsList"
                    async for dct in fetch_all_pages(URL, parser_xml):
                        yield dct
                case "json":
                    URL = f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/json/getStatsList"
                    async for dct in fetch_all_pages(URL, parser_json):
                        yield dct
                case "csv":
                    URL = f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/getSimpleStatsList"
                    async for dct in fetch_all_pages(URL, parser_csv):
                        yield dct
                case _:
                    raise Exception("データタイプが対応していません。")
        except Exception as e:
            self.log.error(f"***{self.get_stats_data_ids.__doc__} => 失敗しました。***: \n{str(e)}")
            raise
        else:
            self.log.info(f"***{self.get_stats_data_ids.__doc__} => 成功しました。***")
        finally:
            pass

    async def write_stats_data_ids_to_file(
        self,
        page_generator: AsyncGenerator[dict],
        data_type_key: str,
        chunk_size: int = 100,
    ) -> bool:
        """統計表IDを指定の件数ごとにファイルへ保存します"""

        def get_info_columns(info: dict, data_type_key: str) -> tuple[str, str]:
            """2列目・3列目を安全に取得します"""
            try:
                match data_type_key:
                    case "xml":
                        col2 = info.get("stat_name")
                        col3 = info.get("title")
                    case "json":
                        col2 = info.get("statistics_name") or info.get("stat_name") or ""
                        col3 = info.get("title") or ""
                    case "csv":
                        col2 = info.get("stat_name") or ""
                        col3 = info.get("category") or ""
                    case _:
                        col2 = ""
                        col3 = ""
            except Exception as e:
                self.log.error(f"***{get_info_columns.__doc__} => 失敗しました。***: \n{str(e)}")
                raise
            else:
                return col2, col3
            finally:
                pass

        result: bool = False
        try:
            folder_p: Path = Path(__file__).parent / "__stats_data_ids__"
            folder_p.mkdir(parents=True, exist_ok=True)
            folder_s: str = str(folder_p)
            self.log.info(f"統計表IDのリストを格納したフォルダ => {folder_s}")
            file_index: int = 1
            buffer: list[str] = []
            async for page in page_generator:
                for stat_id, info in page.items():
                    col2, col3 = get_info_columns(info, data_type_key)
                    buffer.append(f"{stat_id}, {col2}, {col3}")
                    # chunkごとにファイルに保存する
                    if len(buffer) >= chunk_size:
                        date: str = self.obj_of_dt2.convert_for_file_name()
                        file_name: str = f"list_of_stats_data_ids_{date}_{file_index}.csv"
                        file_p: Path = folder_p / file_name
                        file_p.write_text("\n".join(buffer), encoding="utf-8")
                        buffer.clear()
                        file_index += 1
            # 余りをファイルに保存する
            if buffer:
                date: str = self.obj_of_dt2.convert_for_file_name()
                file_name: str = f"list_of_stats_data_ids_{date}_{file_index}.csv"
                file_p: Path = folder_p / file_name
                file_p.write_text("\n".join(buffer), encoding="utf-8")
        except Exception as e:
            self.log.debug(repr(e))
            self.log.error(f"***{self.write_stats_data_ids_to_file.__doc__} => 失敗しました。***: \n{str(e)}")
            raise
        except KeyboardInterrupt:
            raise
        else:
            result = True
            self.log.info(f"***{self.write_stats_data_ids_to_file.__doc__} => 成功しました。***")
        finally:
            pass
        return result

    def get_data_from_api(self) -> DataFrame:
        """APIからデータを取得します"""

        def get_params_of_url() -> dict:
            """APIのURLのパラメータを取得します"""
            params: dict = {
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
            return params

        def with_xml(client: Client) -> DataFrame:
            """XMLでデータを取得します"""
            try:
                self.URL = f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/getStatsData"
                # リクエストを送信する
                res: Response = client.get(self.URL, params=self.params)
                root: Element[str] = et.fromstring(res.text)
                # --- 1. CLASS_INFを辞書化 ---
                mapping: dict = {}
                for obj in root.findall(".//CLASS_OBJ"):
                    obj_id: str = obj.attrib["id"]
                    code_map: dict = {}
                    for cls in obj.findall("CLASS"):
                        code_map[cls.attrib["code"]] = cls.attrib.get("name", cls.attrib["code"])
                    mapping[obj_id] = code_map
                # --- 2. VALUE要素を取り込む ---
                rows: list = []
                for value in root.findall(".//VALUE"):
                    row: dict = {}
                    for key, value in value.attrib.items():
                        if key in mapping:
                            row[key] = mapping[key].get(value, value)
                        else:
                            row[key] = value
                    if isinstance(value, Element):
                        row["値"] = (value.text or "").strip()
                    else:
                        row["値"] = str(value).strip()
                    rows.append(row)
                df: DataFrame = pd.DataFrame(rows)
                # --- 3. 列名をCLASS_OBJのname属性に置換 ---
                id2name: dict = {}
                for obj in root.findall(".//CLASS_OBJ"):
                    obj_id: str = obj.attrib["id"]
                    obj_name: str = obj.attrib.get("name", obj_id)
                    id2name[obj_id] = obj_name
                id2name["unit"] = "単位"
                df.rename(columns=id2name, inplace=True)
                # --- 値列を数値型に変換 ---
                if "値" in df.columns:
                    df["値"] = pd.to_numeric(df["値"], errors="coerce")
            except Exception as e:
                self.log.error(f"***{with_xml.__doc__} => 失敗しました。***: \n{str(e)}")
                # デバッグ用(加工前のデータをクリップボードにコピーする)
                pyperclip.copy(res.text)
                raise
            else:
                return df
            finally:
                pass

        def with_json(client: Client) -> DataFrame:
            """JSONでデータを取得します"""
            try:
                self.URL = f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/json/getStatsData"
                # リクエストを送信する
                res: Response = client.get(self.URL, params=self.params)
                data: Any = res.json()
                # --- CLASS_INF と VALUE を取得 ---
                class_inf: Any = data["GET_STATS_DATA"]["STATISTICAL_DATA"]["CLASS_INF"]["CLASS_OBJ"]
                values: Any = data["GET_STATS_DATA"]["STATISTICAL_DATA"]["DATA_INF"]["VALUE"]
                # --- 列名用のコード(@xxx) → 日本語(@name) の辞書 ---
                col_name_map: dict = {obj["@id"]: obj["@name"] for obj in class_inf}
                # "@unit" を "時間軸" の後に追加
                new_col_map: dict = {}
                for k, v in col_name_map.items():
                    new_col_map[k] = v
                    if k == "time":
                        # "@time"（時間軸）の後
                        new_col_map["unit"] = "単位"
                col_name_map = new_col_map
                # --- コード→日本語辞書を作成 ---
                code_to_name: dict = {}
                for obj in class_inf:
                    cid: str = obj["@id"]
                    cls: Any = obj["CLASS"]
                    if isinstance(cls, list):
                        code_to_name[cid] = {c["@code"]: c["@name"] for c in cls}
                    else:
                        code_to_name[cid] = {cls["@code"]: cls["@name"]}
                # --- VALUE 内のコードを日本語に置換 ---
                translated_rows: list = []
                for value in values:
                    row: dict = {}
                    for k, val in value.items():
                        jp_col: str = ""
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
                df: DataFrame = pd.DataFrame(translated_rows)
                # --- 値列を数値型に変換 ---
                if "値" in df.columns:
                    df["値"] = pd.to_numeric(df["値"], errors="coerce")
            except Exception as e:
                self.log.error(f"***{with_json.__doc__} => 失敗しました。***: \n{str(e)}")
                # デバッグ用(加工前のデータをクリップボードにコピーする)
                pyperclip.copy(json.dumps(res.json(), indent=4, ensure_ascii=False))
                raise
            else:
                return df
            finally:
                pass

        def with_csv(client: Client) -> DataFrame:
            """CSVでデータを取得します"""
            try:
                self.URL = f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/getSimpleStatsData"
                # リクエストを送信する
                res: Response = client.get(self.URL, params=self.params)
                lines: list[str] = res.text.splitlines()
                # VALUE行の位置を探す
                value_idx: int = 0
                for i, line in enumerate(lines):
                    if line.strip().replace('"', '') == "VALUE":
                        value_idx = i
                        break
                if value_idx == 0:
                    raise Exception("CSVに 'VALUE' 行が見つかりませんでした。")
                # --- ヘッダー行 ---
                header_cols: list[str] = [h.strip('"') for h in lines[value_idx + 1].split(',')]
                # --- データ本体をそのまま読み込む ---
                csv_body: str = "\n".join(lines[value_idx + 2 :])
                df: DataFrame = pd.read_csv(StringIO(csv_body), header=None)
                df.columns = header_cols  # 英語+日本語ペア含めて13列のまま
                # --- 列名置換 ---
                rename_map: dict = {}
                drop_cols: list = []
                i: int = 0
                while i < len(header_cols):
                    eng: str = header_cols[i]
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
                self.log.error(f"***{with_csv.__doc__} => 失敗しました。***: \n{str(e)}")
                # デバッグ用(加工前のデータをクリップボードにコピーする)
                pyperclip.copy(res.text)
                raise
            else:
                return df
            finally:
                pass

        df: Optional[DataFrame] = None
        try:
            self.log.info(self.get_data_from_api.__doc__)
            self.params = get_params_of_url()
            # セッションを管理する
            with httpx.Client(timeout=self.TIMEOUT) as client:
                match self.lst_of_data_type[self.KEY]:
                    case "xml":
                        df = with_xml(client)
                    case "json":
                        df = with_json(client)
                    case "csv":
                        df = with_csv(client)
                    case _:
                        raise Exception("データタイプが対応していません。")
        except Exception as e:
            self.log.error(f"***{self.get_data_from_api.__doc__} => 失敗しました。***: \n{str(e)}")
            raise
        else:
            self.log.info(f"***{self.get_data_from_api.__doc__} => 成功しました。***")
            return df
        finally:
            pass

    def filter_data(self, df: DataFrame) -> DataFrame:
        """データをフィルターにかけます"""
        filtered_df: Optional[DataFrame] = None
        try:
            self.log.info(self.filter_data.__doc__)
            match self.lst_of_match[self.KEY]:
                case "部分一致":
                    # 全列で部分一致検索する
                    if len(self.lst_of_keyword) == 1:
                        # 単一キーワード
                        kw: str = str(self.lst_of_keyword[0])
                        filtered_df = df[df.apply(lambda row: row.astype(str).str.contains(kw, case=False, na=False).any(), axis=1)]
                    else:
                        # 複数キーワード
                        match self.lst_of_logic[self.KEY]:
                            case "OR抽出":
                                pattern: str = "|".join(map(str, self.lst_of_keyword))
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
                        kw: str = str(self.lst_of_keyword[0])
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
            self.log.error(f"***{self.filter_data.__doc__} => 失敗しました。***: \n{str(e)}")
            raise
        else:
            self.log.info(f"***{self.filter_data.__doc__} => 成功しました。***")
            return filtered_df
        finally:
            pass

    def show_data(self, df: DataFrame) -> bool:
        """データを表示させます"""
        result: bool = False
        try:
            self.log.info(self.show_data.__doc__)
            match self.order:
                case "先頭":
                    self.log.info(tabulate(df.head(self.DATA_COUNT), headers="keys", tablefmt="pipe", showindex=False))
                case "末尾":
                    self.log.info(tabulate(df.tail(self.DATA_COUNT), headers="keys", tablefmt="pipe", showindex=False))
                case _:
                    raise Exception("その表示順はありません。")
            self.log.info(f"データの取得形式: {self.lst_of_data_type[self.KEY]} => {self.lst_of_data_type[self.DESCRIPTION]}")
            self.log.info(f"検索方式: {self.lst_of_match[self.KEY]} => {self.lst_of_match[self.DESCRIPTION]}")
            self.log.info(f"抽出するキーワード: {", ".join(map(str, self.lst_of_keyword))}")
            if self.lst_of_logic:
                self.log.info(f"抽出方式: {self.lst_of_logic[self.KEY]} => {self.lst_of_logic[self.DESCRIPTION]}")
            self.log.info(f"表示順: {self.order}")
            self.log.info(f"表示件数: {self.DATA_COUNT}")
        except Exception as e:
            self.log.error(f"***{self.show_data.__doc__} => 失敗しました。***: \n{str(e)}")
            raise
        else:
            result = True
            self.log.info(f"***{self.show_data.__doc__} => 成功しました。***")
            return result
        finally:
            pass
