import asyncio
import json
import shutil
import sys
import xml.etree.ElementTree as et
from csv import DictReader
from io import StringIO
from logging import Logger
from pathlib import Path
from typing import Any, AsyncGenerator, Dict, List, Optional, Union
from xml.etree.ElementTree import Element

import clipboard
import httpx
import pandas as pd
from httpx import Client, Response
from pandas import DataFrame
from tabulate import tabulate


class GetGovernmentStatistics:
    """政府の統計データを取得します"""

    def __init__(self, logger: Logger):
        """初期化します"""
        self.log: Logger = logger
        self.log.info(self.__class__.__doc__)
        # 統計表IDの一覧を取得する方法
        self.dct_of_get_type: dict = {
            "非同期": "処理の実行中に待ち時間が発生しても、次の処理に進める方式",
            "同期": "処理の実行中に待ち時間が発生しても、その処理の完了まで次に進まない方式",
        }
        # 取得するデータ形式
        self.dct_of_data_type: dict = {
            "xml": "タグ構造のデータ",
            "json": "キーと値のペアのデータ",
            "csv": "カンマ区切りのデータ",
        }
        # 検索方式
        self.dct_of_match: dict = {
            "部分一致": "フィールドの値にキーワードが含まれている",
            "完全一致": "フィールドの値がキーワードと完全に一致している",
            "検索しない": "なし",
        }
        # 抽出方式(2つ以上のキーワードがある場合)
        self.dct_of_logic: dict = {"OR抽出": "複数のキーワードのいずれかが含まれている", "AND抽出": "複数のキーワードの全てが含まれている"}
        # list変数のキー番号
        self.KEY: int = 0
        # list変数の説明番号
        self.DESCRIPTION: int = 1
        # 統計表IDの一覧を取得する方法
        self.lst_of_get_type: list = []
        # 取得するデータ形式
        self.lst_of_data_type: list = []
        # 検索方式
        self.lst_of_match: list = []
        # 抽出するキーワード
        self.lst_of_keyword: list = []
        # 抽出方式
        self.lst_of_logic: list = []
        # exe化されている場合とそれ以外を切り分ける
        exe_path: Path = Path(sys.executable) if getattr(sys, "frozen", False) else Path(__file__)
        # 統計表IDの一覧のCSVファイルを格納したフォルダ
        self.folder_p_of_ids: Path = exe_path.parent / "__stats_data_ids__"
        self.folder_s_of_ids: str = str(self.folder_p_of_ids)
        self.log.info(f"統計表IDのリストを格納するフォルダ => {self.folder_s_of_ids}")
        # 統計表IDの一覧のCSVファイルのヘッダー
        self.header_of_ids_l: list = ["統計表ID", "統計名", "表題"]
        self.header_of_ids_s: str = ",".join(self.header_of_ids_l)
        # APIのバージョン
        self.VERSION: float = 3.0
        # アプリケーションID
        self.APP_ID: str = ""
        # 統計表ID
        self.STATS_DATA_ID: str = ""
        # dataframeの件数
        self.DATA_COUNT: int = 0

    def _parser_xml(self, res: Response) -> tuple[dict, int]:
        """XMLのデータを解析します(同期版と非同期版で共通)"""
        page_dct: dict = {}
        try:
            root: Element[str] = et.fromstring(res.text)
            table_lst: list[Element[str]] = root.findall(".//TABLE_INF")
            for t in table_lst:
                stat_id: str = (t.attrib.get("id", "") or "") if t is not None else ""
                element_of_stat_name: Optional[Element] = t.find("STAT_NAME")
                stat_name: str = (element_of_stat_name.text or "") if element_of_stat_name is not None else ""
                element_of_title: Optional[Element] = t.find("TITLE")
                title: str = (element_of_title.text or "") if element_of_title is not None else ""
                page_dct[stat_id] = {"stat_name": stat_name, "title": title}
        except Exception as e:
            self.log.error(f"***{self.parser_xml.__doc__} => 失敗しました。***: \n{str(e)}")
            # デバッグ用(加工前のデータをクリップボードにコピーする)
            clipboard.copy(res.text)
        except KeyboardInterrupt:
            sys.exit(0)
        else:
            pass
        finally:
            pass
        return page_dct, len(table_lst)

    def _parser_json(self, res: Response) -> tuple[dict, int]:
        """JSONのデータを解析します(同期版と非同期版で共通)"""
        page_dct: dict = {}
        try:
            data: Any = res.json()
            table_data: Any = data["GET_STATS_LIST"]["DATALIST_INF"]["TABLE_INF"]
            table_lst = [table_data] if isinstance(table_data, dict) else table_data
            for t in table_lst:
                stat_id: str = t.get("@id", "")
                statistics_name: str = t.get("STATISTICS_NAME", {})
                title: str = t.get("TITLE", {})
                page_dct[stat_id] = {
                    "statistics_name": statistics_name,
                    "title": title,
                }
        except Exception as e:
            self.log.error(f"***{self.parser_json.__doc__} => 失敗しました。***: \n{str(e)}")
            # デバッグ用(加工前のデータをクリップボードにコピーする)
            clipboard.copy(json.dumps(data, indent=4, ensure_ascii=False))
        except KeyboardInterrupt:
            sys.exit(0)
        else:
            pass
        finally:
            pass
        return page_dct, len(table_lst)

    def _parser_csv(self, res: Response) -> tuple[dict, int]:
        """CSVのデータを解析します(同期版と非同期版で共通)"""
        page_dct: dict = {}
        row_count: int = 0
        try:
            lines: list = res.text.splitlines()
            # ヘッダー行を探す
            start_idx: int = 0
            for i, line in enumerate(lines):
                if "STAT_INF" in line:
                    # 次の行
                    start_idx = i + 1
                    break
            if start_idx == 0:
                raise Exception("CSVファイルにヘッダー行が見つかりません。")
            csv_text: str = "\n".join(lines[start_idx:])
            reader: DictReader[str] = DictReader(StringIO(csv_text))
            for row in reader:
                row_count += 1
                stat_id: str = row.get("TABLE_INF", "")
                stat_name: str = row.get("STAT_NAME", "")
                title: str = row.get("TITLE", "")
                page_dct[stat_id] = {"stat_name": stat_name, "title": title}
        except Exception as e:
            self.log.error(f"***{self.parser_csv.__doc__} => 失敗しました。***: \n{str(e)}")
            # デバッグ用(加工前のデータをクリップボードにコピーする)
            clipboard.copy(res.text)
        except KeyboardInterrupt:
            sys.exit(0)
        else:
            pass
        finally:
            pass
        return page_dct, row_count

    def _get_stats_data_ids_with_sync(self, dct_of_ids_url: dict) -> List[Dict[str, dict]]:
        """ページを取得します(同期版)"""
        try:
            data_type: str = self.lst_of_data_type[self.KEY]
            parser_map: dict = {"xml": self._parser_xml, "json": self._parser_json, "csv": self._parser_csv}
            parser: Any = parser_map.get(data_type)
            if not parser:
                raise Exception("データタイプが対応していません")
            url: str = dct_of_ids_url[data_type]
            results: list = []
            start: int = 1
            limit: int = 100
            with httpx.Client(timeout=120.0) as client:
                while True:
                    params: dict = {"appId": self.APP_ID, "lang": "J", "limit": limit, "startPosition": start}
                    res: Response = client.get(url, params=params)
                    res.encoding = "utf-8"
                    res.raise_for_status()
                    page_dct, count = parser(res)
                    if count == 0:
                        break
                    results.append(page_dct)
                    start += limit
        except Exception as e:
            self.log.error(f"***{self.get_stats_data_ids_with_sync.__doc__} => 失敗しました。***: \n{str(e)}")
        except KeyboardInterrupt:
            sys.exit(0)
        else:
            pass
        finally:
            pass
        return results

    async def _get_stats_data_ids_with_async(self, dct_of_ids_url: dict) -> AsyncGenerator[Dict[str, dict], None]:
        """ページを取得します(非同期版)"""
        try:
            data_type: str = self.lst_of_data_type[self.KEY]
            parser_map: dict = {"xml": self._parser_xml, "json": self._parser_json, "csv": self._parser_csv}
            parser: Any = parser_map.get(data_type)
            if not parser:
                raise Exception("データタイプが対応していません")
            url: str = dct_of_ids_url[data_type]
            start: int = 1
            limit: int = 100
            async with httpx.AsyncClient(timeout=120.0) as client:
                while True:
                    params: dict = {"appId": self.APP_ID, "lang": "J", "limit": limit, "startPosition": start}
                    res: Response = await client.get(url, params=params)
                    res.encoding = "utf-8"
                    res.raise_for_status()
                    page_dct, count = parser(res)
                    if count == 0:
                        break
                    yield page_dct
                    start += limit
        except Exception as e:
            self.log.error(f"***{self.get_stats_data_ids_with_async.__doc__} => 失敗しました。***: \n{str(e)}")
        except KeyboardInterrupt:
            sys.exit(0)
        else:
            pass
        finally:
            pass

    def write_stats_data_ids_to_file(self, chunk_size: int = 100) -> Union[bool, None]:
        """統計表IDの一覧をCSVファイルに書き込みます (同期版と非同期版を切り替える)"""
        try:
            # 統計表IDの一覧のURL
            dct_of_ids_url: dict = {
                "xml": f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/getStatsList",
                "json": f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/json/getStatsList",
                "csv": f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/getSimpleStatsList",
            }
            match self.lst_of_get_type[self.KEY]:
                case "非同期":
                    try:
                        loop: asyncio.AbstractEventLoop = asyncio.get_running_loop()
                    except RuntimeError:
                        # ループがない場合
                        return self._write_stats_data_ids_to_file_with_async(chunk_size, dct_of_ids_url)
                    else:
                        # ループがある場合
                        return loop.create_task(self._write_stats_data_ids_to_file_with_async(chunk_size, dct_of_ids_url))
                case "同期":
                    return self._write_stats_data_ids_to_file_with_sync(chunk_size, dct_of_ids_url)
                case _:
                    raise Exception("そのような取得方法は、ありません。")
        except Exception as e:
            self.log.error(f"***{self.write_stats_data_ids_to_file.__doc__} => 失敗しました。***: \n{str(e)}")
        except KeyboardInterrupt:
            sys.exit(0)
        finally:
            pass

    def _write_stats_data_ids_to_file_with_sync(self, chunk_size: int = 100, dct_of_ids_url: dict = {}) -> bool:
        """統計表IDの一覧をCSVファイルに書き込みます(同期版)"""
        result: bool = False
        try:
            self.folder_p_of_ids.mkdir(parents=True, exist_ok=True)
            # フォルダの中を空にする
            for e in self.folder_p_of_ids.iterdir():
                if e.is_dir():
                    shutil.rmtree(e)
                else:
                    e.unlink()
            pages: list = self._get_stats_data_ids_with_sync(dct_of_ids_url)
            buffer: list = [self.header_of_ids_s]
            file_index: int = 1
            for page in pages:
                for stat_id, info in page.items():
                    col2: str = info.get("stat_name", info.get("statistics_name", ""))
                    col3: str = info.get("title", "")
                    if col3:
                        # データクレンジング
                        col3 = col3.replace("\u002c", "\u3001").replace("\uff0c", "\u3001")
                    buffer.append(f"{stat_id},{col2},{col3}")
                    if len(buffer) >= chunk_size:
                        file_p_of_ids: Path = self.folder_p_of_ids / f"list_of_stats_data_ids_{file_index}.csv"
                        file_p_of_ids.write_text("\n".join(buffer), encoding="utf-8")
                        buffer.clear()
                        buffer.append(self.header_of_ids_s)
                        file_index += 1
            if len(buffer) > 1:
                file_p_of_ids = self.folder_p_of_ids / f"list_of_stats_data_ids_{file_index}.csv"
                file_p_of_ids.write_text("\n".join(buffer), encoding="utf-8")
        except Exception as e:
            self.log.error(f"***{self._write_stats_data_ids_to_file_with_sync.__doc__} => 失敗しました。***: \n{str(e)}")
        except KeyboardInterrupt:
            sys.exit(0)
        else:
            result = True
            self.log.info(f"***{self._write_stats_data_ids_to_file_with_sync.__doc__} => 成功しました。***")
        finally:
            pass
        return result

    async def _write_stats_data_ids_to_file_with_async(self, chunk_size: int = 100, dct_of_ids_url: dict = {}) -> bool:
        """統計表IDの一覧をCSVファイルに書き込みます(非同期版)"""
        result: bool = False
        try:
            self.folder_p_of_ids.mkdir(parents=True, exist_ok=True)
            # フォルダの中を空にする
            for e in self.folder_p_of_ids.iterdir():
                if e.is_dir():
                    shutil.rmtree(e)
                else:
                    e.unlink()
            buffer: list = [self.header_of_ids_s]
            file_index: int = 1
            async for page in self._get_stats_data_ids_with_async(dct_of_ids_url):
                for stat_id, info in page.items():
                    col2: str = info.get("stat_name", info.get("statistics_name", ""))
                    col3: str = info.get("title", "")
                    if col3:
                        # データクレンジング
                        col3 = col3.replace("\u002c", "\u3001").replace("\uff0c", "\u3001")
                    buffer.append(f"{stat_id},{col2},{col3}")
                    if len(buffer) >= chunk_size:
                        file_p_of_ids: Path = self.folder_p_of_ids / f"list_of_stats_data_ids_{file_index}.csv"
                        file_p_of_ids.write_text("\n".join(buffer), encoding="utf-8")
                        buffer.clear()
                        buffer.append(self.header_of_ids_s)
                        file_index += 1
            if len(buffer) > 1:
                file_p_of_ids: Path = self.folder_p_of_ids / f"list_of_stats_data_ids_{file_index}.csv"
                file_p_of_ids.write_text("\n".join(buffer), encoding="utf-8")
        except Exception as e:
            self.log.error(f"***{self._write_stats_data_ids_to_file_with_async.__doc__} => 失敗しました。***: \n{str(e)}")
        except KeyboardInterrupt:
            sys.exit(0)
        else:
            result = True
            self.log.info(f"***{self._write_stats_data_ids_to_file_with_async.__doc__} => 成功しました。***")
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

        def _with_xml(client: Client, dct_of_params: dict) -> DataFrame:
            """XMLでデータを取得します"""
            try:
                id_url: str = f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/getStatsData"
                # リクエストを送信する
                res: Response = client.get(id_url, params=dct_of_params)
                # 解析して、ルート要素を取得する
                root: Element[str] = et.fromstring(res.text)
                # CLASS_OBJからコードと名称のマッピングを作成する
                mapping: dict = {}
                for obj in root.findall(".//CLASS_OBJ"):
                    obj_id: str = obj.attrib["id"]
                    code_map: dict = {}
                    for cls in obj.findall("CLASS"):
                        # codeをキー、nameを値とする辞書を作成する
                        code_map[cls.attrib["code"]] = cls.attrib.get("name", cls.attrib["code"])
                    mapping[obj_id] = code_map
                # VALUEを取得し、行ごとの辞書に変換する
                rows: list = []
                for element in root.findall(".//VALUE"):
                    row: dict = {}
                    for key, value in element.attrib.items():
                        if key in mapping:
                            row[key] = mapping[key].get(value, value)
                        else:
                            row[key] = value
                    # VALUEのテキストを追加する
                    row["値"] = (element.text or "").strip()
                    rows.append(row)
                df: DataFrame = pd.DataFrame(rows)
                # 列名を日本語に変換する
                id2name: dict = {}
                for obj in root.findall(".//CLASS_OBJ"):
                    obj_id: str = obj.attrib["id"]
                    obj_name: str = obj.attrib.get("name", obj_id)
                    id2name[obj_id] = obj_name
                id2name["unit"] = "単位"
                df.rename(columns=id2name, inplace=True)
                # 値列を数値型に変換する
                if "値" in df.columns:
                    df["値"] = pd.to_numeric(df["値"], errors="coerce")
            except Exception as e:
                self.log.error(f"***{_with_xml.__doc__} => 失敗しました。***: \n{str(e)}")
                raise
            else:
                pass
            finally:
                # デバッグ用(加工前のデータをクリップボードにコピーする)
                clipboard.copy(res.text)
            return df

        def _with_json(client: Client, dct_of_params: dict) -> DataFrame:
            """JSONでデータを取得します"""
            try:
                id_url: str = f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/json/getStatsData"
                # リクエストを送信する
                res: Response = client.get(id_url, params=dct_of_params)
                data: Any = res.json()
                # CLASS_OBJとVALUEを抽出する
                class_inf: Any = data["GET_STATS_DATA"]["STATISTICAL_DATA"]["CLASS_INF"]["CLASS_OBJ"]
                values: Any = data["GET_STATS_DATA"]["STATISTICAL_DATA"]["DATA_INF"]["VALUE"]
                # idと列名の対応表を作成する
                col_name_map: dict = {obj["@id"]: obj["@name"] for obj in class_inf}
                col_name_map["unit"] = "単位"
                # CLASS_OBJ内のコードを日本語名に置換する辞書を作成する
                code_to_name: dict = {}
                for obj in class_inf:
                    cid: str = obj["@id"]
                    cls: Any = obj["CLASS"]
                    if isinstance(cls, list):
                        code_to_name[cid] = {c["@code"]: c["@name"] for c in cls}
                    else:
                        code_to_name[cid] = {cls["@code"]: cls["@name"]}
                # VALUEの各行を日本語に変換する
                translated_rows: list = []
                for value in values:
                    row: dict = {}
                    for k, v in value.items():
                        jp_col: str = ""
                        if k.startswith("@") and k[1:] in code_to_name:
                            jp_col: Any = col_name_map.get(k[1:], k[1:])
                            row[jp_col] = code_to_name[k[1:]].get(v, v)
                        elif k == "@unit":
                            row["単位"] = v
                        elif k == "$":
                            row["値"] = v
                        else:
                            row[k] = v
                    translated_rows.append(row)
                df: DataFrame = pd.DataFrame(translated_rows)
                # 値列を数値型に変換する
                if "値" in df.columns:
                    df["値"] = pd.to_numeric(df["値"], errors="coerce")
            except Exception as e:
                self.log.error(f"***{_with_json.__doc__} => 失敗しました。***: \n{str(e)}")
                raise
            else:
                pass
            finally:
                # デバッグ用(加工前のデータをクリップボードにコピーする)
                clipboard.copy(json.dumps(res.json(), indent=4, ensure_ascii=False))
            return df

        def _with_csv(client: Client, dct_of_params: dict) -> DataFrame:
            """CSVでデータを取得します"""
            try:
                id_url: str = f"http://api.e-stat.go.jp/rest/{self.VERSION}/app/getSimpleStatsData"
                # リクエストを送信する
                res: Response = client.get(id_url, params=dct_of_params)
                lines: list[str] = res.text.splitlines()
                # VALUE行の位置を検索する
                value_idx: int = 0
                for i, line in enumerate(lines):
                    if line.strip().replace('"', '') == "VALUE":
                        value_idx = i
                        break
                if value_idx == 0:
                    raise Exception("CSVに 'VALUE' 行が見つかりませんでした。")
                # ヘッダー行を取得する
                header_cols: list[str] = [h.strip('"') for h in lines[value_idx + 1].split(',')]
                # データ本体を文字列として抽出する
                csv_body: str = "\n".join(lines[value_idx + 2 :])
                df: DataFrame = pd.read_csv(StringIO(csv_body), header=None)
                df.columns = header_cols
                # 列名を日本語に置換し、不要な英語コード列を削除する
                rename_map: dict = {}
                drop_cols: list = []
                i: int = 0
                while i < len(header_cols):
                    eng: str = header_cols[i]
                    if eng.endswith("_code") and i + 1 < len(header_cols):
                        # 英語コード列は削除する
                        drop_cols.append(eng)
                        i += 2
                        continue
                    # 単独列を処理する
                    elif eng == "unit":
                        rename_map[eng] = "単位"
                    elif eng == "value":
                        rename_map[eng] = "値"
                    else:
                        rename_map[eng] = eng
                    i += 1
                df = df.rename(columns=rename_map)
                df = df.drop(columns=drop_cols)
                # 値列を数値型に変換する
                if "値" in df.columns:
                    df["値"] = pd.to_numeric(df["値"], errors="coerce")
            except Exception as e:
                self.log.error(f"***{_with_csv.__doc__} => 失敗しました。***: \n{str(e)}")
                raise
            else:
                pass
            finally:
                # デバッグ用(加工前のデータをクリップボードにコピーする)
                clipboard.copy(res.text)
            return df

        df: Optional[DataFrame] = None
        try:
            self.log.info(self.get_data_from_api.__doc__)
            dct_of_params: dict = get_params_of_url()
            # セッションを管理する
            with httpx.Client(timeout=120.0) as client:
                match self.lst_of_data_type[self.KEY]:
                    case "xml":
                        df = _with_xml(client, dct_of_params)
                    case "json":
                        df = _with_json(client, dct_of_params)
                    case "csv":
                        df = _with_csv(client, dct_of_params)
                    case _:
                        raise Exception("データタイプが対応していません。")
        except Exception as e:
            self.log.error(f"***{self.get_data_from_api.__doc__} => 失敗しました。***: \n{str(e)}")
            raise
        else:
            self.log.info(f"***{self.get_data_from_api.__doc__} => 成功しました。***")
        finally:
            pass
        return df

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
        finally:
            pass
        return filtered_df

    def show_table(self, df: DataFrame) -> bool:
        """表を表示させます"""
        result: bool = False
        try:
            self.log.info(self.show_table.__doc__)
            self.log.info(tabulate(df, headers="keys", tablefmt="pipe", showindex=False))
            self.log.info(f"統計表ID => {self.STATS_DATA_ID}")
            self.log.info("データの取得形式 => " + ": ".join(self.lst_of_data_type))
            self.log.info("検索方式 => " + ": ".join(self.lst_of_match))
            self.log.info("抽出するキーワード => " + (", ".join(map(str, self.lst_of_keyword)) if self.lst_of_keyword else "なし"))
            self.log.info("抽出方式 => " + (": ".join(self.lst_of_logic) if self.lst_of_logic else "なし"))
            self.DATA_COUNT: int = len(df)
            self.log.info(f"表示件数: {self.DATA_COUNT}")
        except Exception as e:
            self.log.error(f"***{self.show_table.__doc__} => 失敗しました。***: \n{str(e)}")
            raise
        else:
            result = True
            self.log.info(f"***{self.show_table.__doc__} => 成功しました。***")
        finally:
            pass
        return result
