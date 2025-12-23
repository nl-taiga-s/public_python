from pathlib import Path

from source.common.common import DatetimeTools, LogTools


def main():
    """主要関数"""
    result: bool = False
    # ログを設定します
    try:
        obj_of_dt2: DatetimeTools = DatetimeTools()
        obj_of_lt: LogTools = LogTools()
        # ログフォルダのパス
        folder_p: Path = Path(__file__).parent / "__log__"
        # ログフォルダが存在しない場合は、作成します
        folder_p.mkdir(parents=True, exist_ok=True)
        # ログファイル名
        file_name: str = f"log_{obj_of_dt2._convert_for_file_name()}.log"
        file_p: Path = folder_p / file_name
        obj_of_lt.file_path_of_log = str(file_p)
        obj_of_lt._setup_file_handler(obj_of_lt.file_path_of_log)
        obj_of_lt._setup_stream_handler()
    except Exception as e:
        print(f"error: \n{str(e)}")
        return result
    else:
        pass
    finally:
        pass
    # 処理の本体部分
