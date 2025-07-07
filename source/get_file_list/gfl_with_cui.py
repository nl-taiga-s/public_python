from pathlib import Path

from source.common.common import DatetimeTools, PathTools
from source.get_file_list.gfl_class import GetFileList


class GFL:
    def __init__(self):
        """初期化します"""
        self.d_of_bool = {
            "yes": ["はい", "1", "Yes", "yes", "Y", "y"],
            "no": ["いいえ", "0", "No", "no", "N", "n"],
        }
        self.obj_of_pt = PathTools()
        self.obj_of_dt2 = DatetimeTools()

    def input_folder_path(self) -> str:
        """フォルダのパスを入力します"""
        while True:
            folder_path = input("ファイルを検索したいフォルダを入力してください。: ")
            if folder_path == "":
                print("処理を中止します。")
                return
            fp = Path(folder_path)
            if fp.exists():
                # 存在する場合
                if fp.is_dir():
                    # フォルダの場合
                    return folder_path

    def input_bool_of_recursive(self) -> bool:
        """フォルダを再帰的に検索するかどうかを入力します"""
        while True:
            str_of_r = input(
                "フォルダを再帰的に検索するかどうかを入力してください。"
                "(Yes => y or No => n): "
            )
            match str_of_r:
                case var if var in self.d_of_bool["yes"]:
                    return True
                case var if var in self.d_of_bool["no"]:
                    return False
                case _:
                    raise ValueError(f"無効な入力です。: {str_of_r}")

    def input_pattern(self) -> str:
        """検索パターンを入力します"""
        pattern = input("ファイルの検索パターンを入力してください。: ")
        return pattern

    def input_bool_of_log(self) -> bool:
        """logファイルを出力するかどうかを入力します"""
        while True:
            str_of_l = input(
                "logファイルを出力するかどうかを入力してください。"
                "(Yes => y or No => n): "
            )
            match str_of_l:
                case var if var in self.d_of_bool["yes"]:
                    return True
                case var if var in self.d_of_bool["no"]:
                    return False
                case _:
                    raise ValueError(f"無効な入力です。: {str_of_l}")

    def output_log_file(self, lst: list):
        """logファイルを出力します"""
        if self.input_bool_of_log():
            try:
                fp_e = Path(__file__)
                fp_l = self.obj_of_pt.get_file_path_of_log(fp_e)
                file_path_of_log = str(fp_l)
                # ファイルに書き出し
                with open(file_path_of_log, "w", encoding="utf-8", newline="") as f:
                    for element in lst:
                        f.write(f"{element},")
                        f.write(f"{self.obj_of_dt2.convert_dt_to_str()}\n")
            except Exception as e:
                print(e)
            else:
                print(f"logファイルを出力しました。: {file_path_of_log}")

    def main(self):
        """メイン関数"""
        try:
            folder_path = self.input_folder_path()
            bool_of_r = self.input_bool_of_recursive()
            obj_of_cls = GetFileList(folder_path, bool_of_r)
            pattern = self.input_pattern()
            obj_of_cls.extract_by_pattern(pattern)
            print(*obj_of_cls.list_file_after, sep="\n")
            self.output_log_file(obj_of_cls.list_file_after)
        except Exception as e:
            print(e)


if __name__ == "__main__":
    obj_with_cui = GFL()
    obj_with_cui.main()
