from pathlib import Path

from source.common.common import PathTools


class COTP:
    def __init__(self):
        """初期化します"""
        self.obj_of_pt = PathTools()

    def input_folder_path(self) -> list:
        """フォルダのパスを入力します"""
        while True:
            folder_path_from = input(
                "ファイルを一括変換するフォルダを指定してください。: "
            )
            folder_path_to = input(
                "一括変換したファイルを格納するフォルダを指定してください。: "
            )

            fp_f = self.obj_of_pt.if_unc_path(folder_path_from)
            fp_t = self.obj_of_pt.if_unc_path(folder_path_to)
            if folder_path_from == "" or folder_path_to == "":
                print("処理を中止します。")
                return None, None
            elif fp_f.exists() and fp_t.exists():
                # 存在する場合
                if fp_f.is_dir() and fp_t.is_dir():
                    # フォルダの場合
                    return folder_path_from, folder_path_to

    def main(self):
        """メイン関数"""
        try:
            from source.convert_office_to_pdf.cotp_class import ConvertOfficeToPdf

            folder_path_from, folder_path_to = self.input_folder_path()
            obj_of_cls = ConvertOfficeToPdf(folder_path_from, folder_path_to)
            obj_of_cls.convert_all()
            fp_e = Path(__file__)
            obj_of_cls.write_log(fp_e)
        except EnvironmentError as e:
            print(e)
        except ValueError as e:
            print(e)
        except Exception as e:
            print(e)


if __name__ == "__main__":
    obj_with_cui = COTP()
    obj_with_cui.main()
