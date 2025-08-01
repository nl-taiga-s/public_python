from source.get_file_list import gfl_with_cui


# テスト関数: 指定されている関数の動作を確認するためのテストをする
def test_func(monkeypatch):
    # 標準入力(ex.input関数)の回数分の戻り値をリストにする
    list_of_inputs = [""]
    # リストをイテレータに変換する
    iter_of_inputs = iter(list_of_inputs)
    # monkeypatchでinputをシミュレートする
    # input関数の引数が
    # ある場合 => lambda _
    # ない場合 => lambda
    # 標準入力の値を無名関数の戻り値のイテレータに置き換える
    monkeypatch.setattr("builtins.input", lambda _: next(iter_of_inputs))
    # 指定された関数を呼び出す
    gfl_with_cui.main()
