from source.pdf_tools import pt_with_cui


# テスト関数: 指定されている関数の動作を確認するためのテストをする
def test_func(monkeypatch):
    # 標準入力(ex.input関数)の回数分の戻り値をリストにする
    list_of_inputs = [""]
    # リストをイテレータに変換する
    iter_of_inputs = iter(list_of_inputs)
    # monkeypatchでinputをシミュレートする
    # 標準入力の値を無名関数の戻り値のイテレータに置き換える
    monkeypatch.setattr("builtins.input", lambda *args: next(iter_of_inputs))
    # 指定された関数を呼び出す
    pt_with_cui.main()
