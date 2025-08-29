import pytest

from source.get_nhk_news import gn2_with_cui


# テスト関数: 指定されている関数の動作を確認するためのテストをする
def test_func(monkeypatch):
    # main を強制的に KeyboardInterrupt させる
    def fake_loop():
        raise KeyboardInterrupt

    monkeypatch.setattr("source.get_nhk_news.gn2_with_cui.main", fake_loop)
    with pytest.raises(KeyboardInterrupt):
        gn2_with_cui.main()
