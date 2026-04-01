import pytest


def test_list_of_dicts():
    from xlutils.normalizer import normalize

    data = [{"a": 1, "b": 2}, {"a": 3, "b": 4}]
    headers, rows = normalize(data)
    assert headers == ["a", "b"]
    assert rows == [[1, 2], [3, 4]]


def test_empty_list():
    from xlutils.normalizer import normalize

    headers, rows = normalize([])
    assert headers == []
    assert rows == []


def test_pandas_dataframe():
    pd = pytest.importorskip("pandas")
    from xlutils.normalizer import normalize

    df = pd.DataFrame({"x": [10, 20], "y": ["foo", "bar"]})
    headers, rows = normalize(df)
    assert headers == ["x", "y"]
    assert rows == [[10, "foo"], [20, "bar"]]


def test_polars_dataframe():
    pl = pytest.importorskip("polars")
    from xlutils.normalizer import normalize

    df = pl.DataFrame({"x": [10, 20], "y": ["foo", "bar"]})
    headers, rows = normalize(df)
    assert headers == ["x", "y"]
    assert rows == [[10, "foo"], [20, "bar"]]


def test_unsupported_type_raises():
    from xlutils.normalizer import normalize

    with pytest.raises(Exception):
        normalize(42)
