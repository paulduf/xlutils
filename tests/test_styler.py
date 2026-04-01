from xlutils import style

DATA = [{"product": "A", "revenue": 100}, {"product": "B", "revenue": 200}]


def test_default_state():
    s = style(DATA)
    assert s._theme.name == "default"
    assert s._column_mode == "fit"
    assert s._freeze_header is True
    assert s._gradient_rules == []


def test_apply_theme_by_name():
    s = style(DATA).apply_theme("minimal")
    assert s._theme.name == "minimal"


def test_apply_theme_object():
    from xlutils.themes import THEME_REGISTRY

    t = THEME_REGISTRY["minimal"]
    s = style(DATA).apply_theme(t)
    assert s._theme is t


def test_unknown_theme_raises():
    import pytest

    with pytest.raises(ValueError, match="Unknown theme"):
        style(DATA).apply_theme("nonexistent")


def test_expand_columns():
    s = style(DATA).expand_columns(20)
    assert s._column_mode == 20


def test_freeze_header():
    s = style(DATA).freeze_header(False)
    assert s._freeze_header is False


def test_color_gradient_accumulates():
    s = (
        style(DATA)
        .color_gradient("revenue")
        .color_gradient("product", "#FF0000", "#00FF00")
    )
    assert len(s._gradient_rules) == 2
    assert s._gradient_rules[0]["column"] == "revenue"
    assert s._gradient_rules[1]["low"] == "#FF0000"


def test_chain_returns_same_instance():
    s = style(DATA)
    assert s.apply_theme("minimal") is s
    assert s.expand_columns("fit") is s
    assert s.freeze_header() is s
    assert s.color_gradient("revenue") is s
