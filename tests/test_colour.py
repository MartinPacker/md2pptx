"""Tests for colour.py: hex/theme colour parsing and application."""

import pytest
from pptx.dml.color import MSO_THEME_COLOR, RGBColor

from colour import parseColour, parseRGB, parseThemeColour, setColour


class FakeFill:
    """Stand-in for a pptx fill/line object that accepts colour assignment."""


@pytest.mark.parametrize(
    "value,expected",
    [
        ("#FF0000", ("RGB", "#FF0000")),
        ("#000000", ("RGB", "#000000")),
        ("#a1b2c3", ("RGB", "#a1b2c3")),
    ],
)
def test_parse_colour_hex(value, expected):
    assert parseColour(value) == expected


@pytest.mark.parametrize(
    "value,expected",
    [
        ("Accent 1", MSO_THEME_COLOR.ACCENT_1),
        ("ACCENT 6", MSO_THEME_COLOR.ACCENT_6),
        ("Text 1", MSO_THEME_COLOR.TEXT_1),
        ("Text 2", MSO_THEME_COLOR.TEXT_2),
        ("Light 1", MSO_THEME_COLOR.LIGHT_1),
        ("Dark 2", MSO_THEME_COLOR.DARK_2),
        ("Background 1", MSO_THEME_COLOR.BACKGROUND_1),
        ("Hyperlink", MSO_THEME_COLOR.HYPERLINK),
        ("Followed Hyperlink", MSO_THEME_COLOR.FOLLOWED_HYPERLINK),
        ("NONE", MSO_THEME_COLOR.NOT_THEME_COLOR),
        ("Mixed", MSO_THEME_COLOR.MIXED),
    ],
)
def test_parse_colour_theme_names(value, expected):
    assert parseColour(value) == ("Theme", expected)


def test_parse_colour_theme_whitespace_and_case_normalised():
    assert parseColour("accent  2") == ("Theme", MSO_THEME_COLOR.ACCENT_2)
    assert parseThemeColour("  light   2 ") == MSO_THEME_COLOR.LIGHT_2


def test_parse_colour_unknown_theme_falls_back_to_accent_1(capsys):
    assert parseColour("bogus name") == ("Theme", MSO_THEME_COLOR.ACCENT_1)
    assert "Invalid theme colour" in capsys.readouterr().out


@pytest.mark.parametrize(
    "value,expected",
    [
        ("#12abEF", (True, "12abEF")),
        ("#000000", (True, "000000")),
        ("#FFFFFF", (True, "FFFFFF")),
        ("#12345", (False, "")),  # too short
        ("#1234567", (True, "123456")),  # regex matches a 6-digit prefix
        ("123456", (False, "")),  # missing leading '#'
        ("#GGGGGG", (False, "")),
        ("", (False, "")),
        ("red", (False, "")),
    ],
)
def test_parse_rgb(value, expected):
    assert parseRGB(value) == expected


def test_set_colour_rgb():
    fill = FakeFill()
    setColour(fill, ("RGB", "#A1B2C3"))
    assert fill.rgb == RGBColor(0xA1, 0xB2, 0xC3)


def test_set_colour_theme():
    fill = FakeFill()
    setColour(fill, ("Theme", MSO_THEME_COLOR.ACCENT_3))
    assert fill.theme_color == MSO_THEME_COLOR.ACCENT_3
