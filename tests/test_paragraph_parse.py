"""Tests for paragraph.parseText, the inline-markup tokeniser.

parseText splits a markdown line into [style, text] fragments where the
style codes include "N" (normal), "B2" (bold), "I" (italic), "C" (code)
and "Link".
"""

from paragraph import parseText


def test_plain_text_is_a_single_normal_fragment():
    assert parseText("hello world") == [["N", "hello world"]]
    assert parseText("plain") == [["N", "plain"]]


def test_bold_span():
    assert parseText("**bold** text") == [["N", ""], ["B2", "bold"], ["N", " text"]]
    assert parseText("text **bold** and *it* end") == [
        ["N", "text "],
        ["B2", "bold"],
        ["N", " and "],
        ["I", "it"],
        ["N", " end"],
    ]


def test_two_bold_spans():
    assert parseText("**a** and **b**") == [
        ["N", ""],
        ["B2", "a"],
        ["N", " and "],
        ["B2", "b"],
    ]


def test_italic_span():
    assert parseText("a *italic* b") == [["N", "a "], ["I", "italic"], ["N", " b"]]
    assert parseText("*a* *b*") == [["N", ""], ["I", "a"], ["N", " "], ["I", "b"]]


def test_triple_asterisk_is_bold():
    assert parseText("***both***") == [["N", ""], ["B2", "both"], ["N", ""]]


def test_nested_italic_inside_bold_stays_bold_text():
    # The tokeniser does not nest: asterisks inside a bold span are dropped
    # and the whole span stays one bold fragment.
    assert parseText("**bold *it* more**") == [["N", ""], ["B2", "bold it more"]]


def test_code_span():
    assert parseText("`code` x") == [["N", ""], ["C", "code"], ["N", " x"]]
    assert parseText("`a` and `b`") == [
        ["N", ""],
        ["C", "a"],
        ["N", " and "],
        ["C", "b"],
    ]


def test_link_fragment_joins_text_and_url_with_sentinel():
    assert parseText("[text](http://x.co)") == [["Link", "text\ufdd8http://x.co"]]


def test_br_becomes_newline():
    assert parseText("a<br/>b") == [["N", "a\nb"]]


def test_escaped_asterisk_becomes_lowast_entity():
    assert parseText("\\*not italic\\*") == [["N", "&lowast;not italic&lowast;"]]


def test_spaced_asterisk_becomes_lowast_entity():
    assert parseText("a * b") == [["N", "a &lowast; b"]]


def test_underscores_are_not_emphasis():
    assert parseText("under_score_word") == [["N", "under_score_word"]]
    assert parseText("_it_ word") == [["N", "_it_ word"]]


def test_tilde_and_caret_markup_not_handled_by_parse_text():
    # Sub/superscript and strikethrough are handled elsewhere; parseText
    # leaves these markers in the normal fragment.
    assert parseText("~~strike~~") == [["N", "~~strike~~"]]
    assert parseText("H~2~O") == [["N", "H~2~O"]]
    assert parseText("x^2^") == [["N", "x^2^"]]


def test_heading_and_list_markers_pass_through():
    assert parseText("# not heading") == [["N", "# not heading"]]
    assert parseText("- dash") == [["N", "- dash"]]
