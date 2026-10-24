"""Tests for funnel.massageFunnelText.

md2pptx internally encodes literal angle brackets as the characters
U+00EC (ì) and U+00ED (í); massageFunnelText converts them back.
"""

from funnel import massageFunnelText


def test_236_becomes_less_than():
    assert massageFunnelText("aìb") == "a<b"


def test_237_becomes_greater_than():
    assert massageFunnelText("aíb") == "a>b"


def test_both_and_repeated():
    assert massageFunnelText("ìtagí") == "<tag>"
    assert massageFunnelText("ììíí") == "<<>>"


def test_plain_text_unchanged():
    assert massageFunnelText("plain <text>") == "plain <text>"
    assert massageFunnelText("") == ""
