"""Tests for card.Card default state."""

from card import Card


def test_card_defaults():
    c = Card()
    assert c.title == ""
    assert c.titleShape is None
    assert c.bullets == ""
    assert c.graphic == ""
    assert c.graphicShape is None
    assert c.graphicDimensions is None
    assert c.graphicTitle is None
    assert c.printableFilename is None
    assert c.mediaInfo is None
    assert c.mediaDimensions is None
    assert c.mediaShape is None
    assert c.mediaURL is None
    assert c.backgroundShape is None
    assert c.backgroundTop is None
    assert c.bodyShape is None
    assert c.bodyTop is None
    assert c.top is None
    assert c.left is None
    assert c.missingCardMedia is None
    assert c.missingCardGraphic is None


def test_card_fields_are_mutable_and_independent():
    a, b = Card(), Card()
    a.title = "Deck"
    a.top = 42
    assert b.title == ""
    assert b.top is None
