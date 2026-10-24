"""Tests for symbols.resolveSymbols: entity and numeric reference handling."""

import pytest

from symbols import resolveSymbols


@pytest.mark.parametrize(
    "text,expected",
    [
        # Named entities from the replacement table
        ("&equals;", "="),
        ("&le;", "≤"),
        ("&ge;", "≥"),
        ("&asymp;", "≈"),
        ("&Delta;", "Δ"),
        ("&delta;", "δ"),
        ("&sim;", "∼"),
        ("&semi;", ";"),
        ("&colon;", ":"),
        ("&comma;", ","),
        ("&amp;", "&"),
        ("&larr;", "←"),
        ("&rarr;", "→"),
        ("&uarr;", "↑"),
        ("&darr;", "↓"),
        ("&harr;", "↔"),
        ("&infin;", "∞"),
        ("&auml;", "ä"),
        ("&Auml;", "Ä"),
        ("&ouml;", "ö"),
        ("&szlig;", "ß"),
        ("&euro;", "€"),
        ("&check;", "✓"),
        ("&hellip;", "…"),
        ("&times;", "×"),
        ("&percnt;", "%"),
        ("&divide;", "÷"),
        ("&forall;", "∀"),
        ("&exist;", "∃"),
        ("&lambda;", "λ"),
        ("&pi;", "π"),
        ("&quot;", '"'),
        ("&apos;", "'"),
        ("&oslash;", "ø"),
        ("&Oslash;", "Ø"),
        # Entities embedded in text
        ("a &le; b", "a ≤ b"),
        ("&Delta;&delta;", "Δδ"),
        ("&euro;5", "€5"),
        ("a &amp; b", "a & b"),
        # Numeric character references (decimal and hex)
        ("&#65;", "A"),
        ("&#x41;", "A"),
        ("&#x48;&#105;", "Hi"),
        ("&#x6a;", "j"),
        # &nbsp; becomes a non-breaking space
        ("x&nbsp;y", "x\xa0y"),
        # &lt; / &gt; map to private-use placeholders, not literal brackets
        ("&lt;tag&gt;", "\ufde1tag\ufde2"),
        # Unknown entities and plain text pass through untouched
        ("&unknown;", "&unknown;"),
        ("plain text", "plain text"),
        ("", ""),
        ("100% & more", "100% & more"),
    ],
)
def test_resolve_symbols(text, expected):
    assert resolveSymbols(text) == expected
