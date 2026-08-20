"""Tokenising of inline maths - $`LaTeX`$ - in paragraph.parseText.

parseText rewrites a line before it tokenises it ("\\#" to an entity reference,
"\\[" to a sentinel, "\\_" to "_") and the tokeniser then drops "*" and "[" when
they match none of its cases, even inside a code span.  LaTeX left in the line
therefore does not survive: "a*b" loses its asterisk, "\\left[" loses its
bracket, and "x**2" closes the code span it is sitting in.

So a formula is lifted out of the line before any of that runs and replaced by a
sentinel.  These tests pin that down, and pin down the places where a formula is
deliberately left as source text instead.
"""

import contextlib
import io
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import globals
import paragraph
import pptx_math
from symbols import resolveSymbols

# MathInserter only compiles the stylesheet it is given, and these tests stop at
# the fragment list without converting anything, so a stylesheet that compiles is
# enough.  MML2OMML.XSL itself ships with Office and cannot be redistributed.
MINIMAL_XSL = """<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:template match="/"/>
</xsl:stylesheet>
"""


def formulae(text):
    """The LaTeX of each formula parseText found, in order."""
    return [f[1] for f in paragraph.parseText(text) if f[0] == "Math"]


class InlineMathTestCase(unittest.TestCase):
    def setUp(self):
        self._directory = tempfile.TemporaryDirectory()
        self.addCleanup(self._directory.cleanup)

        stylesheet = Path(self._directory.name) / "mml2omml.xsl"
        stylesheet.write_text(MINIMAL_XSL, encoding="utf-8")
        self.stylesheet = stylesheet

        globals.processingOptions.setOptionValues("mathxsl", str(stylesheet))
        paragraph._mathInserterCache.clear()
        self.addCleanup(paragraph._mathInserterCache.clear)
        self.addCleanup(
            globals.processingOptions.setOptionValues, "mathxsl", ""
        )


class LaTeXSurvivesTests(InlineMathTestCase):
    """Characters parseText would otherwise rewrite or drop."""

    def test_formula_becomes_its_own_fragment(self):
        self.assertEqual(formulae(r"$`e^{i\pi}+1=0`$"), [r"e^{i\pi}+1=0"])

    def test_single_asterisk_survives(self):
        self.assertEqual(formulae(r"$`a*b`$"), [r"a*b"])

    def test_double_asterisk_survives(self):
        self.assertEqual(formulae(r"$`x**2`$"), [r"x**2"])

    def test_square_bracket_survives(self):
        self.assertEqual(
            formulae(r"$`\left[\frac{a}{b}\right]`$"),
            [r"\left[\frac{a}{b}\right]"],
        )

    def test_escaped_underscore_survives(self):
        self.assertEqual(formulae(r"$`f\_g`$"), [r"f\_g"])

    def test_escaped_hash_survives(self):
        self.assertEqual(formulae(r"$`\#\{x\}`$"), [r"\#\{x\}"])

    def test_surrounding_text_keeps_its_place(self):
        self.assertEqual(
            paragraph.parseText(r"before $`x`$ after"),
            [["N", "before "], ["Math", "x"], ["N", " after"]],
        )

    def test_several_formulae_keep_their_order(self):
        self.assertEqual(formulae(r"$`a`$ and $`b`$"), ["a", "b"])


class DelimiterTests(InlineMathTestCase):
    def test_currency_is_not_a_formula(self):
        """A bare "$" is not a delimiter, so this line is untouched."""
        self.assertEqual(
            paragraph.parseText("costs $5 to $10"), [["N", "costs $5 to $10"]]
        )

    def test_unclosed_opener_does_not_swallow_the_next_formula(self):
        """The LaTeX may not contain a backtick, so an unclosed opener cannot
        pair with the closing delimiter of the formula after it."""
        self.assertEqual(formulae(r"unclosed $` then text $`y`$ after"), ["y"])

    def test_escaped_dollar_inside_a_formula_survives(self):
        """GitHub documents this form, so "$" stays legal inside the LaTeX."""
        self.assertEqual(formulae(r"$`\sqrt{\$4}`$"), [r"\sqrt{\$4}"])

    def test_escaped_delimiter_is_not_a_formula(self):
        self.assertEqual(formulae(r"\$`x`$"), [])

    def test_unclosed_formula_is_left_alone(self):
        self.assertEqual(formulae(r"$`x`"), [])


class EmphasisTests(InlineMathTestCase):
    def test_bold_either_side_of_a_formula_stays_bold(self):
        """Bold accumulates as "B1" and is only emitted as "B2" when the closing
        "**" arrives, so the run before the formula has to be flushed as "B2"."""
        fragments = paragraph.parseText(r"**bold $`x`$ more**")
        self.assertIn(["B2", "bold "], fragments)
        self.assertIn(["B2", " more"], fragments)
        self.assertEqual(formulae(r"**bold $`x`$ more**"), ["x"])

    def test_italic_either_side_of_a_formula_stays_italic(self):
        """Italic needs no mapping: it accumulates as "I" and is emitted as "I".

        Bold does need one, which is why the two cases are both here.
        """
        fragments = paragraph.parseText(r"*italic $`x`$ more*")
        self.assertIn(["I", "italic "], fragments)
        self.assertIn(["I", " more"], fragments)
        self.assertEqual(formulae(r"*italic $`x`$ more*"), ["x"])

    def test_formula_in_superscript_is_still_a_formula(self):
        self.assertEqual(formulae(r"<sup>$`x`$</sup>"), ["x"])

    def test_formula_in_criticmarkup_is_still_a_formula(self):
        self.assertEqual(formulae(r"{++$`x`$++}"), ["x"])


class StructuredFragmentTests(InlineMathTestCase):
    """A span, an abbr and a link read their fragment as structure.

    Each accumulates "something, a separator, then the text" and splits it when
    it closes, so a formula emitted part way through shortens the split.  These
    formulae are left as source text.
    """

    def test_formula_in_link_text_is_left_as_source(self):
        fragments = paragraph.parseText(r"[$`x`$](url)")
        self.assertEqual(formulae(r"[$`x`$](url)"), [])
        self.assertTrue(
            any(f[0] == "Link" and r"$`x`$" in f[1] for f in fragments)
        )

    def test_formula_in_link_url_is_left_as_source(self):
        fragments = paragraph.parseText(r"[text]($`x`$)")
        self.assertEqual(formulae(r"[text]($`x`$)"), [])
        self.assertTrue(
            any(f[0] == "Link" and r"$`x`$" in f[1] for f in fragments)
        )

    def test_formula_in_a_styled_span_is_left_as_source(self):
        """The "<span style=" case only sets spanState once the fragment is
        non-empty, so a span at the start of a line needs its own flag."""
        source = r'<span style="color:red">$`x`$</span>'
        self.assertEqual(formulae(source), [])
        self.assertEqual(
            paragraph.parseText(source),
            [["SpanStyle", ["color:red", r"$`x`$"]]],
        )

    def test_formula_in_an_abbr_title_keeps_the_title(self):
        source = r'<abbr title="$`x`$">A</abbr>'
        self.assertEqual(
            paragraph.parseText(source)[-1], ["Gloss", "A", "A", r"$`x`$"]
        )

    def test_formula_after_a_span_is_a_formula_again(self):
        self.assertEqual(
            formulae(r'<span style="color:red">y</span> $`x`$'), ["x"]
        )

    def test_formula_after_an_abbr_is_a_formula_again(self):
        self.assertEqual(formulae(r'<abbr title="Full">A</abbr> $`x`$'), ["x"])


class AlreadySubstitutedTests(InlineMathTestCase):
    """Titles are resolveSymbols()ed before parseText sees them.

    createSectionSlide, the presentation title and the presentation subtitle all
    resolve symbols first, so a "\\`" or an entity reference inside a formula
    arrives as one of the sentinel noncharacters.  Converting that buries an
    unreadable character in the OMML, so the formula is left alone.
    """

    def test_substituted_escape_is_left_as_source(self):
        text = resolveSymbols(r"$`a\`b`$")
        self.assertEqual(formulae(text), [])
        self.assertEqual(
            paragraph.parseText(text), [["N", "$"], ["C", "a`b"], ["N", "$"]]
        )

    def test_substituted_entity_reference_is_left_as_source(self):
        self.assertEqual(formulae(resolveSymbols(r"$`x &lt; y`$")), [])


class StylesheetAvailabilityTests(InlineMathTestCase):
    def test_install_directory_copy_is_enough(self):
        """No mathxsl option, but a copy beside pptx_math.py - as with a block."""
        globals.processingOptions.setOptionValues("mathxsl", "")
        paragraph._mathInserterCache.clear()

        with mock.patch.object(
            pptx_math,
            "_resolve_mathxsl_path",
            return_value=self.stylesheet,
        ):
            self.assertEqual(formulae(r"$`x`$"), ["x"])

    def test_without_a_stylesheet_the_line_is_left_as_it_was(self):
        """Nothing is lifted, so the line tokenises exactly as it did before."""
        globals.processingOptions.setOptionValues("mathxsl", "")
        paragraph._mathInserterCache.clear()

        with mock.patch.object(
            pptx_math,
            "_resolve_mathxsl_path",
            side_effect=FileNotFoundError("no stylesheet"),
        ):
            with contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(
                    paragraph.parseText(r"$`x`$"),
                    [["N", "$"], ["C", "x"], ["N", "$"]],
                )

    def test_a_line_without_a_formula_never_asks_for_a_stylesheet(self):
        paragraph._mathInserterCache.clear()

        with mock.patch.object(
            pptx_math, "_resolve_mathxsl_path"
        ) as resolve:
            paragraph.parseText("plain text with no maths in it")

        resolve.assert_not_called()


if __name__ == "__main__":
    unittest.main()
