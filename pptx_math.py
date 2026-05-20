"""
pptx_math.py
============
Insert native PowerPoint math (OMML) into slides via python-pptx.

This is a minimal build for md2pptx integration.

Requirements:
    pip install python-pptx latex2mathml lxml
    MML2OMML.XSL  (from a local Microsoft Office installation)

Usage in md2pptx:
    Specify the XSL path in the Markdown metadata header:
        mathxsl: /path/to/MML2OMML.XSL
    Then use inline math in bullet text:
        * The formula $E = mc^2$ is well known.
"""

from __future__ import annotations

from pathlib import Path
from lxml import etree

# ── Namespace URIs ────────────────────────────────────────────────────────────

A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
M   = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A   = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _t(ns: str, local: str) -> str:
    return f"{{{ns}}}{local}"


# ── LaTeX → OMML conversion ───────────────────────────────────────────────────

def _latex_to_omml(latex: str, transform: etree.XSLT) -> etree._Element:
    """Convert LaTeX to an OMML element via MathML and MML2OMML.XSL."""
    import latex2mathml.converter
    mathml_str = latex2mathml.converter.convert(latex)
    mathml     = etree.fromstring(mathml_str.encode())
    result     = transform(mathml)
    root       = result.getroot()
    if root is None:
        raise ValueError(f"XSLT produced no output for: {latex!r}")
    return root


def _extract_oMath(elem: etree._Element) -> etree._Element:
    """
    Return an <m:oMath> element.
    - If the root is already <m:oMath> (typical MML2OMML.XSL output), return it directly.
    - If the root is <m:oMathPara>, return its <m:oMath> child.
    """
    if elem.tag == _t(M, "oMath"):
        return elem
    oMath = elem.find(_t(M, "oMath"))
    if oMath is None:
        raise ValueError(f"No <m:oMath> found in: {etree.tostring(elem)}")
    return oMath


# ── Inline math injection ─────────────────────────────────────────────────────

def inject_inline_math(p_elem: etree._Element, omath_elem: etree._Element) -> None:
    """
    Append <a14:m><m:oMath>...</m:oMath></a14:m> to a paragraph's lxml element.

    Inserted immediately before <a:endParaRPr> so the order matches what
    python-pptx's add_run() produces.

    Parameters
    ----------
    p_elem    : the paragraph's _p attribute (lxml element)
    omath_elem: an <m:oMath> element (inline math, not wrapped in oMathPara)
    """
    a14m = etree.Element(_t(A14, "m"), nsmap={"a14": A14})
    a14m.append(omath_elem)
    endParaRPr = p_elem.find(_t(A, "endParaRPr"))
    if endParaRPr is not None:
        endParaRPr.addprevious(a14m)
    else:
        p_elem.append(a14m)


# ── Public API ────────────────────────────────────────────────────────────────

class MathInserter:
    """
    Convert LaTeX inline math and inject it into a python-pptx paragraph.

    Parameters
    ----------
    xsl_path : path to MML2OMML.XSL (required)
    """

    def __init__(self, xsl_path: str | Path):
        xsl_path = Path(xsl_path)
        if not xsl_path.exists():
            raise FileNotFoundError(f"MML2OMML.XSL not found: {xsl_path}")
        self._transform = etree.XSLT(etree.parse(str(xsl_path)))

    def make_inline_omml(self, latex: str) -> etree._Element:
        """
        Convert a LaTeX expression to an inline <m:oMath> element.
        Pass the result to inject_inline_math() to embed it in a paragraph.
        """
        omml = _latex_to_omml(latex, self._transform)
        return _extract_oMath(omml)
