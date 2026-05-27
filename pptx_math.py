from __future__ import annotations

"""
pptx_math
"""

version = "0.1"

import sys
from pathlib import Path
from lxml import etree

import globals

# ── Namespace URIs ────────────────────────────────────────────────────────────

A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
M   = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
P   = "http://schemas.openxmlformats.org/presentationml/2006/main"
MC  = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def _t(ns: str, local: str) -> str:
    return f"{{{ns}}}{local}"


# ── LaTeX → OMML conversion ───────────────────────────────────────────────────

def _latex_to_omml(latex: str, transform: etree.XSLT) -> etree._Element:
    import latex2mathml.converter
    mathml_str = latex2mathml.converter.convert(latex)
    mathml     = etree.fromstring(mathml_str.encode())
    result     = transform(mathml)
    root       = result.getroot()
    if root is None:
        raise ValueError(f"XSLT produced no output for: {latex!r}")
    return root


def _extract_oMath(elem: etree._Element) -> etree._Element:
    if elem.tag == _t(M, "oMath"):
        return elem
    oMath = elem.find(_t(M, "oMath"))
    if oMath is None:
        raise ValueError(f"No <m:oMath> found in: {etree.tostring(elem)}")
    return oMath


def _add_cambria_math_rpr(omath_elem: etree._Element) -> None:
    """Add <a:rPr><a:latin typeface="Cambria Math"> to each bare <m:r>.

    Google Slides requires explicit Cambria Math declarations on math runs
    to recognise and render OMML equations. MML2OMML.XSL does not add them.
    """
    for mr in omath_elem.iter(_t(M, "r")):
        if mr.find(_t(A, "rPr")) is None:
            rPr   = etree.Element(_t(A, "rPr"))
            latin = etree.SubElement(rPr, _t(A, "latin"))
            latin.set("typeface",   "Cambria Math")
            latin.set("panose",     "02040503050406030204")
            latin.set("pitchFamily","18")
            latin.set("charset",    "0")
            mr.insert(0, rPr)


# ── Math injection ────────────────────────────────────────────────────────────

def inject_inline_math(p_elem: etree._Element, omath_elem: etree._Element) -> None:
    """Append <a14:m><m:oMath>...</m:oMath></a14:m> to a paragraph's lxml element."""
    a14m = etree.Element(_t(A14, "m"), nsmap={"a14": A14})
    a14m.append(omath_elem)
    endParaRPr = p_elem.find(_t(A, "endParaRPr"))
    if endParaRPr is not None:
        endParaRPr.addprevious(a14m)
    else:
        p_elem.append(a14m)


def _rPr(lang: str = "en-US", bold: bool = False, italic: bool = False) -> etree._Element:
    """Build <a:rPr> with Cambria Math font."""
    el = etree.Element(_t(A, "rPr"))
    el.set("lang", lang)
    el.set("b", "1" if bold else "0")
    el.set("i", "1" if italic else "0")
    el.set("smtClean", "0")
    latin = etree.SubElement(el, _t(A, "latin"))
    latin.set("typeface", "Cambria Math")
    latin.set("panose", "02040503050406030204")
    latin.set("pitchFamily", "18")
    latin.set("charset", "0")
    return el


def inject_block_math(p_elem: etree._Element, omath_elem: etree._Element,
                      eq_number: str | None = None) -> None:
    """Inject block math via <m:oMathPara> with optional equation number.

    When eq_number is given, wraps the equation in <m:eqArr> and appends a
    # separator run followed by a <m:d> delimiter — identical to the structure
    PowerPoint produces when the user types '#(n)' in its own equation editor.
    """
    if eq_number:
        eqArr = etree.Element(_t(M, "eqArr"))

        eqArrPr = etree.SubElement(eqArr, _t(M, "eqArrPr"))
        ctrlPr  = etree.SubElement(eqArrPr, _t(M, "ctrlPr"))
        ctrlPr.append(_rPr())

        e_elem = etree.SubElement(eqArr, _t(M, "e"))

        for child in list(omath_elem):
            e_elem.append(child)

        hash_r = etree.SubElement(e_elem, _t(M, "r"))
        hash_r.append(_rPr(italic=True))
        hash_t = etree.SubElement(hash_r, _t(M, "t"))
        hash_t.text = "#"

        d_elem  = etree.SubElement(e_elem, _t(M, "d"))
        dPr     = etree.SubElement(d_elem, _t(M, "dPr"))
        ctrlPr2 = etree.SubElement(dPr, _t(M, "ctrlPr"))
        ctrlPr2.append(_rPr())
        d_e     = etree.SubElement(d_elem, _t(M, "e"))
        num_r   = etree.SubElement(d_e, _t(M, "r"))
        num_r.append(_rPr())
        num_t   = etree.SubElement(num_r, _t(M, "t"))
        num_t.text = eq_number

        omath_elem.append(eqArr)

    oMathPara   = etree.Element(_t(M, "oMathPara"), nsmap={"m": M})
    oMathParaPr = etree.SubElement(oMathPara, _t(M, "oMathParaPr"))
    jc          = etree.SubElement(oMathParaPr, _t(M, "jc"))
    jc.set(_t(M, "val"), "centerGroup")
    oMathPara.append(omath_elem)

    a14m = etree.Element(_t(A14, "m"), nsmap={"a14": A14})
    a14m.append(oMathPara)

    endParaRPr = p_elem.find(_t(A, "endParaRPr"))
    if endParaRPr is not None:
        endParaRPr.addprevious(a14m)
    else:
        p_elem.append(a14m)


def _prep_math_shape(shape) -> None:
    """Remove txBox="1" and fix auto-resize on a math shape."""
    sp = shape._element
    tf = shape.text_frame

    cNvSpPr = sp.find(f'.//{{{P}}}cNvSpPr')
    if cNvSpPr is not None:
        cNvSpPr.attrib.pop('txBox', None)

    bodyPr = tf._txBody.bodyPr
    bodyPr.attrib.pop("wrap", None)
    for child in list(bodyPr):
        if child.tag in (_t(A, "spAutoFit"), _t(A, "normAutofit")):
            bodyPr.remove(child)
    etree.SubElement(bodyPr, _t(A, "noAutofit"))


def _hoist_math_namespaces(shape) -> None:
    """Hoist xmlns:a14 and xmlns:mc to the <p:sp> element.

    Google Slides' PPTX importer uses the presence of xmlns:a14 at the
    shape level to identify OMML-containing shapes.  lxml cannot add
    namespace declarations to existing elements, so we serialise, patch
    the opening tag, re-parse and replace in the slide tree.
    Must be called AFTER all math content has been injected.
    """
    sp = shape._element
    parent = sp.getparent()
    if parent is None or A14 in sp.nsmap.values():
        return

    xml = etree.tostring(sp, encoding="unicode")
    xml = xml.replace(
        "<p:sp ",
        f'<p:sp xmlns:a14="{A14}" xmlns:mc="{MC}" ',
        1,
    )
    new_sp = etree.fromstring(xml.encode())
    idx = list(parent).index(sp)
    parent.remove(sp)
    parent.insert(idx, new_sp)


# ── MathInserter ──────────────────────────────────────────────────────────────

class MathInserter:
    def __init__(self, xsl_path: str | Path):
        xsl_path = Path(xsl_path)
        if not xsl_path.exists():
            raise FileNotFoundError(f"MML2OMML.XSL not found: {xsl_path}")
        self._transform = etree.XSLT(etree.parse(str(xsl_path)))

    def make_inline_omml(self, latex: str) -> etree._Element:
        omml  = _latex_to_omml(latex, self._transform)
        omath = _extract_oMath(omml)
        _add_cambria_math_rpr(omath)
        return omath


# ── PptxMath — block math handler (mirrors RunPython pattern) ─────────────────

class PptxMath:
    def __init__(self):
        pass

    def run(self, prs, slide, renderingRectangle, codeLines, codeType):
        mathxsl_path = globals.processingOptions.getCurrentOption("mathxsl")
        if not mathxsl_path:
            sys.stderr.write("Math block skipped: mathxsl option not set.\n")
            return

        try:
            inserter = MathInserter(mathxsl_path)
        except (FileNotFoundError, Exception) as e:
            sys.stderr.write(f"Math block skipped: {e}\n")
            return

        # Extract optional equation number: "math 1.2.4" → "1.2.4"
        parts     = codeType.split(None, 1)
        eq_number = parts[1].strip() if len(parts) > 1 else None

        latex = "\n".join(codeLines)

        math_box = slide.shapes.add_textbox(
            renderingRectangle.left,
            renderingRectangle.top,
            renderingRectangle.width,
            renderingRectangle.height,
        )

        tf = math_box.text_frame
        _prep_math_shape(math_box)

        p = tf.paragraphs[0]

        pPr = etree.Element(_t(A, "pPr"))
        p._p.insert(0, pPr)

        try:
            omath = inserter.make_inline_omml(latex)
            inject_block_math(p._p, omath, eq_number)
        except Exception as e:
            sys.stderr.write(f"Math conversion failed: {e}\n")
            p.text = latex

        # Must be called after injection: serialises and re-parses the sp
        # element to hoist xmlns:a14 / xmlns:mc to <p:sp> level.
        _hoist_math_namespaces(math_box)
