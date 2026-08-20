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



def inject_block_math(p_elem: etree._Element, omath_elem: etree._Element) -> None:
    """Inject a centered block equation via <m:oMathPara centerGroup>."""
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


def inject_numbered_block_math(p_elem: etree._Element, omath_elem: etree._Element,
                                eq_number: str) -> None:
    """Inject a numbered block equation using DrawingML tab stops.

    The paragraph pPr must already contain a centred tab at shape_width//2
    and a right-aligned tab at shape_width (added by the caller).  Content
    layout: [tab→centre] [inline oMath] [tab→right] [(eq_number)]
    """
    def _append(elem: etree._Element) -> None:
        endParaRPr = p_elem.find(_t(A, "endParaRPr"))
        if endParaRPr is not None:
            endParaRPr.addprevious(elem)
        else:
            p_elem.append(elem)

    # Leading tab: moves insertion point to the centre tab stop
    r_tab1 = etree.Element(_t(A, "r"))
    etree.SubElement(r_tab1, _t(A, "rPr"))
    etree.SubElement(r_tab1, _t(A, "t")).text = "\t"
    _append(r_tab1)

    # Inline equation (no oMathPara wrapper so it stays on the same line)
    a14m = etree.Element(_t(A14, "m"), nsmap={"a14": A14})
    a14m.append(omath_elem)
    _append(a14m)

    # Trailing tab: moves insertion point to the right tab stop
    r_tab2 = etree.Element(_t(A, "r"))
    etree.SubElement(r_tab2, _t(A, "rPr"))
    etree.SubElement(r_tab2, _t(A, "t")).text = "\t"
    _append(r_tab2)

    # Equation number text
    r_num = etree.Element(_t(A, "r"))
    etree.SubElement(r_num, _t(A, "rPr"))
    etree.SubElement(r_num, _t(A, "t")).text = f"({eq_number})"
    _append(r_num)


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

def _resolve_mathxsl_path(
    configured_path: str | Path | None,
    fallback_path: str | Path | None = None,
) -> Path:
    """Return the configured stylesheet or the copy beside md2pptx.

    A user-supplied path takes precedence.  If it is absent or no ``mathxsl``
    option was supplied, look for ``mml2omml.xsl`` in the installation
    directory.
    """
    fallback = (
        Path(fallback_path).expanduser()
        if fallback_path is not None
        else Path(__file__).resolve().with_name("mml2omml.xsl")
    )

    configured = Path(configured_path).expanduser() if configured_path else None
    if configured is not None and configured.is_file():
        return configured
    if fallback.is_file():
        return fallback

    if configured is not None:
        raise FileNotFoundError(
            f"MML2OMML.XSL not found: {configured}; "
            f"installation-directory fallback not found: {fallback}"
        )
    raise FileNotFoundError(
        f"mathxsl option not set and installation-directory fallback "
        f"not found: {fallback}"
    )


class MathInserter:
    def __init__(self, xsl_path: str | Path):
        xsl_path = Path(xsl_path)
        if not xsl_path.exists():
            raise FileNotFoundError(f"MML2OMML.XSL not found: {xsl_path}")
        self._xsl_path = xsl_path
        try:
            self._transform = etree.XSLT(etree.parse(str(xsl_path)))
        except etree.XSLTParseError as e:
            raise ValueError(
                f"{xsl_path} is not a stylesheet libxslt can compile: {e}. "
                f"Copies of MML2OMML.XSL differ; another copy may compile."
            ) from e

    def make_inline_omml(self, latex: str) -> etree._Element:
        omml  = _latex_to_omml(latex, self._transform)
        try:
            omath = _extract_oMath(omml)
        except ValueError as e:
            raise ValueError(
                f"{self._xsl_path} produced no OMML. It has to convert MathML "
                f"into OMML, which is what MML2OMML.XSL does; a stylesheet "
                f"converting the other way matches nothing here and yields an "
                f"empty result. {e}"
            ) from e
        _add_cambria_math_rpr(omath)
        return omath


# ── PptxMath — block math handler (mirrors RunPython pattern) ─────────────────

class PptxMath:
    def __init__(self):
        pass

    def run(self, prs, slide, renderingRectangle, codeLines, codeType):
        try:
            mathxsl_path = _resolve_mathxsl_path(
                globals.processingOptions.getCurrentOption("mathxsl")
            )
            inserter = MathInserter(mathxsl_path)
        except Exception as e:
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
        pPr.set("marL", "0")
        pPr.set("indent", "0")
        etree.SubElement(pPr, _t(A, "buNone"))
        if eq_number:
            # Centre tab at midpoint, right tab at shape right edge
            w = renderingRectangle.width
            tabLst = etree.SubElement(pPr, _t(A, "tabLst"))
            tab_c = etree.SubElement(tabLst, _t(A, "tab"))
            tab_c.set("pos", str(w // 2))
            tab_c.set("algn", "ctr")
            tab_r = etree.SubElement(tabLst, _t(A, "tab"))
            tab_r.set("pos", str(w))
            tab_r.set("algn", "r")

        # Without this the equation inherits the text box default (18pt) and
        # comes out smaller than the surrounding body text.  Setting it on the
        # paragraph's defRPr covers the math runs, the tabs and the equation
        # number alike.  defRPr must follow tabLst in the DrawingML sequence.
        baseTextSize = globals.processingOptions.getCurrentOption("baseTextSize")
        if baseTextSize > 0:
            defRPr = etree.SubElement(pPr, _t(A, "defRPr"))
            defRPr.set("sz", str(int(round(baseTextSize * 100))))

        p._p.insert(0, pPr)

        try:
            omath = inserter.make_inline_omml(latex)
            if eq_number:
                inject_numbered_block_math(p._p, omath, eq_number)
            else:
                inject_block_math(p._p, omath)
        except Exception as e:
            # Name the formula: a deck can hold many math blocks and the
            # literal LaTeX is what the reader will see left in the slide.
            oneLine = " ".join(latex.split())
            if len(oneLine) > 60:
                oneLine = oneLine[:57] + "..."
            sys.stderr.write(f"Math conversion failed for '{oneLine}': {e}\n")
            p.text = latex

        # Must be called after injection: serialises and re-parses the sp
        # element to hoist xmlns:a14 / xmlns:mc to <p:sp> level.
        _hoist_math_namespaces(math_box)
