"""
symbols
"""
import re
import html


# Resolve symbols and unescape any numeric character references
def resolveSymbols(text):
    # h = html.parser.HTMLParser()

    textSplit = re.split("(&\#x?[0-9a-f]{2,6};)", text, flags=re.IGNORECASE)
    text2 = ""
    for t in textSplit:
        if t == "":
            text2 = text2 + t
        elif (t[0:2] == "&#") & (t[-1] == ";"):
            text2 = text2 + html.unescape(t)
        else:
            text2 = text2 + t

    # Replace certain entity references with actual characters
    replacementRules = [
        ("&equals;", "="),
        ("&lt;", u"\uFDE1"),
        ("&gt;", u"\uFDE2"),
        ("&le;", "≤"),
        ("&ge;", "≥"),
        ("&asymp;", "≈"),
        ("&Delta;", "Δ"),
        ("&delta;", "δ"),
        ("&sim;", "∼"),
        ("&nbsp;", chr(160)),
        ("&semi;", ";"),
        ("&colon;", ":"),
        ("&comma;", ","),
        ("&amp;", "&"),
        ("&larr;", "←"),
        ("&rarr;", "→"),
        ("&uarr;", "↑"),
        ("&darr;", "↓"),
        ("&harr;", "↔"),
        ("&varr;", "↕"),
        ("&nwarr;", "↖"),
        ("&nearr;", "↗"),
        ("&swarr;", "↙"),
        ("&searr;", "↘"),
        ("&lsqb;", r"\["),
        ("&rsqb;", r"\]"),
        ("&infin;", "∞"),
        ("&auml;", "ä"),
        ("&Auml;", "Ä"),
        ("&uuml;", "ü"),
        ("&Uuml;", "Ü"),
        ("&ouml;", "ö"),
        ("&Ouml;", "Ö"),
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
        ("&mu;", "μ"),
        ("&nu;", "ν"),
        ("&pi;", "π"),
        ("&rho;", "ρ"),
        ("&dash;", "-"),
        ("\`", u"\uFDE3"),
        ("&grave;", u"\uFDE3"),
        ("&quot;", "\""),
        ("&ldquo;", u"\u201C"),
        ("&rdquo;", u"\u201D"),
        ("&apos;", "'"),
        ("&lsquo;", u"\u2018"),
        ("&rsquo;", u"\u2019"),
        ("&Oslash;",u"\u00D8"),
        ("&oslash;",u"\u00F8"),
    ]
    
    for term, replacement in replacementRules:
        text2 = text2.replace(term, replacement)

    return text2
