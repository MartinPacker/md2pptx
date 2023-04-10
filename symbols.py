"""
symbols
"""
import re


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
    text2 = text2.replace("&equals;", "=")
    text2 = text2.replace("&lt;", chr(236))
    text2 = text2.replace("&gt;", chr(237))
    text2 = text2.replace("&le;", "≤")
    text2 = text2.replace("&ge;", "≥")
    text2 = text2.replace("&asymp;", "≈")
    text2 = text2.replace("&Delta;", "Δ")
    text2 = text2.replace("&delta;", "δ")
    text2 = text2.replace("&sim;", "∼")
    text2 = text2.replace("&nbsp;", chr(160))
    text2 = text2.replace("&semi;", ";")
    text2 = text2.replace("&colon;", ":")
    text2 = text2.replace("&comma;", ",")
    text2 = text2.replace("&amp;", "&")
    text2 = text2.replace("&larr;", "←")
    text2 = text2.replace("&rarr;", "→")
    text2 = text2.replace("&uarr;", "↑")
    text2 = text2.replace("&darr;", "↓")
    text2 = text2.replace("&harr;", "↔")
    text2 = text2.replace("&varr;", "↕")
    text2 = text2.replace("&nwarr;", "↖")
    text2 = text2.replace("&nearr;", "↗")
    text2 = text2.replace("&swarr;", "↙")
    text2 = text2.replace("&searr;", "↘")
    text2 = text2.replace("&lsqb;", "\[")
    text2 = text2.replace("&rsqb;", "\]")
    text2 = text2.replace("&infin;", "∞")
    text2 = text2.replace("&auml;", "ä")
    text2 = text2.replace("&Auml;", "Ä")
    text2 = text2.replace("&uuml;", "ü")
    text2 = text2.replace("&Uuml;", "Ü")
    text2 = text2.replace("&ouml;", "ö")
    text2 = text2.replace("&Ouml;", "Ö")
    text2 = text2.replace("&szlig;", "ß")
    text2 = text2.replace("&euro;", "€")
    text2 = text2.replace("&check;", "✓")
    text2 = text2.replace("&hellip;", "…")
    text2 = text2.replace("&times;", "×")
    text2 = text2.replace("&percnt;", "%")
    text2 = text2.replace("&divide;", "÷")
    text2 = text2.replace("&forall;", "∀")
    text2 = text2.replace("&exist;", "∃")
    text2 = text2.replace("&lambda;", "λ")
    text2 = text2.replace("&mu;", "μ")
    text2 = text2.replace("&nu;", "ν")
    text2 = text2.replace("&pi;", "π")
    text2 = text2.replace("&rho;", "ρ")
    text2 = text2.replace("&dash;", "-")

    return text2
