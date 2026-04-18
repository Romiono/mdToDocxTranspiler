import re

try:
    import latex2mathml.converter as _l2m
    import mathml2omml as _m2o
    from lxml import etree as _etree
    HAS_OMML = True
except ImportError:
    HAS_OMML = False

_M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

_LATEX_PREPROCESS = [
    (re.compile(r'\\displaystyle\s*'), ''),
    (re.compile(r'\\textstyle\s*'),    ''),
    (re.compile(r'\\scriptstyle\s*'),  ''),
    (re.compile(r'\\bigl\s*'),         ''),
    (re.compile(r'\\bigr\s*'),         ''),
    (re.compile(r'\\Bigl\s*'),         ''),
    (re.compile(r'\\Bigr\s*'),         ''),
    (re.compile(r'\\left\s*'),         ''),
    (re.compile(r'\\right\s*'),        ''),
]

# Bug in mathml2omml: <m:groupChrPr> is sometimes closed with </m:groupChr>
_GROUPCHR_BUG = re.compile(
    r'(<m:groupChrPr>(?:(?!</m:groupChrPr>).)*?)</m:groupChr>',
    re.DOTALL,
)


def _preprocess(s: str) -> str:
    for pat, repl in _LATEX_PREPROCESS:
        s = pat.sub(repl, s)
    return s.strip()


def latex_to_omml_elem(latex_str: str):
    if not HAS_OMML:
        return None
    try:
        mathml  = _l2m.convert(_preprocess(latex_str))
        omml    = _m2o.convert(mathml)
        omml    = _GROUPCHR_BUG.sub(r'\1</m:groupChrPr>', omml)
        omml_ns = omml.replace('<m:oMath>', f'<m:oMath xmlns:m="{_M_NS}">')
        return _etree.fromstring(omml_ns.encode('utf-8'))
    except Exception:
        return None
