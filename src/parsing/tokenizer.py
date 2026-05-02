import re
from typing import List

from constants import STRUCTURAL_HEADINGS
from parsing.tokens import (
    Token, StructuralHeadingToken, SectionHeadingToken,
    AppendixHeadingToken, AppendixTitleToken, ParagraphToken,
    ListItemToken, TableToken, FigureToken, DiagramToken, FormulaToken,
    FormulaWhereToken, PageBreakToken, TocToken, EmptyLineToken,
    CodeBlockToken,
)

_SECTION_NUM_RE   = re.compile(r'^(\d+(?:\.\d+)*)\s+(.+)$')
_APPENDIX_LETTER  = re.compile(r'^ПРИЛОЖЕНИЕ\s+([А-ЯA-Z])$')
_LIST_DASH_RE     = re.compile(r'^[-–—]\s+(.+)$')
_LIST_LETTER_RE   = re.compile(r'^([а-яёa-z])\)\s+(.+)$')
_LIST_DIGIT_RE    = re.compile(r'^(\d+)\)\s+(.+)$')
_TABLE_ROW_RE     = re.compile(r'^\|(.+)\|$')
_TABLE_SEP_RE     = re.compile(r'^\|[-\s|:]+\|$')
_FIGURE_RE        = re.compile(r'^!\[([^\]]*)\]\(([^)]+)\)$')
_FORMULA_LINE_RE  = re.compile(r'^\$\$(.+)\$\$\s*(?:\((\S+)\))?$')
_TABLE_CAPTION_RE = re.compile(r'^<!--\s*(Таблица\s+\S+\s*–.+?)\s*-->$')
_NAMED_TABLE_CAP  = re.compile(r'^(Таблица\s+[\w.]+\s*–\s*.+)$')
_PAGEBREAK_RE     = re.compile(r'^---$|^<!-- pagebreak -->$', re.IGNORECASE)
_MERMAID_OPEN_RE   = re.compile(r'^```mermaid\s*(.*)', re.IGNORECASE)
_CODE_BLOCK_OPEN_RE = re.compile(r'^```(\w*)\s*$')

def tokenize(lines: List[str]) -> List[Token]:
    tokens: List[Token] = []
    i = 0
    pending_table_caption = None

    while i < len(lines):
        stripped = lines[i].rstrip()

        if not stripped:
            tokens.append(EmptyLineToken())
            i += 1
            continue

        if _PAGEBREAK_RE.match(stripped):
            tokens.append(PageBreakToken())
            i += 1
            continue

        if stripped.startswith('#'):
            i = _handle_heading(stripped, tokens, i)
            continue

        if stripped.startswith('> '):
            tokens.append(AppendixTitleToken(stripped[2:]))
            i += 1
            continue

        m = _TABLE_CAPTION_RE.match(stripped)
        if m:
            pending_table_caption = m.group(1).strip()
            i += 1
            continue

        m = _NAMED_TABLE_CAP.match(stripped)
        if m and not stripped.startswith('|'):
            pending_table_caption = stripped
            i += 1
            continue

        if _TABLE_ROW_RE.match(stripped):
            i, tok = _handle_table(lines, i, pending_table_caption)
            tokens.append(tok)
            pending_table_caption = None
            continue

        m = _MERMAID_OPEN_RE.match(stripped)
        if m:
            i, tok = _handle_mermaid_block(lines, i, m.group(1).strip())
            tokens.append(tok)
            continue

        m = _CODE_BLOCK_OPEN_RE.match(stripped)
        if m:
            i, tok = _handle_code_block(lines, i, m.group(1).strip())
            tokens.append(tok)
            continue

        m = _FIGURE_RE.match(stripped)
        if m:
            tokens.append(FigureToken(src=m.group(2), alt=m.group(1)))
            i += 1
            continue

        m = _FORMULA_LINE_RE.match(stripped)
        if m:
            tokens.append(FormulaToken(formula=m.group(1), number=m.group(2)))
            i += 1
            continue

        if stripped == '$$':
            i, tok = _handle_multiline_formula(lines, i)
            tokens.append(tok)
            continue

        if stripped.lower().startswith('где ') or stripped.lower() == 'где':
            i, tok = _handle_where_block(lines, i, stripped)
            tokens.append(tok)
            continue

        m = _LIST_DASH_RE.match(stripped)
        if m:
            tokens.append(ListItemToken('dash', '–', m.group(1)))
            i += 1
            continue
        m = _LIST_LETTER_RE.match(stripped)
        if m:
            tokens.append(ListItemToken('letter', m.group(1) + ')', m.group(2)))
            i += 1
            continue
        m = _LIST_DIGIT_RE.match(stripped)
        if m:
            tokens.append(ListItemToken('digit', m.group(1) + ')', m.group(2), indent_level=1))
            i += 1
            continue

        i, tok = _handle_paragraph(lines, i, stripped)
        tokens.append(tok)

    return tokens

def _handle_heading(stripped: str, tokens: list, i: int) -> int:
    m = re.match(r'^(#{1,4})\s+(.+)$', stripped)
    if not m:
        return i + 1
    level   = len(m.group(1))
    content = m.group(2).strip()
    upper   = content.upper()

    if level == 1:
        if upper == 'СОДЕРЖАНИЕ':
            tokens.append(StructuralHeadingToken('СОДЕРЖАНИЕ'))
            tokens.append(TocToken())
        elif upper in STRUCTURAL_HEADINGS:
            tokens.append(StructuralHeadingToken(upper))
        else:
            am = _APPENDIX_LETTER.match(upper)
            if am:
                tokens.append(AppendixHeadingToken(am.group(1)))
            else:
                nm = _SECTION_NUM_RE.match(content)
                if nm:
                    tokens.append(SectionHeadingToken(nm.group(1), nm.group(2), 1))
                else:
                    tokens.append(SectionHeadingToken('', content, 1))
    else:
        nm = _SECTION_NUM_RE.match(content)
        if nm:
            tokens.append(SectionHeadingToken(nm.group(1), nm.group(2), level))
        else:
            tokens.append(SectionHeadingToken('', content, level))
    return i + 1


def _handle_table(lines: list, i: int, caption):
    stripped = lines[i].rstrip()
    header   = [c.strip() for c in stripped.strip('|').split('|')]
    i += 1
    if i < len(lines) and _TABLE_SEP_RE.match(lines[i].rstrip()):
        i += 1
    rows = []
    while i < len(lines) and _TABLE_ROW_RE.match(lines[i].rstrip()):
        rows.append([c.strip() for c in lines[i].rstrip().strip('|').split('|')])
        i += 1
    return i, TableToken(header=header, rows=rows, caption=caption)


def _handle_multiline_formula(lines: list, i: int):
    formula_lines = []
    num = None
    i  += 1
    while i < len(lines) and lines[i].rstrip() != '$$':
        nl    = lines[i].rstrip()
        num_m = re.search(r'\((\S+)\)\s*$', nl)
        if num_m:
            num = num_m.group(1)
            nl  = nl[:num_m.start()].rstrip()
        formula_lines.append(nl)
        i += 1
    if i < len(lines):
        i += 1  # skip closing $$
    return i, FormulaToken(formula='\n'.join(formula_lines), number=num)


def _handle_where_block(lines: list, i: int, stripped: str):
    items = []
    rest  = stripped[4:].strip() if len(stripped) > 4 else ''
    if rest:
        for part in rest.split(';'):
            part = part.strip().rstrip(';')
            if part:
                items.append(part)
    i += 1
    while i < len(lines):
        nl = lines[i].rstrip()
        if not nl:
            break
        if nl.startswith('-') or nl.startswith('–') or re.match(r'^\$', nl):
            cleaned = re.sub(r'^[-–—]\s*', '', nl.strip()).rstrip(';')
            items.append(cleaned)
            i += 1
        else:
            break
    return i, FormulaWhereToken(items=items)


def _handle_code_block(lines: list, i: int, language: str):
    code_lines = []
    i += 1
    while i < len(lines):
        nl = lines[i].rstrip()
        if nl.strip() == '```':
            i += 1
            break
        code_lines.append(nl)
        i += 1
    return i, CodeBlockToken(code='\n'.join(code_lines), language=language)


def _handle_mermaid_block(lines: list, i: int, alt: str):
    code_lines = []
    i += 1
    while i < len(lines):
        nl = lines[i].rstrip()
        if nl.strip() == '```':
            i += 1
            break
        code_lines.append(nl)
        i += 1
    return i, DiagramToken(code='\n'.join(code_lines), alt=alt)


def _handle_paragraph(lines: list, i: int, first_line: str):
    para_lines = [first_line]
    i += 1
    while i < len(lines):
        nl = lines[i].rstrip()
        if not nl:
            break
        if (nl.startswith('#') or nl.startswith('|') or nl.startswith('![')
                or nl.startswith('$$') or nl.startswith('> ')
                or nl.startswith('```')
                or _MERMAID_OPEN_RE.match(nl)
                or _LIST_DASH_RE.match(nl) or _LIST_LETTER_RE.match(nl)
                or _LIST_DIGIT_RE.match(nl)
                or _TABLE_CAPTION_RE.match(nl)
                or _NAMED_TABLE_CAP.match(nl)
                or _PAGEBREAK_RE.match(nl)):
            break
        para_lines.append(nl)
        i += 1
    return i, ParagraphToken(' '.join(para_lines))
