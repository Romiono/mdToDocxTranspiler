from parsing.tokens import (
    EmptyLineToken, PageBreakToken, TocToken,
    StructuralHeadingToken, SectionHeadingToken,
    AppendixHeadingToken, AppendixTitleToken,
    ParagraphToken, ListItemToken, TableToken,
    FigureToken, DiagramToken, FormulaToken, FormulaWhereToken,
    CodeBlockToken,
)
from writer.builder import GostDocxBuilder


def render(tokens: list, builder: GostDocxBuilder) -> None:
    i = 0
    while i < len(tokens):
        tok = tokens[i]

        if isinstance(tok, EmptyLineToken):
            i += 1; continue

        if isinstance(tok, PageBreakToken):
            builder._page_break()
            i += 1; continue

        if isinstance(tok, TocToken):
            builder.add_toc()
            i += 1; continue

        if isinstance(tok, StructuralHeadingToken):
            builder.add_structural_heading(tok.text, page_break_before=(i != 0))
            i += 1; continue

        if isinstance(tok, SectionHeadingToken):
            builder.add_section_heading(tok.number, tok.text, level=tok.level)
            i += 1; continue

        if isinstance(tok, AppendixHeadingToken):
            builder.add_appendix_heading(tok.letter)
            i += 1; continue

        if isinstance(tok, AppendixTitleToken):
            builder.add_appendix_title(tok.text)
            i += 1; continue

        if isinstance(tok, ParagraphToken):
            builder.add_paragraph(tok.text)
            i += 1; continue

        if isinstance(tok, ListItemToken):
            builder.add_list_item(tok.marker_type, tok.marker, tok.text, tok.indent_level)
            i += 1; continue

        if isinstance(tok, TableToken):
            builder.add_table(tok.header, tok.rows, caption=tok.caption)
            i += 1; continue

        if isinstance(tok, CodeBlockToken):
            builder.add_code_block(tok.code, tok.language)
            i += 1; continue

        if isinstance(tok, FigureToken):
            builder.add_figure(tok.src, tok.alt)
            i += 1; continue

        if isinstance(tok, DiagramToken):
            builder.add_diagram(tok.code, tok.alt)
            i += 1; continue

        if isinstance(tok, FormulaToken):
            builder.add_formula(tok.formula, explicit_number=tok.number)
            if i + 1 < len(tokens) and isinstance(tokens[i + 1], FormulaWhereToken):
                i += 1
                builder.add_formula_where(tokens[i].items)
            i += 1; continue

        if isinstance(tok, FormulaWhereToken):
            builder.add_formula_where(tok.items)
            i += 1; continue

        i += 1
