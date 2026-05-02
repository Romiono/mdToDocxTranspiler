from dataclasses import dataclass, field
from typing import List, Optional


class Token:
    pass


@dataclass
class StructuralHeadingToken(Token):
    text: str


@dataclass
class SectionHeadingToken(Token):
    number: str
    text: str
    level: int


@dataclass
class AppendixHeadingToken(Token):
    letter: str


@dataclass
class AppendixTitleToken(Token):
    text: str


@dataclass
class ParagraphToken(Token):
    text: str


@dataclass
class ListItemToken(Token):
    marker_type: str   # 'dash' | 'letter' | 'digit'
    marker: str
    text: str
    indent_level: int = 0


@dataclass
class TableToken(Token):
    header: Optional[List[str]]
    rows: List[List[str]]
    caption: Optional[str] = None


@dataclass
class FigureToken(Token):
    src: str
    alt: str


@dataclass
class DiagramToken(Token):
    code: str
    alt: str


@dataclass
class FormulaToken(Token):
    formula: str
    number: Optional[str] = None


@dataclass
class FormulaWhereToken(Token):
    items: List[str] = field(default_factory=list)


class PageBreakToken(Token):
    pass


class TocToken(Token):
    pass


class EmptyLineToken(Token):
    pass


@dataclass
class CodeBlockToken(Token):
    code: str
    language: str = ''
