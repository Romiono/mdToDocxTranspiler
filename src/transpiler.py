#!/usr/bin/env python3

import sys
from pathlib import Path

from parsing.tokenizer import tokenize
from writer.builder    import GostDocxBuilder
from renderer          import render


def transpile(md_path: str, docx_path: str) -> None:
    text = Path(md_path).read_text(encoding='utf-8')
    if text.startswith('---'):
        end = text.find('\n---', 3)
        if end != -1:
            text = text[end + 4:]

    base_dir = str(Path(md_path).resolve().parent)
    tokens   = tokenize(text.splitlines())
    builder  = GostDocxBuilder(base_dir=base_dir)
    render(tokens, builder)
    builder.save(docx_path)
    print(f'✓ Saved: {docx_path}')


def main() -> None:
    if len(sys.argv) < 2:
        print('Usage: python3 src/transpiler.py input.md [output.docx]')
        sys.exit(1)
    md_path   = sys.argv[1]
    docx_path = sys.argv[2] if len(sys.argv) >= 3 else str(Path(md_path).with_suffix('.docx'))
    transpile(md_path, docx_path)


if __name__ == '__main__':
    main()
