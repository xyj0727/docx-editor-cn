"""Insert a 三线表 (three-line table) block into an unpacked DOCX document.

Generates correct OOXML for a Chinese academic three-line table and optionally
inserts it after a paragraph that contains a given anchor string.

Layout (《岩土工程学报》 CJGE):
  - Top and bottom rules 0.75 pt (w:sz="6"), rule below the header row 0.5 pt (w:sz="4")
  - No vertical rules, no fill; header text not bold
  - Table text 宋体 / Times New Roman 小五 (9 pt), centred, fixed 14 pt line spacing
  - Caption ABOVE the table: "表1 题名" 黑体小五 + optional "Table 1 Title" TNR 小五, centred
  - Default width = one column (4842 DXA = 8.54 cm); full width (双栏表) = 10104 DXA

Usage:
    # Insert after paragraph containing anchor text:
    python scripts/table.py unpacked/ "1" "符号说明" --caption-en "Table 1 Symbols" \\
        --headers "符号,说明" \\
        --rows '[["S","状态空间"],["A","动作空间"]]' \\
        --anchor "以下是符号说明"

    # Print XML snippet only (no file modification):
    python scripts/table.py --caption "表1 示例" \\
        --headers "列1,列2,列3" \\
        --rows '[["a","b","c"],["d","e","f"]]'

    # Custom column widths (DXA; 4842 = one column, 10104 = full width):
    python scripts/table.py unpacked/ "2" "参数说明" \\
        --headers "参数,取值,说明" \\
        --rows '[["lr","0.001","学习率"]]' \\
        --widths "1200,1200,2442"
"""

import argparse
import json
import re
import sys
from pathlib import Path

# ── Layout constants (CJGE: A4, margins 1.65 / 1.53 cm, two columns 8.54 cm) ──
CONTENT_WIDTH = 4842   # one column; 10104 DXA for a full-width table
LATIN, SONG, HEI = 'Times New Roman', 'SimSun', 'SimHei'
SIZE = 18              # 小五, half-points
LINE = 280             # fixed 14 pt

WORD_NS   = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

_THICK_BORDER = 'w:val="single" w:sz="6" w:space="0" w:color="000000"'   # 0.75 pt
_THIN_BORDER  = 'w:val="single" w:sz="4" w:space="0" w:color="000000"'   # 0.5 pt
_NO_BORDER    = 'w:val="none"   w:sz="0"  w:space="0" w:color="auto"'


# ── XML builders ─────────────────────────────────────────────────────────────

def _cell_borders(top: str, bottom: str) -> str:
    return (
        f'<w:tcBorders>'
        f'<w:top {top}/>'
        f'<w:left {_NO_BORDER}/>'
        f'<w:bottom {bottom}/>'
        f'<w:right {_NO_BORDER}/>'
        f'</w:tcBorders>'
    )


def _rpr(east: str = SONG, bold: bool = False) -> str:
    """Run properties in schema order: rFonts, b, bCs, sz, szCs."""
    b = '<w:b/><w:bCs/>' if bold else '<w:b w:val="0"/><w:bCs w:val="0"/>'
    return (f'<w:rFonts w:ascii="{LATIN}" w:hAnsi="{LATIN}" w:cs="{LATIN}" w:eastAsia="{east}"/>'
            f'{b}<w:sz w:val="{SIZE}"/><w:szCs w:val="{SIZE}"/>')


def _cell(text: str, width: int, top: str, bottom: str, bold: bool = False,
          align: str = 'center') -> str:
    bold_open  = _rpr(bold=bold)
    return (
        f'<w:tc>'
        f'  <w:tcPr>'
        f'    <w:tcW w:w="{width}" w:type="dxa"/>'
        f'    {_cell_borders(top, bottom)}'
        f'    <w:vAlign w:val="center"/>'
        f'  </w:tcPr>'
        f'  <w:p>'
        f'    <w:pPr>'
        f'      <w:spacing w:before="0" w:after="0" w:line="{LINE}" w:lineRule="exact"/>'
        f'      <w:ind w:firstLine="0" w:firstLineChars="0"/>'
        f'      <w:jc w:val="{align}"/>'
        f'    </w:pPr>'
        f'    <w:r>'
        f'      <w:rPr>{bold_open}</w:rPr>'
        f'      <w:t xml:space="preserve">{_escape(text)}</w:t>'
        f'    </w:r>'
        f'  </w:p>'
        f'</w:tc>'
    )


def _table_border_xml() -> str:
    """Table-level rules: top and bottom lines only (the header rule is set on the header cells)."""
    nb = f'<w:top {_THICK_BORDER}/><w:left {_NO_BORDER}/><w:bottom {_THICK_BORDER}/><w:right {_NO_BORDER}/><w:insideH {_NO_BORDER}/><w:insideV {_NO_BORDER}/>'
    return f'<w:tblBorders>{nb}</w:tblBorders>'


def _escape(text: str) -> str:
    return (text
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))


def _caption_paragraph(text: str, east: str, before: int) -> str:
    return (
        f'<w:p {WORD_NS}>'
        f'  <w:pPr>'
        f'    <w:spacing w:before="{before}" w:after="0" w:line="{LINE}" w:lineRule="exact"/>'
        f'    <w:ind w:firstLine="0" w:firstLineChars="0"/>'
        f'    <w:jc w:val="center"/>'
        f'  </w:pPr>'
        f'  <w:r>'
        f'    <w:rPr>{_rpr(east)}</w:rPr>'
        f'    <w:t xml:space="preserve">{_escape(text)}</w:t>'
        f'  </w:r>'
        f'</w:p>'
    )


def build_caption_xml(caption: str, caption_en: str = '') -> str:
    """Table caption above the table: 黑体小五 Chinese line + optional TNR 小五 English line."""
    xml = _caption_paragraph(caption, HEI, 60)
    if caption_en:
        xml += _caption_paragraph(caption_en, LATIN, 0)
    return xml


def build_table_xml(headers: list[str], rows: list[list[str]],
                    col_widths: list[int] | None = None) -> str:
    """Return XML for a 三线表.

    Parameters
    ----------
    headers    : column header strings
    rows       : list of data rows (each a list of cell strings)
    col_widths : DXA widths (4842 = one column, 10104 = full width); auto-equal one column if None
    """
    n = len(headers)
    if col_widths is None:
        w = CONTENT_WIDTH // n
        col_widths = [w] * n
        col_widths[-1] = CONTENT_WIDTH - w * (n - 1)

    if len(col_widths) != n:
        raise ValueError(f'col_widths has {len(col_widths)} entries but headers has {n}')
    total = sum(col_widths)
    if total > 10104:
        print(f'[table] Warning: col_widths sum {total} exceeds the full text width 10104 DXA',
              file=sys.stderr)

    # Grid
    grid_cols = ''.join(f'<w:gridCol w:w="{w}"/>' for w in col_widths)

    # Header row: thick top, thin bottom
    header_cells = ''.join(
        _cell(h, col_widths[i], _THICK_BORDER, _THIN_BORDER, bold=False)
        for i, h in enumerate(headers)
    )
    header_row = (
        f'<w:tr>'
        f'  <w:trPr><w:tblHeader/></w:trPr>'
        f'  {header_cells}'
        f'</w:tr>'
    )

    # Body rows
    body_rows = []
    for ri, row in enumerate(rows):
        is_last = ri == len(rows) - 1
        bottom  = _THICK_BORDER if is_last else _NO_BORDER
        cells   = ''.join(
            _cell(str(row[i]) if i < len(row) else '', col_widths[i], _NO_BORDER, bottom)
            for i in range(n)
        )
        body_rows.append(f'<w:tr>{cells}</w:tr>')

    all_rows = header_row + ''.join(body_rows)

    return (
        f'<w:tbl {WORD_NS}>'
        f'  <w:tblPr>'
        f'    <w:tblW w:w="{total}" w:type="dxa"/>'
        f'    <w:jc w:val="center"/>'
        f'    {_table_border_xml()}'
        f'    <w:tblCellMar>'
        f'      <w:top w:w="0" w:type="dxa"/>'
        f'      <w:left w:w="57" w:type="dxa"/>'
        f'      <w:bottom w:w="0" w:type="dxa"/>'
        f'      <w:right w:w="57" w:type="dxa"/>'
        f'    </w:tblCellMar>'
        f'  </w:tblPr>'
        f'  <w:tblGrid>{grid_cols}</w:tblGrid>'
        f'  {all_rows}'
        f'</w:tbl>'
    )


# ── document.xml insertion ────────────────────────────────────────────────────

def insert_after_anchor(doc_xml_path: Path, anchor: str,
                        caption_xml: str, table_xml: str) -> bool:
    """Insert caption + table after the first paragraph whose plain text
    contains *anchor*. Returns True on success."""
    content = doc_xml_path.read_text(encoding='utf-8')
    para_re = re.compile(r'(<w:p\b[^>]*>(?:(?!<w:p\b).)*?</w:p>)', re.DOTALL)

    match_end = None
    for m in para_re.finditer(content):
        plain = re.sub(r'<[^>]+>', '', m.group(0))
        if anchor in plain:
            match_end = m.end()
            break

    if match_end is None:
        return False

    insert = '\n' + caption_xml + '\n' + table_xml
    new_content = content[:match_end] + insert + content[match_end:]
    doc_xml_path.write_text(new_content, encoding='utf-8')
    return True


# ── CLI ───────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description='Insert a 三线表 into an unpacked DOCX document.'
    )
    # Positional (insertion mode)
    parser.add_argument('unpacked_dir', nargs='?',
                        help='Unpacked DOCX directory (omit for snippet output)')
    parser.add_argument('table_id', nargs='?',
                        help='Table number for the caption, e.g. "1" → "表1 <title>"')
    parser.add_argument('title', nargs='?',
                        help='Caption title text, e.g. "符号说明"')

    # Named
    parser.add_argument('--caption',
                        help='Full caption string (overrides table_id + title combination)')
    parser.add_argument('--caption-en', default='',
                        help='English caption line, e.g. "Table 1 Parameters of specimens"')
    parser.add_argument('--headers', required=True,
                        help='Comma-separated column headers, e.g. "符号,说明,备注"')
    parser.add_argument('--rows', required=True,
                        help='JSON array of rows, e.g. \'[["S","状态空间"],["A","动作空间"]]\'')
    parser.add_argument('--widths',
                        help='Comma-separated DXA column widths (4842 = one column, 10104 = full width)')
    parser.add_argument('--anchor',
                        help='Insert after the paragraph containing this text')
    args = parser.parse_args()

    # Resolve caption
    if args.caption:
        caption = args.caption
    elif args.table_id and args.title:
        caption = f'表{args.table_id} {args.title}'
    elif args.table_id:
        caption = f'表{args.table_id}'
    else:
        caption = '表'

    headers = [h.strip() for h in args.headers.split(',')]

    try:
        rows = json.loads(args.rows)
    except json.JSONDecodeError as e:
        print(f'Error: --rows is not valid JSON: {e}', file=sys.stderr)
        sys.exit(1)

    col_widths = None
    if args.widths:
        try:
            col_widths = [int(w.strip()) for w in args.widths.split(',')]
        except ValueError:
            print('Error: --widths must be comma-separated integers', file=sys.stderr)
            sys.exit(1)

    caption_xml = build_caption_xml(caption, args.caption_en)
    table_xml   = build_table_xml(headers, rows, col_widths)

    # Snippet-only mode
    if not args.unpacked_dir:
        print(caption_xml)
        print(table_xml)
        return

    unpacked = Path(args.unpacked_dir)
    doc_xml  = unpacked / 'word' / 'document.xml'
    if not doc_xml.exists():
        print(f'Error: {doc_xml} not found.', file=sys.stderr)
        sys.exit(1)

    if not args.anchor:
        print(caption_xml)
        print(table_xml)
        print('\n[table] No --anchor given; snippet printed above.\n'
              '        Paste into word/document.xml at the desired location.',
              file=sys.stderr)
        return

    print(f'[table] Inserting "{caption}" after paragraph containing: "{args.anchor}"',
          file=sys.stderr)
    ok = insert_after_anchor(doc_xml, args.anchor, caption_xml, table_xml)
    if ok:
        print(f'[table] ✓ Inserted table into {doc_xml}', file=sys.stderr)
    else:
        print(f'[table] ✗ Anchor text not found: "{args.anchor}"\n'
              f'         Snippet printed below — paste manually.', file=sys.stderr)
        print(caption_xml)
        print(table_xml)
        sys.exit(1)


if __name__ == '__main__':
    main()
