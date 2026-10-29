"""Make in-text citation marks superscript in an existing .docx (CJGE rule: "Xu 等[13]" -> "Xu 等^[13]").

Usage: python superscript_citations.py IN.docx OUT.docx

- Matches [n], [n,m,…] and [n-m] / [n–m] (also full-width comma) inside ordinary text runs of body,
  table, header/footer-free document.xml paragraphs.
- Paragraphs that are reference-list entries (text starting with "[n] " or "[n]" followed by an author)
  are skipped, as are runs inside fields, MathType objects and pictures.
- The rest of each run keeps its formatting; only the citation part gets w:vertAlign superscript.
- Prints how many citations were set. Check the result with OfficeCLI.
"""
import copy
import re
import shutil
import sys
import zipfile

from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
q = lambda t: '{%s}%s' % (W, t)
CITE_RE = re.compile(r'\[\d+(?:\s*[-–,，]\s*\d+)*\]')
REF_ENTRY_RE = re.compile(r'^\s*\[\d+\]\s*\S')
AFTER_VERTALIGN = [q(t) for t in ('rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath')]


def para_text(p):
    return ''.join(t.text or '' for t in p.iter(q('t')))


def run_text(r):
    return ''.join((e.text or '') if e.tag == q('t') else ('\t' if e.tag == q('tab') else '') for e in r)


def is_plain_text_run(r):
    if r.getparent() is None or r.getparent().tag != q('p'):
        return False
    if r.find('.//' + q('object')) is not None or r.find('.//' + q('drawing')) is not None:
        return False
    if r.find(q('fldChar')) is not None or r.find(q('instrText')) is not None:
        return False
    return r.find(q('t')) is not None and r.find(q('tab')) is None and r.find(q('br')) is None


def inside_field(p, r):
    """True if run r lies between a fldChar begin and end in paragraph p."""
    depth = 0
    for ch in p:
        if ch is r:
            return depth > 0
        if ch.tag == q('r'):
            fc = ch.find(q('fldChar'))
            if fc is not None:
                t = fc.get(q('fldCharType'))
                depth += 1 if t == 'begin' else (-1 if t == 'end' else 0)
    return False


def clone_run(r, text, superscript):
    n = etree.Element(q('r'))
    rpr = r.find(q('rPr'))
    rpr = copy.deepcopy(rpr) if rpr is not None else etree.Element(q('rPr'))
    for e in rpr.findall(q('vertAlign')):
        rpr.remove(e)
    if superscript:
        va = etree.Element(q('vertAlign'))
        va.set(q('val'), 'superscript')
        pos = len(rpr)
        for i, e in enumerate(rpr):
            if e.tag in AFTER_VERTALIGN:
                pos = i
                break
        rpr.insert(pos, va)
    elif r.find(q('rPr')) is not None and r.find(q('rPr')).find(q('vertAlign')) is not None:
        rpr.insert(len(rpr), copy.deepcopy(r.find(q('rPr')).find(q('vertAlign'))))
    if len(rpr):
        n.append(rpr)
    t = etree.SubElement(n, q('t'))
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return n


def superscript_citations(root):
    count = 0
    for p in root.iter(q('p')):
        if REF_ENTRY_RE.match(para_text(p)):
            continue
        for r in list(p.findall(q('r'))):
            if not is_plain_text_run(r) or inside_field(p, r):
                continue
            txt = run_text(r)
            if not CITE_RE.search(txt):
                continue
            pieces, last = [], 0
            for m in CITE_RE.finditer(txt):
                if m.start() > last:
                    pieces.append((txt[last:m.start()], False))
                pieces.append((m.group(0), True))
                last = m.end()
            if last < len(txt):
                pieces.append((txt[last:], False))
            idx = p.index(r)
            for k, (s, sup) in enumerate(pieces):
                p.insert(idx + k, clone_run(r, s, sup))
                count += sup
            p.remove(r)
    return count


def main(src, dst):
    if src != dst:
        shutil.copyfile(src, dst)
    zin = zipfile.ZipFile(src)
    root = etree.fromstring(zin.read('word/document.xml'))
    n = superscript_citations(root)
    data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    tmp = dst + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, data if item.filename == 'word/document.xml' else zin.read(item.filename))
    zin.close()
    shutil.move(tmp, dst)
    print('superscript citations:', n)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
