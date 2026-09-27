#!/usr/bin/env python3
"""Byte-preservation oracle for GenOffice .docx edits.

Compares two .docx packages: every zip part must be byte-identical except
word/document.xml, and inside document.xml every <w:p> must be byte-identical
except the paragraphs the edit was meant to touch. Prints a JSON verdict.

usage: ooxml_oracle.py before.docx after.docx [--expect-changed 1] [--json]
"""
import json, re, sys, zipfile

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
P_RE = re.compile(r'<w:p[ >].*?</w:p>|<w:p/>', re.S)

def parts(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}

def paragraphs(xml: bytes):
    return P_RE.findall(xml.decode('utf-8'))

def count(xml: str, tag: str):
    return len(re.findall(rf'<{tag}\b', xml))

def main(argv):
    before, after = argv[1], argv[2]
    expect = int(argv[argv.index('--expect-changed') + 1]) if '--expect-changed' in argv else 1
    a, b = parts(before), parts(after)
    part_diffs = sorted(n for n in set(a) | set(b) if not n.endswith('/') and a.get(n) != b.get(n))
    pa, pb = paragraphs(a['word/document.xml']), paragraphs(b['word/document.xml'])
    changed = []
    if len(pa) == len(pb):
        for i, (x, y) in enumerate(zip(pa, pb)):
            if x != y:
                changed.append({'index': i, 'before_bytes': len(x), 'after_bytes': len(y),
                                'text_before': re.sub(r'<[^>]+>', '', x)[:80]})
    doc_a, doc_b = a['word/document.xml'].decode(), b['word/document.xml'].decode()
    counts = {t: [count(doc_a, t), count(doc_b, t)]
              for t in ('w:ins', 'w:del', 'w:proofErr', 'w:bookmarkStart', 'w:r')}
    counts['w16du:dateUtc'] = [doc_a.count('w16du:dateUtc'), doc_b.count('w16du:dateUtc')]
    ids = re.findall(r'<w:(?:ins|del)\b[^>]*w:id="(\d+)"', doc_b)
    verdict = {
        'before': before, 'after': after,
        'paragraphs': len(pa), 'paragraph_count_changed': len(pa) != len(pb),
        'changed_paragraphs': changed, 'changed_paragraph_count': len(changed),
        'expected_changed': expect,
        'other_parts_changed': [p for p in part_diffs if p != 'word/document.xml'],
        'element_counts_before_after': counts,
        'duplicate_revision_ids': sorted({i for i in ids if ids.count(i) > 1}),
    }
    ok = (not verdict['paragraph_count_changed'] and len(changed) <= expect
          and not verdict['other_parts_changed'] and not verdict['duplicate_revision_ids']
          and all(x == y for t, (x, y) in counts.items() if t not in ('w:r',)))
    verdict['byte_confined'] = ok
    print(json.dumps(verdict, indent=2))
    return 0 if ok else 1

if __name__ == '__main__':
    sys.exit(main(sys.argv))
