#!/usr/bin/env python3
"""Build a synthetic .docx that contains the constructs that most often do not
round-trip through GenOffice's editor model: adjacent same-author <w:ins>/<w:del>
runs, a <w:proofErr> marker in an untouched heading, a bookmark that spans
paragraphs, and w16du:dateUtc on revisions. No third-party deps."""
import sys, zipfile

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W16DU = 'http://schemas.microsoft.com/office/word/2023/wordml/word16du'

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}">
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
</w:styles>"""

def run(text):
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'

def ins(i, author, date, text):
    return (f'<w:ins w:id="{i}" w:author="{author}" w:date="{date}" w16du:dateUtc="{date}">'
            f'{run(text)}</w:ins>')

def dele(i, author, date, text):
    return (f'<w:del w:id="{i}" w:author="{author}" w:date="{date}" w16du:dateUtc="{date}">'
            f'<w:r><w:delText xml:space="preserve">{text}</w:delText></w:r></w:del>')

D1 = '2026-09-01T10:00:00Z'
D2 = '2026-09-02T14:30:00Z'

paragraphs = [
    # 0: heading with a grammar proofErr marker (trigger 2) and a bookmark that closes 2 paragraphs later (trigger 3)
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
    '<w:bookmarkStart w:id="0" w:name="_Toc_motion"/>'
    '<w:proofErr w:type="gramStart"/>' + run('MOTION TO ') + run('DISMISS') + '<w:proofErr w:type="gramEnd"/>' + run(' — DRAFT') +
    '</w:p>',
    # 1: the only paragraph an edit will target
    '<w:p>' + run('Edit target paragraph: the word ACME appears here exactly once.') + '</w:p>',
    # 2: plain paragraph closing the bookmark
    '<w:p>' + run('Plaintiff respectfully moves this Court for an order dismissing the complaint.') + '<w:bookmarkEnd w:id="0"/></w:p>',
    # 3: adjacent same-author/same-date insertions (trigger 1) plus deletions
    '<w:p>' + run('The complaint ') + ins(101, 'Author A', D1, 'fails to ') + ins(102, 'Author A', D1, 'state a claim ') +
    dele(103, 'Author A', D1, 'is insufficient ') + run('under Rule 12(b)(6). ') +
    ins(104, 'Author B', D2, 'See Twombly, 550 U.S. 544 (2007).') + '</w:p>',
    # 4: more revisions, different author/day pairs
    '<w:p>' + ins(201, 'Author B', D2, 'Furthermore, ') + ins(202, 'Author B', D2, 'the allegations ') +
    run('are conclusory. ') + dele(203, 'Author A', D1, 'This is obvious. ') + dele(204, 'Author A', D1, 'Clearly. ') + '</w:p>',
] + [
    # 5..: filler untouched body paragraphs
    '<w:p>' + run(f'Untouched body paragraph {n}. Lorem ipsum dolor sit amet, consectetur adipiscing elit.') + '</w:p>'
    for n in range(5, 25)
]

DOCUMENT = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:document xmlns:w="{W}" xmlns:w16du="{W16DU}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w16du">'
            f'<w:body>{"".join(paragraphs)}<w:sectPr/></w:body></w:document>')

def main(out):
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', DOCUMENT)
        z.writestr('word/styles.xml', STYLES)
    print(out, len(paragraphs), 'paragraphs')

if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else 'tracked-motion.docx')
