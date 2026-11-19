"""Build _build/proposal-reference.docx (the Word style template) from Pandoc's default.

Run from the P3-ProposalDraft folder:  python _build/build_reference.py
Needs _build/default-reference.docx, made with:
    quarto pandoc -o _build/default-reference.docx --print-default-data-file reference.docx

What it sets: Calibri body, navy headings, a centered cover page (Title/Subtitle/Author/Date),
the ToC on its own page, APA-style hanging indent for references, a ruled table style,
US Letter with Word's default 1 in margins, and a page-number footer that skips the cover.
"""
import re
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
SRC = HERE / "default-reference.docx"
DST = HERE / "proposal-reference.docx"

# Neutral palette: charcoal headings and grays, no blue. (Names kept so the styles below read the same.)
NAVY = "262626"   # heading 1, title, ToC heading, table rules
BLUE = "4D4D4D"   # heading 2, title rule
INK = "1A1A1A"    # body text
MUTED = "595959"  # subtitle, date, page numbers
RULE = "A6A6A6"   # rule under heading 1
TABLE_FILL = "F2F2F2"


def fonts(name):
    return f'<w:rFonts w:ascii="{name}" w:hAnsi="{name}" w:eastAsia="{name}" w:cs="{name}"/>'


FONT = fonts("Times New Roman")  # body text
HEAD = fonts("Arial")            # headings, title, ToC heading
BODY_PT = 12


def para_style(sid, name, based="Normal", nxt="BodyText", ppr="", rpr="", custom=False, extra=""):
    c = ' w:customStyle="1"' if custom else ""
    return (f'<w:style w:type="paragraph"{c} w:styleId="{sid}"><w:name w:val="{name}"/>'
            f'<w:basedOn w:val="{based}"/><w:next w:val="{nxt}"/>{extra}<w:qFormat/>'
            f'<w:pPr>{ppr}</w:pPr><w:rPr>{rpr}</w:rPr></w:style>')


def sz(pt):
    v = int(pt * 2)
    return f'<w:sz w:val="{v}"/><w:szCs w:val="{v}"/>'


STYLES = {
    "Normal": ('<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>'
               f'<w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr>'
               f'<w:rPr>{FONT}<w:color w:val="{INK}"/>{sz(BODY_PT)}</w:rPr></w:style>'),
    "BodyText": para_style("BodyText", "Body Text", ppr='<w:spacing w:before="0" w:after="160"/>',
                           extra='<w:link w:val="BodyTextChar"/>'),
    "Compact": para_style("Compact", "Compact", based="BodyText", ppr='<w:spacing w:before="0" w:after="0"/>',
                          custom=True),
    # ---- cover page ----
    "Title": para_style("Title", "Title", ppr='<w:spacing w:before="3600" w:after="240" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>'
                        f'<w:pBdr><w:bottom w:val="single" w:sz="12" w:space="12" w:color="{BLUE}"/></w:pBdr>',
                        rpr=f'{HEAD}<w:b/><w:color w:val="{NAVY}"/>{sz(30)}', extra='<w:link w:val="TitleChar"/>'),
    "Subtitle": para_style("Subtitle", "Subtitle", ppr='<w:spacing w:before="120" w:after="1200"/><w:jc w:val="center"/>',
                           rpr=f'{HEAD}<w:color w:val="{MUTED}"/>{sz(14)}', extra='<w:link w:val="SubtitleChar"/>'),
    "Author": para_style("Author", "Author", nxt="Author", ppr='<w:keepNext/><w:keepLines/><w:spacing w:before="0" w:after="40"/><w:jc w:val="center"/>',
                         rpr=f'{FONT}<w:color w:val="{INK}"/>{sz(12)}', custom=True),
    "Date": para_style("Date", "Date", ppr='<w:keepNext/><w:keepLines/><w:spacing w:before="600" w:after="0"/><w:jc w:val="center"/>',
                       rpr=f'{FONT}<w:color w:val="{MUTED}"/>{sz(12)}'),
    # ---- headings ----
    "Heading1": para_style("Heading1", "heading 1", ppr='<w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="120"/>'
                           f'<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="{RULE}"/></w:pBdr><w:outlineLvl w:val="0"/>',
                           rpr=f'{HEAD}<w:b/><w:color w:val="{NAVY}"/>{sz(16)}', extra='<w:link w:val="Heading1Char"/><w:uiPriority w:val="9"/>'),
    "Heading2": para_style("Heading2", "heading 2", ppr='<w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="1"/>',
                           rpr=f'{HEAD}<w:b/><w:color w:val="{BLUE}"/>{sz(13)}', extra='<w:link w:val="Heading2Char"/><w:uiPriority w:val="9"/>'),
    "TOCHeading": para_style("TOCHeading", "TOC Heading", based="Heading1",
                             ppr='<w:pageBreakBefore/><w:spacing w:before="0" w:after="240"/><w:outlineLvl w:val="9"/>',
                             rpr=f'{HEAD}<w:b/><w:color w:val="{NAVY}"/>{sz(16)}', extra='<w:uiPriority w:val="39"/>'),
    # ---- captions (APA puts figure/table labels above the item) ----
    "Caption": para_style("Caption", "caption", ppr='<w:keepNext/><w:spacing w:before="240" w:after="120"/>',
                          rpr=f'<w:color w:val="{INK}"/>{sz(10)}'),
    "ImageCaption": para_style("ImageCaption", "Image Caption", based="Caption", custom=True,
                               ppr='<w:keepNext/><w:spacing w:before="240" w:after="120"/>'),
    "TableCaption": para_style("TableCaption", "Table Caption", based="Caption", custom=True,
                               ppr='<w:keepNext/><w:spacing w:before="240" w:after="120"/>'),
    # ---- references: APA hanging indent ----
    "Bibliography": para_style("Bibliography", "Bibliography", nxt="Bibliography",
                               ppr='<w:spacing w:before="0" w:after="160"/><w:ind w:left="720" w:hanging="720"/>'),
}

TABLE_STYLE = (
    '<w:style w:type="table" w:default="1" w:styleId="Table"><w:name w:val="Table"/><w:basedOn w:val="TableNormal"/><w:qFormat/>'
    '<w:pPr><w:spacing w:before="60" w:after="60" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr>' + sz(10) + '</w:rPr>'
    '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
    f'<w:tblBorders><w:top w:val="single" w:sz="8" w:color="{NAVY}"/><w:bottom w:val="single" w:sz="8" w:color="{NAVY}"/>'
    '<w:insideH w:val="single" w:sz="4" w:color="D5DAE0"/></w:tblBorders>'
    '<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
    # Light fill + bold ink: paragraph-style text color outranks table-style color in Word,
    # so a dark fill with white text doesn't survive.
    '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr>'
    f'<w:tcPr><w:tcBorders><w:bottom w:val="single" w:sz="8" w:color="{NAVY}"/></w:tcBorders>'
    f'<w:shd w:val="clear" w:color="auto" w:fill="{TABLE_FILL}"/><w:vAlign w:val="bottom"/></w:tcPr></w:tblStylePr>'
    '</w:style>'
)

FOOTER = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0"/></w:pPr>'
    f'<w:r><w:rPr><w:color w:val="{MUTED}"/>{sz(10)}</w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
    f'<w:r><w:rPr><w:color w:val="{MUTED}"/>{sz(10)}</w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
    f'<w:r><w:rPr><w:color w:val="{MUTED}"/>{sz(10)}</w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
    f'<w:r><w:rPr><w:color w:val="{MUTED}"/>{sz(10)}</w:rPr><w:t>1</w:t></w:r>'
    f'<w:r><w:rPr><w:color w:val="{MUTED}"/>{sz(10)}</w:rPr><w:fldChar w:fldCharType="end"/></w:r>'
    '</w:p></w:ftr>'
)

# US Letter, Word-default 1 in margins (no custom margins: printer clips page numbers otherwise).
SECTPR = ('<w:sectPr><w:footerReference w:type="default" r:id="rIdFooter1"/>'
          '<w:pgSz w:w="12240" w:h="15840"/>'
          '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
          '<w:titlePg/></w:sectPr>')


def replace_style(xml, sid, new):
    pat = re.compile(r'<w:style [^>]*w:styleId="' + sid + r'"[^>]*>.*?</w:style>', re.S)
    if pat.search(xml):
        return pat.sub(lambda _m: new, xml, count=1)
    return xml.replace("</w:styles>", new + "</w:styles>")


def main():
    with zipfile.ZipFile(SRC) as zin:
        parts = {n: zin.read(n) for n in zin.namelist()}

    styles = parts["word/styles.xml"].decode("utf8")
    styles = re.sub(r'<w:rPrDefault>.*?</w:rPrDefault>',
                    f'<w:rPrDefault><w:rPr>{FONT}{sz(BODY_PT)}<w:lang w:val="en-US"/></w:rPr></w:rPrDefault>', styles, flags=re.S)
    styles = re.sub(r'<w:pPrDefault>.*?</w:pPrDefault>',
                    '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>', styles, flags=re.S)
    for sid, xml in STYLES.items():
        styles = replace_style(styles, sid, xml)
    styles = replace_style(styles, "Table", TABLE_STYLE)
    parts["word/styles.xml"] = styles.encode("utf8")

    doc = parts["word/document.xml"].decode("utf8")
    doc = re.sub(r'<w:sectPr.*?</w:sectPr>', '', doc, flags=re.S)
    doc = doc.replace("</w:body>", SECTPR + "</w:body>")
    parts["word/document.xml"] = doc.encode("utf8")

    parts["word/footer1.xml"] = FOOTER.encode("utf8")
    rels = parts["word/_rels/document.xml.rels"].decode("utf8")
    rels = rels.replace("</Relationships>",
                        '<Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>')
    parts["word/_rels/document.xml.rels"] = rels.encode("utf8")
    ct = parts["[Content_Types].xml"].decode("utf8")
    ct = ct.replace("</Types>", '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')
    parts["[Content_Types].xml"] = ct.encode("utf8")

    with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, data in parts.items():
            zout.writestr(n, data)
    print("wrote", DST)


if __name__ == "__main__":
    main()
