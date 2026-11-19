"""Post-render cleanup for the proposal .docx (runs automatically via _quarto.yml).

1. Quarto wraps every cross-referenced figure/table in a one-cell layout table. Those
   wrappers inherit the ruled "Table" style, which draws stray lines around the charts.
   Real data tables have tblLook firstRow="1"; wrappers have firstRow="0", so we strip
   borders from the wrappers only.
2. Pandoc leaves the Table of Contents as an empty field for Word to fill on open.
   We have Word fill it (and number the pages) and save, so the ToC is populated even
   in viewers that don't update fields (Google Docs, Canvas preview, Outlook).
"""
import os
import re
import subprocess
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

NO_BORDERS = ('<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/>'
              '<w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>')

# 12 pt above and below every captioned figure/table (twips: 20 per pt).
WRAPPER_PADDING = ('<w:tblCellMar><w:top w:w="240" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>'
                   '<w:bottom w:w="240" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>')


def strip_wrapper_borders(path: Path):
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    doc = parts["word/document.xml"].decode("utf8")

    def fix(m):
        tblpr = m.group(0)
        if 'w:firstRow="0"' in tblpr and "<w:tblBorders>" not in tblpr:
            # Borderless wrapper + equal vertical padding. Figures and tables share this
            # wrapper, so both get the same space above and below.
            return tblpr.replace("</w:tblPr>", NO_BORDERS + WRAPPER_PADDING + "</w:tblPr>")
        return tblpr

    doc, _ = re.subn(r"<w:tblPr>.*?</w:tblPr>", fix, doc, flags=re.S)
    # Quarto emits <w:jc left/><w:jc center/> in data-table cells; the second wins and
    # overrides the markdown's :--- alignment. Keep the first (the one the markdown asked for).
    doc = re.sub(r'(<w:jc w:val="(?:left|right)" ?/>)\s*<w:jc w:val="center" ?/>', r"\1", doc)
    # References starts on a new page. "Page break before" on the heading never leaves a
    # blank page (a manual break does when the previous section ends exactly at a page end).
    doc, n = re.subn(r'(<w:p>\s*<w:pPr>\s*<w:pStyle w:val="Heading1" ?/>)((?:(?!</w:p>).)*?>References</w:t>)',
                     r"\1<w:pageBreakBefore/>\2", doc, count=1, flags=re.S)
    if n != 1:
        print("WARNING: References heading not found; no page break added")
    parts["word/document.xml"] = doc.encode("utf8")
    tmp = path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in parts.items():
            z.writestr(n, data)
    os.replace(tmp, path)


def update_toc_with_word(path: Path):
    ps = (
        "$ErrorActionPreference='Stop';"
        "$w=New-Object -ComObject Word.Application; $w.Visible=$false; $w.DisplayAlerts=0;"
        # Pass 1: Convert() leaves Compatibility Mode. Save and close, because Word keeps
        # the pre-conversion layout in memory and would record wrong ToC page numbers.
        f"try {{ $d=$w.Documents.Open('{path}', $false, $false);"
        "$d.Convert(); $d.Save(); $d.Close();"
        # Pass 2: reopen with the final layout, then fill the ToC.
        f"$d=$w.Documents.Open('{path}', $false, $false); $d.Repaginate();"
        "$d.Fields.Update() | Out-Null; $d.TablesOfContents | ForEach-Object { $_.Update() };"
        "$d.Save(); $d.Close() } finally { $w.Quit() }"
    )
    r = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True)
    if r.returncode != 0:
        print("WARNING: Word could not update the ToC; it will fill in when opened in Word.\n", r.stderr)
    else:
        print("ToC updated by Word:", path.name)


outputs = os.environ.get("QUARTO_PROJECT_OUTPUT_FILES", "").split("\n")
for f in filter(None, outputs):
    p = (ROOT / f).resolve() if not os.path.isabs(f) else Path(f)
    if p.suffix.lower() == ".docx" and p.exists():
        strip_wrapper_borders(p)
        update_toc_with_word(p)
