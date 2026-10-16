#!/usr/bin/env python3
"""Builds the test documents in tools/tests/fixtures/ with LibreOffice.

    python3 -I tools/tests/make-fixtures.py            small fixtures (committed)
    python3 -I tools/tests/make-fixtures.py --large    also the large ones for timing (not committed)

The sources are flat OpenDocument files written here (so every detail the
tests check is known), converted by `soffice --headless --convert-to` into
.odt/.docx, .ods/.xlsx/.xls/.csv and .odp/.pptx. LibreOffice runs with a
profile of its own under test-output/, so it never touches a running one.
"""
import base64
import os
import struct
import subprocess
import sys
import zlib

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(HERE, "fixtures")
SRC = os.path.join(OUT, "src")
LARGE = os.path.join(ROOT, "test-output", "large")
PROFILE = os.path.join(ROOT, "test-output", "lo-profile")

NS = ('xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
      'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
      'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
      'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
      'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" '
      'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" '
      'xmlns:xlink="http://www.w3.org/1999/xlink" '
      'xmlns:dc="http://purl.org/dc/elements/1.1/" '
      'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" '
      'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" '
      'xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" '
      'xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" '
      'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"')


def png(width, height, rgb):
    """A plain one-colour PNG, written by hand."""
    raw = b"".join(b"\x00" + bytes(rgb) * width for _ in range(height))

    def chunk(kind, data):
        c = kind + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)) +
            chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def b64png(width, height, rgb):
    return base64.b64encode(png(width, height, rgb)).decode("ascii")


def write(path, text):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)


def text_document():
    image = b64png(40, 20, (200, 30, 30))
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<office:document {NS} office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text">
 <office:meta><dc:title>Fixture Text</dc:title><dc:creator>Documents tests</dc:creator></office:meta>
 <office:styles>
  <style:style style:name="Standard" style:family="paragraph"><style:text-properties fo:font-size="12pt"/></style:style>
  <style:style style:name="Quote" style:family="paragraph" style:parent-style-name="Standard">
   <style:paragraph-properties fo:margin-left="1cm" fo:text-align="center"/>
   <style:text-properties fo:font-style="italic"/>
  </style:style>
  <text:list-style style:name="Numbering">
   <text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/>
   <text:list-level-style-number text:level="2" style:num-format="a" style:num-suffix=")"/>
  </text:list-style>
  <text:list-style style:name="Bullets">
   <text:list-level-style-bullet text:level="1" text:bullet-char="•"/>
  </text:list-style>
 </office:styles>
 <office:automatic-styles>
  <style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>
  <style:style style:name="T2" style:family="text"><style:text-properties fo:color="#c00000" style:text-underline-style="solid"/></style:style>
  <style:style style:name="T3" style:family="text"><style:text-properties style:text-position="super 58%"/></style:style>
  <style:style style:name="Cell1" style:family="table-cell"><style:table-cell-properties fo:background-color="#ddeeff" fo:border="0.5pt solid #000000"/></style:style>
  <style:style style:name="Break" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:break-before="page"/></style:style>
 </office:automatic-styles>
 <office:body><office:text>
  <text:h text:outline-level="1">Fixture heading</text:h>
  <text:p text:style-name="Standard">Plain, <text:span text:style-name="T1">bold words</text:span> and <text:span text:style-name="T2">red underlined</text:span> text, E=mc<text:span text:style-name="T3">2</text:span>.</text:p>
  <text:p text:style-name="Standard">Two<text:s text:c="3"/>spaces and a<text:tab/>tab, a <text:a xlink:type="simple" xlink:href="https://www.webos-ports.org/">link</text:a> and a note<text:note text:id="ftn1" text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><text:p>The footnote text.</text:p></text:note-body></text:note>.</text:p>
  <text:p text:style-name="Quote">A centred quote.</text:p>
  <text:h text:outline-level="2">Lists</text:h>
  <text:list text:style-name="Numbering">
   <text:list-item><text:p>First item</text:p></text:list-item>
   <text:list-item><text:p>Second item</text:p>
    <text:list><text:list-item><text:p>Nested item</text:p></text:list-item></text:list>
   </text:list-item>
  </text:list>
  <text:list text:style-name="Bullets">
   <text:list-item><text:p>Bullet one</text:p></text:list-item>
   <text:list-item><text:p>Bullet two</text:p></text:list-item>
  </text:list>
  <text:h text:outline-level="2">Table</text:h>
  <table:table table:name="Table1">
   <table:table-column table:number-columns-repeated="3"/>
   <table:table-header-rows>
    <table:table-row>
     <table:table-cell table:style-name="Cell1" office:value-type="string"><text:p>Fruit</text:p></table:table-cell>
     <table:table-cell table:style-name="Cell1" office:value-type="string"><text:p>Colour</text:p></table:table-cell>
     <table:table-cell table:style-name="Cell1" office:value-type="string"><text:p>Count</text:p></table:table-cell>
    </table:table-row>
   </table:table-header-rows>
   <table:table-row>
    <table:table-cell office:value-type="string"><text:p>Apple</text:p></table:table-cell>
    <table:table-cell office:value-type="string"><text:p>Green</text:p></table:table-cell>
    <table:table-cell office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell>
   </table:table-row>
   <table:table-row>
    <table:table-cell table:number-columns-spanned="2" office:value-type="string"><text:p>Spanning two columns</text:p></table:table-cell>
    <table:covered-table-cell/>
    <table:table-cell office:value-type="float" office:value="7"><text:p>7</text:p></table:table-cell>
   </table:table-row>
  </table:table>
  <text:p text:style-name="Standard">A picture: <draw:frame draw:name="Red" text:anchor-type="as-char" svg:width="2cm" svg:height="1cm"><draw:image><office:binary-data>{image}</office:binary-data></draw:image><svg:title>Red box</svg:title></draw:frame></text:p>
  <text:p text:style-name="Break">Second page: findme appears here, and FINDME again, and Findmé with an accent.</text:p>
 </office:text></office:body>
</office:document>
"""


def spreadsheet():
    rows = []
    rows.append('<table:table-row>' + ''.join(
        f'<table:table-cell office:value-type="string"><text:p>{h}</text:p></table:table-cell>'
        for h in ("Item", "Price", "Qty", "Total")) + '</table:table-row>')
    for i, (name, price, qty) in enumerate([("Pen", 1.5, 4), ("Book", 12.25, 2), ("Lamp", 30, 1)], start=2):
        rows.append(
            '<table:table-row>'
            f'<table:table-cell office:value-type="string"><text:p>{name}</text:p></table:table-cell>'
            f'<table:table-cell office:value-type="float" office:value="{price}"><text:p>{price}</text:p></table:table-cell>'
            f'<table:table-cell office:value-type="float" office:value="{qty}"><text:p>{qty}</text:p></table:table-cell>'
            f'<table:table-cell table:formula="of:=[.B{i}]*[.C{i}]" office:value-type="float" office:value="{price * qty}"><text:p>{price * qty}</text:p></table:table-cell>'
            '</table:table-row>')
    rows.append('<table:table-row><table:table-cell table:number-columns-spanned="3" office:value-type="string"><text:p>Merged total label</text:p></table:table-cell><table:covered-table-cell table:number-columns-repeated="2"/>'
                '<table:table-cell table:formula="of:=SUM([.D2:.D4])" office:value-type="float" office:value="60.5"><text:p>60.5</text:p></table:table-cell></table:table-row>')
    second = ('<table:table-row><table:table-cell office:value-type="string"><text:p>Second sheet</text:p></table:table-cell></table:table-row>'
              '<table:table-row><table:table-cell office:value-type="string"><text:p>findme in sheet two</text:p></table:table-cell></table:table-row>')
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<office:document {NS} office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
 <office:automatic-styles>
  <style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="4cm"/></style:style>
  <style:style style:name="co2" style:family="table-column"><style:table-column-properties style:column-width="2cm"/></style:style>
 </office:automatic-styles>
 <office:body><office:spreadsheet>
  <table:table table:name="Prices"><table:table-column table:style-name="co1"/><table:table-column table:style-name="co2" table:number-columns-repeated="3"/>{''.join(rows)}</table:table>
  <table:table table:name="Notes"><table:table-column/>{second}</table:table>
 </office:spreadsheet></office:body>
</office:document>
"""


def presentation():
    image = b64png(64, 48, (30, 120, 200))
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<office:document {NS} office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.presentation">
 <office:styles>
  <style:style style:name="Default-title" style:family="presentation"><style:graphic-properties draw:fill="none" draw:stroke="none"/><style:text-properties fo:font-size="40pt"/></style:style>
  <style:style style:name="Default-outline1" style:family="presentation"><style:graphic-properties draw:fill="none" draw:stroke="none"/><style:text-properties fo:font-size="28pt"/></style:style>
 </office:styles>
 <office:automatic-styles>
  <style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="28cm" fo:page-height="15.75cm"/></style:page-layout>
  <style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#ffcc00" draw:stroke="none"/></style:style>
  <style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties draw:fill="solid" draw:fill-color="#eeeeff"/></style:style>
 </office:automatic-styles>
 <office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1"/></office:master-styles>
 <office:body><office:presentation>
  <draw:page draw:name="Title" draw:master-page-name="Default" draw:style-name="dp1">
   <draw:frame presentation:style-name="Default-title" presentation:class="title" svg:x="2cm" svg:y="1cm" svg:width="24cm" svg:height="3cm"><draw:text-box><text:p>Fixture slides</text:p></draw:text-box></draw:frame>
   <draw:frame presentation:style-name="Default-outline1" presentation:class="outline" svg:x="2cm" svg:y="5cm" svg:width="24cm" svg:height="8cm"><draw:text-box>
    <text:list><text:list-item><text:p>First point</text:p></text:list-item><text:list-item><text:p>Second point</text:p></text:list-item></text:list>
   </draw:text-box></draw:frame>
  </draw:page>
  <draw:page draw:name="Picture" draw:master-page-name="Default">
   <draw:frame svg:x="2cm" svg:y="2cm" svg:width="8cm" svg:height="6cm"><draw:image><office:binary-data>{image}</office:binary-data></draw:image></draw:frame>
   <draw:custom-shape draw:style-name="gr1" svg:x="12cm" svg:y="2cm" svg:width="6cm" svg:height="4cm"><text:p>Yellow box</text:p><draw:enhanced-geometry draw:type="rectangle"/></draw:custom-shape>
  </draw:page>
  <draw:page draw:name="Last" draw:master-page-name="Default">
   <draw:frame svg:x="2cm" svg:y="2cm" svg:width="20cm" svg:height="3cm"><draw:text-box><text:p>findme on the last slide</text:p></draw:text-box></draw:frame>
  </draw:page>
 </office:presentation></office:body>
</office:document>
"""


def markdown():
    return """# Markdown fixture

Some *emphasis*, **strong** and `code`, a [link](https://www.webos-ports.org/) and
<script>alert("not run")</script> raw HTML shown as text.

- one
- two
  - nested

1. first
2. second

| Name | Value |
|------|------:|
| a    | 1     |

```
code block <b>kept</b>
```

> quoted findme
"""


def large_text(paragraphs):
    body = []
    for i in range(paragraphs):
        if i % 50 == 0:
            body.append(f'<text:h text:outline-level="1">Chapter {i // 50 + 1}</text:h>')
        body.append('<text:p>' + (f'Paragraph {i}: ' + 'The quick brown fox jumps over the lazy dog. ' * 8) + '</text:p>')
        if i % 100 == 99:
            cells = ''.join('<table:table-row>' + ''.join(
                f'<table:table-cell office:value-type="string"><text:p>r{r}c{c}</text:p></table:table-cell>' for c in range(5)) +
                '</table:table-row>' for r in range(10))
            body.append(f'<table:table table:name="T{i}"><table:table-column table:number-columns-repeated="5"/>{cells}</table:table>')
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<office:document {NS} office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text">
 <office:body><office:text>{''.join(body)}</office:text></office:body>
</office:document>
"""


def large_csv(rows, cols):
    lines = [",".join(f"Col{c}" for c in range(cols))]
    for r in range(rows):
        lines.append(",".join(str(r * cols + c) if c % 3 else f"row {r} text {c}" for c in range(cols)))
    return "\n".join(lines) + "\n"


def convert(src, fmt, outdir):
    os.makedirs(PROFILE, exist_ok=True)
    cmd = ["soffice", "-env:UserInstallation=file://" + PROFILE, "--headless", "--convert-to", fmt, "--outdir", outdir, src]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)


def main():
    os.makedirs(SRC, exist_ok=True)
    write(os.path.join(SRC, "sample.fodt"), text_document())
    write(os.path.join(SRC, "sample.fods"), spreadsheet())
    write(os.path.join(SRC, "sample.fodp"), presentation())
    write(os.path.join(OUT, "sample.md"), markdown())
    write(os.path.join(OUT, "sample.txt"), "Plain text fixture.\nSecond line with findme.\nUnicode: éèê €\n")
    write(os.path.join(OUT, "semicolon.csv"), 'Name;Amount;Note\n"Smith; J.";12,5;"said ""hi"""\nJones;7;\n')
    for fmt in ("odt", "docx"):
        convert(os.path.join(SRC, "sample.fodt"), fmt, OUT)
    for fmt in ("ods", "xlsx", "xls"):
        convert(os.path.join(SRC, "sample.fods"), fmt, OUT)
    for fmt in ("odp", "pptx"):
        convert(os.path.join(SRC, "sample.fodp"), fmt, OUT)
    if "--large" in sys.argv:
        os.makedirs(LARGE, exist_ok=True)
        write(os.path.join(LARGE, "large.fodt"), large_text(3000))
        convert(os.path.join(LARGE, "large.fodt"), "docx", LARGE)
        convert(os.path.join(LARGE, "large.fodt"), "odt", LARGE)
        write(os.path.join(LARGE, "large.csv"), large_csv(50000, 12))
        convert(os.path.join(LARGE, "large.csv"), "xlsx", LARGE)
        convert(os.path.join(LARGE, "large.csv"), "ods", LARGE)
    for name in sorted(os.listdir(OUT)):
        p = os.path.join(OUT, name)
        if os.path.isfile(p):
            print(f"{os.path.getsize(p):>9}  {name}")


if __name__ == "__main__":
    main()
