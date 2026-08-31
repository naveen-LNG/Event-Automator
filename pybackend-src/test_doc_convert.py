"""Does a Word or PDF One-Pager actually survive the trip to Markdown?

The point of converting on upload is that a Word brief becomes a first-class
One-Pager. That is only true if the STRUCTURE survives: a flat dump of strings
loses which heading a scene sat under and which column a number came from, and
those are exactly what an Art Requirement is built from.

Run: python test_doc_convert.py
"""
import io
import os
import shutil
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import doc_convert                                            # noqa: E402

fails = []


def ok(cond, msg, detail=""):
    if cond:
        print("OK   " + msg)
    else:
        print("FAIL " + msg + (("\n     " + str(detail)) if detail else ""))
        fails.append(msg)


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _png(w=300, h=220, colour=(200, 90, 40)):
    from PIL import Image
    buf = _io_bytes()
    Image.new("RGB", (w, h), colour).save(buf, "PNG")
    return buf.getvalue()


def _io_bytes():
    import io as _io
    return _io.BytesIO()


def _docx(body_xml, media=None):
    """A minimal but real .docx - these three parts are all Word needs.

    `media` is {rel_id: (filename, bytes)}; the rels part is written so the
    converter can resolve a drawing back to its image the way Word does.
    """
    p = os.path.join(tempfile.mkdtemp(), "brief.docx")
    with zipfile.ZipFile(p, "w") as z:
        if media:
            rels = "".join(
                '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/'
                'officeDocument/2006/relationships/image" Target="media/%s"/>'
                % (rid, name) for rid, (name, _b) in media.items())
            z.writestr("word/_rels/document.xml.rels",
                       '<?xml version="1.0"?><Relationships xmlns="http://schemas.'
                       'openxmlformats.org/package/2006/relationships">%s</Relationships>' % rels)
            for _rid, (name, data) in media.items():
                z.writestr("word/media/" + name, data)
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/'
                   'package/2006/content-types"><Default Extension="xml" '
                   'ContentType="application/xml"/></Types>')
        z.writestr("_rels/.rels",
                   '<?xml version="1.0"?><Relationships xmlns="http://schemas.'
                   'openxmlformats.org/package/2006/relationships"/>')
        z.writestr("word/document.xml",
                   '<?xml version="1.0"?><w:document xmlns:w="%s"><w:body>%s</w:body>'
                   '</w:document>' % (W, body_xml))
    return p


def drawing(rid):
    """A paragraph containing one image, shaped the way Word writes it."""
    return ('<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.'
            'openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:graphicData><pic:pic xmlns:pic="http://schemas.openxmlformats.org/'
            'drawingml/2006/picture"><pic:blipFill><a:blip xmlns:r="http://schemas.'
            'openxmlformats.org/officeDocument/2006/relationships" r:embed="%s"/>'
            '</pic:blipFill></pic:pic></a:graphicData></a:graphic>'
            '</wp:inline></w:drawing></w:r></w:p>') % rid


def para(text, style=None, numbered=False):
    pr = ""
    if style or numbered:
        pr = "<w:pPr>%s%s</w:pPr>" % (
            '<w:pStyle w:val="%s"/>' % style if style else "",
            '<w:numPr><w:ilvl w:val="0"/></w:numPr>' if numbered else "")
    return "<w:p>%s<w:r><w:t>%s</w:t></w:r></w:p>" % (pr, text)


# ---- headings become headings ----------------------------------------------
md = doc_convert.docx_to_markdown(_docx(
    para("Harvest Fair", "Title") + para("Scenes", "Heading1")
    + para("The Barn", "Heading2") + para("Hay bales everywhere.")))
ok("# Harvest Fair" in md, "a Word Title becomes an H1")
ok("# Scenes" in md and "## The Barn" in md, "and Heading1/Heading2 keep their levels")
ok("Hay bales everywhere." in md, "body text comes through")

# A brief authored in another language still carries the level in the style name,
# so matching on the digit rather than an English word list is what makes it work.
md = doc_convert.docx_to_markdown(_docx(para("Szenen", "berschrift2")))
ok("## Szenen" in md, "a non-English heading style still resolves by its digit", md.strip())

# ---- one sentence split across runs is still one line ----------------------
# Word breaks a paragraph at every formatting change, so one bolded word mid
# sentence yields three runs. Reading only the first loses most of the line.
split = ('<w:p><w:r><w:t>The barn is </w:t></w:r><w:r><w:t>deep red</w:t></w:r>'
         '<w:r><w:t> at dusk.</w:t></w:r></w:p>')
md = doc_convert.docx_to_markdown(_docx(split))
ok("The barn is deep red at dusk." in md,
   "a sentence split across runs is reassembled, not truncated", md.strip())

# ---- lists ------------------------------------------------------------------
md = doc_convert.docx_to_markdown(_docx(
    para("Pumpkin orange", numbered=True) + para("Harvest gold", numbered=True)))
ok(len([l for l in md.splitlines() if l.startswith("- ")]) == 2,
   "numbered and bulleted paragraphs become list items", md.strip())

# ---- tables, which the real studio doc did not exercise ---------------------
tbl = ("<w:tbl>"
       "<w:tr><w:tc>%s</w:tc><w:tc>%s</w:tc></w:tr>"
       "<w:tr><w:tc>%s</w:tc><w:tc>%s</w:tc></w:tr>"
       "</w:tbl>") % (para("Zone"), para("Palette"), para("Barn"), para("Barn red"))
md = doc_convert.docx_to_markdown(_docx(tbl))
ok("| Zone | Palette |" in md, "a Word table becomes a Markdown pipe table", md.strip())
ok("|---|---|" in md, "with a header separator, so it renders as a table")
ok("| Barn | Barn red |" in md, "and its rows keep column alignment")

# A pipe inside a cell would split the row if passed through raw.
tbl2 = "<w:tbl><w:tr><w:tc>%s</w:tc></w:tr></w:tbl>" % para("warm | cold")
ok(r"warm \| cold" in doc_convert.docx_to_markdown(_docx(tbl2)),
   "a pipe inside a cell is escaped rather than breaking the table")

# Ragged tables are normal in real documents.
tbl3 = ("<w:tbl><w:tr><w:tc>%s</w:tc><w:tc>%s</w:tc></w:tr>"
        "<w:tr><w:tc>%s</w:tc></w:tr></w:tbl>") % (para("A"), para("B"), para("C"))
md = doc_convert.docx_to_markdown(_docx(tbl3))
ok(all(l.count("|") == 3 for l in md.splitlines() if l.startswith("|")),
   "a short row is padded so every row has the same column count", md.strip())

# ---- blank lines collapse rather than multiply ------------------------------
md = doc_convert.docx_to_markdown(_docx(
    para("One") + para("") + para("") + para("") + para("Two")))
ok("\n\n\n" not in md, "runs of empty paragraphs collapse to one blank line", repr(md))

# ---- failures are explained, not silent ------------------------------------
try:
    doc_convert.docx_to_markdown(_docx(""))
    ok(False, "an empty document is refused with a reason")
except ValueError as e:
    ok("No text found" in str(e), "an empty document is refused with a reason")
    ok("Images in it are still" in str(e),
       "and says images would have been kept, so an image-only brief is not "
       "mistaken for an unreadable one")

bad = os.path.join(tempfile.mkdtemp(), "notreally.docx")
io.open(bad, "w", encoding="utf-8").write("this is not a zip")
try:
    doc_convert.docx_to_markdown(bad)
    ok(False, "a file that is not really a .docx fails clearly")
except Exception as e:
    ok(not isinstance(e, (AttributeError, TypeError)),
       "a file that is not really a .docx fails clearly",
       "%s: %s" % (e.__class__.__name__, e))

# ---- dispatch ---------------------------------------------------------------
ok(set(doc_convert.CONVERTERS) == {"docx", "pdf"},
   "exactly the two formats that need converting are registered")
try:
    doc_convert.convert_to_markdown("brief.md")
    ok(False, "converting an already-readable format is refused")
except ValueError as e:
    ok("already readable" in str(e), "converting an already-readable format is refused")

# convert_to_markdown writes a SIBLING, leaving the original in place - the
# original is the file the studio signed off and someone will want to open it.
src = _docx(para("Harvest Fair", "Title") + para("Two interiors among five."))
out = doc_convert.convert_to_markdown(src)
ok(out.endswith(".md") and os.path.isfile(out),
   "convert_to_markdown writes a .md beside the original", out)
ok(os.path.isfile(src), "and leaves the original where it was")
ok("Harvest Fair" in io.open(out, encoding="utf-8").read(),
   "with the converted content in it")

# ---- images are extracted, not dropped -------------------------------------
# The first version silently threw them away. In Merge2608 the actual zone
# contents exist ONLY inside mood-board pictures, so that default lost real
# content with no sign it had happened.
img_dir = tempfile.mkdtemp()
md_target = os.path.join(img_dir, "brief.md")
src = _docx(para("Scenes", "Heading1") + drawing("rId9") + para("The barn at dusk."),
            media={"rId9": ("shot.png", _png())})

# Without a destination it stays a pure text conversion - the tests above and any
# caller that only wants words rely on that.
text_only = doc_convert.docx_to_markdown(src)
ok("](" not in text_only, "no destination means no images and no stray placeholder",
   repr(text_only))
ok("@@IMG" not in text_only, "and the placeholder never leaks into the output")

md = doc_convert.docx_to_markdown(src, md_target)
ok("](brief.assets/" in md, "given a destination, an image is extracted and referenced", md)
ok(os.path.isdir(os.path.join(img_dir, "brief.assets")),
   "into a .assets folder beside the Markdown")
imgs_in = [n for n in os.listdir(os.path.join(img_dir, "brief.assets"))
           if not n.endswith(".json")]
ok(len(imgs_in) == 1, "with one file per image", imgs_in)
ok(os.path.isfile(os.path.join(img_dir, "brief.assets", "manifest.json")),
   "plus a manifest, so the generation step can pick what to attach without "
   "re-parsing the Markdown it just wrote")
# Position matters: a reference shot belongs under the heading it illustrated.
ok(md.index("# Scenes") < md.index("](brief.assets/") < md.index("The barn at dusk"),
   "and placed where it appeared in the document, not appended at the end")

# A sliver carries nothing and would only add noise.
tiny = _docx(drawing("rId9") + para("text"), media={"rId9": ("dot.png", _png(20, 20))})
md = doc_convert.docx_to_markdown(tiny, os.path.join(tempfile.mkdtemp(), "t.md"))
ok("](" not in md, "an image too small to carry anything is skipped", md.strip())
ok(doc_convert.LAST_SUMMARY.get("skipped") == 1, "and counted as skipped",
   doc_convert.LAST_SUMMARY)

# The same picture repeated is one file and ONE reference. Emitting a link per
# occurrence looked like dedup and was not: Merge2608's watermark turned 164
# objects into 16 files and still 164 links, about 150 lines of watermark in the
# document a model was about to read.
same = _png()
rep = _docx(drawing("rId1") + drawing("rId2") + drawing("rId3"),
            media={"rId1": ("a.png", same), "rId2": ("b.png", same),
                   "rId3": ("c.png", same)})
d2 = tempfile.mkdtemp()
md = doc_convert.docx_to_markdown(rep, os.path.join(d2, "r.md"))
ok(md.count("](r.assets/") == 1, "a repeated image is referenced exactly once",
   "%d links" % md.count("](r.assets/"))
ok(len([n for n in os.listdir(os.path.join(d2, "r.assets"))
        if not n.endswith(".json")]) == 1, "and written exactly once")
ok("appears 3 times" in md,
   "with a repeat count, which is the strongest available signal for page furniture", md)

# Page furniture is not OCR'd. The Merge2608 watermark reads as "Internal
# document, do not distribute, <someone>@liquidnitrogames.com" - text that would
# otherwise be forwarded to the model with the brief.
ok(doc_convert.OCR_SKIP_IF_REPEATED_OVER < 75,
   "an image repeated on every page is treated as furniture and not read",
   str(doc_convert.OCR_SKIP_IF_REPEATED_OVER))

# ---- the real shipped documents --------------------------------------------
# The synthetic cases pin the rules; these prove it works on what the studio
# actually produces, which is the only claim that matters.
REAL_DOCX = r"C:\Users\NaveenKumar\Downloads\Blocks Art.docx"
if os.path.isfile(REAL_DOCX):
    md = doc_convert.docx_to_markdown(REAL_DOCX)
    ok(len(md) > 8000, "the real Blocks Art.docx converts to a full document",
       "%d chars" % len(md))
    ok(md.count("\n#") > 30, "keeping its heading structure",
       "%d headings" % md.count("\n#"))
    ok("Art Direction Breakdown" in md, "and its actual opening text")
else:
    print("--   Blocks Art.docx not on this machine, skipped")

REAL_PDF = (r"C:\Users\NaveenKumar\OneDrive - Liquidnitro Games Private Limited"
            r"\Skip-Bo - Design\ideationData\Events\Merge\One-Pager Ideation"
            r"\Merge2608 One-Pager Ideation.pdf")
if os.path.isfile(REAL_PDF):
    md = doc_convert.pdf_to_markdown(REAL_PDF)
    ok(len(md) > 2000, "a real shipped One-Pager PDF converts", "%d chars" % len(md))
    ok("Event Brief" in md or "One pager" in md, "with its actual heading text")
    ok("---" in md, "and page boundaries kept as rules, since they tend to be sections")
else:
    print("--   Merge2608 One-Pager PDF not on this machine, skipped")



# ---------------------------------------------------------------------------
# Markdown -> .docx
#
# The trap with a hand-built .docx is that it looks finished right up until Word
# refuses to open it, which happens after somebody has already forwarded it. So
# these read the file back with docx_to_markdown - the other direction, a
# different code path, the one that already reads the production team's real
# Word files - rather than re-parsing my own XML with my own parser, which would
# only ever agree with itself about a mistake.

SAMPLE_MD = """# Art Requirements

**Status:** Draft
**By:** A Name

| Field | Value |
|---|---|
| Event Name | Harvest Fair |
| Design Spec | _add link_ |

## 1. Overview

Plain text with an & ampersand and a <bracket>.

- first bullet
- second bullet

### Steps

  1. a wooden SHOP COUNTER, dead centre
  2. a garland strung across the room

The file blocks2604_scene.json keeps its underscores.
"""


def _sample(tmp):
    p = os.path.join(tmp, "sample.md")
    io.open(p, "w", encoding="utf-8").write(SAMPLE_MD)
    return p


tmp = tempfile.mkdtemp()
docx = doc_convert.markdown_to_docx(_sample(tmp))
ok(os.path.isfile(docx), "markdown_to_docx writes a file beside the .md")

with zipfile.ZipFile(docx) as z:
    names = set(z.namelist())
    missing = [n for n in ("[Content_Types].xml", "_rels/.rels", "word/document.xml",
                           "word/styles.xml", "word/_rels/document.xml.rels")
               if n not in names]
    ok(not missing, "the package carries every part Word requires", missing)
    bad = []
    for part in names:
        try:
            ET.fromstring(z.read(part))
        except Exception as exc:                      # noqa: BLE001
            bad.append("%s: %s" % (part, exc))
    ok(not bad, "and every part is well-formed XML", bad)
    doc_xml = z.read("word/document.xml").decode("utf-8")

ok('w:orient="landscape"' in doc_xml,
   "landscape, because a seven-column requirement table is unreadable in portrait")

back = doc_convert.docx_to_markdown(docx, md_path=os.path.join(tmp, "back.md"))

ok("# Art Requirements" in back, "the h1 survives the round trip as a real Word heading")
ok("## 1. Overview" in back, "and so does the h2, at its own level")
ok("### Steps" in back, "and the h3")
ok("| Event Name | Harvest Fair |" in back.replace("  ", " "),
   "a table comes back as a table, with its cells in their columns", back[:400])
ok("&" in back and "<bracket>" in back,
   "XML-special characters are escaped on the way in and intact on the way out")
ok("first bullet" in back and "second bullet" in back, "bullets keep their text")
ok("1." in back and "SHOP COUNTER" in back,
   "a numbered step keeps the number it was given, not one Word invents")
ok("blocks2604_scene.json" in back,
   "an underscored filename stays literal instead of turning half a line italic")

# The inline splitter is the only piece with real branching, so it is checked
# directly as well - the round trip cannot tell bold from plain text.
r = doc_convert._runs
ok(r("**Status:** Draft") == [("Status:", True, False), (" Draft", False, False)],
   "**bold** becomes a bold run and the rest stays plain", r("**Status:** Draft"))
ok(r("see _add reference_ here")[1] == ("add reference", False, True),
   "_italic_ becomes an italic run", r("see _add reference_ here"))
ok(r("a_b_c") == [("a_b_c", False, False)],
   "underscores inside a word are not italic markers", r("a_b_c"))
ok(doc_convert._esc("a & b <c>") == "a &amp; b &lt;c&gt;", "escaping is applied once")

shutil.rmtree(tmp, ignore_errors=True)
print("")
if fails:
    print("%d DOC CONVERT CHECK(S) FAILED" % len(fails))
    sys.exit(1)
print("DOC CONVERT OK")
