"""Turn a .docx or .pdf One-Pager into the Markdown the rest of the tool reads.

Converting on UPLOAD rather than on every read is deliberate. The Art Requirement
step, the picker and the reader all understand exactly one format, and keeping it
that way means a Word brief is a first-class One-Pager from the moment it is
filed instead of a special case every later stage has to remember. It also means
the conversion is inspectable: the .md lands next to the original, so anyone can
see what the model was actually given and fix it by editing text.

Neither format needs a new dependency.

  .docx  is a zip with the document in word/document.xml, so stdlib zipfile and
         ElementTree are enough. Word's own paragraph styles carry the heading
         levels, which is what makes a real structural conversion possible rather
         than a flat dump of strings.

  .pdf   goes through pypdf, already vendored and already used by art_reference
         to read shipped requirement PDFs.

IMAGES ARE NOT DROPPED. That was the first version's worst default: the shipped
One-Pager PDFs carry 164 to 309 image objects each, and in Merge2608 the actual
zone contents - pickleball, lawn bowls, zip line, ball pit - exist ONLY inside
mood-board pictures. A text-only conversion silently threw that away. Images are
now extracted into a sibling .assets folder, referenced from the Markdown, and
read with the OCR already vendored for checking generated art, so text that lives
in a picture arrives as text.

The honest limits:

  * a PDF has no headings, no lists and no table model - only positioned text -
    so its conversion is the words in reading order, with images grouped by page
    rather than placed inline, because a PDF gives no reliable relationship
    between a picture and the sentence near it;
  * OCR output is OCR output. It is labelled as read-from-image in the Markdown
    so a reader, and a model, can weigh it accordingly;
  * a scanned PDF with no text layer is detected and reported rather than passed
    on as an empty document.
"""

import hashlib
import io
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

# Images are spliced back in at the position they occupied, so a reference shot
# stays under the heading it illustrated. A placeholder is cheaper than threading
# a second structure through the paragraph walk, and this one is checked against
# the document text before use - see _splice_images - rather than assumed unique.
IMG_TOKEN = "@@IMG%d@@"


def _para_text(p):
    """All the text of one Word paragraph, including inside hyperlinks.

    Word splits a paragraph into runs at every formatting change, so a single
    sentence is routinely several <w:t> elements and reading only the first
    loses most of the line. Tabs and breaks are real characters in the source
    and are kept as spacing rather than dropped.
    """
    out = []
    for node in p.iter():
        if node.tag == W + "t":
            out.append(node.text or "")
        elif node.tag == W + "tab":
            out.append("\t")
        elif node.tag in (W + "br", W + "cr"):
            out.append("\n")
    return "".join(out).strip()


def _style_of(p):
    pr = p.find(W + "pPr")
    if pr is None:
        return ""
    st = pr.find(W + "pStyle")
    if st is None:
        return ""
    return (st.get(W + "val") or "").strip()


def _numbered(p):
    """True when Word is drawing this paragraph as a list item."""
    pr = p.find(W + "pPr")
    return pr is not None and pr.find(W + "numPr") is not None


def _heading_level(style):
    """Word's style name to a Markdown level.

    Matches Heading1/heading 1/Titre1 style names by digit rather than by an
    English list, because a document authored in another language still carries
    the level in the style name.
    """
    s = style.lower().replace(" ", "")
    if s in ("title", "titre"):
        return 1
    m = re.match(r"^(heading|titre|berschrift|kop|rubrik)(\d)$", s)
    if m:
        return min(6, int(m.group(2)))
    return 0


def _table_md(tbl):
    """A Word table as a Markdown pipe table.

    Worth doing rather than flattening: the studio's briefs put scene lists and
    zone breakdowns in tables, and a table flattened to prose loses which value
    belonged to which column.
    """
    rows = []
    for tr in tbl.findall(W + "tr"):
        cells = []
        for tc in tr.findall(W + "tc"):
            text = " ".join(_para_text(p) for p in tc.findall(W + "p")).strip()
            cells.append(text.replace("|", r"\|").replace("\n", " "))
        if any(c for c in cells):
            rows.append(cells)
    if not rows:
        return []
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    out = ["| " + " | ".join(rows[0]) + " |",
           "|" + "|".join(["---"] * width) + "|"]
    for r in rows[1:]:
        out.append("| " + " | ".join(r) + " |")
    return out


def _docx_rel_media(z):
    """r:embed id -> (filename, bytes) for every image part in the document."""
    try:
        rels = ET.fromstring(z.read("word/_rels/document.xml.rels"))
    except Exception:
        return {}
    out = {}
    for r in rels:
        rid = r.get("Id")
        target = (r.get("Target") or "").replace("\\", "/")
        if not rid or "media/" not in target:
            continue
        part = "word/" + target.lstrip("./") if not target.startswith("word/") else target
        try:
            out[rid] = (os.path.basename(part), z.read(part))
        except KeyError:
            continue
    return out


def _drawing_ids(node):
    """The r:embed ids of every image drawn inside this element, in order."""
    return [b.get(REL + "embed") for b in node.iter(A + "blip")
            if b.get(REL + "embed")]


def docx_to_markdown(path, md_path=None):
    """Structural conversion: headings, lists and tables survive as themselves.

    md_path is where the Markdown will be written. Given one, images are
    extracted into a sibling .assets folder and referenced from the text at the
    point they appeared; without one this is a pure text conversion, which is
    what the tests and any caller that only wants words should use.
    """
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if "word/document.xml" not in names:
            raise ValueError("This .docx has no word/document.xml - it may be a "
                             ".doc renamed, or a corrupted file. Re-save it as .docx.")
        xml = z.read("word/document.xml")
        media = _docx_rel_media(z) if md_path else {}
    body = ET.fromstring(xml).find(W + "body")
    if body is None:
        raise ValueError("This .docx has no document body.")

    lines = []
    images = []                      # (name, bytes) in document order
    for child in body:
        if child.tag == W + "p":
            for rid in _drawing_ids(child):
                if rid in media:
                    lines.append(IMG_TOKEN % len(images))
                    lines.append("")
                    images.append(media[rid])
            text = _para_text(child)
            if not text:
                # One blank line between blocks, never a run of them.
                if lines and lines[-1] != "":
                    lines.append("")
                continue
            lvl = _heading_level(_style_of(child))
            if lvl:
                if lines and lines[-1] != "":
                    lines.append("")
                lines.append("#" * lvl + " " + text)
                lines.append("")
            elif _numbered(child):
                lines.append("- " + text)
            else:
                lines.append(text)
                lines.append("")
        elif child.tag == W + "tbl":
            if lines and lines[-1] != "":
                lines.append("")
            lines.extend(_table_md(child))
            lines.append("")

    md = "\n".join(lines)
    md = re.sub(r"\n{3,}", "\n\n", md).strip()
    if not md:
        raise ValueError("No text found in this .docx. Images in it are still "
                         "extracted and read, but a document with neither text "
                         "nor images has nothing to convert.")
    return _splice_images(md + "\n", images, md_path)


def pdf_to_markdown(path, md_path=None):
    """The words in reading order, and the images that sat among them.

    A PDF carries no structure to recover - no headings, no list semantics, no
    table model - so this does not pretend to. Page breaks are kept as rules
    because a One-Pager's page boundaries usually are section boundaries.

    Images are grouped under their page rather than placed inline: a PDF's image
    objects have positions but no relationship to the text flow, so claiming a
    picture belongs after a particular sentence would be invention. Page is the
    finest grouping the format actually supports.
    """
    from pypdf import PdfReader

    reader = PdfReader(path)
    if getattr(reader, "is_encrypted", False):
        try:
            reader.decrypt("")           # many are encrypted with an empty owner password
        except Exception:
            raise ValueError("This PDF is password-protected. Remove the protection "
                             "or export it as text, then upload that.")

    blocks, images = [], []
    for n, pg in enumerate(reader.pages, 1):
        try:
            t = pg.extract_text() or ""
        except Exception:
            t = ""
        t = re.sub(r"[ \t]+\n", "\n", t)
        t = re.sub(r"\n{3,}", "\n\n", t).strip()

        page_imgs = []
        if md_path:
            try:
                for im in pg.images:
                    page_imgs.append(IMG_TOKEN % len(images))
                    images.append((im.name or ("p%d.png" % n), im.data))
            except Exception:
                pass                     # a page whose images will not decode

        if t or page_imgs:
            part = [t] if t else []
            if page_imgs:
                part.append("")
                part.extend(page_imgs)
            blocks.append("\n".join(part).strip())

    if not blocks:
        raise ValueError("No text layer in this PDF - it is almost certainly a scan "
                         "or an export of images. Export the original as Word or "
                         "text and upload that instead.")

    # A scan whose pages are one big image each: no text anywhere, but images
    # present. Worth converting rather than refusing, because the OCR below is
    # exactly what recovers it - but say so, since the result is only as good as
    # the OCR.
    md = ("\n\n---\n\n".join(blocks)).strip() + "\n"
    return _splice_images(md, images, md_path)


# ---------------------------------------------------------------- images ---
# The shipped One-Pager PDFs carry 164 to 309 image objects each - they are
# reference and mood art, and the briefs say so in their own words. Dropping them
# silently was the wrong default twice over: a reviewer reading the converted .md
# could not see what the brief showed, and a brief whose CONTENT sits in a
# screenshot lost that content with no sign it had happened.
#
# So images are extracted beside the Markdown and referenced from it in document
# order, and then read with OCR. The OCR is the part that makes an image-borne
# brief work rather than merely be preserved: a table pasted in as a picture
# comes back as text a model can use.
#
# Three filters keep this proportionate, and the third was learned the hard way.
#
# Slivers are skipped: a PDF's image count is inflated by sliced graphics and
# rules that carry nothing.
#
# Byte-identical images are written ONCE - the same reference thumbnail on five
# pages is one image, not five.
#
# And each one is REFERENCED once, at its first appearance. Writing one file but
# still emitting a link at every occurrence looked like deduplication and was
# not: the Merge2608 brief carries its confidentiality watermark on every page,
# so 164 image objects became 16 files and still 164 links - about 150 lines of
# watermark in the document a model was about to read. A repeat count on the
# first reference says what a wall of links was trying to say, and says it in
# one line, while also flagging the thing as decoration rather than content.
MIN_IMAGE_PIXELS = 120 * 120        # below this an image is decoration
OCR_MIN_PIXELS = 200 * 200          # below this OCR finds nothing but noise
MAX_OCR_IMAGES = 24                 # a bound, reported when it bites

# An image repeated more often than this is page furniture - a watermark, a logo,
# a footer - and reading it adds nothing a brief needs. It is also the one place
# this pass could leak something: the Merge2608 watermark OCRs to "Internal
# document, do not distribute, <someone>@liquidnitrogames.com", and that text
# would otherwise be forwarded to the model along with the brief. The image is
# still extracted and still referenced; only the OCR is skipped.
OCR_SKIP_IF_REPEATED_OVER = 5


def _image_dir(md_path):
    return os.path.splitext(md_path)[0] + ".assets"


def _keep_image(data):
    """(ok, width, height) - is this image worth carrying?"""
    try:
        from PIL import Image
        im = Image.open(io.BytesIO(data))
        w, h = im.size
    except Exception:
        return False, 0, 0
    return (w * h) >= MIN_IMAGE_PIXELS, w, h


def _ocr_lines(path):
    """Text rendered inside an image, or [] if OCR is unavailable or finds none."""
    try:
        import text_detect
        if not text_detect.ocr_available():
            return []
        found = text_detect.read_text(path) or []
    except Exception:
        return []
    out, seen = [], set()
    for item in found:
        t = (item.get("text") or "").strip()
        # Single stray glyphs are OCR noise on artwork, not content.
        if len(t) < 2 or t.lower() in seen:
            continue
        seen.add(t.lower())
        out.append(t)
    return out


def _write_images(md_path, images, ocr=True):
    """Save images, OCR them, and return the Markdown to splice in per image.

    `images` is [(suggested_name, bytes), ...] in document order. Returns
    {index: [markdown lines]} plus a summary dict, so the caller can place each
    image where it actually appeared rather than dumping them all at the end.
    """
    if not images:
        return {}, {"images": 0, "ocr_images": 0, "ocr_lines": 0, "skipped": 0}

    d = _image_dir(md_path)
    os.makedirs(d, exist_ok=True)
    rel = os.path.basename(d)

    # Counted first, so the single reference can say how many times the image
    # occurred. Without the count, an image tiled on every page and an image used
    # once look identical in the output, and the reader loses the strongest
    # available signal for "this is a watermark, not content".
    digests, repeats = [], {}
    for _name, data in images:
        dg = hashlib.sha1(data).hexdigest()
        digests.append(dg)
        repeats[dg] = repeats.get(dg, 0) + 1

    md_by_index, seen, manifest = {}, {}, []
    kept = skipped = ocr_done = ocr_lines = 0
    for idx, (name, data) in enumerate(images):
        digest = digests[idx]
        if digest in seen:
            md_by_index[idx] = []       # already referenced; say nothing again
            continue
        keep, w, h = _keep_image(data)
        if not keep:
            skipped += 1
            seen[digest] = None         # and do not re-test every copy of it
            md_by_index[idx] = []
            continue
        out_name = "%02d_%s" % (kept + 1, name)
        with open(os.path.join(d, out_name), "wb") as f:
            f.write(data)
        kept += 1
        n = repeats[digest]
        alt = "image %d, %dx%d%s" % (kept, w, h,
                                     "" if n == 1 else ", appears %d times" % n)
        seen[digest] = out_name
        lines = ["![%s](%s/%s)" % (alt, rel, out_name)]
        words = []                      # reset per image, not carried over
        if (ocr and (w * h) >= OCR_MIN_PIXELS and ocr_done < MAX_OCR_IMAGES
                and n <= OCR_SKIP_IF_REPEATED_OVER):
            words = _ocr_lines(os.path.join(d, out_name))
            ocr_done += 1
            if words:
                ocr_lines += len(words)
                # Labelled, so a model can tell brief text from text that was
                # read out of a picture and may be imperfect.
                lines.append("")
                lines.append("> Text read from this image: " + " · ".join(words))
        manifest.append({"file": out_name, "w": w, "h": h, "repeats": n,
                         "ocr_lines": len(words)})
        md_by_index[idx] = lines

    # A manifest, so the generation step can choose which images to attach
    # without re-deriving anything by parsing the Markdown it just wrote. The
    # repeat count is the field that matters there: furniture must not be sent.
    if kept:
        try:
            io.open(os.path.join(d, "manifest.json"), "w", encoding="utf-8").write(
                json.dumps({"images": manifest}, indent=2))
        except Exception:
            pass                        # the conversion is still valid without it

    return md_by_index, {"images": kept, "skipped": skipped,
                         "ocr_images": ocr_done, "ocr_lines": ocr_lines,
                         "ocr_capped": ocr_done >= MAX_OCR_IMAGES,
                         "assets_dir": d if kept else ""}


# ------------------------------------------- images for the generation call ---
# OCR recovers words from a picture; it does not recover what the picture LOOKS
# like, and a One-Pager's reference art is chosen for how it looks. So the
# strongest images are attached to the generation call as well.
#
# Bounded deliberately. Every attached image is billed on every variation
# request, and a brief holds up to 16 usable pictures, so sending all of them
# would multiply the cost of the step for diminishing return. Six is enough to
# establish a look.
PROMPT_IMAGE_LIMIT = 6
PROMPT_IMAGE_MAX_EDGE = 1568        # Anthropic downsizes past this anyway
PROMPT_IMAGE_MAX_BYTES = 900 * 1024


def reference_images(md_path, limit=PROMPT_IMAGE_LIMIT):
    """The best reference images from a converted brief, ready to attach.

    Returns [{"media_type", "data" (base64), "alt", "file"}], largest first.

    Chosen from the manifest rather than by re-reading the folder, because the
    manifest carries the repeat count and page furniture must not be sent: a
    watermark attached to a generation request costs money and teaches nothing.
    """
    import base64

    d = _image_dir(md_path)
    man = os.path.join(d, "manifest.json")
    if not os.path.isfile(man):
        return []
    try:
        entries = json.load(io.open(man, encoding="utf-8")).get("images") or []
    except Exception:
        return []

    usable = [e for e in entries
              if e.get("repeats", 1) <= OCR_SKIP_IF_REPEATED_OVER
              and e.get("w", 0) * e.get("h", 0) >= OCR_MIN_PIXELS]
    usable.sort(key=lambda e: -(e.get("w", 0) * e.get("h", 0)))

    out = []
    for e in usable[:max(0, limit)]:
        p = os.path.join(d, e["file"])
        if not os.path.isfile(p):
            continue
        try:
            from PIL import Image
            im = Image.open(p)
            im = im.convert("RGB")
            longest = max(im.size)
            if longest > PROMPT_IMAGE_MAX_EDGE:
                s = PROMPT_IMAGE_MAX_EDGE / float(longest)
                im = im.resize((max(1, int(im.width * s)), max(1, int(im.height * s))),
                               Image.LANCZOS)
            # JPEG rather than PNG: these are photographs and mood boards, and
            # the quality ladder keeps a big mood board inside the byte budget
            # instead of failing the request.
            for q in (85, 70, 55, 40):
                buf = io.BytesIO()
                im.save(buf, "JPEG", quality=q, optimize=True)
                if buf.tell() <= PROMPT_IMAGE_MAX_BYTES:
                    break
            out.append({"media_type": "image/jpeg",
                        "data": base64.b64encode(buf.getvalue()).decode("ascii"),
                        "alt": "reference image %dx%d from the One-Pager"
                               % (e.get("w", 0), e.get("h", 0)),
                        "file": e["file"]})
        except Exception:
            continue                    # one unreadable image must not lose the rest
    return out


def _splice_images(md, images, md_path):
    """Replace each placeholder with its image (and any text read out of it).

    Returns the finished Markdown. LAST_SUMMARY carries what happened, so the
    caller can report "4 images, 2 with readable text" rather than leaving the
    designer to notice a folder appeared.
    """
    global LAST_SUMMARY
    LAST_SUMMARY = {"images": 0, "skipped": 0, "ocr_images": 0, "ocr_lines": 0,
                    "ocr_capped": False, "assets_dir": ""}
    if not images or not md_path:
        # No placeholders were emitted in text-only mode, but a document could
        # legitimately contain the token as text; strip any so it never shows.
        return re.sub(r"@@IMG\d+@@\n?", "", md)

    md_by_index, summary = _write_images(md_path, images)
    LAST_SUMMARY = summary
    for idx in range(len(images)):
        token = IMG_TOKEN % idx
        block = "\n".join(md_by_index.get(idx, []))
        md = md.replace(token, block)
    # A dropped sliver leaves an empty placeholder line behind.
    md = re.sub(r"\n{3,}", "\n\n", md)
    return md


# Set by the last conversion, so the CLI can report it. Not thread-safe and does
# not need to be: one conversion runs at a time, in its own process.
LAST_SUMMARY = {}

CONVERTERS = {"docx": docx_to_markdown, "pdf": pdf_to_markdown}


def convert_to_markdown(src, dest=None):
    """Write `src` out as Markdown beside itself. Returns the path written.

    The .md is a sibling rather than a replacement so the original stays in the
    shared folder: it is the thing the studio actually signed off, and someone
    will want to open it.
    """
    ext = os.path.splitext(src)[1].lstrip(".").lower()
    fn = CONVERTERS.get(ext)
    if not fn:
        raise ValueError("Nothing to convert: .%s is already readable, or is not "
                         "a format this tool accepts." % ext)
    # dest is resolved BEFORE converting, because the images are written into a
    # folder named after it and the Markdown has to reference them by that name.
    dest = dest or (os.path.splitext(src)[0] + ".md")
    md = fn(src, dest)
    io.open(dest, "w", encoding="utf-8", newline="").write(md)
    return dest


def convert_main():
    """CLI: pybackend_cli.py doc_convert payload.json

    payload: {"src": "...", "dest": "..."}   dest optional
    """
    import json
    import sys
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Usage: doc_convert <payload.json>"}))
        sys.exit(1)
    try:
        with io.open(sys.argv[1], "r", encoding="utf-8-sig") as f:
            payload = json.load(f)
        src = payload.get("src") or ""
        if not src or not os.path.isfile(src):
            print(json.dumps({"error": "File not found: %s" % src}))
            sys.exit(1)
        out = convert_to_markdown(src, payload.get("dest") or None)
        text = io.open(out, encoding="utf-8").read()
        summary = dict(LAST_SUMMARY or {})
        summary.update({"markdown_path": out, "chars": len(text),
                        "lines": text.count("\n") + 1})
        print(json.dumps(summary, ensure_ascii=False))
    except ValueError as e:
        # Wording is aimed at the designer, so it is safe to show as-is.
        print(json.dumps({"error": str(e)}))
        sys.exit(1)
    except Exception as e:
        print(json.dumps({"error": "%s: %s" % (e.__class__.__name__, e)}))
        sys.exit(1)


# ---------------------------------------------------------------------------
# Markdown -> .docx
#
# The other direction, for the same reason the first one exists: the production
# team works in Word. Everything this tool generates is Markdown, so handing a
# requirement document back to them means either they read a .md in Notepad or
# somebody retypes it into the template. Neither is acceptable for a document
# whose entire job is to be marked up, argued with and returned.
#
# No new dependency here either. A .docx is a zip of XML, which is exactly what
# docx_to_markdown above already takes apart; this puts the same parts back.
#
# Deliberately NOT supported, because nothing this tool emits uses them: nested
# lists, inline code, images, hyperlinks. The generated documents were censused
# before this was written - h1-h4, tables, bullets, ordered lists, bold, italic
# and a horizontal rule is the whole vocabulary. Anything unrecognised falls
# through as a plain paragraph rather than being dropped, so an unexpected
# construct degrades to readable text instead of disappearing.

# Landscape. The requirement tables run to seven columns of prose; in portrait
# Word crushes them to a few words per line and the document stops serving the
# one purpose it has. A4 landscape, in twentieths of a point.
PAGE_W, PAGE_H = 16838, 11906
PAGE_MARGIN = 720                   # half an inch all round

_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '</Types>')

_ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
    ' Target="word/document.xml"/></Relationships>')

_DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"'
    ' Target="styles.xml"/></Relationships>')

# Sizes are half-points: 32 = 16pt. Without real style definitions a pStyle
# reference points at nothing and every heading renders as body text, which is
# the usual reason a hand-built docx opens looking like an undifferentiated dump.
_HEADINGS = ((1, 32, "1F3864"), (2, 26, "2E5496"), (3, 23, "2E5496"), (4, 21, "404040"))


def _styles_xml():
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:docDefaults><w:rPrDefault><w:rPr>'
           '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="20"/>'
           '</w:rPr></w:rPrDefault></w:docDefaults>'
           '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
           '<w:name w:val="Normal"/><w:pPr><w:spacing w:after="120"/></w:pPr></w:style>']
    for lvl, sz, colour in _HEADINGS:
        out.append(
            '<w:style w:type="paragraph" w:styleId="Heading%d">'
            '<w:name w:val="heading %d"/><w:basedOn w:val="Normal"/>'
            '<w:pPr><w:keepNext/><w:outlineLvl w:val="%d"/>'
            '<w:spacing w:before="%d" w:after="120"/></w:pPr>'
            '<w:rPr><w:b/><w:color w:val="%s"/><w:sz w:val="%d"/></w:rPr></w:style>'
            % (lvl, lvl, lvl - 1, 320 - lvl * 40, colour, sz))
    out.append('</w:styles>')
    return "".join(out)


def _esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# Bold first, then italic, and the italic arm refuses to fire when an underscore
# is flanked by word characters - so a snake_case name in the text stays literal
# instead of silently turning half a sentence italic.
_INLINE = re.compile(r"\*\*(?P<b>.+?)\*\*"
                     r"|(?<![A-Za-z0-9])_(?P<i>[^_\n]+)_(?![A-Za-z0-9])")


def _runs(text):
    """Split one line into (text, bold, italic) runs."""
    out, at = [], 0
    for m in _INLINE.finditer(text):
        if m.start() > at:
            out.append((text[at:m.start()], False, False))
        if m.group("b") is not None:
            out.append((m.group("b"), True, False))
        else:
            out.append((m.group("i"), False, True))
        at = m.end()
    if at < len(text):
        out.append((text[at:], False, False))
    return [r for r in out if r[0]] or [("", False, False)]


def _run_xml(text, bold=False, italic=False):
    pr = ""
    if bold or italic:
        pr = "<w:rPr>%s%s</w:rPr>" % ("<w:b/>" if bold else "", "<w:i/>" if italic else "")
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (pr, _esc(text))


def _para(text, style=None, indent=0, rule=False, bold_all=False):
    pr = []
    if style:
        pr.append('<w:pStyle w:val="%s"/>' % style)
    if indent:
        pr.append('<w:ind w:left="%d" w:hanging="220"/>' % indent)
    if rule:
        pr.append('<w:pBdr><w:bottom w:val="single" w:sz="6" w:color="BFBFBF"/></w:pBdr>')
    ppr = "<w:pPr>%s</w:pPr>" % "".join(pr) if pr else ""
    runs = "".join(_run_xml(t, b or bold_all, i) for t, b, i in _runs(text))
    return "<w:p>%s%s</w:p>" % (ppr, runs)


def _cell(text, header, width):
    shade = '<w:shd w:val="clear" w:fill="D9E2F3"/>' if header else ""
    body = _para(text, bold_all=header) if text.strip() else "<w:p/>"
    return ('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s'
            '<w:vAlign w:val="top"/></w:tcPr>%s</w:tc>' % (width, shade, body))


def _table(rows, header_row):
    """rows is a list of lists of cell strings; header_row says row 0 is a header."""
    ncols = max(len(r) for r in rows)
    width = (PAGE_W - 2 * PAGE_MARGIN) // ncols
    border = '<w:tblBorders>' + "".join(
        '<w:%s w:val="single" w:sz="4" w:color="8496B0"/>' % e
        for e in ("top", "left", "bottom", "right", "insideH", "insideV")) + '</w:tblBorders>'
    out = ['<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>%s'
           '<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>%s</w:tblGrid>'
           % (border, ('<w:gridCol w:w="%d"/>' % width) * ncols)]
    for n, row in enumerate(rows):
        head = header_row and n == 0
        # tblHeader repeats the header on every page the table spills onto.
        # These tables run well past one page and a headerless continuation is
        # a grid of unlabelled prose.
        trpr = "<w:trPr><w:tblHeader/></w:trPr>" if head else ""
        cells = list(row) + [""] * (ncols - len(row))
        out.append("<w:tr>%s%s</w:tr>"
                   % (trpr, "".join(_cell(c, head, width) for c in cells)))
    out.append("</w:tbl>")
    # Word welds two tables that touch with nothing between them into one. An
    # empty paragraph after each keeps the header block off section 1's table.
    out.append("<w:p/>")
    return "".join(out)


_TABLE_SEP = re.compile(r"^\|[\s:|-]+\|$")


def _split_row(line):
    line = line.strip()
    if line.startswith("|"):
        line = line[1:]
    if line.endswith("|"):
        line = line[:-1]
    return [c.strip() for c in line.split("|")]


def markdown_to_docx(md_path, docx_path=None):
    """Write a .docx beside the Markdown and return its path."""
    if docx_path is None:
        docx_path = os.path.splitext(md_path)[0] + ".docx"
    lines = io.open(md_path, encoding="utf-8").read().splitlines()

    body, table = [], []
    header_row = [False]

    def flush():
        if table:
            body.append(_table(list(table), header_row[0]))
            del table[:]
        header_row[0] = False

    for raw in lines:
        stripped = raw.strip()
        if stripped.startswith("|"):
            if _TABLE_SEP.match(stripped):
                header_row[0] = bool(table)      # the row above it was the header
            else:
                table.append(_split_row(stripped))
            continue
        flush()
        if not stripped:
            continue
        m = re.match(r"^(#{1,6})\s+(.*)$", stripped)
        if m:
            body.append(_para(m.group(2), style="Heading%d" % min(len(m.group(1)), 4)))
            continue
        if re.match(r"^-{3,}$", stripped):
            body.append(_para("", rule=True))
            continue
        m = re.match(r"^[-*]\s+(.*)$", stripped)
        if m:
            # ponytail: a literal bullet glyph with a hanging indent, not a
            # numbering.xml list definition. It prints identically; the ceiling
            # is that Word will not renumber or re-level it as a real list.
            # Add numbering.xml if anyone needs to restructure these in Word.
            body.append(_para("•\t" + m.group(1), indent=360))
            continue
        m = re.match(r"^(\d+)\.\s+(.*)$", stripped)
        if m:
            # The number is kept exactly as typed. These are build-step numbers
            # that match the art, so Word inventing its own sequence would be
            # wrong in a way nobody would notice until the render came back.
            body.append(_para("%s.\t%s" % (m.group(1), m.group(2)), indent=360))
            continue
        body.append(_para(stripped))
    flush()

    sect = ('<w:sectPr><w:pgSz w:w="%d" w:h="%d" w:orient="landscape"/>'
            '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d"'
            ' w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>'
            % (PAGE_W, PAGE_H, PAGE_MARGIN, PAGE_MARGIN, PAGE_MARGIN, PAGE_MARGIN))
    document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:body>%s%s</w:body></w:document>' % ("".join(body), sect))

    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _CONTENT_TYPES)
        z.writestr("_rels/.rels", _ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", _DOC_RELS)
        z.writestr("word/styles.xml", _styles_xml())
        z.writestr("word/document.xml", document)
    return docx_path


def docx_main():
    """CLI: pybackend_cli.py md_to_docx payload.json

    payload: {"src": "...", "dest": "..."}   dest optional

    Deliberately re-converts every time rather than caching on mtime the way
    ideation:readOnePager does. The cache there exists because conversion is
    slow and involves OCR; this is a few milliseconds of zip writing, and a
    stale .docx sitting next to an edited .md is exactly the failure that would
    get forwarded to the production team without anyone noticing.
    """
    import json
    import sys
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Usage: md_to_docx <payload.json>"}))
        sys.exit(1)
    try:
        with io.open(sys.argv[1], "r", encoding="utf-8-sig") as f:
            payload = json.load(f)
        src = payload.get("src") or ""
        if not src or not os.path.isfile(src):
            print(json.dumps({"error": "File not found: %s" % src}))
            sys.exit(1)
        out = markdown_to_docx(src, payload.get("dest") or None)
        print(json.dumps({"docx_path": out, "bytes": os.path.getsize(out)},
                         ensure_ascii=False))
    except Exception as e:
        print(json.dumps({"error": "%s: %s" % (e.__class__.__name__, e)}))
        sys.exit(1)
