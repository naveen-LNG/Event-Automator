"""Do the brief's reference images actually reach the model?

OCR recovers a mood board's WORDS. It does not recover what the board looks
like, and that is the half of a One-Pager the art direction is derived from. So
the images are attached to the generation call - and this pins the shape of that
call, because a wrong content block fails at the API with a 400 rather than
anywhere a suite would notice.

No API call is made. Both providers are intercepted and their outgoing request
inspected, so this costs nothing and needs no key.

Run: python test_prompt_images.py
"""
import io
import json
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import doc_convert                                            # noqa: E402
import ideation_generator as ig                               # noqa: E402

fails = []


def ok(cond, msg, detail=""):
    if cond:
        print("OK   " + msg)
    else:
        print("FAIL " + msg + (("\n     " + str(detail)) if detail else ""))
        fails.append(msg)


def _png(w, h, colour=(180, 90, 40)):
    from PIL import Image
    b = io.BytesIO()
    Image.new("RGB", (w, h), colour).save(b, "PNG")
    return b.getvalue()


# ---- a converted brief with a big reference image and a repeated watermark ---
d = tempfile.mkdtemp()
md_path = os.path.join(d, "brief.md")
assets = doc_convert._image_dir(md_path)
os.makedirs(assets)
big = _png(900, 1200, (60, 120, 200))
mark = _png(400, 300, (200, 200, 200))
io.open(os.path.join(assets, "01_mark.png"), "wb").write(mark)
io.open(os.path.join(assets, "02_big.png"), "wb").write(big)
io.open(os.path.join(assets, "03_mid.png"), "wb").write(_png(400, 500, (200, 90, 40)))
io.open(os.path.join(assets, "manifest.json"), "w", encoding="utf-8").write(json.dumps({
    "images": [
        {"file": "01_mark.png", "w": 400, "h": 300, "repeats": 75, "ocr_lines": 0},
        {"file": "02_big.png", "w": 900, "h": 1200, "repeats": 1, "ocr_lines": 9},
        {"file": "03_mid.png", "w": 400, "h": 500, "repeats": 1, "ocr_lines": 0},
        {"file": "04_gone.png", "w": 800, "h": 800, "repeats": 1, "ocr_lines": 0},
    ]}))

imgs = doc_convert.reference_images(md_path)
names = [i["file"] for i in imgs]
ok("01_mark.png" not in names,
   "page furniture is never attached - a watermark costs money and teaches nothing",
   names)
ok(names[0] == "02_big.png",
   "the largest image goes first, so a truncated set keeps the most informative one",
   names)
ok("04_gone.png" not in names, "a manifest entry with no file on disk is skipped")
ok(all(i["media_type"] == "image/jpeg" for i in imgs),
   "everything is re-encoded to one format the API accepts")

from PIL import Image                                          # noqa: E402
import base64                                                  # noqa: E402
first = Image.open(io.BytesIO(base64.b64decode(imgs[0]["data"])))
ok(max(first.size) <= doc_convert.PROMPT_IMAGE_MAX_EDGE,
   "an oversized image is downscaled before sending, rather than billed at full size",
   str(first.size))
ok(all(len(base64.b64decode(i["data"])) <= doc_convert.PROMPT_IMAGE_MAX_BYTES for i in imgs),
   "and every attachment is inside the byte budget")

lim = doc_convert.reference_images(md_path, limit=1)
ok(len(lim) == 1, "the count is bounded - every image is billed on every variation")
ok(doc_convert.reference_images(md_path, limit=0) == [],
   "and a limit of zero attaches nothing")
ok(doc_convert.reference_images(os.path.join(tempfile.mkdtemp(), "none.md")) == [],
   "a brief with no converted images asks for nothing")


# ---- the Anthropic request shape -------------------------------------------
# Verified against the documented format: image blocks with a base64 source,
# placed BEFORE the text block, so the model sees the pictures and then the
# instruction about them.
class _FakeStream:
    def __init__(self, captured):
        self.captured = captured

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def get_final_message(self):
        class _M:
            content = [type("B", (), {"type": "text", "text": '{"ok": true}'})()]
        return _M()


captured = {}


class _FakeMessages:
    def stream(self, **kw):
        captured.update(kw)
        return _FakeStream(captured)


class _FakeClient:
    messages = _FakeMessages()


ig._client = lambda key: _FakeClient()
ig._extract_json = lambda resp: {"ok": True}

ig._call_claude("k", "sys", "the brief text", {"type": "object"}, 1000, images=imgs)
content = captured["messages"][0]["content"]
ok(isinstance(content, list), "with images, the user content is a block list", type(content))
ok(sum(1 for b in content if b["type"] == "image") == len(imgs),
   "one image block per attachment")
ok(content[-1]["type"] == "text" and content[-1]["text"] == "the brief text",
   "the text block comes LAST, after the images")
src = content[0]["source"]
ok(src["type"] == "base64" and src["media_type"] == "image/jpeg" and src["data"],
   "each image carries a base64 source with its media type", src.get("type"))

captured.clear()
ig._call_claude("k", "sys", "just words", {"type": "object"}, 1000)
ok(captured["messages"][0]["content"] == "just words",
   "without images the message stays a plain string - the theme generator's path "
   "is untouched by any of this")


# ---- the Ollama request shape ----------------------------------------------
# Ollama takes base64 images on the message itself. The configured fallback is
# usually a text model, which ignores them rather than failing.
sent = {}


class _Resp:
    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def read(self):
        return json.dumps({"message": {"content": '{"ok": true}'}}).encode()


def _fake_urlopen(req, timeout=0):
    sent["body"] = json.loads(req.data.decode("utf-8"))
    return _Resp()


import urllib.request                                          # noqa: E402
urllib.request.urlopen = _fake_urlopen

ig._call_ollama("http://x", "m", "sys", "words", {"type": "object"}, images=imgs)
user = [m for m in sent["body"]["messages"] if m["role"] == "user"][0]
ok(len(user.get("images") or []) == len(imgs),
   "Ollama gets the images on the user message, where its API expects them")
ok(user["content"] == "words", "and the text is unchanged beside them")

sent.clear()
ig._call_ollama("http://x", "m", "sys", "words", {"type": "object"})
user = [m for m in sent["body"]["messages"] if m["role"] == "user"][0]
ok("images" not in user, "and no images key at all when there are none to send")

print("")
if fails:
    print("%d PROMPT IMAGE CHECK(S) FAILED" % len(fails))
    sys.exit(1)
print("PROMPT IMAGES OK")
