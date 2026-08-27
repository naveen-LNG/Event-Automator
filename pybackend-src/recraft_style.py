"""Create and store a Recraft custom style from the shipped Sort floors.

Why this exists. Style transfer by attaching reference images is what produced
the reference-bleed failure: with five Sort floors attached, NanoBananaAIO
composited the cat room into a glaze room instead of imitating it (see
Sort_RENDER_STYLE.md 5.9). Recraft attacks the same problem from the other end -
you upload the references ONCE, Recraft returns a style_id, and every subsequent
generation cites that id with NO images attached. There is nothing to bleed from,
so the failure mode stops being something to gate for and becomes impossible.

Two hard limits from the node's own description: at most 5 reference images, and
5 MB total across them. The five shipped floors are 8.27 MB, so they are
downscaled before upload - style capture wants palette, finish and texture, not
resolution.

Graph: LoadImage xN -> RecraftCreateStyleNode -> PreviewAny.
PreviewAny is needed because RecraftCreateStyleNode is not an output node, and
ComfyUI only publishes output nodes' results to /history - without a sink the
style_id is computed and thrown away.

The style_id is the deliverable. It is written to
ideationData/StyleReference/recraft_style.json and is reusable indefinitely:
feed it to RecraftStyleV3InfiniteStyleLibrary to get a RECRAFT_V3_STYLE, which
plugs into RecraftTextToImageNode.recraft_style.
"""

import json
import os
import time
import urllib.error
import uuid

import comfy_client as cc

MAX_REFS = 5
MAX_TOTAL_BYTES = 5 * 1024 * 1024
SHRINK_MAX_DIM = 800          # keeps five PNGs comfortably inside 5 MB
BASE_STYLE = "digital_illustration"   # the other option, realistic_image, is wrong for us

STYLE_FILE = "recraft_style.json"

# ComfyUI's V3 autogrow inputs flatten to per-slot sockets in the API format,
# namespaced by the autogrow field and 0-based: "images.image0", "images.image1".
# Nothing in /object_info says so - the schema only exposes the template and its
# "image" prefix. It was recovered from the validation error, which names the
# socket it wanted in extra_info.input_name.
#
# The rest are kept as fallbacks in case a ComfyUI upgrade changes the flattening.
# Probing costs nothing: a name mismatch fails validation, which happens before
# execution, so a wrong guess never reaches Recraft and never spends credits.
SOCKET_PATTERNS = ["images.image%d:0-based", "image%d", "image_%d",
                   "image%d:0-based", "images-as-list"]


def _shrink(src_paths, out_dir, max_dim=SHRINK_MAX_DIM):
    """Downscaled copies that fit Recraft's 5 MB total budget."""
    from PIL import Image
    os.makedirs(out_dir, exist_ok=True)
    out = []
    for p in src_paths:
        im = Image.open(p).convert("RGB")   # flatten: alpha is not part of the style
        w, h = im.size
        s = min(1.0, float(max_dim) / max(w, h))
        if s < 1.0:
            im = im.resize((max(1, int(w * s)), max(1, int(h * s))), Image.LANCZOS)
        d = os.path.join(out_dir, "style_" + os.path.basename(p))
        im.save(d, "PNG", optimize=True)
        out.append(d)
    return out


def _build(ref_names, pattern, base_style=BASE_STYLE):
    wf = {}
    inputs = {"style": base_style}
    for i, ref in enumerate(ref_names):
        nid = str(200 + i)
        wf[nid] = {"class_type": "LoadImage", "inputs": {"image": ref, "upload": "image"}}
        if pattern == "images-as-list":
            inputs.setdefault("images", []).append([nid, 0])
        elif pattern.endswith(":0-based"):
            inputs[pattern[:-len(":0-based")] % i] = [nid, 0]
        else:
            inputs[pattern % (i + 1)] = [nid, 0]
    wf["1"] = {"class_type": "RecraftCreateStyleNode", "inputs": inputs}
    wf["2"] = {"class_type": "PreviewAny", "inputs": {"source": ["1", 0]}}
    return wf


def _queue(base_url, wf):
    """Returns (prompt_id, None) or (None, error_text)."""
    try:
        res = cc._post_json(base_url, "/prompt",
                            {"prompt": wf, "client_id": str(uuid.uuid4())}, timeout=60)
    except urllib.error.HTTPError as e:
        return None, e.read().decode("utf-8", "replace")[:2000]
    except Exception as e:
        return None, "%s: %s" % (e.__class__.__name__, e)
    return res.get("prompt_id"), None


def _harvest(base_url, pid, timeout_s=240):
    """Poll /history for the style_id PreviewAny printed."""
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        time.sleep(cc.POLL_INTERVAL_S)
        try:
            rec = (cc._get(base_url, "/history/%s" % pid, timeout=30) or {}).get(pid)
        except Exception:
            continue
        if not rec:
            continue
        for msg in (rec.get("status") or {}).get("messages") or []:
            if msg and msg[0] == "execution_error":
                d = msg[1] if len(msg) > 1 else {}
                return None, "%s: %s" % (d.get("node_type", "?"),
                                         d.get("exception_message", ""))
        outs = rec.get("outputs") or {}
        for node_out in outs.values():
            # PreviewAny publishes under "text"; be liberal about the key since
            # this is the one part of the contract not visible in /object_info.
            for key in ("text", "string", "value", "result", "style_id"):
                v = node_out.get(key)
                if v:
                    s = v[0] if isinstance(v, (list, tuple)) else v
                    return str(s).strip(), None
        if outs:
            return None, "finished but published no text output: %s" % json.dumps(outs)[:400]
        if (rec.get("status") or {}).get("status_str") == "error":
            return None, "error with no detail"
    return None, "timed out after %ds" % timeout_s


def create(base_url, ref_paths, style_dir, base_style=BASE_STYLE, log_fn=None):
    """Upload references, create the style, persist the id. Returns a dict."""
    log = log_fn or (lambda m: None)
    refs = list(ref_paths)[:MAX_REFS]
    if not refs:
        return {"ok": False, "error": "no reference images given"}

    # A temp dir, not style_dir: the downscaled copies are throwaway derivatives
    # and style_dir lives in the SVN/OneDrive tree.
    import shutil
    import tempfile
    work = tempfile.mkdtemp(prefix="recraft_style_")
    try:
        small = _shrink(refs, work)
        total = sum(os.path.getsize(p) for p in small)
        log("references: %d, downscaled to %.2f MB (cap %.0f MB)"
            % (len(small), total / 1048576.0, MAX_TOTAL_BYTES / 1048576.0))
        if total > MAX_TOTAL_BYTES:
            return {"ok": False, "error": "references still %.2f MB after downscaling; "
                                          "lower SHRINK_MAX_DIM" % (total / 1048576.0)}
        names = [cc.upload_image(base_url, p) for p in small]
    finally:
        shutil.rmtree(work, ignore_errors=True)
    log("uploaded: " + ", ".join(names))

    last = ""
    for pat in SOCKET_PATTERNS:
        pid, err = _queue(base_url, _build(names, pat, base_style))
        if pid:
            log("socket naming '%s' accepted (prompt %s)" % (pat, pid))
            sid, herr = _harvest(base_url, pid)
            if sid:
                rec = {"style_id": sid, "base_style": base_style,
                       "sources": [os.path.basename(p) for p in refs],
                       "shrunk_to_px": SHRINK_MAX_DIM,
                       "socket_pattern": pat,
                       "created_utc": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())}
                dest = os.path.join(style_dir, STYLE_FILE)
                with open(dest, "w", encoding="utf-8") as f:
                    json.dump(rec, f, indent=2)
                    f.write("\n")
                rec["ok"] = True
                rec["saved_to"] = dest
                return rec
            if "unauthorized" in (herr or "").lower() or "login" in (herr or "").lower():
                return {"ok": False, "unauthorized": True, "error":
                        "The graph is correct and the node executed, but this ComfyUI host is "
                        "not signed in to comfy.org, so every api_node class is refused "
                        "(167 of the 947 installed - all of Recraft, FLUX 2 Pro/Max, Seedream, "
                        "Grok, GPT-Image-1, Ideogram, Stability, Reve, Kling, Veo, and Comfy's "
                        "own Gemini nodes). NanoBananaAIO is unaffected because it is a "
                        "third-party node carrying its own key. To unblock: sign the host into "
                        "comfy.org (Settings -> Comfy API key) and make sure the account has "
                        "credits, then re-run. Server said: %s" % herr}
            return {"ok": False, "error": "style creation ran but returned nothing: %s" % herr}
        last = err or ""
        log("socket naming '%s' rejected: %s" % (pat, last[:220].replace("\n", " ")))
    return {"ok": False, "error": "no socket naming accepted. Last error: %s" % last[:900]}


def load(style_dir):
    """The stored style record, or None."""
    p = os.path.join(style_dir, STYLE_FILE)
    if not os.path.isfile(p):
        return None
    try:
        with open(p, encoding="utf-8") as f:
            r = json.load(f)
        return r if r.get("style_id") else None
    except Exception:
        return None


def build_workflow(prompt, style_id, negative_prompt="", size="1024x1365",
                   seed=0, filename_prefix="eventautomator/recraft"):
    """Text-to-image citing the stored style. No reference images attached.

    That absence is the point: with no images in the graph there is nothing for
    the model to composite, which is the failure this whole path exists to avoid.
    """
    wf = {
        "1": {"class_type": "RecraftStyleV3InfiniteStyleLibrary",
              "inputs": {"style_id": style_id}},
        "2": {"class_type": "RecraftTextToImageNode",
              "inputs": {"prompt": prompt, "size": size, "n": 1, "seed": int(seed),
                         "recraft_style": ["1", 0]}},
        "3": {"class_type": "SaveImage",
              "inputs": {"images": ["2", 0], "filename_prefix": filename_prefix}},
    }
    if negative_prompt:
        wf["2"]["inputs"]["negative_prompt"] = negative_prompt
    return wf
