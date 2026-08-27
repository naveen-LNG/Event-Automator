"""ComfyUI client for reference-image generation.

Implements the studio's approved Nano Banana workflow (Art Style.json) against
the ComfyUI host, e.g. https://skipbo.comfyui.liquidnitro.games.

    LoadImage(style ref) -> NanoBananaAIO.image_1..6 -> ImageScale -> SaveImage

Node names, option lists and defaults below are NOT guesses - they were read
from that server's live /object_info and from the approved workflow export:

    NanoBananaAIO   "Nano Banana AIO"
        model_name  : gemini-3-pro-image-preview   (only option)
        image_count : INT      use_search : BOOL
        image_1..6  : IMAGE    (reference inputs)
        aspect_ratio: 1:1 2:3 3:2 3:4 4:3 4:5 5:4 9:16 16:9 21:9 Auto
        image_size  : 1K 2K 4K            temperature: FLOAT
    ImageScale      upscale_method: nearest-exact bilinear area bicubic lanczos
    SaveImage       filename_prefix

The approved workflow's settings are the defaults here (use_search False,
temperature 0.4, image_size 2K, aspect Auto, lanczos 1115x994).

NanoBananaAIO forwards the prompt and the attached reference images to Google
via the ComfyUI host, so anything in a prompt leaves our network through it.

Stdlib only (urllib) - the bundled python has no requests.
"""

import json
import mimetypes
import os
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid

MODEL_NAME = "gemini-3-pro-image-preview"       # NanoBananaAIO's only option
NODE = "NanoBananaAIO"

# ROOM_RENDER_STYLE.md section 8: "A generation that returns ~5 KB is a blank;
# retry it. Guard on output size." A real 2K room render is hundreds of KB.
MIN_IMAGE_BYTES = 20_000
BLANK_RETRIES = 2

POLL_INTERVAL_S = 2.0
DEFAULT_TIMEOUT_S = 420          # 2K Nano Banana renders are not fast


class ComfyError(RuntimeError):
    """Message is intended to be shown to the designer as-is."""


def _url(base, path):
    return base.rstrip("/") + path


def _get(base, path, timeout=30):
    req = urllib.request.Request(_url(base, path), method="GET")
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def _post_json(base, path, payload, timeout=60):
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(_url(base, path), data=data,
                                 headers={"Content-Type": "application/json"},
                                 method="POST")
    with urllib.request.urlopen(req, timeout=timeout) as r:
        body = r.read().decode("utf-8")
        return json.loads(body) if body.strip() else {}


# ---------------------------------------------------------------- health ---
def check(base_url, timeout=20):
    """Three states, because each needs a different fix:

      unreachable - wrong URL / VPN / host down    -> fix the setting
      wedged      - status endpoints answer in ms   -> the host's ComfyUI worker
                    but POST /prompt never returns     must be restarted
      ok          - queueing works

    'wedged' is not hypothetical: a ComfyUI worker whose executor has locked up
    still serves /system_stats and /object_info in ~80ms while POST /prompt and
    POST /interrupt hang forever. A plain ping calls that server healthy and
    then every generation silently times out.
    """
    info = {"reachable": False, "queue_ok": False, "node_present": False,
            "state": "unreachable", "detail": "", "version": ""}
    try:
        stats = _get(base_url, "/system_stats", timeout=timeout)
        info["reachable"] = True
        info["version"] = (stats.get("system") or {}).get("comfyui_version", "")
    except Exception as e:
        info["detail"] = f"Cannot reach ComfyUI at {base_url} ({e.__class__.__name__}: {e})"
        return info

    try:
        oi = _get(base_url, "/object_info/" + NODE, timeout=timeout)
        info["node_present"] = bool(oi.get(NODE))
    except Exception as e:
        info["state"] = "no_node"
        info["detail"] = f"Reachable, but could not read the {NODE} node ({e})"
        return info

    if not info["node_present"]:
        info["state"] = "no_node"
        info["detail"] = (f"ComfyUI is up but has no {NODE} node - the Nano Banana "
                          "custom nodes are not installed on this host.")
        return info

    # The only honest liveness test is asking it to accept work. A healthy
    # server rejects an empty prompt instantly (HTTP 400); a wedged one hangs.
    # Nothing real is queued either way.
    try:
        _post_json(base_url, "/prompt", {"prompt": {}}, timeout=timeout)
        info["queue_ok"] = True
        info["state"] = "ok"
    except urllib.error.HTTPError:
        info["queue_ok"] = True                  # 400 => handler alive
        info["state"] = "ok"
    except Exception as e:
        info["state"] = "wedged"
        info["detail"] = (
            f"ComfyUI answers status requests but POST /prompt never responds "
            f"({e.__class__.__name__}). The worker's execution queue is stuck - "
            "/prompt and /interrupt hang while /system_stats stays fast. "
            "Restarting the ComfyUI worker on the host clears it; nothing can be "
            "queued until then.")
        return info

    info["detail"] = f"ComfyUI {info['version']} ready; {NODE} available."
    return info


# ---------------------------------------------------------------- upload ---
def upload_image(base_url, local_path, timeout=180):
    """POST /upload/image as multipart, so LoadImage can reference the file.

    Hand-rolled multipart because the bundled python has no requests.
    Returns the server-side name (including any subfolder) for LoadImage.
    """
    if not os.path.isfile(local_path):
        raise ComfyError(f"Style reference image not found: {local_path}")
    name = os.path.basename(local_path)
    ctype = mimetypes.guess_type(name)[0] or "application/octet-stream"
    with open(local_path, "rb") as f:
        blob = f.read()

    boundary = "----EventAutomator" + uuid.uuid4().hex
    pre = (f"--{boundary}\r\n"
           f'Content-Disposition: form-data; name="image"; filename="{name}"\r\n'
           f"Content-Type: {ctype}\r\n\r\n").encode("utf-8")
    mid = (f"\r\n--{boundary}\r\n"
           f'Content-Disposition: form-data; name="overwrite"\r\n\r\ntrue\r\n'
           f"--{boundary}--\r\n").encode("utf-8")
    body = pre + blob + mid

    req = urllib.request.Request(
        _url(base_url, "/upload/image"), data=body,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}",
                 "Content-Length": str(len(body))}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            res = json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        raise ComfyError(f"Upload of {name} rejected (HTTP {e.code}): "
                         f"{e.read().decode('utf-8', 'replace')[:300]}")
    except Exception as e:
        raise ComfyError(f"Upload of {name} failed ({e.__class__.__name__}: {e})")

    sub = res.get("subfolder") or ""
    fn = res.get("name") or name
    return f"{sub}/{fn}" if sub else fn


# -------------------------------------------------------------- workflow ---
BASE_SCALE = (1115, 994)              # what the approved Art Style.json exported at
BASE_PIXELS = BASE_SCALE[0] * BASE_SCALE[1]


def scale_for(aspect_ratio, scale_w=None, scale_h=None):
    """Output size for this aspect, keeping the exported pixel budget.

    The workflow ends in an ImageScale, and its size used to be hardcoded to
    1115x994. That silently overrode aspect_ratio: a 3:4 request reached the
    model, came back portrait, and was then resampled to a near-square 1115x994
    - so every render in the tool has been landscape regardless of what was
    asked for. Deriving the size from the ratio is what makes the parameter mean
    anything.

    An explicit scale_w/scale_h still wins, and "Auto" (or anything unparseable)
    keeps the old size, so nothing that relied on it changes.
    """
    if scale_w and scale_h:
        return int(scale_w), int(scale_h)
    try:
        aw, ah = [float(x) for x in str(aspect_ratio).split(":")]
        if aw <= 0 or ah <= 0:
            raise ValueError(aspect_ratio)
    except (ValueError, TypeError):
        return BASE_SCALE
    ratio = aw / ah
    h = (BASE_PIXELS / ratio) ** 0.5
    return int(round(h * ratio)), int(round(h))


def build_workflow(prompt, ref_names=(), image_count=1, use_search=False,
                   aspect_ratio="Auto", image_size="2K", temperature=0.4,
                   scale_w=None, scale_h=None, filename_prefix="eventautomator/ref"):
    """API-format graph mirroring the approved Art Style.json export.

    ref_names are server-side filenames from upload_image(); each becomes a
    LoadImage feeding one of NanoBananaAIO's image_1..image_6 slots. Section 8
    of ROOM_RENDER_STYLE.md is explicit that an attached render-quality
    reference beats any prose description of finish, so this is the path that
    matters - a prompt-only call is the degraded mode, not the intended one.
    """
    ref_names = list(ref_names)[:6]
    wf = {}
    nb_inputs = {
        "model_name": MODEL_NAME,
        "prompt": prompt,
        "image_count": int(image_count),
        "use_search": bool(use_search),
        "aspect_ratio": aspect_ratio,
        "image_size": image_size,
        "temperature": float(temperature),
    }
    for i, ref in enumerate(ref_names, start=1):
        nid = str(100 + i)
        wf[nid] = {"class_type": "LoadImage",
                   "inputs": {"image": ref, "upload": "image"}}
        nb_inputs[f"image_{i}"] = [nid, 0]

    wf["1"] = {"class_type": NODE, "inputs": nb_inputs}
    out_w, out_h = scale_for(aspect_ratio, scale_w, scale_h)
    wf["2"] = {"class_type": "ImageScale",
               "inputs": {"image": ["1", 0], "upscale_method": "lanczos",
                          "width": out_w, "height": out_h,
                          "crop": "disabled"}}
    wf["3"] = {"class_type": "SaveImage",
               "inputs": {"images": ["2", 0], "filename_prefix": filename_prefix}}
    return wf


def _download(base_url, item, dest_path, timeout=180):
    q = urllib.parse.urlencode({"filename": item.get("filename", ""),
                                "subfolder": item.get("subfolder", ""),
                                "type": item.get("type", "output")})
    req = urllib.request.Request(_url(base_url, "/view?" + q), method="GET")
    with urllib.request.urlopen(req, timeout=timeout) as r:
        blob = r.read()
    os.makedirs(os.path.dirname(os.path.abspath(dest_path)), exist_ok=True)
    with open(dest_path, "wb") as f:
        f.write(blob)
    return len(blob)


def _run_once(base_url, wf, dest_path, timeout_s):
    try:
        res = _post_json(base_url, "/prompt",
                         {"prompt": wf, "client_id": str(uuid.uuid4())}, timeout=60)
    except urllib.error.HTTPError as e:
        return {"ok": False, "error": f"ComfyUI rejected the workflow (HTTP {e.code}): "
                                      f"{e.read().decode('utf-8', 'replace')[:600]}"}
    except Exception as e:
        return {"ok": False, "error":
                f"Could not queue the image ({e.__class__.__name__}: {e}). If status "
                "pages load but this keeps timing out, the ComfyUI worker's queue is "
                "wedged and the host needs a restart."}

    pid = res.get("prompt_id")
    if not pid:
        return {"ok": False, "error": f"POST /prompt returned no prompt_id: {res}"}

    deadline = time.time() + timeout_s
    while time.time() < deadline:
        time.sleep(POLL_INTERVAL_S)
        try:
            rec = _get(base_url, f"/history/{pid}", timeout=30).get(pid)
        except Exception:
            continue                                    # transient; keep polling
        if not rec:
            continue

        status = rec.get("status") or {}
        for msg in status.get("messages") or []:
            if msg and msg[0] == "execution_error":
                d = msg[1] if len(msg) > 1 else {}
                return {"ok": False, "error":
                        f"ComfyUI failed in node {d.get('node_type', '?')}: "
                        f"{d.get('exception_message', '')}".strip()}

        for node_out in (rec.get("outputs") or {}).values():
            for item in node_out.get("images") or []:
                try:
                    size = _download(base_url, item, dest_path)
                except Exception as e:
                    return {"ok": False, "error": f"Rendered but could not download: {e}"}
                return {"ok": True, "path": dest_path, "bytes": size, "prompt_id": pid}

        if status.get("status_str") == "error":
            return {"ok": False, "error": "ComfyUI reported an error with no detail."}

    return {"ok": False, "error": f"Timed out after {timeout_s}s waiting for ComfyUI."}


def generate(base_url, prompt, dest_path, ref_paths=(), timeout_s=DEFAULT_TIMEOUT_S,
             log_fn=None, **wf_kwargs):
    """Render one image and save it to dest_path.

    Never raises for an ordinary failure: images are an addition to the
    document, and a ComfyUI outage must not throw away the art requirement the
    designer just waited on. Returns {"ok": False, "error": ...} instead.
    """
    say = log_fn or (lambda m: None)

    ref_names = []
    for p in ref_paths or ():
        try:
            ref_names.append(upload_image(base_url, p))
            say(f"style reference uploaded: {os.path.basename(p)}")
        except ComfyError as e:
            return {"ok": False, "error": str(e)}

    wf = build_workflow(prompt, ref_names=ref_names, **wf_kwargs)

    last = None
    for attempt in range(1 + BLANK_RETRIES):
        last = _run_once(base_url, wf, dest_path, timeout_s)
        if not last.get("ok"):
            return last
        # Blank-output guard (ROOM_RENDER_STYLE.md section 8).
        if last["bytes"] >= MIN_IMAGE_BYTES:
            say(f"image saved: {os.path.basename(dest_path)} ({last['bytes'] // 1024} KB)")
            return last
        say(f"blank output ({last['bytes']} bytes) - retrying "
            f"{attempt + 1}/{BLANK_RETRIES}")

    return {"ok": False,
            "error": f"ComfyUI returned a blank image ({last['bytes']} bytes) "
                     f"after {BLANK_RETRIES} retries."}
