"""Browser checks for the upload step of the proposal page (Nov 11 2026).

Run from the repo root:  python3 tests/upload_page_test.py
Needs Playwright with Chromium (a test-only dependency, not in requirements.txt). The app is
started in-process on a free port. No OpenAI calls: the extract routes are answered by the test.

Why: Dinari Hospitality, Nov 11 2026. Clicking Extract & Analyze showed only "Load failed".
Railway's log had the session created (200) and the upload cut off by Safari 50 ms in (499), at
the same byte on both tries: the end of the first PDF, where the second PDF had to be read. The
page kept a pointer to each file on disk and the browser read it at upload time, so a file that
was moved, renamed, re-saved, dragged out of an email or not downloaded from a cloud folder
killed the upload, and every retry failed the same way.

  1. a file changed or deleted on disk after it was added still uploads, byte for byte
  2. a file that cannot be read is named when it is added, and the others are kept
  3. a dropped connection gives a plain message, never the browser's "Failed to fetch" / "Load failed"
  4. one or two missed status checks do not abort a running extraction
  5. a session cleared by a restart says so instead of waiting ten minutes
  6. the same file can be picked again after it is removed
"""
import hashlib
import json
import os
import sys
import tempfile
import threading

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
UPLOADS = tempfile.mkdtemp(prefix="upload_page_test_")
os.environ["UPLOAD_DIR"] = UPLOADS
os.environ.setdefault("OPENAI_API_KEY", "test")
for _k in ("TELEGRAM_TOKEN", "TELEGRAM_BOT_TOKEN"):
    os.environ.pop(_k, None)

import web_app  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402
from werkzeug.serving import make_server  # noqa: E402

PASS = 0


def check(cond, msg):
    global PASS
    if not cond:
        raise AssertionError(msg)
    PASS += 1


srv = make_server("127.0.0.1", 0, web_app.app, threaded=True)
BASE = f"http://127.0.0.1:{srv.server_port}"
threading.Thread(target=srv.serve_forever, daemon=True).start()

FAST_POLL = "const _st = window.setTimeout; window.setTimeout = (fn, ms, ...a) => _st(fn, ms === 5000 ? 30 : ms, ...a);"
EXTRACTED = {"status": "extracted", "summary": {}, "data": {"client_info": {}, "coverages": {}, "locations": []}}
work = tempfile.mkdtemp(prefix="upload_page_files_")


def make_pdf(name, size, seed):
    path = os.path.join(work, name)
    with open(path, "wb") as f:
        f.write(b"%PDF-1.4\n" + hashlib.sha256(seed).digest() * (size // 32))
    return path


def sha(path_or_bytes):
    data = path_or_bytes if isinstance(path_or_bytes, bytes) else open(path_or_bytes, "rb").read()
    return hashlib.sha256(data).hexdigest()


def new_page(browser, statuses=None, init=""):
    """statuses: list of replies for GET /api/extract/<sid>/status, in order; the last one repeats.
    'abort' drops the request; a dict is the JSON body; (code, dict) sets the status code."""
    page = browser.new_page()
    page.add_init_script(FAST_POLL + init)
    seq = list(statuses or [EXTRACTED])

    def on_extract(route):
        if route.request.url.endswith("/status"):
            item = seq.pop(0) if len(seq) > 1 else seq[0]
            if item == "abort":
                return route.abort()
            code, body = item if isinstance(item, tuple) else (200, item)
            return route.fulfill(status=code, content_type="application/json", body=json.dumps(body))
        return route.fulfill(status=200, content_type="application/json", body='{"status": "extracting"}')
    page.route("**/api/extract/**", on_extract)
    page.goto(BASE + "/")
    return page


def toasts(page):
    return " | ".join(page.locator(".toast").all_inner_texts())


def add_files(page, paths, expect):
    """Attach real on-disk files, the way a person picking them does. Playwright's own
    set_input_files hands the page in-memory copies, which can never go stale, so it would not
    show the fault: with files attached this way the page as it was before this change fails
    check 1 with net::ERR_UPLOAD_FILE_CHANGED ("Failed to fetch", Chrome's "Load failed")."""
    before = page.evaluate("uploadedFiles.length")
    cdp = page.context.new_cdp_session(page)
    root = cdp.send("DOM.getDocument")["root"]["nodeId"]
    node = cdp.send("DOM.querySelector", {"nodeId": root, "selector": "#file-input"})["nodeId"]
    cdp.send("DOM.setFileInputFiles", {"files": list(paths), "nodeId": node})
    try:
        page.wait_for_function(f"uploadedFiles.length > {before}", timeout=1500)
    except Exception:      # this Chromium build did not fire the change event itself
        page.evaluate("document.getElementById('file-input').dispatchEvent(new Event('change', {bubbles: true}))")
    page.wait_for_function(f"uploadedFiles.length === {expect}")
    cdp.detach()


with sync_playwright() as pw:
    browser = pw.chromium.launch()

    # 1. files go stale after they are added: one deleted, one rewritten
    a = make_pdf("JASMIN QUOTE LETTER.pdf", 106_000, b"a")
    b = make_pdf("Jasmin - HotelBound Quote.pdf", 1_796_000, b"b")
    want = {os.path.basename(a): sha(a), os.path.basename(b): sha(b)}
    page = new_page(browser)
    page.fill("#client-name", "Stale File Hotel LLC")
    add_files(page, [a, b], 2)
    os.remove(a)
    with open(b, "wb") as f:
        f.write(b"re-saved after it was added")
    page.click("#btn-extract")
    page.wait_for_selector(".toast-success", timeout=20000)
    check("Extraction complete!" in toasts(page), toasts(page))
    sess = [s for s in json.load(open(os.path.join(UPLOADS, "_sessions.json"))).values()
            if s["client_name"] == "Stale File Hotel LLC"]
    check(len(sess) == 1 and len(sess[0]["files"]) == 2, sess)
    got = {f["filename"]: sha(f["path"]) for f in sess[0]["files"]}
    check(got == want, "the server received both files exactly as they were when added")
    page.close()

    # 2. a file the browser cannot read is named at once; the readable one is kept
    good = make_pdf("Good Quote.pdf", 40_000, b"g")
    bad = make_pdf("Unreadable Quote.pdf", 40_000, b"u")
    page = new_page(browser, init="""
        const _ab = File.prototype.arrayBuffer;
        File.prototype.arrayBuffer = function () {
            return this.name.startsWith('Unreadable') ? Promise.reject(new DOMException('gone', 'NotReadableError'))
                                                      : _ab.call(this);
        };""")
    page.set_input_files("#file-input", [good, bad])
    page.wait_for_selector(".toast-error")
    t = toasts(page)
    check("Could not read Unreadable Quote.pdf" in t and "Downloads or Desktop" in t, t)
    check(page.evaluate("uploadedFiles.map(f => f.name)") == ["Good Quote.pdf"], "only the readable file is listed")
    check(page.locator(".file-item").count() == 1 and not page.is_disabled("#btn-extract"), "list and button")
    page.close()

    # 3. the upload cannot reach the server -> plain words, and the files stay for a retry
    page = new_page(browser)
    page.route("**/api/upload/**", lambda route: route.abort())
    page.fill("#client-name", "No Connection Hotel LLC")
    add_files(page, [good], 1)
    page.click("#btn-extract")
    page.wait_for_selector(".toast-error")
    t = toasts(page)
    check("The upload did not reach the proposal builder" in t, t)
    check("Failed to fetch" not in t and "Load failed" not in t, t)
    check(page.evaluate("uploadedFiles.length") == 1, "files kept for the retry")
    page.unroute("**/api/upload/**")
    page.click("#btn-extract")
    page.wait_for_selector(".toast-success", timeout=20000)
    check("Extraction complete!" in toasts(page), "the retry goes through without re-adding the files")
    page.close()

    # 4. two missed status checks, then the extraction finishes
    page = new_page(browser, statuses=["abort", "abort", {"status": "extracting", "progress": "Reading"}, EXTRACTED])
    page.fill("#client-name", "Blip Hotel LLC")
    add_files(page, [good], 1)
    page.click("#btn-extract")
    page.wait_for_selector(".toast-success", timeout=20000)
    check("Extraction complete!" in toasts(page) and page.locator(".toast-error").count() == 0, toasts(page))
    page.close()

    # 4b. the connection stays down -> plain words after six misses
    page = new_page(browser, statuses=["abort"])
    page.fill("#client-name", "Down Hotel LLC")
    add_files(page, [good], 1)
    page.click("#btn-extract")
    page.wait_for_selector(".toast-error", timeout=20000)
    t = toasts(page)
    check("Lost the connection while the quotes were being read" in t and "Failed to fetch" not in t, t)
    page.close()

    # 5. the session is gone (the app restarted) -> said so on the first check
    page = new_page(browser, statuses=[(404, {"error": "Session not found"})])
    page.fill("#client-name", "Restart Hotel LLC")
    add_files(page, [good], 1)
    page.click("#btn-extract")
    page.wait_for_selector(".toast-error", timeout=20000)
    check("restarted and cleared this session" in toasts(page), toasts(page))
    page.close()

    # 6. remove a file, pick the same one again
    page = new_page(browser)
    add_files(page, [good], 1)
    check(page.evaluate("document.getElementById('file-input').value") == "", "input cleared after adding")
    page.evaluate("removeFile(0)")
    check(page.evaluate("uploadedFiles.length") == 0 and page.is_disabled("#btn-extract"), "removed")
    add_files(page, [good], 1)
    check(page.locator(".file-item").count() == 1, "same file added again")
    page.close()

    browser.close()

srv.shutdown()
print(f"upload_page_test: {PASS} checks passed")
