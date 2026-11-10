"""C gate (Koordinator 4.2/4.3, 2026-11-10): batch sources must be strict subfolders of OneDrive Pictures (real path, no
junction/symlink on the way), every file pinned (P) and local. Pictures is a temp folder here; +P works on any NTFS folder,
so pinning is exercised for real.

The server under test is loaded from a FILE, never by bare import name, so `python -I -m pytest tests/test_storage_gate.py`
from the hub folder tests the staged server deterministically: QA_HUB_SERVER_DIR, else _staging/ when it holds a
hatshop_server.py, else the hub folder (live)."""
import collections
import ctypes
import hashlib
import importlib.util
import io
import json
import os
import subprocess
import sys
import tempfile
import threading
import time
from http.server import HTTPServer

import pytest
import requests

HUB_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SERVER_DIR = os.environ.get("QA_HUB_SERVER_DIR") or (
    os.path.join(HUB_DIR, "_staging") if os.path.isfile(os.path.join(HUB_DIR, "_staging", "hatshop_server.py")) else HUB_DIR)
os.environ.setdefault("QA_HUB_REVIEWS_DIR", tempfile.mkdtemp(prefix="qa_hub_reviews_storage_"))
_spec = importlib.util.spec_from_file_location("hatshop_server_under_test", os.path.join(SERVER_DIR, "hatshop_server.py"))
hatshop_server = importlib.util.module_from_spec(_spec)
sys.modules[_spec.name] = hatshop_server
_spec.loader.exec_module(hatshop_server)

PORT = 18097
BASE_URL = f"http://127.0.0.1:{PORT}"
PINNED = 0x00080000
GB = 10 ** 9
Usage = collections.namedtuple("Usage", "total used free")


def attrs(p):
    return os.stat(p).st_file_attributes


def junction(link, target):
    r = subprocess.run(["cmd", "/c", "mklink", "/J", str(link), str(target)], capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr


def test_server_under_test_is_the_selected_file():
    assert os.path.normcase(os.path.dirname(os.path.abspath(hatshop_server.__file__))) == os.path.normcase(os.path.abspath(SERVER_DIR))


@pytest.fixture()
def env(tmp_path, monkeypatch):
    pics = tmp_path / "Pictures"
    outside = tmp_path / "Elsewhere"
    pics.mkdir()
    outside.mkdir()
    monkeypatch.setattr(hatshop_server, "PICTURES_ROOT", str(pics))
    monkeypatch.setattr(hatshop_server, "BATCHES_FILE", str(tmp_path / "batches_config.json"))
    monkeypatch.setattr(hatshop_server, "load_calendar_projects", lambda: [])
    monkeypatch.setattr(hatshop_server, "_users_cache", None)
    monkeypatch.setenv("QA_HUB_USERS_FILE", str(tmp_path / "no-users.json"))
    hatshop_server._STORAGE_WARNINGS.clear()
    yield pics, outside
    hatshop_server._STORAGE_WARNINGS.clear()


def make_batch(root, name, files=("P1/a.jpg",)):
    d = root / name
    d.mkdir(parents=True, exist_ok=True)
    for f in files:
        (d / f).parent.mkdir(parents=True, exist_ok=True)
        (d / f).write_bytes(b"x")
    return d


def test_in_pictures_path_forms(env):
    pics, outside = env
    b = make_batch(pics, "a-long-folder-name-for-8dot3")
    (pics.parent / "Pictures2" / "b").mkdir(parents=True)
    ok = hatshop_server.in_pictures
    assert ok(str(b)) and ok(str(pics / "not-yet-created"))
    assert ok(str(b).upper())  # case
    assert ok(str(b) + "\\") and ok(str(b).replace("\\", "/"))  # trailing slash, forward slashes
    assert ok(str(pics / "x" / ".." / b.name))
    assert not ok(str(pics))  # the Pictures root itself is never a batch
    assert not ok(str(pics) + "\\")
    assert not ok(str(b / ".." / ".." / "Elsewhere"))  # .. traversal out
    assert not ok(str(pics.parent / "Pictures2" / "b"))  # sibling sharing the prefix
    assert not ok(str(pics) + "-x\\b")
    assert not ok(str(outside / "b"))
    assert not ok("")
    buf = ctypes.create_unicode_buffer(1024)
    n = ctypes.windll.kernel32.GetShortPathNameW(os.path.realpath(b), buf, 1024)
    if not n or buf.value == os.path.realpath(b):
        pytest.skip("no 8.3 short names on this volume")
    assert "~" in buf.value and ok(buf.value)  # 8.3 short name resolves to the same folder


def test_junction_inside_pictures_is_outside_and_never_walked(env):
    pics, outside = env
    secret = make_batch(outside, "secret", ("s.jpg",))
    junction(pics / "jx", outside)
    assert not hatshop_server.in_pictures(str(pics / "jx"))
    assert not hatshop_server.in_pictures(str(pics / "jx" / "secret"))
    assert hatshop_server.storage_issues(str(pics / "jx" / "secret"))["outside_pictures"] is True
    assert hatshop_server.pin_batch_folder(str(pics / "jx" / "secret")) == "utanför Pictures"
    b = make_batch(pics, "b")
    junction(b / "link", secret)  # a junction INSIDE a real batch folder
    issue = hatshop_server.storage_issues(str(b))
    assert issue["links"] == 1 and "genvägar" in issue["message"]
    assert issue["unpinned"] == 3  # b, b/P1, b/P1/a.jpg: nothing behind the junction is counted
    assert hatshop_server.pin_batch_folder(str(b)) == ""
    assert attrs(b / "P1" / "a.jpg") & PINNED
    assert not attrs(secret / "s.jpg") & PINNED and not attrs(secret) & PINNED  # walk did not follow the junction
    assert hatshop_server.storage_issues(str(b))["links"] == 1  # still warned


def test_junction_to_pictures_root_is_refused(env):
    """Blind gate v2 R1: the link check ran after the "reached Pictures" exit, so a junction resolving to the root passed."""
    pics, outside = env
    make_batch(pics, "b")
    junction(pics / "jr", pics)  # inside Pictures -> Pictures root
    junction(outside / "jp", pics)  # outside Pictures -> Pictures root
    assert not hatshop_server.in_pictures(str(pics / "jr"))
    assert not hatshop_server.in_pictures(str(pics / "jr" / "b"))
    assert not hatshop_server.in_pictures(str(outside / "jp" / "b"))
    assert hatshop_server.in_pictures(str(pics / "b"))


def test_removed_sibling_junction_is_not_walked_or_pinned(env):
    """Blind gate v2 R2: "<batch> Removed" as a junction to an outside folder got +P on the outside file."""
    pics, outside = env
    secret = make_batch(outside, "secret", ("s.jpg",))
    b = make_batch(pics, "b")
    junction(pics / "b Removed", secret)
    issue = hatshop_server.storage_issues(str(b))
    assert issue["links"] == 1 and "genvägar" in issue["message"]
    assert hatshop_server.pin_batch_folder(str(b)) == ""
    assert attrs(b / "P1" / "a.jpg") & PINNED
    assert not attrs(secret / "s.jpg") & PINNED and not attrs(secret) & PINNED
    assert hatshop_server.storage_issues(str(b))["links"] == 1


def test_sharepoint_free_refusal_by_real_path_and_links(env):
    """Blind gate v2 R8: sharepoint_sync free compared paths as text; a junction/8.3 alias could reach a hub source."""
    spec = importlib.util.spec_from_file_location("sharepoint_sync_under_test", os.path.join(HUB_DIR, "scripts", "sharepoint_sync.py"))
    sync = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(sync)
    pics, outside = env
    src = make_batch(pics, "26-11-10-src")
    lib = outside / "lib"
    plain = make_batch(lib, "plain")
    batches = [{"id": "src", "path": str(src)}, {"id": "empty", "path": ""}]
    assert sync.free_refusal(str(plain), batches) == ""
    junction(lib / "alias", src)  # the library folder IS a hub source through a junction
    assert "junction" in sync.free_refusal(str(lib / "alias"), batches)
    assert "junction" in sync.free_refusal(str(lib), batches)  # contains one: attrib /S could follow it
    assert "overlaps hub source src" in sync.free_refusal(str(src), batches)
    assert "overlaps hub source src" in sync.free_refusal(str(pics), batches)  # above a source
    (pics / "26-11-10-src Removed").mkdir()
    assert "overlaps hub source src" in sync.free_refusal(str(pics / "26-11-10-src Removed"), batches)
    buf = ctypes.create_unicode_buffer(1024)
    if ctypes.windll.kernel32.GetShortPathNameW(os.path.realpath(src), buf, 1024) and buf.value != os.path.realpath(src):
        assert "overlaps hub source src" in sync.free_refusal(buf.value, batches)  # 8.3 alias of the source


def test_storage_issues_outside_and_unpinned_then_pinned(env):
    pics, outside = env
    out = make_batch(outside, "o")
    assert hatshop_server.storage_issues(str(out))["outside_pictures"] is True
    b = make_batch(pics, "b", ("P1/a.jpg", "P2/b.jpg", "desktop.ini"))
    (pics / "b Removed" / "P1").mkdir(parents=True)
    (pics / "b Removed" / "P1" / "c.jpg").write_bytes(b"x")
    issue = hatshop_server.storage_issues(str(b))
    assert issue and not issue["outside_pictures"] and issue["cloud_only"] == 0
    assert issue["unpinned"] == 3 + 3 + 2  # 3 folders + 2 files (desktop.ini ignored) + Removed: 2 folders + 1 file
    assert issue["images"] == 3 and issue["images_unpinned"] == 3
    assert hatshop_server.pin_batch_folder(str(b)) == ""
    assert hatshop_server.storage_issues(str(b)) is None
    assert attrs(b) & PINNED and attrs(b / "P1" / "a.jpg") & PINNED and attrs(pics / "b Removed" / "P1" / "c.jpg") & PINNED


def test_long_path_file_without_p_is_flagged_and_pinned(env):
    """26-10-09-Charrone-x (2026-11-10): attrib.exe skipped >= 260-char paths and exited 0, leaving 36 images 0x420
    (neither P nor U). The gate must count them as not pinned and pin_batch_folder must pin them."""
    pics, _outside = env
    b = pics / "b"
    deep = "\\\\?\\" + str(b / ("d" * 120) / ("f" * 120 + ".jpg"))
    os.makedirs(os.path.dirname(deep))
    with open(deep, "wb") as f:
        f.write(b"x")
    assert len(deep) - 4 >= 260 and not attrs(deep) & PINNED and not attrs(deep) & 0x00100000
    issue = hatshop_server.storage_issues(str(b))
    assert issue["images"] == 1 and issue["images_unpinned"] == 1 and "1 av 1 bildfiler" in issue["message"]
    assert hatshop_server.pin_batch_folder(str(b)) == ""
    assert attrs(deep) & PINNED and hatshop_server.storage_issues(str(b)) is None


def test_pin_refuses_outside_root_and_low_disk(env, monkeypatch):
    pics, outside = env
    assert hatshop_server.pin_batch_folder(str(make_batch(outside, "o"))) == "utanför Pictures"
    assert hatshop_server.pin_batch_folder(str(pics)) == "utanför Pictures"
    b = make_batch(pics, "b")
    monkeypatch.setattr(hatshop_server, "_STORAGE_MIN_FREE_GB", 10 ** 9)
    assert "fäster inte" in hatshop_server.pin_batch_folder(str(b))
    assert not attrs(b) & PINNED


def test_disk_guard_counts_bytes_to_download(env, monkeypatch):
    pics, _outside = env
    b = make_batch(pics, "b")
    monkeypatch.setattr(hatshop_server.shutil, "disk_usage", lambda _p: Usage(100 * GB, 70 * GB, 30 * GB))
    monkeypatch.setattr(hatshop_server, "_hydration_bytes", lambda _p: 15 * GB)
    err = hatshop_server.pin_batch_folder(str(b))
    assert "15.0 GB ska laddas ned" in err and "fäster inte" in err and not attrs(b) & PINNED  # 30 - 15 < 20
    monkeypatch.setattr(hatshop_server, "_hydration_bytes", lambda _p: 5 * GB)
    assert "redan köat" in hatshop_server.pin_batch_folder(str(b), reserved_bytes=6 * GB)  # 30 - 6 - 5 < 20
    assert hatshop_server.pin_batch_folder(str(b)) == "" and attrs(b) & PINNED  # 30 - 5 >= 20


def write_batches(*pairs):
    with open(hatshop_server.BATCHES_FILE, "w", encoding="utf-8") as f:
        json.dump([{"id": i, "path": str(p)} for i, p in pairs], f)


def test_scan_reserves_queued_bytes_between_batches(env, monkeypatch):
    pics, _outside = env
    a, b = make_batch(pics, "a"), make_batch(pics, "b")
    write_batches(("a", a), ("b", b))
    monkeypatch.setattr(hatshop_server.shutil, "disk_usage", lambda _p: Usage(100 * GB, 65 * GB, 35 * GB))
    monkeypatch.setattr(hatshop_server, "_hydration_bytes", lambda _p: 8 * GB)
    monkeypatch.setattr(hatshop_server, "_STORAGE_MAX_SCAN_BYTES", 100 * GB)  # only the disk guard decides here
    hatshop_server.storage_scan_once()
    assert attrs(a / "P1" / "a.jpg") & PINNED and "a" not in hatshop_server._STORAGE_WARNINGS  # 35 - 0 - 8 >= 20
    assert not attrs(b / "P1" / "a.jpg") & PINNED  # 35 - 8 - 8 < 20
    assert "redan köat" in hatshop_server._STORAGE_WARNINGS["b"]["pin_error"]


def test_scan_caps_bytes_per_scan(env, monkeypatch):
    pics, _outside = env
    a, b = make_batch(pics, "a"), make_batch(pics, "b")
    write_batches(("a", a), ("b", b))
    monkeypatch.setattr(hatshop_server.shutil, "disk_usage", lambda _p: Usage(2000 * GB, 0, 1000 * GB))
    monkeypatch.setattr(hatshop_server, "_hydration_bytes", lambda _p: 8 * GB)
    monkeypatch.setattr(hatshop_server, "_STORAGE_MAX_SCAN_BYTES", 10 * GB)
    hatshop_server.storage_scan_once()
    assert attrs(a / "P1" / "a.jpg") & PINNED and not attrs(b / "P1" / "a.jpg") & PINNED
    assert "nästa skanning" in hatshop_server._STORAGE_WARNINGS["b"]["pin_error"]


def test_scan_pins_cloud_only_even_when_nothing_is_unpinned(env, monkeypatch):
    pics, _outside = env
    a = make_batch(pics, "a")
    write_batches(("a", a))
    calls = []
    real_issues = hatshop_server.storage_issues
    monkeypatch.setattr(hatshop_server, "storage_issues", lambda p: (
        {"outside_pictures": False, "cloud_only": 4, "unpinned": 0, "message": "4 fil(er) bara i molnet"} if not calls else real_issues(p)))
    monkeypatch.setattr(hatshop_server, "pin_batch_folder", lambda p, reserved_bytes=0: calls.append(p) or "")
    hatshop_server.storage_scan_once()
    assert calls == [str(a)]


def test_scan_pins_inside_and_warns_outside(env, capsys):
    pics, outside = env
    b = make_batch(pics, "b")
    o = make_batch(outside, "o")
    write_batches(("in", b), ("out", o))
    hatshop_server.storage_scan_once()
    assert "in" not in hatshop_server._STORAGE_WARNINGS
    assert hatshop_server._STORAGE_WARNINGS["out"]["outside_pictures"] is True
    assert attrs(b / "P1" / "a.jpg") & PINNED
    log = capsys.readouterr().out
    assert "[storage-gate] out:" in log
    hatshop_server.storage_scan_once()
    assert "[storage-gate]" not in capsys.readouterr().out  # no repeat line while the state is unchanged


@pytest.fixture()
def server(env):
    httpd = HTTPServer(("127.0.0.1", PORT), hatshop_server.HubHTTPRequestHandler)
    t = threading.Thread(target=httpd.serve_forever, daemon=True)
    t.start()
    time.sleep(0.3)
    yield env
    httpd.shutdown()
    httpd.server_close()


def add(path, name):
    return requests.post(f"{BASE_URL}/api/batches/add", json={"path": str(path), "name": name})


def saved_paths():
    if not os.path.exists(hatshop_server.BATCHES_FILE):
        return []
    return [b.get("path") for b in json.load(open(hatshop_server.BATCHES_FILE, encoding="utf-8"))]


def test_add_outside_root_and_junction_refused(server):
    pics, outside = server
    o = make_batch(outside, "o")
    junction(pics / "jx", outside)
    refused = (o, pics, pics / "jx" / "o")
    for p in refused:
        r = add(p, "o")
        assert r.status_code == 400 and "Pictures" in r.json()["message"], p
    assert not {str(p) for p in refused} & set(saved_paths())  # (a missing config is seeded with a default batch)


def test_add_pin_failure_is_ok_with_warning(server, monkeypatch):
    pics, _outside = server
    b = make_batch(pics, "26-11-10-boom")

    def boom(_p, reserved_bytes=0):
        raise OSError("disk borta")
    monkeypatch.setattr(hatshop_server, "pin_batch_folder", boom)
    r = add(b, "26-11-10-boom")
    assert r.status_code == 200 and r.json()["status"] == "ok"
    assert "disk borta" in r.json()["batch"]["storage_warning"]["pin_error"]
    assert str(b) in saved_paths()


def cp1252_stdout(monkeypatch):
    out = io.TextIOWrapper(io.BytesIO(), encoding="cp1252")  # like qa_watchdog.ps1's redirect without PYTHONUTF8
    monkeypatch.setattr(sys, "stdout", out)
    return out


def test_add_pin_error_outside_cp1252_is_ok_with_warning(server, monkeypatch):
    """Blind gate v3 R5: the [storage-gate] print after save_batches raised UnicodeEncodeError -> HTTP 400."""
    pics, _outside = server
    b = make_batch(pics, "26-11-10-emoji")

    def boom(_p, reserved_bytes=0):
        raise OSError("disk → borta \U0001F4A5")
    monkeypatch.setattr(hatshop_server, "pin_batch_folder", boom)
    out = cp1252_stdout(monkeypatch)
    r = add(b, "26-11-10-emoji")
    assert r.status_code == 200 and r.json()["status"] == "ok"
    assert "\U0001F4A5" in r.json()["batch"]["storage_warning"]["pin_error"]
    out.flush()
    assert b"[storage-gate] added" in out.buffer.getvalue()


def test_scan_continues_after_unencodable_log_line(env, monkeypatch):
    pics, _outside = env
    a, b = make_batch(pics, "a"), make_batch(pics, "b")
    write_batches(("a", a), ("b", b))
    calls = []
    monkeypatch.setattr(hatshop_server, "pin_batch_folder",
                        lambda p, reserved_bytes=0: calls.append(p) or ("fel → \U0001F4A5" if p == str(a) else ""))
    out = cp1252_stdout(monkeypatch)
    hatshop_server.storage_scan_once()
    assert calls == [str(a), str(b)]  # the first batch's log line did not abort the scan
    assert "\U0001F4A5" in hatshop_server._STORAGE_WARNINGS["a"]["pin_error"]
    out.flush()
    assert b"[storage-gate] a:" in out.buffer.getvalue() and b"[storage-gate] b:" in out.buffer.getvalue()


def test_add_inside_pictures_pins_and_warning_is_admin_only(server, tmp_path):
    pics, outside = server
    with open(hatshop_server.BATCHES_FILE, "w", encoding="utf-8") as f:
        json.dump([{"id": "old-out", "name": "old", "path": str(make_batch(outside, "old")), "org": "cama",
                    "published": True}], f)
    b = make_batch(pics, "26-11-10-new")
    r = add(b, "26-11-10-new")
    assert r.status_code == 200 and r.json()["status"] == "ok"
    assert "storage_warning" not in r.json()["batch"]
    assert attrs(b) & PINNED and attrs(b / "P1" / "a.jpg") & PINNED
    hatshop_server.storage_scan_once(pin=False)
    listed = {x["id"]: x for x in requests.get(f"{BASE_URL}/api/batches?all=1").json()["batches"]}
    assert listed["old-out"]["storage_warning"]["outside_pictures"] is True
    assert "storage_warning" not in listed[r.json()["batch"]["id"]]
    # logged-in non-admin via the tunnel: sees the published batch, never the warning
    with open(os.environ["QA_HUB_USERS_FILE"], "w", encoding="utf-8") as f:
        json.dump({"enforce": True, "users": [{"name": "Tittare", "role": "viewer", "org": "cama",
                                                "token_sha256": hashlib.sha256(b"tok-viewer").hexdigest()}]}, f)
    hatshop_server._users_cache = None
    tunnel = requests.get(f"{BASE_URL}/api/batches", headers={"CF-Connecting-IP": "1.1.1.1", "Cookie": "qa_session=tok-viewer"})
    assert tunnel.status_code == 200
    seen = tunnel.json()["batches"]
    assert [x["id"] for x in seen] == ["old-out"]
    assert all("storage_warning" not in x for x in seen)
