"""C gate (Koordinator 4.2, 2026-11-10): batch sources must be in OneDrive Pictures, pinned and local.
Pictures is a temp folder here; attrib +P works on any NTFS folder, so pinning is exercised for real."""
import json
import os
import shutil
import subprocess
import tempfile
import threading
import time
from http.server import HTTPServer

import pytest
import requests

os.environ.setdefault("QA_HUB_REVIEWS_DIR", tempfile.mkdtemp(prefix="qa_hub_reviews_storage_"))
import hatshop_server  # noqa: E402

PORT = 18097
BASE_URL = f"http://127.0.0.1:{PORT}"
PINNED = 0x00080000


def attrs(p):
    return os.stat(p).st_file_attributes


@pytest.fixture()
def env(tmp_path, monkeypatch):
    pics = tmp_path / "Pictures"
    outside = tmp_path / "Elsewhere"
    pics.mkdir()
    outside.mkdir()
    monkeypatch.setattr(hatshop_server, "PICTURES_ROOT", str(pics))
    monkeypatch.setattr(hatshop_server, "BATCHES_FILE", str(tmp_path / "batches_config.json"))
    monkeypatch.setattr(hatshop_server, "load_calendar_projects", lambda: [])
    monkeypatch.setenv("QA_HUB_USERS_FILE", str(tmp_path / "no-users.json"))
    hatshop_server._STORAGE_WARNINGS.clear()
    yield pics, outside
    hatshop_server._STORAGE_WARNINGS.clear()
    subprocess.run(["attrib", "-P", str(tmp_path / "*"), "/S", "/D"], capture_output=True)


def make_batch(root, name, files=("P1/a.jpg",)):
    d = root / name
    for f in files:
        (d / f).parent.mkdir(parents=True, exist_ok=True)
        (d / f).write_bytes(b"x")
    return d


def test_in_pictures(env):
    pics, outside = env
    assert hatshop_server.in_pictures(str(pics / "b"))
    assert hatshop_server.in_pictures(str(pics))
    assert not hatshop_server.in_pictures(str(outside / "b"))
    assert not hatshop_server.in_pictures(str(pics) + "-x\\b")  # prefix trick is not "inside"


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
    assert hatshop_server.pin_batch_folder(str(b)) == ""
    assert hatshop_server.storage_issues(str(b)) is None
    assert attrs(b) & PINNED and attrs(b / "P1" / "a.jpg") & PINNED and attrs(pics / "b Removed" / "P1" / "c.jpg") & PINNED


def test_pin_refuses_outside_and_low_disk(env, monkeypatch):
    pics, outside = env
    assert hatshop_server.pin_batch_folder(str(make_batch(outside, "o"))) == "utanför Pictures"
    b = make_batch(pics, "b")
    monkeypatch.setattr(hatshop_server, "_STORAGE_MIN_FREE_GB", 10 ** 9)
    assert "fäster inte" in hatshop_server.pin_batch_folder(str(b))
    assert not attrs(b) & PINNED


def test_scan_pins_inside_and_warns_outside(env, capsys):
    pics, outside = env
    b = make_batch(pics, "b")
    o = make_batch(outside, "o")
    with open(hatshop_server.BATCHES_FILE, "w", encoding="utf-8") as f:
        json.dump([{"id": "in", "path": str(b)}, {"id": "out", "path": str(o)}], f)
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


def test_add_outside_pictures_refused(server):
    pics, outside = server
    o = make_batch(outside, "o")
    r = requests.post(f"{BASE_URL}/api/batches/add", json={"path": str(o), "name": "o"})
    assert r.status_code == 400 and "Pictures" in r.json()["message"]
    assert not os.path.exists(hatshop_server.BATCHES_FILE) or all(
        b.get("path") != str(o) for b in json.load(open(hatshop_server.BATCHES_FILE, encoding="utf-8")))


def test_add_inside_pictures_pins_and_lists_warning_for_admin(server):
    pics, outside = server
    with open(hatshop_server.BATCHES_FILE, "w", encoding="utf-8") as f:
        json.dump([{"id": "old-out", "name": "old", "path": str(make_batch(outside, "old")), "published": True}], f)
    b = make_batch(pics, "26-11-10-new")
    r = requests.post(f"{BASE_URL}/api/batches/add", json={"path": str(b), "name": "26-11-10-new"})
    assert r.status_code == 200 and r.json()["status"] == "ok"
    assert "storage_warning" not in r.json()["batch"]
    assert attrs(b) & PINNED and attrs(b / "P1" / "a.jpg") & PINNED
    hatshop_server.storage_scan_once(pin=False)
    listed = {x["id"]: x for x in requests.get(f"{BASE_URL}/api/batches?all=1").json()["batches"]}
    assert listed["old-out"]["storage_warning"]["outside_pictures"] is True
    assert "storage_warning" not in listed[r.json()["batch"]["id"]]
    tunnel = requests.get(f"{BASE_URL}/api/batches", headers={"CF-Connecting-IP": "1.1.1.1"})
    if tunnel.status_code == 200:  # non-admins never see the warning
        assert all("storage_warning" not in x for x in tunnel.json().get("batches", []))
