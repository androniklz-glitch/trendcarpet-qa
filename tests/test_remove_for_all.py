import os
import json
import socket
import threading
import time
import subprocess
import pytest
import requests
from http.server import HTTPServer

os.environ["QA_HUB_NOTIFY_SCRIPT"] = ""

import hatshop_server
hatshop_server.NOTIFY_SCRIPT = ""

from hatshop_server import (
    removed_root_for,
    remove_shot,
    undo_remove,
    list_removed,
    scan_batch_products,
    viewer_may_remove,
    LOW_SHOT_THRESHOLD,
    HubHTTPRequestHandler,
)


def get_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def test_own_undo_restores_own_file(tmp_path):
    """
    1. Users A and B each remove one shot;
       undo_remove(batch, user="A") restores A's file, not B's.
    """
    base_dir = tmp_path / "batch_ab"
    base_dir.mkdir()

    p_dir = base_dir / "1500px" / "Product-AB"
    p_dir.mkdir(parents=True)
    (p_dir / "01.jpg").write_bytes(b"content_a")
    (p_dir / "02.jpg").write_bytes(b"content_b")
    (p_dir / "03.jpg").write_bytes(b"content_c")

    batch = {
        "id": "batch-ab",
        "path": str(base_dir),
    }

    # User A removes 01.jpg
    res_a = remove_shot(batch, "Product-AB", "01.jpg", "A")
    assert res_a["status"] == "ok"
    assert not (p_dir / "01.jpg").exists()

    # User B removes 02.jpg
    res_b = remove_shot(batch, "Product-AB", "02.jpg", "B")
    assert res_b["status"] == "ok"
    assert not (p_dir / "02.jpg").exists()

    # User A calls undo_remove with user="A"
    undo_res = undo_remove(batch, user="A")
    assert undo_res["status"] == "ok"
    assert undo_res["op_id"] == res_a["op_id"]

    # Verify A's file is restored, but B's file is NOT restored
    assert (p_dir / "01.jpg").exists()
    assert (p_dir / "01.jpg").read_bytes() == b"content_a"
    assert not (p_dir / "02.jpg").exists()

    # A's undo_count should now be 0
    assert undo_res["undo_count"] == 0

    # User B calls undo_remove with user="B" -> restores B's file
    undo_res_b = undo_remove(batch, user="B")
    assert undo_res_b["status"] == "ok"
    assert undo_res_b["op_id"] == res_b["op_id"]
    assert (p_dir / "02.jpg").exists()
    assert (p_dir / "02.jpg").read_bytes() == b"content_b"


def test_low_shot_warning_and_notification(tmp_path, monkeypatch):
    """
    2. Product with 3 shots:
       remove 1 -> no low_warning;
       remove a 2nd -> low_warning.remaining == 1 and exactly one Popen call;
       remove the 3rd -> remaining == 0 and NO second Popen call (already alerted).
    """
    base_dir = tmp_path / "batch_warn"
    base_dir.mkdir()

    p_dir = base_dir / "1500px" / "Product-Warn"
    p_dir.mkdir(parents=True)
    (p_dir / "01.jpg").write_bytes(b"img_1")
    (p_dir / "02.jpg").write_bytes(b"img_2")
    (p_dir / "03.jpg").write_bytes(b"img_3")

    batch = {
        "id": "batch-warn",
        "path": str(base_dir),
    }

    # Dummy notify script
    notify_dummy = tmp_path / "mock_notify.mjs"
    notify_dummy.write_text("// dummy notify", encoding="utf-8")
    monkeypatch.setattr(hatshop_server, "NOTIFY_SCRIPT", str(notify_dummy))

    popen_calls = []

    def mock_popen(args, *a, **kw):
        popen_calls.append(args)
        class DummyProc:
            pass
        return DummyProc()

    monkeypatch.setattr(subprocess, "Popen", mock_popen)

    # 1st remove -> 2 shots remaining (not < LOW_SHOT_THRESHOLD)
    res1 = remove_shot(batch, "Product-Warn", "01.jpg", "Alice")
    assert res1["status"] == "ok"
    assert "low_warning" not in res1
    assert len(popen_calls) == 0

    # 2nd remove -> 1 shot remaining (< LOW_SHOT_THRESHOLD)
    res2 = remove_shot(batch, "Product-Warn", "02.jpg", "Alice")
    assert res2["status"] == "ok"
    assert "low_warning" in res2
    assert res2["low_warning"]["remaining"] == 1
    assert res2["low_warning"]["product_id"] == "Product-Warn"
    assert len(popen_calls) == 1
    cmd = popen_calls[0]
    assert cmd[0] == "node"
    assert cmd[1] == str(notify_dummy)
    assert "[QA-hubb 1.0] Få bilder kvar:" in cmd[2]
    assert "Product-Warn" in cmd[2]
    assert "1 bild(er) kvar" in cmd[2]
    assert "(Alice)" in cmd[2]

    # Verify _qa_low_alerts.json records the product
    removed_root = removed_root_for(base_dir)
    alerts_file = os.path.join(removed_root, "_qa_low_alerts.json")
    assert os.path.isfile(alerts_file)
    with open(alerts_file, "r", encoding="utf-8") as f:
        alerts = json.load(f)
    assert "Product-Warn" in alerts

    # 3rd remove -> 0 shots remaining (< LOW_SHOT_THRESHOLD), but already alerted -> no second Popen call
    res3 = remove_shot(batch, "Product-Warn", "03.jpg", "Alice")
    assert res3["status"] == "ok"
    assert "low_warning" in res3
    assert res3["low_warning"]["remaining"] == 0
    assert len(popen_calls) == 1  # No second call!


def test_undo_restores_and_cleans_low_alert(tmp_path, monkeypatch):
    """
    3. Undo twice (back to 2 shots) -> the alert key is gone from _qa_low_alerts.json.
    """
    base_dir = tmp_path / "batch_undo_warn"
    base_dir.mkdir()

    p_dir = base_dir / "1500px" / "Product-Clean"
    p_dir.mkdir(parents=True)
    (p_dir / "01.jpg").write_bytes(b"img_1")
    (p_dir / "02.jpg").write_bytes(b"img_2")
    (p_dir / "03.jpg").write_bytes(b"img_3")

    batch = {
        "id": "batch-undo-warn",
        "path": str(base_dir),
    }

    notify_dummy = tmp_path / "mock_notify.mjs"
    notify_dummy.write_text("// dummy", encoding="utf-8")
    monkeypatch.setattr(hatshop_server, "NOTIFY_SCRIPT", str(notify_dummy))
    monkeypatch.setattr(subprocess, "Popen", lambda *a, **kw: None)

    # Remove all 3 shots
    remove_shot(batch, "Product-Clean", "01.jpg", "Alice")
    remove_shot(batch, "Product-Clean", "02.jpg", "Alice")
    remove_shot(batch, "Product-Clean", "03.jpg", "Alice")

    removed_root = removed_root_for(base_dir)
    alerts_file = os.path.join(removed_root, "_qa_low_alerts.json")
    assert os.path.isfile(alerts_file)
    with open(alerts_file, "r", encoding="utf-8") as f:
        alerts = json.load(f)
    assert "Product-Clean" in alerts

    # Undo 1 -> 1 shot restored (1 shot remaining, which is < 2)
    undo_remove(batch)
    with open(alerts_file, "r", encoding="utf-8") as f:
        alerts = json.load(f)
    assert "Product-Clean" in alerts

    # Undo 2 -> 2 shots restored (2 shots remaining, which is >= LOW_SHOT_THRESHOLD)
    undo_remove(batch)
    with open(alerts_file, "r", encoding="utf-8") as f:
        alerts = json.load(f)
    assert "Product-Clean" not in alerts


def test_can_remove_request_logic():
    """
    4. can_remove_request logic:
       reviewer "" or "Anonym" -> False for a non-admin; "Amalia" -> True.
    """
    # Non-admin tests
    assert viewer_may_remove(is_admin=False, reviewer="") is False
    assert viewer_may_remove(is_admin=False, reviewer="   ") is False
    assert viewer_may_remove(is_admin=False, reviewer="Anonym") is False
    assert viewer_may_remove(is_admin=False, reviewer=" Anonym ") is False
    assert viewer_may_remove(is_admin=False, reviewer=None) is False
    assert viewer_may_remove(is_admin=False, reviewer="Amalia") is True

    # Admin tests: always True
    assert viewer_may_remove(is_admin=True, reviewer="") is True
    assert viewer_may_remove(is_admin=True, reviewer="Anonym") is True
    assert viewer_may_remove(is_admin=True, reviewer="Amalia") is True

    # Test via dummy handler instance with can_remove_request
    class DummyHandler:
        def __init__(self, admin):
            self.admin = admin
        def is_admin_request(self, query=None, reviewer=None):
            return self.admin
        def can_remove_request(self, query=None, reviewer=None):
            return HubHTTPRequestHandler.can_remove_request(self, query, reviewer)

    non_admin_handler = DummyHandler(admin=False)
    assert non_admin_handler.can_remove_request({}, "") is False
    assert non_admin_handler.can_remove_request({}, "Anonym") is False
    assert non_admin_handler.can_remove_request({}, "Amalia") is True

    admin_handler = DummyHandler(admin=True)
    assert admin_handler.can_remove_request({}, "") is True
    assert admin_handler.can_remove_request({}, "Anonym") is True
    assert admin_handler.can_remove_request({}, "Amalia") is True


def test_http_viewer_routes(tmp_path, monkeypatch):
    """
    Full HTTP endpoint verification for named viewers:
    - POST /api/image/remove: non-admin without name -> 403, with name -> 200
    - GET /api/batches: contains can_remove: True
    - GET /api/products: contains can_remove: True
    - GET /api/image/removed: non-admin without name -> 200 {"undo_count": 0, "ops": []}
    - GET /api/image/removed?reviewer=Amalia: 200 with Amalia's ops
    - POST /api/image/undo_remove: non-admin with name -> 200
    """
    base_dir = tmp_path / "batch_http_viewer"
    base_dir.mkdir()

    p_dir = base_dir / "1500px" / "Product-HTTP"
    p_dir.mkdir(parents=True)
    (p_dir / "01.jpg").write_bytes(b"content_1")
    (p_dir / "02.jpg").write_bytes(b"content_2")
    (p_dir / "03.jpg").write_bytes(b"content_3")

    test_batch = {
        "id": "viewer-batch",
        "name": "Viewer Batch",
        "path": str(base_dir),
        "org": "cama",
        "published": True
    }

    monkeypatch.setattr(hatshop_server, "load_batches", lambda: [test_batch])

    port = get_free_port()
    base_url = f"http://127.0.0.1:{port}"

    server = HTTPServer(("127.0.0.1", port), HubHTTPRequestHandler)
    thread = threading.Thread(target=server.serve_forever)
    thread.daemon = True
    thread.start()

    time.sleep(0.2)

    try:
        tunnel_header = {"CF-Connecting-IP": "203.0.113.195"}  # Non-admin

        # GET /api/batches -> can_remove: True
        r_batches = requests.get(f"{base_url}/api/batches", headers=tunnel_header)
        assert r_batches.status_code == 200
        assert r_batches.json().get("can_remove") is True

        # GET /api/products -> can_remove: True
        r_prods = requests.get(f"{base_url}/api/products?batch=viewer-batch", headers=tunnel_header)
        assert r_prods.status_code == 200
        assert r_prods.json().get("can_remove") is True

        # Non-admin with empty name / Anonym POST /api/image/remove -> 403
        r_403 = requests.post(
            f"{base_url}/api/image/remove?batch=viewer-batch",
            json={"product_id": "Product-HTTP", "filename": "01.jpg", "reviewer": "Anonym"},
            headers=tunnel_header
        )
        assert r_403.status_code == 403
        assert r_403.json()["message"] == "Skriv ditt namn i hubben först för att kunna flytta ut bilder."

        # Non-admin with named reviewer "Amalia" POST /api/image/remove -> 200
        r_ok = requests.post(
            f"{base_url}/api/image/remove?batch=viewer-batch",
            json={"product_id": "Product-HTTP", "filename": "01.jpg", "reviewer": "Amalia"},
            headers=tunnel_header
        )
        assert r_ok.status_code == 200
        data_ok = r_ok.json()
        assert data_ok["status"] == "ok"
        assert data_ok["undo_count"] == 1

        # GET /api/image/removed without reviewer -> 200 {"undo_count": 0, "ops": []}
        r_unnamed = requests.get(f"{base_url}/api/image/removed?batch=viewer-batch", headers=tunnel_header)
        assert r_unnamed.status_code == 200
        assert r_unnamed.json() == {"undo_count": 0, "ops": []}

        # GET /api/image/removed?reviewer=Amalia -> 200 with 1 op
        r_named = requests.get(f"{base_url}/api/image/removed?batch=viewer-batch&reviewer=Amalia", headers=tunnel_header)
        assert r_named.status_code == 200
        data_named = r_named.json()
        assert data_named["undo_count"] == 1
        assert len(data_named["ops"]) == 1

        # POST /api/image/undo_remove with reviewer="Amalia" -> 200
        r_undo = requests.post(
            f"{base_url}/api/image/undo_remove?batch=viewer-batch",
            json={"reviewer": "Amalia"},
            headers=tunnel_header
        )
        assert r_undo.status_code == 200
        assert r_undo.json()["status"] == "ok"
        assert (p_dir / "01.jpg").exists()

    finally:
        server.shutdown()
        server.server_close()
