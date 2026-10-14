import os
import json
import socket
import threading
import time
import pytest
import requests
from http.server import HTTPServer

import hatshop_server
from hatshop_server import (
    removed_root_for,
    remove_shot,
    undo_remove,
    set_remove_reason,
    REMOVE_REASONS,
    HubHTTPRequestHandler,
)


def get_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def test_remove_shot_and_set_reason(tmp_path):
    """
    1. remove a shot, then set_remove_reason(batch, op_id, "fel_farg", "för blå", "Andro")
       -> the log op has reason "fel_farg", note "för blå", reason_by "Andro", reason_ts.
    """
    base_dir = tmp_path / "batch_1"
    base_dir.mkdir()

    p_dir = base_dir / "1500px" / "Test-Cap-Blue"
    p_dir.mkdir(parents=True)
    (p_dir / "01.jpg").write_bytes(b"image_blue_content")

    batch = {
        "id": "test-batch-1",
        "path": str(base_dir),
    }

    res_remove = remove_shot(batch, "Test-Cap-Blue", "01.jpg", "Andro")
    assert res_remove["status"] == "ok"
    op_id = res_remove["op_id"]

    res_reason = set_remove_reason(batch, op_id, "fel_farg", "för blå", "Andro")
    assert res_reason["status"] == "ok"

    log_path = os.path.join(removed_root_for(base_dir), "_qa_removed_log.json")
    assert os.path.isfile(log_path)
    with open(log_path, "r", encoding="utf-8") as f:
        log_data = json.load(f)

    assert len(log_data) == 1
    op = log_data[0]
    assert op["op_id"] == op_id
    assert op["reason"] == "fel_farg"
    assert op["note"] == "för blå"
    assert op["reason_by"] == "Andro"
    assert "reason_ts" in op and isinstance(op["reason_ts"], str)


def test_unknown_reason_and_unknown_op_id_raise_value_error(tmp_path):
    """
    2. unknown reason -> ValueError; unknown op_id -> ValueError.
    """
    base_dir = tmp_path / "batch_2"
    base_dir.mkdir()

    p_dir = base_dir / "1500px" / "Test-Cap-Red"
    p_dir.mkdir(parents=True)
    (p_dir / "01.jpg").write_bytes(b"image_red_content")

    batch = {
        "id": "test-batch-2",
        "path": str(base_dir),
    }

    res_remove = remove_shot(batch, "Test-Cap-Red", "01.jpg", "Reviewer")
    op_id = res_remove["op_id"]

    # Unknown reason -> ValueError("Okänd orsak")
    with pytest.raises(ValueError, match="Okänd orsak"):
        set_remove_reason(batch, op_id, "invalid_reason_key", "note", "Reviewer")

    # Unknown op_id -> ValueError("Åtgärden hittades inte")
    with pytest.raises(ValueError, match="Åtgärden hittades inte"):
        set_remove_reason(batch, "non-existent-op-id", "fel_farg", "note", "Reviewer")

    # Empty reason is accepted
    res_empty = set_remove_reason(batch, op_id, "", "bara en notering", "Reviewer")
    assert res_empty["status"] == "ok"


def test_note_longer_than_300_chars_is_cut_to_300(tmp_path):
    """
    3. note longer than 300 chars is cut to 300.
    """
    base_dir = tmp_path / "batch_3"
    base_dir.mkdir()

    p_dir = base_dir / "1500px" / "Test-Cap-Long"
    p_dir.mkdir(parents=True)
    (p_dir / "01.jpg").write_bytes(b"image_long_content")

    batch = {
        "id": "test-batch-3",
        "path": str(base_dir),
    }

    res_remove = remove_shot(batch, "Test-Cap-Long", "01.jpg", "UserLong")
    op_id = res_remove["op_id"]

    long_note = "A" * 350
    res_reason = set_remove_reason(batch, op_id, "styling", long_note, "UserLong")
    assert res_reason["status"] == "ok"

    log_path = os.path.join(removed_root_for(base_dir), "_qa_removed_log.json")
    with open(log_path, "r", encoding="utf-8") as f:
        log_data = json.load(f)

    op = log_data[0]
    assert len(op["note"]) == 300
    assert op["note"] == "A" * 300


def test_undo_after_setting_reason_still_works(tmp_path):
    """
    4. undo after setting a reason still works (existing undo_remove).
    """
    base_dir = tmp_path / "batch_4"
    base_dir.mkdir()

    p_dir = base_dir / "1500px" / "Test-Cap-Undo"
    p_dir.mkdir(parents=True)
    img_file = p_dir / "01.jpg"
    img_file.write_bytes(b"undo_test_content")

    batch = {
        "id": "test-batch-4",
        "path": str(base_dir),
    }

    res_remove = remove_shot(batch, "Test-Cap-Undo", "01.jpg", "UserUndo")
    op_id = res_remove["op_id"]
    assert not img_file.exists()

    # Tag reason
    set_remove_reason(batch, op_id, "rummet", "för mörk miljö", "UserUndo")

    # Undo
    res_undo = undo_remove(batch)
    assert res_undo["status"] == "ok"
    assert res_undo["op_id"] == op_id
    assert res_undo["undo_count"] == 0
    assert img_file.is_file()
    assert img_file.read_bytes() == b"undo_test_content"

    # Verify log still has reason info and undone is True
    log_path = os.path.join(removed_root_for(base_dir), "_qa_removed_log.json")
    with open(log_path, "r", encoding="utf-8") as f:
        log_data = json.load(f)
    assert log_data[0]["undone"] is True
    assert log_data[0]["reason"] == "rummet"
    assert log_data[0]["note"] == "för mörk miljö"


def test_http_remove_reason_routes(tmp_path, monkeypatch):
    """
    Test HTTP endpoints:
    - GET /api/image/remove_reasons (no admin needed, returns REMOVE_REASONS)
    - POST /api/image/remove_reason?batch=<id> (admin vs non-admin, 404, 400, 200)
    """
    base_dir = tmp_path / "batch_http_reason"
    base_dir.mkdir()

    p_dir = base_dir / "1500px" / "Test-Cap-Http"
    p_dir.mkdir(parents=True)
    (p_dir / "01.jpg").write_bytes(b"http_img_content")

    test_batch = {
        "id": "http-reason-batch",
        "name": "HTTP Reason Batch",
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
        tunnel_header = {"CF-Connecting-IP": "203.0.113.195"}  # Non-admin tunnel header

        # 1. GET /api/image/remove_reasons (no admin needed)
        r_reasons = requests.get(f"{base_url}/api/image/remove_reasons", headers=tunnel_header)
        assert r_reasons.status_code == 200
        reasons_data = r_reasons.json()
        assert reasons_data == REMOVE_REASONS
        assert "anatomi" in reasons_data
        assert "fel_farg" in reasons_data

        # Remove image first to get an op_id
        r_remove = requests.post(
            f"{base_url}/api/image/remove?batch=http-reason-batch",
            json={"product_id": "Test-Cap-Http", "filename": "01.jpg", "reviewer": "AdminUser"}
        )
        assert r_remove.status_code == 200
        op_id = r_remove.json()["op_id"]

        # 2. Non-admin POST /api/image/remove_reason -> 403
        r_forbidden = requests.post(
            f"{base_url}/api/image/remove_reason?batch=http-reason-batch",
            json={"op_id": op_id, "reason": "fransar", "note": "trasig", "reviewer": "ReviewerUser"},
            headers=tunnel_header
        )
        assert r_forbidden.status_code == 403

        # 3. Unknown batch -> 404
        r_notfound = requests.post(
            f"{base_url}/api/image/remove_reason?batch=non-existent-batch",
            json={"op_id": op_id, "reason": "fransar", "reviewer": "AdminUser"}
        )
        assert r_notfound.status_code == 404

        # 4. Unknown reason -> 400
        r_bad_reason = requests.post(
            f"{base_url}/api/image/remove_reason?batch=http-reason-batch",
            json={"op_id": op_id, "reason": "not_valid", "reviewer": "AdminUser"}
        )
        assert r_bad_reason.status_code == 400
        assert r_bad_reason.json()["message"] == "Okänd orsak"

        # 5. Success -> 200
        r_ok = requests.post(
            f"{base_url}/api/image/remove_reason?batch=http-reason-batch",
            json={"op_id": op_id, "reason": "fransar", "note": "fransig kant", "reviewer": "AdminUser"}
        )
        assert r_ok.status_code == 200
        assert r_ok.json()["status"] == "ok"

    finally:
        server.shutdown()
        server.server_close()
