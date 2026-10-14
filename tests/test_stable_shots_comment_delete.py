import os
import json
import socket
import threading
import time
import pytest
import requests
from http.server import HTTPServer
from PIL import Image

os.environ["QA_HUB_NOTIFY_SCRIPT"] = ""

import hatshop_server
hatshop_server.NOTIFY_SCRIPT = ""

from hatshop_server import (
    scan_batch_products,
    remove_shot,
    undo_remove,
    save_reviews,
    load_reviews,
    HubHTTPRequestHandler,
)


@pytest.fixture(autouse=True)
def disable_notify(monkeypatch):
    monkeypatch.setenv("QA_HUB_NOTIFY_SCRIPT", "")
    monkeypatch.setattr(hatshop_server, "NOTIFY_SCRIPT", "")


def make_tiny_png(path):
    os.makedirs(os.path.dirname(str(path)), exist_ok=True)
    img = Image.new("RGB", (8, 8), color=(120, 140, 160))
    img.save(str(path), "PNG")


def get_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _snapshot_shots(batch, product_id=None):
    products = scan_batch_products(batch)
    assert len(products) >= 1
    if product_id is not None:
        prod = next(p for p in products if p["id"] == product_id)
    else:
        assert len(products) == 1
        prod = products[0]
    shots = prod["variants"]["Standard"]
    return prod["id"], {s["filename"]: (s["shot"], s["base_shot"]) for s in shots}


def _assert_remove_middle_and_undo_keeps_stable_ids(batch, middle_filename, product_id=None):
    pid, initial_map = _snapshot_shots(batch, product_id=product_id)
    assert len(initial_map) == 3
    assert middle_filename in initial_map

    # Remove the middle file
    res_rm = remove_shot(batch, pid, middle_filename, "TestReviewer")
    assert res_rm["status"] == "ok"

    # Rescan after remove: remaining files must keep identical (shot, base_shot)
    _, after_remove_map = _snapshot_shots(batch, product_id=pid)
    expected_remaining = {
        fn: val for fn, val in initial_map.items() if fn != middle_filename
    }
    assert after_remove_map == expected_remaining

    # Undo remove and rescan: all values must equal the original
    res_undo = undo_remove(batch, op_id=res_rm["op_id"])
    assert res_undo["status"] == "ok"

    _, after_undo_map = _snapshot_shots(batch, product_id=pid)
    assert after_undo_map == initial_map


def test_stable_shots_layout_a_flat_folder(tmp_path):
    """(a) A flat folder with keps-a-black-01.png, -02, -03."""
    base_dir = tmp_path / "batch_layout_a"
    for fn in ["keps-a-black-01.png", "keps-a-black-02.png", "keps-a-black-03.png"]:
        make_tiny_png(base_dir / fn)

    batch = {"id": "batch-a", "path": str(base_dir)}
    _assert_remove_middle_and_undo_keeps_stable_ids(batch, "keps-a-black-02.png")


def test_stable_shots_layout_b_flow_names(tmp_path):
    """(b) A product folder with Flow names tc-x-LS01.png, tc-x-LS02.png, tc-x-LS03.png."""
    base_dir = tmp_path / "batch_layout_b"
    p_dir = base_dir / "tc-x"
    for fn in ["tc-x-LS01.png", "tc-x-LS02.png", "tc-x-LS03.png"]:
        make_tiny_png(p_dir / fn)

    batch = {"id": "batch-b", "path": str(base_dir)}
    _assert_remove_middle_and_undo_keeps_stable_ids(batch, "tc-x-LS02.png")


def test_stable_shots_layout_c_upload_order(tmp_path):
    """(c) A product folder with upload_order.txt listing 3 files."""
    base_dir = tmp_path / "batch_layout_c"
    p_dir = base_dir / "ProdUpload"
    files = ["first_img.png", "second_img.png", "third_img.png"]
    for fn in files:
        make_tiny_png(p_dir / fn)

    upload_order_lines = [
        "ORDNING | FILNAMN | ORIGINAL | TYP",
        "---",
        "01 | first_img.png | orig1.png | GALLERI",
        "02 | second_img.png | orig2.png | GALLERI",
        "03 | third_img.png | orig3.png | GALLERI",
    ]
    (p_dir / "upload_order.txt").write_text("\n".join(upload_order_lines), encoding="utf-8")

    batch = {"id": "batch-c", "path": str(base_dir)}
    _assert_remove_middle_and_undo_keeps_stable_ids(batch, "second_img.png")


def test_stable_shots_layout_d_no_digits(tmp_path):
    """(d) A product folder with 3 files whose names have NO digits (alpha.png, beta.png, gamma.png)."""
    base_dir = tmp_path / "batch_layout_d"
    p_dir = base_dir / "ProdNoDigits"
    for fn in ["alpha.png", "beta.png", "gamma.png"]:
        make_tiny_png(p_dir / fn)

    batch = {"id": "batch-d", "path": str(base_dir)}
    _assert_remove_middle_and_undo_keeps_stable_ids(batch, "beta.png")


def test_stable_shots_interior_prefix_no_digits(tmp_path):
    """Product folder with AI interior/lifestyle prefix files without shot numbers."""
    base_dir = tmp_path / "batch_layout_interior"
    p_dir = base_dir / "ProdInterior"
    for fn in ["interior_alpha.png", "interior_beta.png", "interior_gamma.png"]:
        make_tiny_png(p_dir / fn)

    batch = {"id": "batch-interior", "path": str(base_dir)}
    _assert_remove_middle_and_undo_keeps_stable_ids(batch, "interior_beta.png")


def test_comment_delete_by_typed_reviewer_http(tmp_path, monkeypatch):
    """
    Comment delete via HTTP harness:
    - author "Lisa", reviewer "lisa" -> deleted (200);
    - reviewer "Per" -> 403 and the comment remains;
    - empty reviewer -> 403 and the comment remains.
    """
    base_dir = tmp_path / "batch_comments"
    p_dir = base_dir / "ProdA"
    make_tiny_png(p_dir / "ProdA-01.png")

    reviews_dir = tmp_path / "reviews_store"
    reviews_dir.mkdir(parents=True, exist_ok=True)
    monkeypatch.setattr(hatshop_server, "REVIEWS_DIR", str(reviews_dir))

    test_batch = {
        "id": "comment-test-batch",
        "name": "Comment Test Batch",
        "path": str(base_dir),
        "org": "cama",
        "published": True,
    }
    monkeypatch.setattr(hatshop_server, "load_batches", lambda: [test_batch])

    # Seed a comment by author "Lisa"
    initial_reviews = {
        "ProdA": {
            "status": "pending",
            "flags": {},
            "comments": [
                {
                    "id": "c_lisa_1",
                    "author": "Lisa",
                    "text": "Snygg bild!",
                    "shots": ["01"],
                    "time": "12:00",
                }
            ],
        }
    }
    save_reviews("comment-test-batch", initial_reviews)

    port = get_free_port()
    base_url = f"http://127.0.0.1:{port}"
    server = HTTPServer(("127.0.0.1", port), HubHTTPRequestHandler)
    thread = threading.Thread(target=server.serve_forever)
    thread.daemon = True
    thread.start()
    time.sleep(0.15)

    try:
        tunnel_header = {"CF-Connecting-IP": "203.0.113.195"}

        # 1. reviewer "Per" -> 403 and the comment remains
        r_per = requests.post(
            f"{base_url}/api/comment/delete?batch=comment-test-batch",
            json={"product_id": "ProdA", "comment_id": "c_lisa_1", "reviewer": "Per"},
            headers=tunnel_header,
        )
        assert r_per.status_code == 403
        assert r_per.json()["error"] == "Du kan bara ta bort dina egna kommentarer."
        revs_after_per = load_reviews("comment-test-batch")
        assert len(revs_after_per["ProdA"]["comments"]) == 1
        assert revs_after_per["ProdA"]["comments"][0]["id"] == "c_lisa_1"

        # 2. empty reviewer -> 403 and the comment remains
        r_empty = requests.post(
            f"{base_url}/api/comment/delete?batch=comment-test-batch",
            json={"product_id": "ProdA", "comment_id": "c_lisa_1", "reviewer": ""},
            headers=tunnel_header,
        )
        assert r_empty.status_code == 403
        assert r_empty.json()["error"] == "Du kan bara ta bort dina egna kommentarer."
        revs_after_empty = load_reviews("comment-test-batch")
        assert len(revs_after_empty["ProdA"]["comments"]) == 1

        # 3. author "Lisa", reviewer "lisa" (case-insensitive) -> 200 and deleted
        r_lisa = requests.post(
            f"{base_url}/api/comment/delete?batch=comment-test-batch",
            json={"product_id": "ProdA", "comment_id": "c_lisa_1", "reviewer": "lisa"},
            headers=tunnel_header,
        )
        assert r_lisa.status_code == 200
        assert r_lisa.json()["status"] == "ok"
        revs_after_lisa = load_reviews("comment-test-batch")
        assert revs_after_lisa["ProdA"]["comments"] == []
    finally:
        server.shutdown()
        server.server_close()
