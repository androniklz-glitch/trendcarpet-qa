"""Manifest batches (products_manifest.json): a product whose files were all moved out must not be listed.

Loads the server from QA_HUB_SERVER_DIR, else _staging/ when it holds a hatshop_server.py, else the hub root
(same rule as test_storage_gate.py). Run from the hub: `python -I -B -m pytest tests/test_manifest_scan.py`.
"""
import importlib.util
import json
import os

import pytest

HUB_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SERVER_DIR = os.environ.get("QA_HUB_SERVER_DIR") or (
    os.path.join(HUB_DIR, "_staging") if os.path.isfile(os.path.join(HUB_DIR, "_staging", "hatshop_server.py")) else HUB_DIR
)
_spec = importlib.util.spec_from_file_location("hatshop_server_manifest_test", os.path.join(SERVER_DIR, "hatshop_server.py"))
hatshop_server = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(hatshop_server)


def _touch(path, data=b"x"):
    with open(path, "wb") as f:
        f.write(data)


@pytest.fixture
def manifest_batch(tmp_path):
    batch_dir = tmp_path / "test-video-batch"
    batch_dir.mkdir()
    manifest = [
        {"id": "01_full", "title": "01 full", "category": "Test",
         "images": [{"filename": "01_full_15s.mp4", "kind": "Video"}, {"filename": "01_full-01.jpg", "kind": "Studio"}]},
        {"id": "02_moved_out", "title": "02 moved out", "category": "Test",
         "images": [{"filename": "02_moved_out_15s.mp4"}, {"filename": "02_moved_out-01.jpg"}]},
        {"id": "03_partial", "title": "03 partial", "category": "Test",
         "images": [{"filename": "03_partial_15s.mp4"}, {"filename": "03_partial-01.jpg"}]},
    ]
    (batch_dir / "products_manifest.json").write_text(json.dumps(manifest), encoding="utf-8")
    for fn in ("01_full_15s.mp4", "01_full-01.jpg", "03_partial-01.jpg"):
        _touch(batch_dir / fn)
    return {"id": "test-video-batch-manifest-scan", "path": str(batch_dir), "category": "mattor_inredning", "org": "cama"}


def test_manifest_product_with_no_files_is_not_listed(manifest_batch):
    ids = [p["id"] for p in hatshop_server.scan_batch_products(manifest_batch)]
    assert ids == ["01_full", "03_partial"]


def test_manifest_keeps_file_order_video_first(manifest_batch):
    prods = {p["id"]: p for p in hatshop_server.scan_batch_products(manifest_batch)}
    shots = [s["filename"] for v in prods["01_full"]["variants"].values() for s in v]
    assert shots == ["01_full_15s.mp4", "01_full-01.jpg"]
    partial = [s["filename"] for v in prods["03_partial"]["variants"].values() for s in v]
    assert partial == ["03_partial-01.jpg"]


def test_manifest_product_returns_when_file_is_back(manifest_batch):
    _touch(os.path.join(manifest_batch["path"], "02_moved_out-01.jpg"))
    ids = [p["id"] for p in hatshop_server.scan_batch_products(manifest_batch)]
    assert ids == ["01_full", "02_moved_out", "03_partial"]
