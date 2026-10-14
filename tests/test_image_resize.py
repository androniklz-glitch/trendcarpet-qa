import io
import os
import socket
import threading
import time
import pytest
import requests
from http.server import HTTPServer
from PIL import Image

import hatshop_server
from hatshop_server import HubHTTPRequestHandler


def get_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def test_image_resize(tmp_path, monkeypatch):
    base_dir = tmp_path / "batch_resize"
    base_dir.mkdir()

    # Create a 2000x1500 JPEG image
    img_path = base_dir / "sample.jpg"
    orig_img = Image.new("RGB", (2000, 1500), color=(120, 140, 160))
    orig_img.save(str(img_path), "JPEG", quality=90)
    orig_bytes = img_path.read_bytes()

    test_batch = {
        "id": "resize-test-batch",
        "name": "Resize Test Batch",
        "path": str(base_dir),
        "org": "cama",
        "published": True,
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
        # 1. w=1200: longest side is 1200
        r_1200 = requests.get(f"{base_url}/image/resize-test-batch/sample.jpg?w=1200")
        assert r_1200.status_code == 200
        img_1200 = Image.open(io.BytesIO(r_1200.content))
        assert max(img_1200.size) == 1200
        assert img_1200.size == (1200, 900)

        # 2. w=600: longest side is 600
        r_600 = requests.get(f"{base_url}/image/resize-test-batch/sample.jpg?w=600")
        assert r_600.status_code == 200
        img_600 = Image.open(io.BytesIO(r_600.content))
        assert max(img_600.size) == 600
        assert img_600.size == (600, 450)

        # 3. w=5000: out-of-range value falls back to 160
        r_5000 = requests.get(f"{base_url}/image/resize-test-batch/sample.jpg?w=5000")
        assert r_5000.status_code == 200
        img_5000 = Image.open(io.BytesIO(r_5000.content))
        assert max(img_5000.size) == 160
        assert img_5000.size == (160, 120)

        # 4. No w: serves the original bytes
        r_orig = requests.get(f"{base_url}/image/resize-test-batch/sample.jpg")
        assert r_orig.status_code == 200
        assert r_orig.content == orig_bytes
        img_orig = Image.open(io.BytesIO(r_orig.content))
        assert img_orig.size == (2000, 1500)

    finally:
        server.shutdown()
        server.server_close()
