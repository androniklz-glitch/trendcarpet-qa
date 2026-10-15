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


def test_thumb_format_and_transparency(tmp_path, monkeypatch):
    base_dir = tmp_path / "batch_thumb"
    base_dir.mkdir()

    # 1. An opaque RGBA PNG (alpha 255 everywhere, 2000x2000, noisy content)
    # Generate 200x200 random RGB texture upscaled to 2000x2000 for realistic photographic noise
    # that compresses well in JPEG but poorly in PNG
    small = Image.frombytes("RGB", (200, 200), os.urandom(200 * 200 * 3))
    opaque_img = small.resize((2000, 2000), Image.Resampling.BICUBIC).convert("RGBA")
    assert opaque_img.getchannel("A").getextrema() == (255, 255)

    opaque_path = base_dir / "opaque_sample.png"
    opaque_img.save(str(opaque_path), "PNG")
    opaque_orig_bytes = opaque_path.read_bytes()

    # 2. An RGBA PNG with a transparent region
    trans_img = Image.new("RGBA", (800, 800), color=(0, 0, 0, 0))
    # Fill center with non-transparent color
    for x in range(200, 600):
        for y in range(200, 600):
            trans_img.putpixel((x, y), (255, 50, 50, 255))
    assert trans_img.getchannel("A").getextrema()[0] < 255

    trans_path = base_dir / "transparent_sample.png"
    trans_img.save(str(trans_path), "PNG")

    test_batch = {
        "id": "thumb-format-batch",
        "name": "Thumb Format Batch",
        "path": str(base_dir),
        "org": "cama",
        "published": True,
    }

    cache_dir = tmp_path / "cache"
    cache_dir.mkdir()
    monkeypatch.setattr(hatshop_server, "load_batches", lambda: [test_batch])
    monkeypatch.setattr(hatshop_server, "THUMB_CACHE_DIR", str(cache_dir))

    port = get_free_port()
    base_url = f"http://127.0.0.1:{port}"

    server = HTTPServer(("127.0.0.1", port), HubHTTPRequestHandler)
    thread = threading.Thread(target=server.serve_forever)
    thread.daemon = True
    thread.start()

    time.sleep(0.2)

    try:
        # Case 1: Opaque RGBA requested with w=600
        # -> Content-Type image/jpeg, width <= 600, body < 25% of PNG size
        r_opaque = requests.get(f"{base_url}/image/thumb-format-batch/opaque_sample.png?w=600")
        assert r_opaque.status_code == 200
        assert r_opaque.headers.get("Content-Type") == "image/jpeg"
        
        decoded_opaque = Image.open(io.BytesIO(r_opaque.content))
        assert max(decoded_opaque.size) <= 600
        assert decoded_opaque.size[0] <= 600
        
        png_buf = io.BytesIO()
        decoded_opaque.save(png_buf, "PNG")
        png_size = png_buf.tell()
        assert len(r_opaque.content) < 0.25 * png_size

        # Case 2: Transparent RGBA PNG requested with w=600
        # -> Content-Type image/webp, decoded image keeps alpha < 255 somewhere
        r_trans = requests.get(f"{base_url}/image/thumb-format-batch/transparent_sample.png?w=600")
        assert r_trans.status_code == 200
        assert r_trans.headers.get("Content-Type") == "image/webp"

        decoded_trans = Image.open(io.BytesIO(r_trans.content))
        assert decoded_trans.mode in ("RGBA", "LA") or "A" in decoded_trans.getbands()
        assert decoded_trans.convert("RGBA").getchannel("A").getextrema()[0] < 255

        # Case 3: Same opaque file with download=1&w=600
        # -> original bytes, unchanged
        r_down = requests.get(f"{base_url}/image/thumb-format-batch/opaque_sample.png?download=1&w=600")
        assert r_down.status_code == 200
        assert r_down.content == opaque_orig_bytes

    finally:
        server.shutdown()
        server.server_close()
