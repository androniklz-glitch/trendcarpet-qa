import io
import json
import os
import socket
import threading
import time
import urllib.parse
from http.server import HTTPServer

import pytest
import requests
from PIL import Image

import hatshop_server
from hatshop_server import (
    HubHTTPRequestHandler,
    get_ratings,
    load_ratings,
    set_rating,
)


def get_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def test_ratings_reviewer_isolation(tmp_path, monkeypatch):
    """
    1. A sets 4 stars on x, B sets 2 on x -> get_ratings(batch, "A") has only A's 4; B sees only 2.
    """
    monkeypatch.setattr(hatshop_server, "REVIEWS_DIR", str(tmp_path))

    batch = {
        "id": "test-batch-isolation",
        "path": str(tmp_path),
    }

    # A sets 4 stars on x (product_id="prod_x", filename="01.jpg")
    set_rating(batch, "A", "prod_x", "01.jpg", 4)
    # B sets 2 stars on x
    set_rating(batch, "B", "prod_x", "01.jpg", 2)

    ratings_a = get_ratings(batch, "A")
    ratings_b = get_ratings(batch, "B")

    key = "prod_x|01.jpg"

    # get_ratings(batch, "A") has only A's 4
    assert key in ratings_a
    assert ratings_a[key]["stars"] == 4
    assert len(ratings_a) == 1

    # B sees only 2
    assert key in ratings_b
    assert ratings_b[key]["stars"] == 2
    assert len(ratings_b) == 1

    # Unknown reviewer sees empty dict
    assert get_ratings(batch, "C") == {}
    # Anonymous reviewer sees empty dict
    assert get_ratings(batch, "Anonym") == {}
    assert get_ratings(batch, "") == {}

    # Check file on disk
    ratings_file = tmp_path / "ratings_test-batch-isolation.json"
    assert ratings_file.exists()
    disk_data = json.loads(ratings_file.read_text(encoding="utf-8"))
    assert "A" in disk_data and disk_data["A"][key]["stars"] == 4
    assert "B" in disk_data and disk_data["B"][key]["stars"] == 2


def test_ratings_stars_zero_removes_and_invalid_raises(tmp_path, monkeypatch):
    """
    2. stars 0 removes the entry; stars 6 -> ValueError.
    """
    monkeypatch.setattr(hatshop_server, "REVIEWS_DIR", str(tmp_path))

    batch = {
        "id": "test-batch-removal",
        "path": str(tmp_path),
    }

    key = "prod_x|01.jpg"

    # Set rating 5
    set_rating(batch, "Alice", "prod_x", "01.jpg", 5)
    assert get_ratings(batch, "Alice")[key]["stars"] == 5

    # stars 0 removes the entry
    set_rating(batch, "Alice", "prod_x", "01.jpg", 0)
    assert key not in get_ratings(batch, "Alice")
    assert get_ratings(batch, "Alice") == {}

    # Removing non-existent key does not error
    set_rating(batch, "Alice", "prod_x", "02.jpg", 0)

    # stars 6 -> ValueError("Ogiltigt betyg")
    with pytest.raises(ValueError, match="Ogiltigt betyg"):
        set_rating(batch, "Alice", "prod_x", "01.jpg", 6)

    # Negative stars -> ValueError
    with pytest.raises(ValueError, match="Ogiltigt betyg"):
        set_rating(batch, "Alice", "prod_x", "01.jpg", -1)

    # Boolean -> ValueError
    with pytest.raises(ValueError, match="Ogiltigt betyg"):
        set_rating(batch, "Alice", "prod_x", "01.jpg", True)

    # String non-integer -> ValueError
    with pytest.raises(ValueError, match="Ogiltigt betyg"):
        set_rating(batch, "Alice", "prod_x", "01.jpg", "invalid")


def test_http_image_download_original(tmp_path, monkeypatch):
    """
    3. GET /image/...&download=1&w=600 returns the original bytes and a Content-Disposition header
       containing attachment.
    """
    base_dir = tmp_path / "batch_download"
    base_dir.mkdir()

    # Create a 2000x1500 JPEG image
    img_name = "test original photo.jpg"
    img_path = base_dir / img_name
    orig_img = Image.new("RGB", (2000, 1500), color=(100, 150, 200))
    orig_img.save(str(img_path), "JPEG", quality=95)
    orig_bytes = img_path.read_bytes()

    test_batch = {
        "id": "download-test-batch",
        "name": "Download Test Batch",
        "path": str(base_dir),
        "org": "cama",
        "published": True,
    }

    monkeypatch.setattr(hatshop_server, "load_batches", lambda: [test_batch])
    monkeypatch.setattr(hatshop_server, "REVIEWS_DIR", str(tmp_path))

    port = get_free_port()
    base_url = f"http://127.0.0.1:{port}"

    server = HTTPServer(("127.0.0.1", port), HubHTTPRequestHandler)
    thread = threading.Thread(target=server.serve_forever)
    thread.daemon = True
    thread.start()

    time.sleep(0.2)

    try:
        url_quoted_fn = urllib.parse.quote(img_name)

        # 1. GET with download=1 and w=600: ignores w=600, serves original bytes and Content-Disposition
        r_down = requests.get(f"{base_url}/image/download-test-batch/{url_quoted_fn}?w=600&download=1")
        assert r_down.status_code == 200
        assert r_down.content == orig_bytes
        cd_header = r_down.headers.get("Content-Disposition", "")
        assert "attachment" in cd_header
        assert f"filename*=UTF-8''{url_quoted_fn}" in cd_header

        # Verify image size is full original (2000x1500), NOT resized
        img_down = Image.open(io.BytesIO(r_down.content))
        assert img_down.size == (2000, 1500)

        # 2. GET with w=600 without download=1: resized, NO attachment header
        r_resized = requests.get(f"{base_url}/image/download-test-batch/{url_quoted_fn}?w=600")
        assert r_resized.status_code == 200
        assert "attachment" not in r_resized.headers.get("Content-Disposition", "")
        img_resized = Image.open(io.BytesIO(r_resized.content))
        assert max(img_resized.size) == 600

    finally:
        server.shutdown()
        server.server_close()


def test_http_ratings_routes(tmp_path, monkeypatch):
    """
    Test GET /api/ratings and POST /api/ratings routes.
    - POST without reviewer / Anonym -> 403
    - POST with reviewer -> 200
    - GET returns {ratings: {key: stars}} for requested reviewer only
    - Never return other reviewers' ratings
    """
    base_dir = tmp_path / "batch_api"
    base_dir.mkdir()

    test_batch = {
        "id": "ratings-http-batch",
        "name": "Ratings HTTP Batch",
        "path": str(base_dir),
        "org": "cama",
        "published": True,
    }

    monkeypatch.setattr(hatshop_server, "load_batches", lambda: [test_batch])
    monkeypatch.setattr(hatshop_server, "REVIEWS_DIR", str(tmp_path))

    port = get_free_port()
    base_url = f"http://127.0.0.1:{port}"

    server = HTTPServer(("127.0.0.1", port), HubHTTPRequestHandler)
    thread = threading.Thread(target=server.serve_forever)
    thread.daemon = True
    thread.start()

    time.sleep(0.2)

    try:
        tunnel_header = {"CF-Connecting-IP": "203.0.113.195"}

        # 1. POST without reviewer -> 403
        r_no_name = requests.post(
            f"{base_url}/api/ratings?batch=ratings-http-batch",
            json={"product_id": "P1", "filename": "01.jpg", "stars": 4, "reviewer": ""},
            headers=tunnel_header,
        )
        assert r_no_name.status_code == 403
        assert r_no_name.json()["message"] == "Skriv ditt namn i hubben först för att kunna sätta betyg."

        # 2. POST with reviewer="Anonym" -> 403
        r_anonym = requests.post(
            f"{base_url}/api/ratings?batch=ratings-http-batch",
            json={"product_id": "P1", "filename": "01.jpg", "stars": 4, "reviewer": "Anonym"},
            headers=tunnel_header,
        )
        assert r_anonym.status_code == 403
        assert r_anonym.json()["message"] == "Skriv ditt namn i hubben först för att kunna sätta betyg."

        # 3. POST by Alice -> 200
        r_alice = requests.post(
            f"{base_url}/api/ratings?batch=ratings-http-batch",
            json={"product_id": "P1", "filename": "01.jpg", "stars": 4, "reviewer": "Alice"},
            headers=tunnel_header,
        )
        assert r_alice.status_code == 200
        assert r_alice.json() == {"status": "ok"}

        # 4. POST by Bob -> 200
        r_bob = requests.post(
            f"{base_url}/api/ratings?batch=ratings-http-batch",
            json={"product_id": "P1", "filename": "01.jpg", "stars": 2, "reviewer": "Bob"},
            headers=tunnel_header,
        )
        assert r_bob.status_code == 200
        assert r_bob.json() == {"status": "ok"}

        # 5. GET by Alice: only sees 4
        g_alice = requests.get(f"{base_url}/api/ratings?batch=ratings-http-batch&reviewer=Alice", headers=tunnel_header)
        assert g_alice.status_code == 200
        assert g_alice.json() == {"ratings": {"P1|01.jpg": 4}}

        # 6. GET by Bob: only sees 2
        g_bob = requests.get(f"{base_url}/api/ratings?batch=ratings-http-batch&reviewer=Bob", headers=tunnel_header)
        assert g_bob.status_code == 200
        assert g_bob.json() == {"ratings": {"P1|01.jpg": 2}}

        # 7. GET by Anonym or empty -> empty ratings dict
        g_anon = requests.get(f"{base_url}/api/ratings?batch=ratings-http-batch&reviewer=Anonym", headers=tunnel_header)
        assert g_anon.status_code == 200
        assert g_anon.json() == {"ratings": {}}

        g_empty = requests.get(f"{base_url}/api/ratings?batch=ratings-http-batch&reviewer=", headers=tunnel_header)
        assert g_empty.status_code == 200
        assert g_empty.json() == {"ratings": {}}

        # 8. POST invalid stars (6) -> 400
        r_bad_stars = requests.post(
            f"{base_url}/api/ratings?batch=ratings-http-batch",
            json={"product_id": "P1", "filename": "01.jpg", "stars": 6, "reviewer": "Alice"},
            headers=tunnel_header,
        )
        assert r_bad_stars.status_code == 400
        assert r_bad_stars.json()["message"] == "Ogiltigt betyg"

        # 9. POST stars 0 removes rating -> GET returns {}
        r_del = requests.post(
            f"{base_url}/api/ratings?batch=ratings-http-batch",
            json={"product_id": "P1", "filename": "01.jpg", "stars": 0, "reviewer": "Alice"},
            headers=tunnel_header,
        )
        assert r_del.status_code == 200
        g_alice_after = requests.get(f"{base_url}/api/ratings?batch=ratings-http-batch&reviewer=Alice", headers=tunnel_header)
        assert g_alice_after.json() == {"ratings": {}}
        # Bob's rating is untouched
        g_bob_after = requests.get(f"{base_url}/api/ratings?batch=ratings-http-batch&reviewer=Bob", headers=tunnel_header)
        assert g_bob_after.json() == {"ratings": {"P1|01.jpg": 2}}

        # 10. Unknown batch -> 404
        r_404 = requests.get(f"{base_url}/api/ratings?batch=unknown-batch&reviewer=Alice")
        assert r_404.status_code == 404

    finally:
        server.shutdown()
        server.server_close()
