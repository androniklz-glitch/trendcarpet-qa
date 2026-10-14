import os
import json
import socket
import threading
import time
import shutil
import pytest
import requests
from http.server import HTTPServer

import hatshop_server
from hatshop_server import (
    removed_root_for,
    remove_shot,
    undo_remove,
    list_removed,
    scan_batch_products,
    HubHTTPRequestHandler,
)


def get_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def test_removed_root_for_helper():
    p1 = r"C:\Pictures\26-10-08-Curtain new-x"
    expected1 = r"C:\Pictures\26-10-08-Curtain new-x Removed"
    assert removed_root_for(p1) == expected1

    # Trailing slashes
    p2 = r"C:\Pictures\26-10-08-Curtain new-x\\"
    assert removed_root_for(p2) == expected1

    # Forward slashes
    p3 = "C:/Pictures/26-10-08-Curtain new-x"
    assert removed_root_for(p3) == expected1


def test_remove_shot_3_variant_tree(tmp_path):
    """
    (a) a remove in a 3-variant tree moves 3 files into <tmp>/<batch> Removed/<same rel_path>,
        and the product then lists 0 of that shot in each variant
    """
    base_dir = tmp_path / "batch_a"
    base_dir.mkdir()

    variant_names = ["1500px", "1500px iphone", "Original"]
    product_name = "Keps-Garda-Ranger-Army-Cap-dark-grey"

    for v in variant_names:
        p_dir = base_dir / v / product_name
        p_dir.mkdir(parents=True)
        (p_dir / "01.jpg").write_bytes(b"image_01_content")
        (p_dir / "02.jpg").write_bytes(b"image_02_content")

    batch = {
        "id": "test-batch-a",
        "path": str(base_dir),
    }

    prods_before = scan_batch_products(batch)
    assert len(prods_before) == 1
    assert prods_before[0]["id"] == product_name
    for v in variant_names:
        assert len(prods_before[0]["variants"][v]) == 2

    # Remove shot 01.jpg
    res = remove_shot(batch, product_name, "01.jpg", "admin_user")
    assert res["status"] == "ok"
    assert "op_id" in res
    assert len(res["moved"]) == 3
    assert res["undo_count"] == 1
    removed_root = res["removed_root"]
    assert removed_root == str(tmp_path / "batch_a Removed")

    # Verify moved files exist in Removed with same rel_path
    for v in variant_names:
        expected_rel = f"{v}/{product_name}/01.jpg"
        assert expected_rel in res["moved"]
        target_path = os.path.join(removed_root, expected_rel)
        assert os.path.isfile(target_path)
        with open(target_path, "rb") as f:
            assert f.read() == b"image_01_content"
        # Original should no longer exist
        assert not (base_dir / v / product_name / "01.jpg").exists()

    # Product should now list 0 of that shot in each variant
    prods_after = scan_batch_products(batch)
    assert len(prods_after) == 1
    for v in variant_names:
        shots = prods_after[0]["variants"][v]
        assert len(shots) == 1
        assert shots[0]["filename"] == "02.jpg"
        assert all(s["filename"] != "01.jpg" for s in shots)


def test_undo_puts_all_back(tmp_path):
    """
    (b) undo puts all 3 back and undo_count goes back to 0
    """
    base_dir = tmp_path / "batch_b"
    base_dir.mkdir()

    variant_names = ["1500px", "1500px iphone", "Original"]
    product_name = "Keps-Garda-Ranger-Army-Cap-dark-grey"

    for v in variant_names:
        p_dir = base_dir / v / product_name
        p_dir.mkdir(parents=True)
        (p_dir / "01.jpg").write_bytes(b"content_01")
        (p_dir / "02.jpg").write_bytes(b"content_02")

    batch = {
        "id": "test-batch-b",
        "path": str(base_dir),
    }

    res_remove = remove_shot(batch, product_name, "01.jpg", "admin_user")
    assert res_remove["undo_count"] == 1

    # Undo
    res_undo = undo_remove(batch)
    assert res_undo["status"] == "ok"
    assert res_undo["op_id"] == res_remove["op_id"]
    assert res_undo["undo_count"] == 0
    assert len(res_undo["restored"]) == 3

    # Verify all 3 files are back in original location
    for v in variant_names:
        orig_file = base_dir / v / product_name / "01.jpg"
        assert orig_file.is_file()
        assert orig_file.read_bytes() == b"content_01"

    # Verify scan lists both shots again
    prods = scan_batch_products(batch)
    for v in variant_names:
        filenames = [s["filename"] for s in prods[0]["variants"][v]]
        assert "01.jpg" in filenames
        assert "02.jpg" in filenames


def test_two_removes_two_undos_lifo(tmp_path):
    """
    (c) two removes then two undos restore in LIFO order
    """
    base_dir = tmp_path / "batch_c"
    base_dir.mkdir()

    product_name = "Keps-Garda-Ranger-Army-Cap-dark-grey"
    p_dir = base_dir / "1500px" / product_name
    p_dir.mkdir(parents=True)
    (p_dir / "01.jpg").write_bytes(b"c1")
    (p_dir / "02.jpg").write_bytes(b"c2")
    (p_dir / "03.jpg").write_bytes(b"c3")

    batch = {"id": "test-batch-c", "path": str(base_dir)}

    op1 = remove_shot(batch, product_name, "01.jpg", "user1")
    assert op1["undo_count"] == 1

    op2 = remove_shot(batch, product_name, "02.jpg", "user2")
    assert op2["undo_count"] == 2

    # Check list_removed order (newest first)
    listing = list_removed(batch)
    assert listing["undo_count"] == 2
    assert [op["op_id"] for op in listing["ops"]] == [op2["op_id"], op1["op_id"]]
    assert listing["ops"][0]["filename"] == "02.jpg"
    assert listing["ops"][1]["filename"] == "01.jpg"

    # First undo: should restore op2 ("02.jpg")
    undo1 = undo_remove(batch)
    assert undo1["op_id"] == op2["op_id"]
    assert undo1["undo_count"] == 1
    assert (p_dir / "02.jpg").is_file()
    assert not (p_dir / "01.jpg").exists()

    # Second undo: should restore op1 ("01.jpg")
    undo2 = undo_remove(batch)
    assert undo2["op_id"] == op1["op_id"]
    assert undo2["undo_count"] == 0
    assert (p_dir / "01.jpg").is_file()

    # Third undo: nothing left to undo -> ValueError
    with pytest.raises(ValueError, match="Inget att ångra"):
        undo_remove(batch)


def test_name_collision_gives_dash_one(tmp_path):
    """
    (d) a name collision in Removed gives -1
    """
    base_dir = tmp_path / "batch_d"
    base_dir.mkdir()

    product_name = "Keps-Garda-Ranger-Army-Cap-dark-grey"
    p_dir = base_dir / "1500px" / product_name
    p_dir.mkdir(parents=True)
    (p_dir / "01.jpg").write_bytes(b"fresh_01")

    # Pre-create the collision target in Removed folder
    removed_dir = tmp_path / "batch_d Removed" / "1500px" / product_name
    removed_dir.mkdir(parents=True)
    existing_file = removed_dir / "01.jpg"
    existing_file.write_bytes(b"already_existing_in_removed")

    batch = {"id": "test-batch-d", "path": str(base_dir)}

    res = remove_shot(batch, product_name, "01.jpg", "user_d")
    assert res["status"] == "ok"

    # Pre-existing file in Removed must NOT be overwritten
    assert existing_file.read_bytes() == b"already_existing_in_removed"

    # Collided file must have -1 inserted before extension
    collided_target = removed_dir / "01-1.jpg"
    assert collided_target.is_file()
    assert collided_target.read_bytes() == b"fresh_01"

    # Undo must restore from 01-1.jpg back to 01.jpg
    res_undo = undo_remove(batch)
    assert res_undo["status"] == "ok"
    assert (p_dir / "01.jpg").read_bytes() == b"fresh_01"
    assert not collided_target.exists()
    assert existing_file.read_bytes() == b"already_existing_in_removed"


def test_manifest_batch_raises_value_error(tmp_path):
    """
    (e) a manifest batch raises ValueError and moves nothing
    """
    base_dir = tmp_path / "batch_e"
    base_dir.mkdir()

    p_dir = base_dir / "1500px" / "ProdE"
    p_dir.mkdir(parents=True)
    test_img = p_dir / "01.jpg"
    test_img.write_bytes(b"data_e")

    # Put products_manifest.json in base_dir
    manifest_file = base_dir / "products_manifest.json"
    manifest_file.write_text("[]", encoding="utf-8")

    batch = {"id": "test-batch-e", "path": str(base_dir)}

    with pytest.raises(ValueError, match="Mappen använder manifest – flytt stöds inte"):
        remove_shot(batch, "ProdE", "01.jpg", "user_e")

    # Verify nothing was moved
    assert test_img.is_file()
    removed_dir = tmp_path / "batch_e Removed"
    assert not removed_dir.exists()


def test_undo_raises_file_exists_error_moves_nothing(tmp_path):
    """
    (f) undo raises FileExistsError when the original exists again and moves nothing
    """
    base_dir = tmp_path / "batch_f"
    base_dir.mkdir()

    product_name = "Keps-Garda-Ranger-Army-Cap-dark-grey"
    v1_dir = base_dir / "1500px" / product_name
    v2_dir = base_dir / "Original" / product_name
    v1_dir.mkdir(parents=True)
    v2_dir.mkdir(parents=True)
    (v1_dir / "01.jpg").write_bytes(b"v1_data")
    (v2_dir / "01.jpg").write_bytes(b"v2_data")

    batch = {"id": "test-batch-f", "path": str(base_dir)}

    res_remove = remove_shot(batch, product_name, "01.jpg", "user_f")
    assert res_remove["status"] == "ok"

    # Now create an obstacle at one of the destination paths
    (v1_dir / "01.jpg").write_bytes(b"obstacle_v1")

    # Calling undo_remove should raise FileExistsError before moving any file
    with pytest.raises(FileExistsError, match="Originalfilen finns redan:"):
        undo_remove(batch)

    # v1_dir/01.jpg should still contain obstacle_v1
    assert (v1_dir / "01.jpg").read_bytes() == b"obstacle_v1"
    # v2_dir/01.jpg should NOT have been moved back
    assert not (v2_dir / "01.jpg").exists()

    # The files in Removed should still exist
    removed_root = res_remove["removed_root"]
    assert os.path.isfile(os.path.join(removed_root, "1500px", product_name, "01.jpg"))
    assert os.path.isfile(os.path.join(removed_root, "Original", product_name, "01.jpg"))

    # Op in log must still be active (not undone)
    listing = list_removed(batch)
    assert listing["undo_count"] == 1


def test_model_color_tree_0330(tmp_path):
    """
    (g) a 0330 model/color tree (1500px/Lumira/Lumia-röd-x/a.jpg) works
    """
    base_dir = tmp_path / "batch_g"
    base_dir.mkdir()

    for v in ["1500px", "Original"]:
        img_dir = base_dir / v / "Lumira" / "Lumia-röd-x"
        img_dir.mkdir(parents=True)
        (img_dir / "a.jpg").write_bytes(b"image_a")
        (img_dir / "b.jpg").write_bytes(b"image_b")

    batch = {
        "id": "test-nested-batch-g",
        "path": str(base_dir)
    }

    products = scan_batch_products(batch)
    assert len(products) == 1
    prod = products[0]
    assert prod["id"] == "Lumira-Lumia-röd-x"

    res = remove_shot(batch, "Lumira-Lumia-röd-x", "a.jpg", "admin")
    assert res["status"] == "ok"
    assert len(res["moved"]) == 2

    # Removed destination check
    removed_root = res["removed_root"]
    assert os.path.isfile(os.path.join(removed_root, "1500px", "Lumira", "Lumia-röd-x", "a.jpg"))
    assert os.path.isfile(os.path.join(removed_root, "Original", "Lumira", "Lumia-röd-x", "a.jpg"))

    # Product in batch now has only b.jpg
    prods_after = scan_batch_products(batch)
    assert len(prods_after) == 1
    for v in ["1500px", "Original"]:
        filenames = [s["filename"] for s in prods_after[0]["variants"][v]]
        assert "a.jpg" not in filenames
        assert "b.jpg" in filenames

    # Undo
    undo_res = undo_remove(batch)
    assert undo_res["status"] == "ok"
    assert undo_res["undo_count"] == 0
    assert os.path.isfile(str(base_dir / "1500px" / "Lumira" / "Lumia-röd-x" / "a.jpg"))
    assert os.path.isfile(str(base_dir / "Original" / "Lumira" / "Lumia-röd-x" / "a.jpg"))


def test_http_image_routes(tmp_path, monkeypatch):
    """
    Test HTTP endpoints:
    - POST /api/image/remove?batch=<id> (admin vs non-admin, 404 on unknown batch, 400 on ValueError, 200 on success)
    - GET /api/image/removed?batch=<id> (admin vs non-admin)
    - POST /api/image/undo_remove?batch=<id> (admin vs non-admin, 409 on FileExistsError, 200 on success)
    """
    base_dir = tmp_path / "batch_http"
    base_dir.mkdir()

    p_dir = base_dir / "1500px" / "Keps-Garda-Ranger-Army-Cap-dark-grey"
    p_dir.mkdir(parents=True)
    (p_dir / "01.jpg").write_bytes(b"http_img_1")
    (p_dir / "02.jpg").write_bytes(b"http_img_2")

    test_batch = {
        "id": "http-test-batch",
        "name": "HTTP Test Batch",
        "path": str(base_dir),
        "org": "cama",
        "published": True
    }

    # Mock load_batches to return our test_batch
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

        # 1. Non-admin without name POST /api/image/remove -> 403
        r_forbidden = requests.post(
            f"{base_url}/api/image/remove?batch=http-test-batch",
            json={"product_id": "Keps-Garda-Ranger-Army-Cap-dark-grey", "filename": "01.jpg", "reviewer": ""},
            headers=tunnel_header
        )
        assert r_forbidden.status_code == 403
        assert r_forbidden.json()["message"] == "Skriv ditt namn i hubben först för att kunna flytta ut bilder."

        # 1b. Non-admin with named reviewer "Amalia" POST /api/image/remove -> 200
        r_amalia = requests.post(
            f"{base_url}/api/image/remove?batch=http-test-batch",
            json={"product_id": "Keps-Garda-Ranger-Army-Cap-dark-grey", "filename": "02.jpg", "reviewer": "Amalia"},
            headers=tunnel_header
        )
        assert r_amalia.status_code == 200
        # Restore 02.jpg so subsequent tests remain unaffected
        r_amalia_undo = requests.post(
            f"{base_url}/api/image/undo_remove?batch=http-test-batch",
            json={"reviewer": "Amalia"},
            headers=tunnel_header
        )
        assert r_amalia_undo.status_code == 200

        # 2. Unknown batch POST /api/image/remove -> 404
        r_notfound = requests.post(
            f"{base_url}/api/image/remove?batch=unknown-batch",
            json={"product_id": "Keps-Garda-Ranger-Army-Cap-dark-grey", "filename": "01.jpg"}
        )
        assert r_notfound.status_code == 404

        # 3. Missing / invalid product POST /api/image/remove -> 400
        r_bad = requests.post(
            f"{base_url}/api/image/remove?batch=http-test-batch",
            json={"product_id": "NonExistentProduct", "filename": "01.jpg"}
        )
        assert r_bad.status_code == 400
        assert r_bad.json()["message"] == "Produkten hittades inte"

        # 4. Valid admin POST /api/image/remove -> 200
        r_ok = requests.post(
            f"{base_url}/api/image/remove?batch=http-test-batch",
            json={"product_id": "Keps-Garda-Ranger-Army-Cap-dark-grey", "filename": "01.jpg", "reviewer": "AdminUser"}
        )
        assert r_ok.status_code == 200
        data_ok = r_ok.json()
        assert data_ok["status"] == "ok"
        assert data_ok["undo_count"] == 1
        op_id = data_ok["op_id"]

        # 5. Non-admin GET /api/image/removed -> 200 {"undo_count": 0, "ops": []}
        r_get_nonadmin = requests.get(
            f"{base_url}/api/image/removed?batch=http-test-batch",
            headers=tunnel_header
        )
        assert r_get_nonadmin.status_code == 200
        assert r_get_nonadmin.json() == {"undo_count": 0, "ops": []}

        # 6. Admin GET /api/image/removed -> 200 with list_removed
        r_get_admin = requests.get(f"{base_url}/api/image/removed?batch=http-test-batch")
        assert r_get_admin.status_code == 200
        data_removed = r_get_admin.json()
        assert data_removed["undo_count"] == 1
        assert len(data_removed["ops"]) == 1
        assert data_removed["ops"][0]["op_id"] == op_id

        # 7. Non-admin POST /api/image/undo_remove without name -> 403
        r_undo_forbidden = requests.post(
            f"{base_url}/api/image/undo_remove?batch=http-test-batch",
            json={"op_id": op_id, "reviewer": ""},
            headers=tunnel_header
        )
        assert r_undo_forbidden.status_code == 403
        assert r_undo_forbidden.json()["message"] == "Skriv ditt namn i hubben först för att kunna flytta ut bilder."

        # 8. Obstacle in original place -> 409 Conflict
        (p_dir / "01.jpg").write_bytes(b"obstacle")
        r_undo_conflict = requests.post(
            f"{base_url}/api/image/undo_remove?batch=http-test-batch",
            json={"op_id": op_id}
        )
        assert r_undo_conflict.status_code == 409
        assert "Originalfilen finns redan:" in r_undo_conflict.json()["message"]

        # Remove obstacle and undo successfully -> 200
        (p_dir / "01.jpg").unlink()
        r_undo_ok = requests.post(
            f"{base_url}/api/image/undo_remove?batch=http-test-batch",
            json={"op_id": op_id}
        )
        assert r_undo_ok.status_code == 200
        assert r_undo_ok.json()["status"] == "ok"
        assert r_undo_ok.json()["undo_count"] == 0
        assert (p_dir / "01.jpg").read_bytes() == b"http_img_1"

    finally:
        server.shutdown()
        server.server_close()
