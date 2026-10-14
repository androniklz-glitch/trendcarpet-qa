import os
import pytest
from PIL import Image

from hatshop_server import scan_batch_products, remove_shot


def create_tiny_png(path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    img = Image.new("RGBA", (1, 1), color=(255, 0, 0, 255))
    img.save(str(path), "PNG")


def test_root_plus_subfolder_both_show(tmp_path):
    """
    1. Root caps-lab-01-001.png, caps-lab-01-003.png + B/caps-lab-02-002.png
       -> a product with id B (1 shot) exists AND the total number of shots
          over all products is 3.
    """
    base_dir = tmp_path / "batch_root_plus_subfolder"
    base_dir.mkdir()

    create_tiny_png(base_dir / "caps-lab-01-001.png")
    create_tiny_png(base_dir / "caps-lab-01-003.png")
    create_tiny_png(base_dir / "B" / "caps-lab-02-002.png")

    batch = {"id": "t", "path": str(base_dir)}
    products = scan_batch_products(batch)

    # A product with id B (1 shot) exists
    prod_b = next((p for p in products if p["id"] == "B"), None)
    assert prod_b is not None
    assert len(prod_b["variants"]["Standard"]) == 1

    # Total number of shots over all products is 3
    total_shots = sum(len(shots) for p in products for shots in p["variants"].values())
    assert total_shots == 3

    # Product caps-lab also exists with 2 shots
    prod_caps = next((p for p in products if p["id"] == "caps-lab"), None)
    assert prod_caps is not None
    assert len(prod_caps["variants"]["Standard"]) == 2


def test_pure_flat_folder(tmp_path):
    """
    2. Pure flat folder (only the two root files) -> the same products/shots
       as before the change (assert total shots 2 and no product B).
    """
    base_dir = tmp_path / "batch_pure_flat"
    base_dir.mkdir()

    create_tiny_png(base_dir / "caps-lab-01-001.png")
    create_tiny_png(base_dir / "caps-lab-01-003.png")

    batch = {"id": "t", "path": str(base_dir)}
    products = scan_batch_products(batch)

    total_shots = sum(len(shots) for p in products for shots in p["variants"].values())
    assert total_shots == 2
    assert not any(p["id"] == "B" for p in products)
    assert any(p["id"] == "caps-lab" for p in products)


def test_subfolder_only(tmp_path):
    """
    3. Subfolder-only (B/x-01.png, C/y-01.png) -> exactly products B and C.
    """
    base_dir = tmp_path / "batch_subfolder_only"
    base_dir.mkdir()

    create_tiny_png(base_dir / "B" / "x-01.png")
    create_tiny_png(base_dir / "C" / "y-01.png")

    batch = {"id": "t", "path": str(base_dir)}
    products = scan_batch_products(batch)

    prod_ids = sorted(p["id"] for p in products)
    assert prod_ids == ["B", "C"]
    total_shots = sum(len(shots) for p in products for shots in p["variants"].values())
    assert total_shots == 2


def test_key_collision_prefixed_with_rot(tmp_path):
    """
    Collision test: if a flat product key collides with an existing
    products_map key, prefix the flat key with rot- (id, folder_name and key).
    Verify remove_shot still works on the root shot.
    """
    base_dir = tmp_path / "batch_collision"
    base_dir.mkdir()

    create_tiny_png(base_dir / "B" / "shot-01.png")
    create_tiny_png(base_dir / "B-01.png")

    batch = {"id": "t", "path": str(base_dir)}
    products = scan_batch_products(batch)

    prod_ids = sorted(p["id"] for p in products)
    assert "B" in prod_ids
    assert "rot-B" in prod_ids
    assert len(products) == 2

    prod_rot_b = next(p for p in products if p["id"] == "rot-B")
    assert prod_rot_b["folder_name"] == "rot-B"
    assert prod_rot_b["variants"]["Standard"][0]["rel_path"] == "B-01.png"

    # Verify remove_shot works for the loose root shot
    res = remove_shot(batch, "rot-B", "B-01.png", "admin")
    assert res["status"] == "ok"
    assert (tmp_path / "batch_collision Removed" / "B-01.png").is_file()
