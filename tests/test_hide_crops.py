import os
import shutil
import pytest
from PIL import Image

import hatshop_server
from hatshop_server import (
    scan_batch_products,
    remove_shot,
    undo_remove,
    hidden_crop_dirs,
    crop_siblings,
    removed_root_for,
    CROP_FORMAT_DIRS,
    MASTER_FORMAT_DIR,
)


@pytest.fixture(autouse=True)
def disable_notify(monkeypatch):
    monkeypatch.setenv("QA_HUB_NOTIFY_SCRIPT", "")
    monkeypatch.setattr(hatshop_server, "NOTIFY_SCRIPT", "")



def make_tiny_png(path):
    os.makedirs(os.path.dirname(str(path)), exist_ok=True)
    img = Image.new("RGB", (10, 10), color=(100, 150, 200))
    img.save(str(path), "PNG")


def test_hidden_crop_dirs_helper():
    # Only hides when MASTER_FORMAT_DIR (9x16) is present
    dirs = ["1x1", "3x4", "4x5", "9x16"]
    assert hidden_crop_dirs(dirs) == {"3x4", "4x5"}

    # Case-insensitive check
    dirs_upper = ["1X1", "3X4", "4X5", "9X16"]
    assert hidden_crop_dirs(dirs_upper) == {"3X4", "4X5"}

    # Without 9x16 sibling, crop dirs stay visible
    dirs_no_master = ["1x1", "3x4", "4x5"]
    assert hidden_crop_dirs(dirs_no_master) == set()

    # Empty list
    assert hidden_crop_dirs([]) == set()


def test_crop_siblings_helper(tmp_path):
    base_dir = tmp_path / "batch_crops_helper"
    p_dir = base_dir / "ModelA"
    
    make_tiny_png(p_dir / "9x16" / "photo.png")
    make_tiny_png(p_dir / "3x4" / "photo-3x4.png")
    make_tiny_png(p_dir / "4x5" / "photo-4x5.jpg")
    make_tiny_png(p_dir / "3x4" / "other-3x4.png")

    master_path = str(p_dir / "9x16" / "photo.png")
    sibs = crop_siblings(master_path, batch_path=str(base_dir))

    norm_sibs = [os.path.normcase(os.path.abspath(s)) for s in sibs]
    assert os.path.normcase(os.path.abspath(str(p_dir / "3x4" / "photo-3x4.png"))) in norm_sibs
    assert os.path.normcase(os.path.abspath(str(p_dir / "4x5" / "photo-4x5.jpg"))) in norm_sibs
    assert os.path.normcase(os.path.abspath(str(p_dir / "3x4" / "other-3x4.png"))) not in norm_sibs

    # If parent is not 9x16, returns empty list
    not_master = str(p_dir / "3x4" / "photo-3x4.png")
    assert crop_siblings(not_master, batch_path=str(base_dir)) == []


def test_hide_crops_batch_workflow(tmp_path):
    """
    1. Batch P/{1x1,3x4,4x5,9x16} with 2 files in 9x16 (a.png, b.png)
       + crops a-3x4.png, a-4x5.png, b-3x4.png, b-4x5.png + 1x1/c.png
       -> product ids contain P-9x16 and P-1x1, and NO id ending in -3x4 or -4x5.
    3. remove_shot of a.png on P-9x16 -> a.png, a-3x4.png, a-4x5.png are all gone
       from the batch and exist under the Removed root; b-* untouched.
    4. undo_remove -> all three are back.
    """
    base_dir = tmp_path / "batch_main"
    p_dir = base_dir / "P"

    make_tiny_png(p_dir / "9x16" / "a.png")
    make_tiny_png(p_dir / "9x16" / "b.png")
    make_tiny_png(p_dir / "3x4" / "a-3x4.png")
    make_tiny_png(p_dir / "3x4" / "b-3x4.png")
    make_tiny_png(p_dir / "4x5" / "a-4x5.png")
    make_tiny_png(p_dir / "4x5" / "b-4x5.png")
    make_tiny_png(p_dir / "1x1" / "c.png")

    batch = {
        "id": "test-batch-main",
        "path": str(base_dir),
    }

    # Step 1: scan_batch_products hides 3x4 and 4x5 folders next to 9x16
    products = scan_batch_products(batch)
    p_ids = [p["id"] for p in products]

    assert "P-9x16" in p_ids
    assert "P-1x1" in p_ids
    assert not any(pid.endswith("-3x4") or pid.endswith("-4x5") for pid in p_ids)
    assert len(p_ids) == 2

    # Step 3: remove_shot of a.png on P-9x16 moves a.png, a-3x4.png, a-4x5.png
    res_remove = remove_shot(batch, "P-9x16", "a.png", "admin_user")
    assert res_remove["status"] == "ok"
    assert "op_id" in res_remove
    assert len(res_remove["moved"]) == 3
    assert "P/9x16/a.png" in res_remove["moved"]
    assert "P/3x4/a-3x4.png" in res_remove["moved"]
    assert "P/4x5/a-4x5.png" in res_remove["moved"]

    removed_root = removed_root_for(base_dir)
    assert os.path.normpath(res_remove["removed_root"]) == os.path.normpath(removed_root)

    # a.png, a-3x4.png, a-4x5.png are all gone from the batch
    assert not (p_dir / "9x16" / "a.png").exists()
    assert not (p_dir / "3x4" / "a-3x4.png").exists()
    assert not (p_dir / "4x5" / "a-4x5.png").exists()

    # and exist under the Removed root
    assert (base_dir.parent / f"{base_dir.name} Removed" / "P" / "9x16" / "a.png").exists()
    assert (base_dir.parent / f"{base_dir.name} Removed" / "P" / "3x4" / "a-3x4.png").exists()
    assert (base_dir.parent / f"{base_dir.name} Removed" / "P" / "4x5" / "a-4x5.png").exists()

    # b-* and c untouched
    assert (p_dir / "9x16" / "b.png").exists()
    assert (p_dir / "3x4" / "b-3x4.png").exists()
    assert (p_dir / "4x5" / "b-4x5.png").exists()
    assert (p_dir / "1x1" / "c.png").exists()

    # Step 4: undo_remove brings all three back
    res_undo = undo_remove(batch, op_id=res_remove["op_id"])
    assert res_undo["status"] == "ok"
    assert res_undo["undo_count"] == 0

    assert (p_dir / "9x16" / "a.png").exists()
    assert (p_dir / "3x4" / "a-3x4.png").exists()
    assert (p_dir / "4x5" / "a-4x5.png").exists()

    # Verify scan again includes a.png in P-9x16
    prods_restored = scan_batch_products(batch)
    p_9x16 = next(p for p in prods_restored if p["id"] == "P-9x16")
    filenames = [s["filename"] for s in p_9x16["variants"]["Standard"]]
    assert "a.png" in filenames
    assert "b.png" in filenames


def test_scan_crop_without_9x16_stays_visible(tmp_path):
    """
    2. Batch with only P/3x4/x.png (no 9x16) -> product P-3x4 is present.
    """
    base_dir = tmp_path / "batch_solo_crop"
    p_dir = base_dir / "P"
    make_tiny_png(p_dir / "3x4" / "x.png")

    batch = {
        "id": "test-batch-solo",
        "path": str(base_dir),
    }

    products = scan_batch_products(batch)
    p_ids = [p["id"] for p in products]

    assert "P-3x4" in p_ids
    assert len(products) == 1


def test_crop_same_stem_without_suffix(tmp_path):
    """
    Verify crop files named <stem><ext> (identical stem to master) are also moved and restored.
    """
    base_dir = tmp_path / "batch_exact_stem"
    p_dir = base_dir / "P"

    make_tiny_png(p_dir / "9x16" / "shot01.png")
    # Same stem without -3x4 suffix:
    make_tiny_png(p_dir / "3x4" / "shot01.png")
    make_tiny_png(p_dir / "4x5" / "shot01.png")

    batch = {
        "id": "test-batch-exact",
        "path": str(base_dir),
    }

    products = scan_batch_products(batch)
    assert len(products) == 1
    assert products[0]["id"] == "P-9x16"

    res = remove_shot(batch, "P-9x16", "shot01.png", "admin_user")
    assert res["status"] == "ok"
    assert len(res["moved"]) == 3

    assert not (p_dir / "9x16" / "shot01.png").exists()
    assert not (p_dir / "3x4" / "shot01.png").exists()
    assert not (p_dir / "4x5" / "shot01.png").exists()

    res_undo = undo_remove(batch)
    assert res_undo["status"] == "ok"

    assert (p_dir / "9x16" / "shot01.png").exists()
    assert (p_dir / "3x4" / "shot01.png").exists()
    assert (p_dir / "4x5" / "shot01.png").exists()


def test_variant_subdir_hides_crops_when_9x16_present(tmp_path):
    """
    Test variant subdirs branch (e.g. 1500px / P / {1x1, 3x4, 4x5, 9x16})
    """
    base_dir = tmp_path / "batch_variant_subdirs"
    v_dir = base_dir / "1500px" / "P"

    make_tiny_png(v_dir / "9x16" / "v1.png")
    make_tiny_png(v_dir / "3x4" / "v1-3x4.png")
    make_tiny_png(v_dir / "4x5" / "v1-4x5.png")
    make_tiny_png(v_dir / "1x1" / "v1-1x1.png")

    batch = {
        "id": "test-batch-variant-subdirs",
        "path": str(base_dir),
    }

    products = scan_batch_products(batch)
    p_ids = [p["id"] for p in products]

    assert "P-9x16" in p_ids
    assert "P-1x1" in p_ids
    assert not any(pid.endswith("-3x4") or pid.endswith("-4x5") for pid in p_ids)
