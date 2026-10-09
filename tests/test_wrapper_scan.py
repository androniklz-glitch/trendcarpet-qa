import os
import pytest
from hatshop_server import scan_batch_products, find_batch_by_path


def test_wrapper_scan_tree(tmp_path):
    base_dir = tmp_path / "batch_wrapper"
    base_dir.mkdir()
    wrapper_dir = base_dir / "levererat"
    wrapper_dir.mkdir()

    variant_names = ["1500px", "1500px iphone", "Original"]
    product_names = [
        "Hattar-Garda-Napoli-Fedora-Wool-Hat-gra",
        "Keps-Garda-Ranger-Army-Cap-dark-grey"
    ]

    for v in variant_names:
        v_dir = wrapper_dir / v
        v_dir.mkdir()
        for p in product_names:
            p_dir = v_dir / p
            p_dir.mkdir()
            (p_dir / "01.jpg").write_bytes(b"dummy_image_data")
            (p_dir / "02.jpg").write_bytes(b"dummy_image_data")

    batch = {
        "id": "test-wrapper-batch",
        "path": str(base_dir)
    }

    products = scan_batch_products(batch)
    assert len(products) == 2

    prod_ids = [p["id"] for p in products]
    assert "Hattar-Garda-Napoli-Fedora-Wool-Hat-gra" in prod_ids
    assert "Keps-Garda-Ranger-Army-Cap-dark-grey" in prod_ids

    for p in products:
        assert sorted(p["variants"].keys()) == sorted(variant_names)
        for v_name, shots in p["variants"].items():
            assert len(shots) == 2
            for s in shots:
                assert s["rel_path"].startswith("levererat/"), f"rel_path '{s['rel_path']}' should start with levererat/"


def test_flat_scan_tree(tmp_path):
    base_dir = tmp_path / "batch_flat"
    base_dir.mkdir()

    variant_names = ["1500px", "Original"]
    product_name = "Keps-Garda-Ranger-Army-Cap-dark-grey"

    for v in variant_names:
        v_dir = base_dir / v
        v_dir.mkdir()
        p_dir = v_dir / product_name
        p_dir.mkdir()
        (p_dir / "01.jpg").write_bytes(b"dummy_image_data")

    batch = {
        "id": "test-flat-batch",
        "path": str(base_dir)
    }

    products = scan_batch_products(batch)
    assert len(products) == 1
    assert products[0]["id"] == product_name
    assert sorted(products[0]["variants"].keys()) == sorted(variant_names)

    for v_name, shots in products[0]["variants"].items():
        assert len(shots) == 1
        for s in shots:
            assert not s["rel_path"].startswith("levererat/")
            assert s["rel_path"].startswith(f"{v_name}/")


def test_find_batch_by_path():
    batches = [
        {"id": "b1", "name": "Batch 1", "path": r"C:\Pictures\26-10-08-HS-Hatshop-rest"},
        {"id": "b2", "name": "Batch 2", "path": r"C:\Pictures\26-08-28-Hatshop-Black-River pt1-x"}
    ]

    # Exact match
    res1 = find_batch_by_path(batches, r"C:\Pictures\26-10-08-HS-Hatshop-rest")
    assert res1 is not None
    assert res1["id"] == "b1"

    # Case-insensitive and trailing slash normalization match
    res2 = find_batch_by_path(batches, r"c:\pictures\26-10-08-hs-hatshop-rest\\")
    assert res2 is not None
    assert res2["id"] == "b1"

    # Forward slash vs backslash normalization match
    res3 = find_batch_by_path(batches, "C:/Pictures/26-08-28-Hatshop-Black-River pt1-x")
    assert res3 is not None
    assert res3["id"] == "b2"

    # Non-existing path
    assert find_batch_by_path(batches, r"C:\Pictures\Non-Existent") is None
    assert find_batch_by_path([], r"C:\Pictures\26-10-08-HS-Hatshop-rest") is None
    assert find_batch_by_path(batches, "") is None
    assert find_batch_by_path(batches, None) is None


def test_variant_model_color_subfolders(tmp_path):
    base_dir = tmp_path / "batch_nested_variant"
    base_dir.mkdir()

    for v in ["1500px", "Original"]:
        img_dir = base_dir / v / "Lumira" / "Lumia-röd-x"
        img_dir.mkdir(parents=True)
        (img_dir / "a.jpg").write_bytes(b"dummy_image_data")

    batch = {
        "id": "test-nested-variant-batch",
        "path": str(base_dir)
    }

    products = scan_batch_products(batch)
    assert len(products) == 1
    prod = products[0]
    assert prod["id"] == "Lumira-Lumia-röd-x"
    assert sorted(prod["variants"].keys()) == ["1500px", "Original"]
    assert prod["variants"]["1500px"][0]["rel_path"] == "1500px/Lumira/Lumia-röd-x/a.jpg"
    assert prod["variants"]["Original"][0]["rel_path"] == "Original/Lumira/Lumia-röd-x/a.jpg"


def test_variant_model_color_wrapped_tree(tmp_path):
    base_dir = tmp_path / "batch_wrapped_nested_variant"
    base_dir.mkdir()

    for v in ["1500px", "Original"]:
        img_dir = base_dir / "levererat" / v / "Lumira" / "Lumia-röd-x"
        img_dir.mkdir(parents=True)
        (img_dir / "a.jpg").write_bytes(b"dummy_image_data")

    batch = {
        "id": "test-wrapped-nested-variant-batch",
        "path": str(base_dir)
    }

    products = scan_batch_products(batch)
    assert len(products) == 1
    prod = products[0]
    assert prod["id"] == "Lumira-Lumia-röd-x"
    assert sorted(prod["variants"].keys()) == ["1500px", "Original"]
    for s in prod["variants"]["1500px"]:
        assert s["rel_path"].startswith("levererat/1500px/Lumira/")
    assert prod["variants"]["1500px"][0]["rel_path"] == "levererat/1500px/Lumira/Lumia-röd-x/a.jpg"
    assert prod["variants"]["Original"][0]["rel_path"] == "levererat/Original/Lumira/Lumia-röd-x/a.jpg"

