"""Offline tests for scripts/inbox_ingest.py (loaded by file; tmp dirs stand in for the inbox and Pictures; no hub server)."""
import datetime
import importlib.util
import json
import os
import types

import pytest

HUB = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_spec = importlib.util.spec_from_file_location("ii", os.path.join(HUB, "scripts", "inbox_ingest.py"))
ii = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(ii)

NOW = datetime.datetime(2026, 11, 10, 14, 5)


@pytest.fixture
def env(tmp_path, monkeypatch):
    pics = tmp_path / "Pictures"
    (pics / "QA-hubb-batcher").mkdir(parents=True)
    inbox = tmp_path / "hub" / "_inkorg"
    inbox.mkdir(parents=True)
    cfg = tmp_path / "hub" / "batches_config.json"
    cfg.write_text("[]", encoding="utf-8")
    monkeypatch.setattr(ii, "PICTURES", str(pics))
    monkeypatch.setattr(ii, "BATCHER_DIR", str(pics / "QA-hubb-batcher"))
    monkeypatch.setattr(ii, "BATCHES_FILE", str(cfg))
    monkeypatch.setattr(ii, "CONFIG_BACKUP_DIR", str(tmp_path / "backup"))
    return types.SimpleNamespace(pics=pics, batcher=pics / "QA-hubb-batcher", inbox=inbox, cfg=cfg)


def put(folder, files):
    for rel, data in files.items():
        p = folder / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(data.encode())


def test_bad_id_and_missing_source(env):
    assert ii.plan("../evil", str(env.inbox), [])["status"] == "error_bad_batch_id"
    assert ii.plan("Has Space", str(env.inbox), [])["status"] == "error_bad_batch_id"
    assert ii.plan("not-there", str(env.inbox), [])["status"] == "error_source_missing"


def test_no_images_is_refused(env):
    put(env.inbox / "b-1", {"notes.txt": "x"})
    assert ii.plan("b-1", str(env.inbox), [])["status"] == "error_no_images"


def test_new_delivery_copies_to_batcher(env):
    put(env.inbox / "b-1", {"p1/a.jpg": "A", "p1/b.jpg": "B"})
    rep = ii.plan("b-1", str(env.inbox), [], NOW)
    assert rep["status"] == "would_copy" and rep["entry"] == "new"
    assert rep["dst"] == str(env.batcher / "b-1")
    assert rep["files"] == {"main": 2, "removed": 0}


def test_copy_verified_includes_removed_and_skips_desktop_ini(env):
    put(env.inbox / "b-1", {"a.jpg": "A", "desktop.ini": "x"})
    put(env.inbox / "b-1 Removed", {"r.jpg": "R"})
    ok, counts = ii.copy_verified(str(env.inbox / "b-1"), str(env.batcher / "b-1"))
    assert ok and counts == {"b-1": 1, "b-1 Removed": 1}
    assert (env.batcher / "b-1 Removed" / "r.jpg").read_bytes() == b"R"
    assert not (env.batcher / "b-1" / "desktop.ini").exists()
    assert (env.inbox / "b-1" / "a.jpg").exists()  # source untouched


def test_same_delivery_again_is_reused(env):
    put(env.inbox / "b-1", {"a.jpg": "A", "b.jpg": "B"})
    put(env.batcher / "b-1", {"a.jpg": "A", "b.jpg": "B"})
    batches = [{"id": "b-1", "path": str(env.batcher / "b-1")}]
    rep = ii.plan("b-1", str(env.inbox), batches, NOW)
    assert rep["status"] == "would_reuse" and rep["dst"] == str(env.batcher / "b-1")


def test_reviewer_removals_still_count_as_same_delivery(env):
    put(env.inbox / "b-1", {"a.jpg": "A", "b.jpg": "B"})
    put(env.batcher / "b-1", {"a.jpg": "A"})
    put(env.batcher / "b-1 Removed", {"b.jpg": "B"})
    rep = ii.plan("b-1", str(env.inbox), [{"id": "b-1", "path": str(env.batcher / "b-1")}], NOW)
    assert rep["status"] == "would_reuse"


def test_rebuilt_delivery_goes_to_new_version(env):
    put(env.inbox / "b-1", {"a.jpg": "A2", "c.jpg": "C"})
    put(env.batcher / "b-1", {"a.jpg": "A", "b.jpg": "B"})
    rep = ii.plan("b-1", str(env.inbox), [{"id": "b-1", "path": str(env.batcher / "b-1")}], NOW)
    assert rep["status"] == "would_copy" and rep["entry"] == "existing"
    assert rep["dst"] == str(env.batcher / "b-1__v261110-1405")


def test_fewer_images_is_not_same_delivery(env):
    """The copy still shows b.jpg which the new delivery dropped -> new version, not reuse."""
    put(env.inbox / "b-1", {"a.jpg": "A"})
    put(env.batcher / "b-1", {"a.jpg": "A", "b.jpg": "B"})
    rep = ii.plan("b-1", str(env.inbox), [], NOW)
    assert rep["status"] == "would_copy" and rep["dst"].endswith("__v261110-1405")


def test_dropped_duplicate_image_is_not_same_delivery(env):
    """Sovrum v8 holds every image twice (root + product folder): dropping the root copy must still be a new delivery."""
    put(env.inbox / "b-1", {"p/a.jpg": "A"})
    put(env.batcher / "b-1", {"a.jpg": "A", "p/a.jpg": "A"})
    rep = ii.plan("b-1", str(env.inbox), [], NOW)
    assert rep["status"] == "would_copy" and rep["dst"].endswith("__v261110-1405")


def test_same_content_other_path_is_not_same_delivery(env):
    put(env.inbox / "b-1", {"q/a.jpg": "A"})
    put(env.batcher / "b-1", {"p/a.jpg": "A"})
    assert ii.plan("b-1", str(env.inbox), [], NOW)["status"] == "would_copy"


def test_entry_pointing_to_other_pictures_copy_is_reused(env):
    other = env.pics / "26-11-09-x"
    put(other, {"a.jpg": "A"})
    put(env.inbox / "b-1", {"a.jpg": "A"})
    rep = ii.plan("b-1", str(env.inbox), [{"id": "b-1", "path": str(other)}], NOW)
    assert rep["status"] == "would_reuse" and rep["dst"] == str(other)


def test_version_collision_is_an_error(env):
    put(env.inbox / "b-1", {"a.jpg": "NEW"})
    put(env.batcher / "b-1", {"a.jpg": "OLD"})
    put(env.batcher / "b-1__v261110-1405", {"a.jpg": "OTHER"})
    assert ii.plan("b-1", str(env.inbox), [], NOW)["status"] == "error_collision"


def test_dst_owned_by_other_batch_is_an_error(env):
    put(env.inbox / "b-1", {"a.jpg": "A"})
    rep = ii.plan("b-1", str(env.inbox), [{"id": "other", "path": str(env.batcher / "b-1")}], NOW)
    assert rep["status"] == "error_path_owned_by_other"


def test_register_new_and_repoint(env):
    batches = [{"id": "b-1", "path": "C:/old", "category": "kepsar"}]
    assert ii.register(batches, "b-1", "C:/new", {})["action"] == "repointed"
    assert batches[0]["path"] == "C:/new" and batches[0]["category"] == "kepsar"
    assert ii.register(batches, "b-1", "C:/new", {})["action"] == "unchanged"
    rec = ii.register(batches, "b-2", "C:/b2", {"category": "mattor_inredning", "org": "evil", "name": "B2"})
    assert rec["action"] == "registered"
    assert batches[1]["org"] == "cama" and batches[1]["published"] is False and batches[1]["name"] == "B2"


def test_ingest_end_to_end_without_server(env, monkeypatch, capsys):
    put(env.inbox / "b-1", {"p/a.jpg": "A"})
    (env.inbox / "b-1.json").write_text(json.dumps({"category": "mattor_inredning", "name": "B1"}), encoding="utf-8")
    monkeypatch.setattr(ii, "live_count", lambda bid: {"status": 200, "products": 1, "shots": 1})
    monkeypatch.setattr(ii.subprocess, "run", lambda *a, **k: types.SimpleNamespace(returncode=0, stdout='{"verdict": "ALL_IMAGES_P"}'))
    args = types.SimpleNamespace(batch="b-1", name=None, category=None, org=None, published=False, dry_run=False)
    assert ii.cmd_ingest(args, str(env.inbox)) == 0
    cfg = json.loads(env.cfg.read_text(encoding="utf-8"))
    assert cfg[0]["id"] == "b-1" and cfg[0]["path"] == str(env.batcher / "b-1") and cfg[0]["name"] == "B1"
    assert (env.batcher / "b-1" / "p" / "a.jpg").read_bytes() == b"A"
    log = (env.inbox / "_ingested.jsonl").read_text(encoding="utf-8").splitlines()
    assert json.loads(log[0])["status"] == "copied_verified"
    # same delivery again: no copy, config unchanged
    assert ii.cmd_ingest(args, str(env.inbox)) == 0
    assert json.loads(capsys.readouterr().out.strip().splitlines()[-1])["config"]["action"] == "unchanged"


def test_dry_run_writes_nothing(env, monkeypatch):
    put(env.inbox / "b-1", {"a.jpg": "A"})
    monkeypatch.setattr(ii, "live_count", lambda bid: {"status": 404})
    args = types.SimpleNamespace(batch="b-1", name=None, category="kepsar", org=None, published=False, dry_run=True)
    assert ii.cmd_ingest(args, str(env.inbox)) == 0
    assert not (env.batcher / "b-1").exists() and env.cfg.read_text(encoding="utf-8") == "[]"
    assert not (env.inbox / "_ingested.jsonl").exists()


def test_new_batch_without_category_refused(env, monkeypatch):
    put(env.inbox / "b-1", {"a.jpg": "A"})
    args = types.SimpleNamespace(batch="b-1", name=None, category=None, org=None, published=False, dry_run=True)
    assert ii.cmd_ingest(args, str(env.inbox)) == 2
