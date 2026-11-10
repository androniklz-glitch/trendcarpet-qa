"""Offline tests for scripts/pictures_godkanda.py (loaded by file; no hub server, no OneDrive)."""
import importlib.util
import json
import os

import pytest

HUB = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_spec = importlib.util.spec_from_file_location("pg", os.path.join(HUB, "scripts", "pictures_godkanda.py"))
pg = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(pg)


def rec(tmp, bid, pid, name, files, stars=None, org="cama"):
    d = tmp / bid
    d.mkdir(exist_ok=True)
    shots = []
    for f in files:
        (d / f).write_bytes(f"{bid}/{f}".encode())
        shots.append({"src": str(d / f), "filename": f, "gold": False})
    return {"batch": {"id": bid, "path": str(d), "org": org, "category": "mattor_inredning"},
            "product": {"id": pid, "display_name": name}, "shots": shots, "stars": stars or {}}


@pytest.mark.parametrize("pid,name,cat", [
    ("218976", "[218976] Badrumsmatta - Alvona (creme, 50 x 80 cm, 218976)", "Mattor - Badrum"),
    ("219500", "[219500] Ryamattor - Bianca Hand Woven Wool Premium (Creme, 219500 — 30 cm Ullpäls)", "Mattor - Rya & Shaggy"),
    ("Kuddfodral-Avafors-Wool-Bubble-beige", "Kuddfodral-Avafors-Wool-Bubble-beige", "Kuddar"),
    ("Hattar-Garda-Napoli-Fedora-Wool-Hat-gra", "Garda", "Hattar"),
    ("Keps-Gårda-Hunter-Army-Cap-Blå", "Gårda", "Kepsar"),
    ("Wiltonmatta-Auckland-terrakotta", "Auckland", "Mattor - Wilton & Print"),
    ("4101+4102", "[4101+4102] baddsoffa-kap-verde-gul", "Rum & Möbler"),
])
def test_category(pid, name, cat):
    assert pg.category_for({"id": "x"}, {"id": pid, "display_name": name}) == cat


def test_product_name_keeps_model_and_sku():
    p = {"id": "219500", "display_name": "[219500] Ryamattor - Bianca Hand Woven Wool Premium (Creme, 219500 — 30 cm Ullpäls)"}
    assert pg.sku_for(p) == "219500"
    n = pg.product_name(p, "219500")
    assert n.startswith("Ryamattor - Bianca") and n.endswith("(219500)") and n.count("219500") == 1


def test_same_sku_in_two_batches_one_folder_no_name_collision(tmp_path):
    a = rec(tmp_path, "26-11-05-a", "219500", "[219500] Ryamattor - Bianca", ["01.jpg", "02.jpg"])
    b = rec(tmp_path, "26-11-09-b", "219500", "[219500] Ryamattor - Bianca", ["01.jpg"])
    items, products = pg.build_plan([a, b], str(tmp_path / "T"))
    assert len(products) == 1
    imgs = [i["dest_rel"] for i in items if i["kind"] == "image"]
    assert len(imgs) == 3 and len({x.lower() for x in imgs}) == 3
    assert {os.path.basename(x) for x in imgs} == {"261105_01.jpg", "261105_02.jpg", "261109_01.jpg"}


def test_cover_is_best_rated(tmp_path):
    r = rec(tmp_path, "26-11-06-c", "218976", "[218976] Badrumsmatta - Alvona", ["01.jpg", "02.png"],
            stars={"218976|02.png": 4.5, "218976|01.jpg": 3.0})
    items, _ = pg.build_plan([r], str(tmp_path / "T"))
    cover = [i for i in items if i["kind"] == "cover"]
    assert len(cover) == 1 and cover[0]["src"].endswith("02.png") and cover[0]["dest_rel"].endswith("00_omslag.png")


def test_stars_mean_over_reviewers_ignores_junk():
    s = pg.stars_by_key({"A": {"p|1.jpg": {"stars": 5}}, "B": {"p|1.jpg": {"stars": 3}, "p|2.jpg": {"stars": True}}, "C": []})
    assert s == {"p|1.jpg": 4.0}


def test_paths_stay_under_max(tmp_path):
    long = "x" * 200 + ".jpg"
    r = rec(tmp_path, "26-11-06-d", "218976", "[218976] " + "Badrumsmatta " * 10, [long])
    target = "C:\\" + "T" * 76  # same length as the real target (79); build_plan never touches the target on disk
    items, _ = pg.build_plan([r], target)
    assert all(len(os.path.join(target, i["dest_rel"])) <= pg.MAX_PATH for i in items)


class FakeSS:
    class SelectionError(Exception):
        pass

    def __init__(self, batches, products):
        self._b, self._p = batches, products
        self.hub = self

    def select_batches(self, args):
        return list(self._b), []  # simulates a broken first org filter: Reforma comes through

    def org_decision(self, b):
        return None if b.get("org") == "cama" else "excluded_org_reforma"

    def local_gate(self, b, pictures):
        pass

    def _count_removed(self, root):
        return 0

    def load_ratings(self, bid):
        return {}

    def load_batches(self):
        return list(self._b)

    def scan_batch_products(self, b):
        return self._p[b["id"]]


def test_reforma_never_collected_even_if_first_filter_fails(tmp_path):
    for bid in ("cama-1", "ref-1"):
        (tmp_path / bid).mkdir()
        (tmp_path / bid / "a.jpg").write_bytes(b"x")
    batches = [{"id": "cama-1", "path": str(tmp_path / "cama-1"), "org": "cama"},
               {"id": "ref-1", "path": str(tmp_path / "ref-1"), "org": "reforma"}]
    prod = lambda: [{"id": "p", "variants": {"Original": [{"rel_path": "a.jpg", "filename": "a.jpg"}]}}]  # noqa: E731
    ss = FakeSS(batches, {"cama-1": prod(), "ref-1": prod()})
    _b, records, alarms, _r, skipped = pg.collect(ss, str(tmp_path), str(tmp_path / "T"))
    assert [r["batch"]["id"] for r in records] == ["cama-1"] and not alarms
    assert {"batch": "ref-1", "reason": "excluded_org_reforma"} in skipped


def test_guard_refuses_hub_batch_inside_target(tmp_path):
    target = tmp_path / "T"
    (target / "b").mkdir(parents=True)
    ss = FakeSS([{"id": "b", "path": str(target / "b"), "org": "cama"}], {})
    with pytest.raises(pg.Alarm):
        pg.guard_target(ss, str(target))


def test_sync_item_copy_unchanged_overwrite_conflict(tmp_path, monkeypatch):
    monkeypatch.setattr(pg, "MANIFEST_FILE", str(tmp_path / "m" / "manifest.jsonl"))
    src = tmp_path / "s.jpg"
    src.write_bytes(b"one")
    target, latest = str(tmp_path / "T"), {}
    it = {"src": str(src), "dest_rel": os.path.join("Kat", "Prod", "s.jpg"), "batch": "b", "kind": "image"}
    assert pg.sync_item(it, target, latest, False) == "copied"
    assert pg.sync_item(it, target, latest, False) == "unchanged"
    src.write_bytes(b"two-changed")
    assert pg.sync_item(it, target, latest, False) == "overwritten"
    assert open(os.path.join(target, it["dest_rel"]), "rb").read() == b"two-changed"
    with open(os.path.join(target, it["dest_rel"]), "wb") as f:
        f.write(b"edited by colleague")
    src.write_bytes(b"three")
    assert pg.sync_item(it, target, latest, False) == "conflict"
    assert open(os.path.join(target, it["dest_rel"]), "rb").read() == b"edited by colleague"
    assert len([json.loads(x) for x in open(pg.MANIFEST_FILE, encoding="utf-8")]) == 2


def test_dest_outside_target_refused(tmp_path):
    with pytest.raises(pg.Alarm):
        pg.dest_path(str(tmp_path / "T"), os.path.join("..", "evil.jpg"))
