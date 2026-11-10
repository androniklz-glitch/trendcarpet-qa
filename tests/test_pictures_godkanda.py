"""Offline tests for scripts/pictures_godkanda.py (loaded by file; no hub server, no OneDrive)."""
import importlib.util
import json
import os
import re

import pytest

HUB = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_spec = importlib.util.spec_from_file_location("pg", os.path.join(HUB, "scripts", "pictures_godkanda.py"))
pg = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(pg)
REF = {"skus": {"502850", "644-65"}, "handles": ["baddsoffa-kap-verde-gul-mlm-502850", "byra-myshult-120x40cm-natur-2374-1-natur"],
       "file": "fixture"}  # Reforma list fixture (the real one is never read by tests)


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
    items, products, _ba = pg.build_plan([a, b], str(tmp_path / "T"))
    assert len(products) == 1
    imgs = [i["dest_rel"] for i in items if i["kind"] == "image"]
    assert len(imgs) == 3 and len({x.lower() for x in imgs}) == 3
    assert {os.path.basename(x) for x in imgs} == {"261105_01.jpg", "261105_02.jpg", "261109_01.jpg"}


def test_cover_is_best_rated(tmp_path):
    r = rec(tmp_path, "26-11-06-c", "218976", "[218976] Badrumsmatta - Alvona", ["01.jpg", "02.png"],
            stars={"218976|02.png": 4.5, "218976|01.jpg": 3.0})
    items, _, _ba = pg.build_plan([r], str(tmp_path / "T"))
    cover = [i for i in items if i["kind"] == "cover"]
    assert len(cover) == 1 and cover[0]["src"].endswith("02.png") and cover[0]["dest_rel"].endswith("00_omslag.png")


def test_stars_mean_over_reviewers_ignores_junk():
    s = pg.stars_by_key({"A": {"p|1.jpg": {"stars": 5}}, "B": {"p|1.jpg": {"stars": 3}, "p|2.jpg": {"stars": True}}, "C": []})
    assert s == {"p|1.jpg": 4.0}


def test_paths_stay_under_max(tmp_path):
    long = "x" * 200 + ".jpg"
    r = rec(tmp_path, "26-11-06-d", "218976", "[218976] " + "Badrumsmatta " * 10, [long])
    target = "C:\\" + "T" * 76  # same length as the real target (79); build_plan never touches the target on disk
    items, _, _ba = pg.build_plan([r], target)
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
    _b, records, alarms, _r, skipped, _h = pg.collect(ss, str(tmp_path), str(tmp_path / "T"), REF)
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


@pytest.mark.parametrize("pid,dn,ctx,name,cat", [
    ("Ryamattor-Merida-Soft-Shaggy-Bla", "Merida", "", "Ryamattor Merida Soft Shaggy Bla", "Mattor - Rya & Shaggy"),
    ("In-Out-5281-Brown-Matta-carpet-teppich", "Out", "", "In Out 5281 Brown Matta", "Mattor - Inomhus & Utomhus"),
    ("Trasmatta - Happy (gulmulti)", "Trasmatta", "", "Trasmatta Happy (gulmulti)", "Mattor - Trasmattor"),
    ("Inomhusutomhus-matta-Arco-Recycled-PET-creme", "matta", "", "Inomhusutomhus matta Arco Recycled PET creme",
     "Mattor - Inomhus & Utomhus"),
    ("Viskosmatta-Jodhpur-Special-Luxury-Edition-Ljusbeige", "Jodhpur", "", "Viskosmatta Jodhpur Special Luxury Edition Ljusbeige",
     "Mattor - Viskos"),
    ("1500px", "1500px", "barnmatta-ice-hockey-bla-carpet-wollteppich-ullteppe fix", "barnmatta ice hockey bla fix 1500px",
     "Mattor - Barn"),
])
def test_generic_display_name_uses_id_or_folder(pid, dn, ctx, name, cat):
    p = {"id": pid, "display_name": dn}
    assert pg.product_name(p, "", ctx) == name
    assert pg.category_for({"id": "x"}, p, ctx) == cat


def test_generic_subfolder_products_in_two_folders_stay_apart(tmp_path):
    b = {"id": "26-10-08-mattor-mix-fix", "path": str(tmp_path / "b"), "org": "cama", "category": "mattor_inredning"}
    recs = []
    for folder in ("barnmatta-ice-hockey-bla fix", "wiltonmatta-levin fix"):
        d = tmp_path / "b" / folder / "1500px"
        d.mkdir(parents=True)
        (d / "a.jpg").write_bytes(folder.encode())
        recs.append({"batch": b, "product": {"id": "1500px", "display_name": "1500px"},
                     "shots": [{"src": str(d / "a.jpg"), "filename": "a.jpg", "gold": False}], "stars": {}})
    _items, products, _ba = pg.build_plan(recs, str(tmp_path / "t"))
    folders = sorted(e["folder_rel"] for e in products.values())
    assert folders == [os.path.join("Trendcarpet", "Mattor - Barn", "barnmatta ice hockey bla fix 1500px"),
                       os.path.join("Trendcarpet", "Mattor - Wilton & Print", "wiltonmatta levin fix 1500px")]


def test_sku_paren_balanced_and_same_name_gets_colour(tmp_path):
    p = {"id": "218976", "display_name": "[218976] Badrumsmatta - Alvona (creme, 50 x 80 cm, 218976)"}
    assert pg.product_name(p, "218976") == "Badrumsmatta - Alvona (creme, 50 x 80 cm) (218976)"
    b = {"id": "26-09-23-hs-black-river-x", "path": str(tmp_path / "b"), "org": "cama", "category": "kepsar"}
    recs = []
    for col in ("Khaki", "Svart"):
        d = tmp_path / "b" / col
        d.mkdir(parents=True)
        (d / "a.jpg").write_bytes(col.encode())
        recs.append({"batch": b, "product": {"id": f"Keps-Black-River-Legacy-Trucker-{col}", "display_name": "Black River Legacy Trucker",
                                             "color": col}, "shots": [{"src": str(d / "a.jpg"), "filename": "a.jpg", "gold": False}], "stars": {}})
    _items, products, _ba = pg.build_plan(recs, str(tmp_path / "t"))
    assert sorted(e["folder_rel"] for e in products.values()) == [os.path.join("Hatshop", "Kepsar", "Black River Legacy Trucker Khaki"),
                                                                   os.path.join("Hatshop", "Kepsar", "Black River Legacy Trucker Svart")]


# ---- 0569 points 8/11/12/13: brand level, Borttagna mirror, LÄS MIG text, Reforma content filter ----

def test_brand_for_metadata_first_doubt_is_alarm():
    assert pg.brand_for({"id": "a", "category": "kepsar"}, "Kepsar") == "Hatshop"
    assert pg.brand_for({"id": "a", "category": "mattor_inredning"}, "Gardiner") == "Trendcarpet"
    for batch, cat in (({"id": "a", "category": "okänd"}, "Gardiner"), ({"id": "a"}, "Kepsar"),
                       ({"id": "a", "category": "kepsar"}, "Gardiner"), ({"id": "a", "category": "mattor_inredning"}, "Kepsar")):
        with pytest.raises(pg.Alarm):
            pg.brand_for(batch, cat)


def test_build_plan_brand_level_and_two_brand_product_excluded(tmp_path):
    a = rec(tmp_path, "26-11-05-a", "219500", "[219500] Ryamattor - Bianca", ["01.jpg"])
    k = rec(tmp_path, "26-11-06-k", "Keps-Gårda-Hunter-Army-Cap-Blå", "Gårda", ["01.jpg"])
    k["batch"]["category"] = "kepsar"
    x1 = rec(tmp_path, "26-11-07-x", "218976", "[218976] Badrumsmatta - Alvona", ["01.jpg"])
    x2 = rec(tmp_path, "26-11-08-y", "218976", "[218976] Badrumsmatta - Alvona", ["02.jpg"])
    x2["batch"]["category"] = "kepsar"  # same product key from a kepsar batch: brand contradicts -> excluded
    items, products, alarms = pg.build_plan([a, k, x1, x2], str(tmp_path / "T"))
    assert {e["brand"] for e in products.values()} == {"Trendcarpet", "Hatshop"}
    assert all(i["dest_rel"].split(os.sep)[0] in ("Trendcarpet", "Hatshop") for i in items)
    assert [x["product_key"] for x in alarms] == ["218976"] and "218976" not in products
    assert not any("Alvona" in i["dest_rel"] for i in items)


@pytest.fixture
def mirror_env(tmp_path, monkeypatch):
    monkeypatch.setattr(pg, "MANIFEST_FILE", str(tmp_path / "m" / "manifest.jsonl"))
    monkeypatch.setattr(pg, "MOVES_FILE", str(tmp_path / "m" / "moves.jsonl"))
    return str(tmp_path / "Godk"), str(tmp_path / "Bort")


def run_sync(recs, target, bort, ok=("26-11-05-a",), live=("26-11-05-a",), dry=False, allow_mass=False):
    """Same order as cmd_plan_or_sync: build_plan -> move_back -> sync_item -> mirror_removed."""
    items, _p, _ba = pg.build_plan(recs, target)
    latest = pg.load_manifest()
    back = pg.move_back(items, latest, target, bort, dry)
    for it in items:
        pg.sync_item(it, target, latest, dry)
    return items, back, pg.mirror_removed(latest, items, target, bort, ok, live, dry, allow_mass)


def moves():
    return [json.loads(x) for x in open(pg.MOVES_FILE, encoding="utf-8")] if os.path.isfile(pg.MOVES_FILE) else []


def two_image_rec(tmp_path):
    return rec(tmp_path, "26-11-05-a", "219500", "[219500] Ryamattor - Bianca", ["01.jpg", "02.jpg"],
               stars={"219500|01.jpg": 5})


def test_mirror_moves_dropped_image_then_back(tmp_path, mirror_env):
    target, bort = mirror_env
    r = two_image_rec(tmp_path)
    items, _b, m = run_sync([r], target, bort)
    assert not m["moved"] and len(items) == 3
    dropped = next(i for i in items if i["dest_rel"].endswith("261105_02.jpg"))
    digest = pg.sha256(os.path.join(target, dropped["dest_rel"]))
    r2 = dict(r, shots=r["shots"][:1])
    _i, _b, m = run_sync([r2], target, bort)
    assert [x["dest_rel"] for x in m["moved"]] == [dropped["dest_rel"]]
    assert not os.path.exists(os.path.join(target, dropped["dest_rel"]))
    assert pg.sha256(os.path.join(bort, dropped["dest_rel"])) == digest
    assert pg.load_manifest()[dropped["dest_rel"].lower()]["state"] == "borttagen"
    assert [x["action"] for x in moves()] == ["to_borttagna"] and moves()[0]["sha256_after"] == digest
    _i, back, m = run_sync([r], target, bort)  # image returns to the selection
    assert [b["status"] for b in back] == ["moved_back"] and not m["moved"]
    assert pg.sha256(os.path.join(target, dropped["dest_rel"])) == digest
    assert not os.path.exists(os.path.join(bort, dropped["dest_rel"]))
    assert pg.is_active(pg.load_manifest()[dropped["dest_rel"].lower()])
    assert [x["action"] for x in moves()] == ["to_borttagna", "back"]


def test_mirror_empty_product_folder_goes_cover_too(tmp_path, mirror_env):
    target, bort = mirror_env
    items, _b, _m = run_sync([two_image_rec(tmp_path)], target, bort)
    prod_dir = os.path.join(target, os.path.dirname(items[0]["dest_rel"]))
    assert os.path.isdir(prod_dir)
    _i, _b, m = run_sync([], target, bort)
    assert len(m["moved"]) == 3 and any(x["dest_rel"].endswith("00_omslag.jpg") for x in m["moved"])
    assert not os.path.exists(prod_dir) and not os.path.exists(os.path.join(target, "Trendcarpet")) and os.path.isdir(target)
    assert all(os.path.isfile(os.path.join(bort, i["dest_rel"])) for i in items)


def test_mirror_dry_run_changes_nothing(tmp_path, mirror_env):
    target, bort = mirror_env
    r = two_image_rec(tmp_path)
    run_sync([r], target, bort)
    man = open(pg.MANIFEST_FILE, encoding="utf-8").read()
    files = sorted(os.path.join(d, f) for d, _s, fs in os.walk(target) for f in fs)
    _i, _b, m = run_sync([dict(r, shots=r["shots"][:1])], target, bort, dry=True)
    assert len(m["would_move"]) == 1 and not m["moved"]
    assert sorted(os.path.join(d, f) for d, _s, fs in os.walk(target) for f in fs) == files
    assert open(pg.MANIFEST_FILE, encoding="utf-8").read() == man and not moves() and not os.path.exists(bort)


def test_mirror_mass_guard(tmp_path, mirror_env, monkeypatch):
    target, bort = mirror_env
    monkeypatch.setattr(pg, "MASS_MOVE_MIN", 1)
    run_sync([two_image_rec(tmp_path)], target, bort)
    _i, _b, m = run_sync([], target, bort)  # 3 candidates > max(1, 0.1 * 3)
    assert not m["moved"] and any("mass move refused: 3 of 3" in a["message"] for a in m["alarms"])
    assert len(os.listdir(os.path.join(target, "Trendcarpet", "Mattor - Rya & Shaggy"))) == 1
    _i, _b, m = run_sync([], target, bort, allow_mass=True)
    assert len(m["moved"]) == 3 and not m["alarms"]


def test_mirror_held_live_alarmed_batch_and_unpublished_moves(tmp_path, mirror_env):
    target, bort = mirror_env
    run_sync([two_image_rec(tmp_path)], target, bort)
    _i, _b, m = run_sync([], target, bort, ok=(), live=("26-11-05-a",))  # live but alarmed (local_gate) -> held
    assert len(m["held"]) == 3 and not m["moved"] and not os.path.exists(bort)
    _i, _b, m = run_sync([], target, bort, ok=(), live=())  # batch unpublished -> moved
    assert len(m["moved"]) == 3 and {x["reason"] for x in m["moved"]} == {"batch avpublicerad"}


def test_mirror_destination_clash_gets_timestamp(tmp_path, mirror_env):
    target, bort = mirror_env
    r = two_image_rec(tmp_path)
    items, _b, _m = run_sync([r], target, bort)
    rel = next(i["dest_rel"] for i in items if i["dest_rel"].endswith("261105_02.jpg"))
    os.makedirs(os.path.dirname(os.path.join(bort, rel)))
    with open(os.path.join(bort, rel), "wb") as f:
        f.write(b"older removal")
    _i, _b, m = run_sync([dict(r, shots=r["shots"][:1])], target, bort)
    assert len(m["moved"]) == 1
    assert open(os.path.join(bort, rel), "rb").read() == b"older removal"
    stamped = [f for f in os.listdir(os.path.dirname(os.path.join(bort, rel))) if f.startswith("261105_02_")]
    assert len(stamped) == 1 and pg.load_manifest()[rel.lower()]["borttagen_rel"].endswith(stamped[0])


def test_move_back_skips_when_source_changed(tmp_path, mirror_env):
    target, bort = mirror_env
    r = two_image_rec(tmp_path)
    run_sync([r], target, bort)
    run_sync([dict(r, shots=r["shots"][:1])], target, bort)
    with open(r["shots"][1]["src"], "wb") as f:
        f.write(b"new version of 02")
    items, back, _m = run_sync([r], target, bort)
    rel = next(i["dest_rel"] for i in items if i["dest_rel"].endswith("261105_02.jpg"))
    assert [b["status"] for b in back] == ["recopy"]
    assert open(os.path.join(target, rel), "rb").read() == b"new version of 02" and os.path.isfile(os.path.join(bort, rel))


def test_mirror_reforma_copy_held_with_alarm(tmp_path, mirror_env):
    target, bort = mirror_env
    items, _p, _ba = pg.build_plan([two_image_rec(tmp_path)], target)
    latest = pg.load_manifest()
    for it in items:
        pg.sync_item(it, target, latest, False)
    m = pg.mirror_removed(latest, [], target, bort, (), ("26-11-05-a",), False, reforma_batches={"26-11-05-a"})
    assert len(m["held"]) == 3 and m["alarms"] and not m["moved"] and not os.path.exists(bort)


def test_readme_descriptive_only(tmp_path):
    a = rec(tmp_path, "26-11-05-a", "219500", "[219500] Ryamattor - Bianca", ["01.jpg"])
    k = rec(tmp_path, "26-11-06-k", "Keps-Gårda-Hunter-Army-Cap-Blå", "Gårda", ["01.jpg"])
    k["batch"]["category"] = "kepsar"
    _i, products, _ba = pg.build_plan([a, k], str(tmp_path / "T"))
    t = pg.readme_text(products)
    for must in ("uppdateras automatiskt", "Varumärke > Kategori > Produkt", "Hatshop", "Trendcarpet", "00_omslag", "ÅÅMMDD",
                 "underkända", "Frågor om innehållet: Andronik Lindgren."):
        assert must in t
    assert not re.search(r"(?i)\b(ändra|redigera|radera|byt namn|lägg till|flytta|borttagna|removed)\b", t)


@pytest.mark.parametrize("product,why", [
    ({"id": "4101+4102", "display_name": "[4101+4102] baddsoffa-kap-verde-gul"}, "möbel"),
    ({"id": "644-65", "display_name": "Lampa"}, "Reforma-SKU"),
    ({"id": "x", "display_name": "Bord (502850)"}, "Reforma-SKU"),
    ({"id": "byra-myshult-120x40cm"}, "möbel"),
    ({"id": "spegel-myshult-natur"}, ""),  # handle prefix must match a real handle at a word boundary
    ({"id": "219500", "display_name": "[219500] Ryamattor - Bianca"}, ""),
    ({"id": "Gardiner-Lumia-Linne-Beige", "display_name": "Lumia"}, ""),
])
def test_reforma_product_hit(product, why):
    hit = pg.reforma_product_hit(product, REF)
    assert hit.startswith(why) if why else hit == ""


def test_reforma_handle_prefix_without_furniture_word():
    ref = {"skus": set(), "handles": ["vas-kap-verde-gul-mlm-1"], "file": "fixture"}
    assert pg.reforma_product_hit({"id": "vas-kap-verde"}, ref).startswith("Reforma-handle")
    assert pg.reforma_product_hit({"id": "vas-kap"}, ref) == ""  # < 3 words never matches


def test_collect_reforma_content_filter_zero_files(tmp_path):
    bids = ("26-10-19-ag-reforma-interiors", "26-10-01-flow-vibes-rum", "26-10-02-flow-vibes-mattor")
    for bid in bids:
        (tmp_path / bid).mkdir()
        for f in ("a.jpg", "b.jpg"):
            (tmp_path / bid / f).write_bytes(f.encode())
    batches = [{"id": b, "path": str(tmp_path / b), "org": "cama", "category": "mattor_inredning"} for b in bids]
    v = lambda f: {"Original": [{"rel_path": f, "filename": f}]}  # noqa: E731
    prods = {b: [{"id": "Wiltonmatta-Auckland", "variants": v("a.jpg")}] for b in bids}
    prods["26-10-02-flow-vibes-mattor"].append({"id": "soffbord-ek", "variants": v("b.jpg")})
    _b, records, _a, _r, _s, hits = pg.collect(FakeSS(batches, prods), str(tmp_path), str(tmp_path / "T"), REF)
    assert [(r["batch"]["id"], r["product"]["id"]) for r in records] == [("26-10-02-flow-vibes-mattor", "Wiltonmatta-Auckland")]
    assert {h["batch"] for h in hits if h["product"] is None} == set(bids[:2])
    assert [h["product"] for h in hits if h["product"]] == ["soffbord-ek"]
    items, _p, _ba = pg.build_plan(records, str(tmp_path / "T"))
    assert {i["batch"] for i in items} == {"26-10-02-flow-vibes-mattor"} and not any(i["src"].endswith("b.jpg") for i in items)


def test_load_reforma_fails_closed(tmp_path):
    with pytest.raises(pg.Alarm):
        pg.load_reforma(str(tmp_path / "missing.json"))
    p = tmp_path / "small.json"
    p.write_text(json.dumps({"products": [{"handle": "a", "skus": ["123456"]}]}), encoding="utf-8")
    with pytest.raises(pg.Alarm):
        pg.load_reforma(str(p))
