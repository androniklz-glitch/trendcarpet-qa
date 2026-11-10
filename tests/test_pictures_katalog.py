"""Offline tests for scripts/pictures_katalog.py (fake target tree in tmp; no hub server, no OneDrive)."""
import json
import os
import re
import subprocess
import sys

from PIL import Image

HUB = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(HUB, "scripts"))
import pictures_katalog as k  # noqa: E402


def img(path, mode="RGB", size=(300, 200), color=(120, 90, 60)):
    path.parent.mkdir(parents=True, exist_ok=True)
    Image.new(mode, size, color if mode == "RGB" else (*color, 0)).save(path)


def make_tree(tmp, n_tc=13):
    t = tmp / "QA-hubb Godkända"
    for i in range(n_tc):  # 13 products = 2 pages in one category
        p = t / "Trendcarpet" / "Mattor - Ull" / f"Matta Ull {i:02d} (2{i:05d})"
        img(p / "00_omslag.jpg")
        img(p / "261110_a.jpg")
        img(p / "261110_b.png")
    p = t / "Trendcarpet" / "Kuddar" / "Kudde Łódź utan omslag"
    img(p / "261110_a.jpg")
    p = t / "Hatshop" / "Kepsar" / "Keps Svart"
    img(p / "00_omslag.png", mode="RGBA")  # transparent hero must render on white, not crash
    img(p / "261110_a.png")
    (t / "LÄS MIG.txt").write_text("x", encoding="utf-8")
    return t


def pdf_pages(path):
    return len(re.findall(rb"/Type\s*/Page(?![a-z])", open(path, "rb").read()))


def snapshot(root):
    return sorted((os.path.relpath(os.path.join(d, n), root), os.path.getsize(os.path.join(d, n)))
                  for d, _ds, ns in os.walk(root) for n in ns)


def test_scan_counts_exclude_cover_and_flag_missing(tmp_path):
    tree = k.scan(str(make_tree(tmp_path)))
    assert set(tree) == {"Hatshop", "Trendcarpet"}
    ull = tree["Trendcarpet"]["Mattor - Ull"]
    assert len(ull) == 13 and all(p["count"] == 2 for p in ull)
    assert ull[0]["name"] == "Matta Ull 00 (200000)"  # sorted by name
    s = k.summary(tree)
    assert s["Trendcarpet"]["missing_cover"] == ["Kuddar\\Kudde Łódź utan omslag"]
    assert s["Trendcarpet"]["products"] == 14 and s["Trendcarpet"]["images"] == 27
    assert s["Hatshop"] == {**s["Hatshop"], "products": 1, "images": 1, "pages": 2, "missing_cover": []}


def test_layout_page_numbers():
    cats = {"B": [{}] * 13, "A": [{}] * 12, "C": []}
    assert k.layout(cats) == [("A", 2, 1), ("B", 3, 2), ("C", 5, 1)]


def test_plural():
    assert (k.nb(1, "bild", "bilder"), k.nb(2, "bild", "bilder"), k.nb(0, "bild", "bilder")) == ("1 bild", "2 bilder", "0 bilder")


def test_text_fit_ellipsis():
    d = k.ImageDraw.Draw(Image.new("RGB", (10, 10)))
    out = k.text_fit(d, "ett mycket långt produktnamn som aldrig får plats på två rader alls", k.font(22, True), 120, 2)
    assert len(out) == 2 and out[-1].endswith("…")
    assert k.text_fit(d, "Kort", k.font(22), 400) == ["Kort"]


def test_cli_plan_writes_nothing(tmp_path):
    t = make_tree(tmp_path)
    before = snapshot(t)
    r = subprocess.run([sys.executable, "-I", "-B", "-X", "utf8", os.path.join(HUB, "scripts", "pictures_katalog.py"), "plan",
                        "--target", str(t)], capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stderr
    res = json.loads(r.stdout)
    assert [os.path.basename(p) for p in res["would_write"]] == ["Katalog Hatshop.pdf", "Katalog Trendcarpet.pdf"]
    assert snapshot(t) == before


def test_cli_build_only_adds_two_pdfs(tmp_path):
    t = make_tree(tmp_path)
    before = snapshot(t)
    r = subprocess.run([sys.executable, "-I", "-B", "-X", "utf8", os.path.join(HUB, "scripts", "pictures_katalog.py"), "build",
                        "--target", str(t)], capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stderr
    res = json.loads(r.stdout)
    after = snapshot(t)
    added = sorted(set(n for n, _s in after) - set(n for n, _s in before))
    assert added == ["Katalog Hatshop.pdf", "Katalog Trendcarpet.pdf"]  # no .tmp left, nothing else touched
    assert [x for x in after if x[0] not in added] == before
    # Trendcarpet: cover + Kuddar 1 page + Mattor - Ull 2 pages = 4; Hatshop: cover + Kepsar = 2
    assert pdf_pages(t / "Katalog Trendcarpet.pdf") == 4 == res["brands"]["Trendcarpet"]["pdf_pages"]
    assert pdf_pages(t / "Katalog Hatshop.pdf") == 2 == res["brands"]["Hatshop"]["pdf_pages"]


def test_cli_build_out_dir_leaves_target_untouched(tmp_path):
    t = make_tree(tmp_path)
    before = snapshot(t)
    out = tmp_path / "preview"
    r = subprocess.run([sys.executable, "-I", "-B", "-X", "utf8", os.path.join(HUB, "scripts", "pictures_katalog.py"), "build",
                        "--target", str(t), "--out", str(out)], capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stderr
    assert snapshot(t) == before
    assert sorted(os.listdir(out)) == ["Katalog Hatshop.pdf", "Katalog Trendcarpet.pdf"]


def test_missing_target_is_alarm(tmp_path):
    r = subprocess.run([sys.executable, "-I", "-B", "-X", "utf8", os.path.join(HUB, "scripts", "pictures_katalog.py"), "plan",
                        "--target", str(tmp_path / "nope")], capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 3 and "ALARM" in r.stdout


def test_unreadable_cover_renders_placeholder(tmp_path):
    t = make_tree(tmp_path, n_tc=1)
    bad = t / "Trendcarpet" / "Mattor - Ull" / "Matta Ull 00 (200000)" / "00_omslag.jpg"
    bad.write_bytes(b"not an image")
    assert k.thumb(str(bad), (100, 100)) is None
    n = k.render_brand("Trendcarpet", k.scan(str(t))["Trendcarpet"], str(tmp_path / "x.pdf"))
    assert n == 3 and pdf_pages(tmp_path / "x.pdf") == 3
