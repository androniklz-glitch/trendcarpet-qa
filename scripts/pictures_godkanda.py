"""Distribution copy of approved hub images in the CEO's OneDrive Pictures (CEO 2026-11-10 "ja kör på pictures-mappen"):
  <OneDrive>\\Pictures\\QA-hubb Godkända\\<Brand>\\<Swedish category>\\<Product name (SKU)>\\<yymmdd>_<file>
Brand = Hatshop | Trendcarpet from the batch metadata (`category` kepsar | mattor_inredning; CEO "DEN SNYGGA MAPPEN KAN
FÖRST VARA INDELAD I HATSHOP RESPEKTIVE TRENDCARPET"); unknown metadata or a contradiction with the product category = ALARM,
the product is never placed. Plus 00_omslag.<ext> (best image by hub rating) per product and LÄS MIG.txt at the root
(descriptive only: brand -> category -> product -> count).
Selection (CEO-RÄTTELSE 5): every shot the hub shows in a batch folder; only images in "<batch> Removed" are underkända
(Removed is never scanned). Reforma is NEVER copied: org filter twice plus a content filter (batch id/name "reforma", the
flow-vibes-rum batch, furniture products, product id/SKU/handle in the Reforma shop list REFORMA_PRODUCTS_FILE; an unreadable
list = ALARM). A product in several batches gets ONE folder (key = SKU, else the product id); file names carry the batch
date so they never collide.
The target is a copy, never a hub source: the run refuses (exit 3) if any batches_config path lies inside it or inside
Pictures\\QA-hubb Borttagna, and `free` (+U) refuses hub sources (sharepoint_sync.free_refusal). Sources are read only,
never freed. Idempotent: unchanged files are skipped by stat signature, a changed source overwrites its copy (temp name + sha256
check + os.replace; a killed run's *.godk-tmp is removed next sync), a copy edited at the target is left alone and reported as conflict.
Mirror (CEO "ja ta bort removed automatiskt och rensa removed till separat mapp"): a copy whose image left the selection
(Removed, or batch unpublished) is MOVED, never deleted, to Pictures\\QA-hubb Borttagna\\<same relative path> (sha256
before == after), and moved back when the image returns; product folders that become empty are removed (rmdir only).
Copies of a live batch that failed its gate are held, never moved. A freed (+U) or cloud-only copy is pinned (attrib -U +P)
and moved on a later run once local; ALARM only after 3 runs in a row (pin_holds.json). A product already in the target
keeps its folder when a newer batch titles it differently (manifest product_key / '(SKU)'). More than max(50, 10 %) moves at
once is refused unless --allow-mass-move. Every move is logged in moves.jsonl, every copy/move in manifest.jsonl (~/qa-hub-backup/pictures-godkanda/).

Commands (hub python, cwd = hub dir):
  plan                dry-run: products + files + GB per brand and category, Removed excluded, ALARM batches, Reforma
                      exclusions, what would move to / back from Borttagna (read-only)
  sync [--allow-mass-move]   move back, copy/overwrite per plan, move left-selection copies to Borttagna, write LÄS MIG.txt
                      + pictures_godkanda_links.json, append manifest + moves log
  verify              target + Borttagna vs manifest by metadata (no download, +U safe) + C: free >= 20 GB; deviation -> exit 3
  restore <dest-prefix> [--dry-run]   copy back from the hub source when its sha256 equals the manifest
  free                +U on product folders OneDrive reports as uploaded (target only)
  reforma-scan        read-only: every published batch, every Reforma hit (strong = excluded) + weak model-word matches
Exit: 0 OK, 1 copy error, 2 usage, 3 ALARM, 4 lock held (a dead pid's or > 30 min old lock is removed, logged)."""
import argparse
import ctypes
import datetime
import hashlib
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
import urllib.parse

HUB_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TARGET_NAME = "QA-hubb Godkända"
BORTTAGNA_NAME = "QA-hubb Borttagna"
MANIFEST_FILE = os.path.join(os.path.expanduser("~"), "qa-hub-backup", "pictures-godkanda", "manifest.jsonl")
MOVES_FILE = os.path.join(os.path.expanduser("~"), "qa-hub-backup", "pictures-godkanda", "moves.jsonl")
MASS_MOVE_MIN, MASS_MOVE_FRACTION = 50, 0.10
BRANDS = {"kepsar": "Hatshop", "mattor_inredning": "Trendcarpet"}
HATSHOP_CATEGORIES = {"Kepsar", "Hattar"}
# Reforma shop list (reformahome.myshopify.com products.json, 2026-11-03; the dump the Reforma threads use, e.g.
# scripts/reforma/matrixify/diep_skugga_shopify.py). Read only; never copied into the hub.
REFORMA_PRODUCTS_FILE = os.path.join(os.path.expanduser("~"), "scratch", "Projects", "Reforma_FTP_gamla_sajten", "live_audit",
                                     "26-11-03-pos1", "products-2026-11-03.json")
REFORMA_BATCH_NEEDLES = ("reforma", "flow-vibes-rum", "flow-vibes-mattor")  # flow-vibes-rum: CEO 2026-11-10 "det låter som en
# reforma-mapp?"; flow-vibes-mattor: Koordinator 4.4 2026-11-10 (whole batch, all 13 product ids are exact Reforma handles)
FURNITURE_NEEDLES = ("soffbord", "baddsoffa", "soffa-", "sangbord", "sängbord", "matbord", "byra-", "tv-bank", "skank",
                     "fåtölj", "fatolj", "hylla", "stol-")  # Trendcarpet/Hatshop sell no furniture: suspected Reforma
LINKS_FILE = os.path.join(HUB_DIR, "pictures_godkanda_links.json")
LOCK_FILE = os.path.join(HUB_DIR, ".pictures_godkanda.lock")
HUB_URL = "https://trendcarpet-qa.vercel.app"
ONEDRIVE_WEB = "https://trendcarpet-my.sharepoint.com/personal/andronik_lindgren_trendcarpet_se/Documents"
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".gif", ".avif", ".heic"}
MAX_PATH = 255
MIN_FREE_GB = 20
ATTR_PINNED, ATTR_UNPINNED = 0x80000, 0x100000
ATTR_CLOUD_ONLY = 0x1000 | 0x40000 | 0x400000  # OFFLINE | RECALL_ON_OPEN | RECALL_ON_DATA_ACCESS: data not on disk
PIN_ALARM_RUNS = 3
TMP_SUFFIX = ".godk-tmp"  # copy_atomic's temp name next to the target file (4.11 N2)
LOCK_STALE_S = 30 * 60  # a lock older than this (or of a dead pid) is stale (4.10 gate point 1)

# (category, batch-id needles, product needles); first hit wins. Product text = product id + display name + colour.
CATEGORIES = [
    ("Mattor - Trappsteg", ("trappsteg",), ("trappsteg", "halvmåne", "halvmane")),
    ("Mattor - Badrum", ("badrumsmatt",), ("badrumsmatta", "alvona")),
    ("Kuddar", (), ("kuddfodral", "pillow")),
    ("Gardiner", ("curtain", "gardiner"), ("gardin", "curtain", "lumia", "lumira")),
    ("Hattar", (), ("hattar", "fedora", "wool-hat")),
    ("Kepsar", ("kepsar", "caps-", "army-h", "black-river", "garda-ranger"), ("keps", "trucker", "army-cap", "legacy")),
    ("Mattor - Trasmattor", (), ("trasmatta",)),
    ("Mattor - Inomhus & Utomhus", (), ("inomhusutomhus", "inomhus-utomhus", "in-out", "recycled-pet", "recycled pet")),
    ("Mattor - Viskos", (), ("viskos",)),
    ("Mattor - Rya & Shaggy", (),("rya", "shaggy", "merida", "bianca", "amalia", "kanha", "batumi", "aranga", "charonne",
                                   "charrone", "ullpäls")),
    ("Mattor - Barn", (), ("barnmatta", "basketball")),
    ("Mattor - Wilton & Print", ("print-",), ("wilton", "rund matta", "frome", "auckland", "barcelona", "alexandra", "artena")),
    ("Mattor - Ull", (), ("ullmatta", "agra", "avafors", "gingham", "wool")),
    ("Mattor - Bomull", (), ("bomullsmatta", "husvik", "bomull")),
    ("Mattor - Jute", (), ("jute",)),
]
SKU_RE = re.compile(r"(?<![0-9A-Za-z])(IK\d{9}|2\d{5}|26O\d{9})(?![0-9])", re.I)
NOISE_RE = re.compile(r"[-_ ]?(carpet|teppich|wollteppich|ullteppe|vorhang|curtain)\b", re.I)
FORBIDDEN = '<>:"/\\|?*'
GENERIC_NAMES = {"out", "original", "1500px", "1500px iphone", "matta", "alla per format", "alla per format 1x1", "trasmatta",
                 "bomullsmatta", "jutematta", "ullmatta", "ryamatta", "ryamattor", "wiltonmatta", "rund matta", "kontaktkartor"}


class Alarm(Exception):
    pass


def lp(path):
    """Long-path form for Win32 file APIs (Explorer-length paths are kept <= MAX_PATH anyway)."""
    p = os.path.abspath(path)
    return p if p.startswith("\\\\?\\") or os.name != "nt" else "\\\\?\\" + p


def category_for(batch, product, context=""):
    bid = (batch.get("id") or "").lower()
    text = " ".join([str(product.get(k) or "") for k in ("id", "display_name", "color")] + [context]).lower()
    for name, batch_needles, prod_needles in CATEGORIES:
        if any(n in bid for n in batch_needles) or any(n in text for n in prod_needles):
            return name
    return "Kepsar" if (batch.get("category") or "") == "kepsar" else "Mattor - Övriga"


def brand_for(batch, category):
    """Brand from the batch metadata first, the product category only as a cross-check: doubt = ALARM, never placed."""
    brand = BRANDS.get(str(batch.get("category") or "").strip().lower())
    if not brand:
        raise Alarm(f"batch {batch.get('id')}: batch category {batch.get('category')!r} gives no brand")
    if (category in HATSHOP_CATEGORIES) != (brand == "Hatshop"):
        raise Alarm(f"batch {batch.get('id')}: brand {brand} (batch metadata) contradicts category {category}")
    return brand


def load_reforma(path=None):
    """{'skus': set, 'handles': list} from the Reforma shop dump; missing/unreadable/implausibly small = ALARM (fail closed)."""
    path = path or REFORMA_PRODUCTS_FILE
    try:
        with open(lp(path), encoding="utf-8") as f:
            prods = json.load(f)["products"]
    except (OSError, ValueError, KeyError, TypeError) as exc:
        raise Alarm(f"Reforma product list unreadable ({path}): {exc}")
    if len(prods) < 100:
        raise Alarm(f"Reforma product list has only {len(prods)} products: {path}")
    skus = {str(s).strip().lower() for p in prods for s in (p.get("skus") or []) if s and len(str(s).strip()) >= 4}
    return {"skus": skus, "handles": sorted({str(p.get("handle") or "").lower() for p in prods} - {""}), "file": path}


def _product_text(product, context=""):
    return " ".join([str(product.get(k) or "") for k in ("id", "display_name", "color")] + [context]).lower()


def reforma_batch_hit(batch):
    if (batch.get("org") or "").strip().lower() == "reforma":
        return "org=reforma"
    text = f"{batch.get('id') or ''} {batch.get('name') or ''}".lower()
    hit = next((n for n in REFORMA_BATCH_NEEDLES if n in text), None)
    return f"batch-namn innehåller '{hit}'" if hit else ""


def reforma_product_hit(product, reforma, context=""):
    """Strong product-level Reforma match (excluded): furniture, exact SKU, or the product id slug equals / prefixes a
    Reforma handle at a word boundary (>= 3 words, e.g. 'baddsoffa-kap-verde-gul' -> 'baddsoffa-kap-verde-gul-mlm-502850')."""
    text = _product_text(product, context)
    hit = next((n for n in FURNITURE_NEEDLES if n in text), None)
    if hit and category_for({"id": ""}, product, context) == "Mattor - Övriga":  # a rug in a room with an armchair is no furniture

        return f"möbel ('{hit}')"
    raw = [str(product.get(k) or "") for k in ("id", "display_name")] + [sku_for(product)]
    cands = {x.strip().lower() for x in raw if x.strip()}
    cands |= {t for x in raw for t in re.split(r"[\s,+()\[\]/_]+", x.lower()) if len(t) >= 4}
    sku = next((c for c in sorted(cands) if c in reforma["skus"]), None)
    if sku:
        return f"Reforma-SKU '{sku}'"
    slug = re.sub(r"[\s_]+", "-", str(product.get("id") or "").strip().lower())
    if slug.count("-") >= 2:
        h = next((h for h in reforma["handles"] if h == slug or h.startswith(slug + "-")), None)
        if h:
            return f"Reforma-handle '{h}'"
    return ""


def reforma_weak_hits(product, reforma, context=""):
    """Model words (2nd word of a Reforma handle, >= 5 letters) found in the product text: reported, never excluded."""
    words = reforma.setdefault("_models", {w for h in reforma["handles"] for w in h.split("-")[1:2] if w.isalpha() and len(w) >= 5})
    return sorted(set(re.findall(r"[a-zåäö]{5,}", _product_text(product, context))) & words)


def sku_for(product):
    for k in ("id", "display_name"):
        m = SKU_RE.search(str(product.get(k) or ""))
        if m:
            return m.group(1).upper()
    return ""


def clean_component(s, limit):
    s = "".join("_" if c in FORBIDDEN or ord(c) < 32 else c for c in s)
    s = re.sub(r"\s+", " ", s).strip(" .-_")
    return s[:limit].rstrip(" .-_") or "Produkt"


def _slug_words(s):
    s = NOISE_RE.sub("", str(s or "")).replace("-", " ").replace("_", " ")
    return re.sub(r"\s+", " ", s).strip()


def is_generic(name):
    n = _slug_words(name).lower()
    return len(n.split()) <= 1 or n in GENERIC_NAMES


def base_name(product, context=""):
    """Readable product name before SKU/length handling. display_name is often only the model ("Merida", "Matta", "Out")
    while the id holds type + model + colour, and some old batches expose sub-folders ("1500px", "Original") as products:
    then fall back to the id, and when that is generic too, prefix the product's top folder in the batch (context)."""
    name = re.sub(r"^\[[^\]]*\]\s*", "", str(product.get("display_name") or product.get("id") or ""))
    if not is_generic(name):
        return name if " " in name.strip() else _slug_words(name)
    pid = _slug_words(product.get("id"))
    if pid and not is_generic(pid):
        return pid
    ctx = _slug_words(context)
    return f"{ctx} {pid or _slug_words(name)}".strip() if ctx else (pid or name)


def product_name(product, sku, context=""):
    name = base_name(product, context)
    for sep in (" (", " — ") if len(name) > 50 else ():  # " - " separates type and model ("Ryamattor - Bianca"): keep
        if sep in name[8:]:
            name = name[:8 + name[8:].index(sep)]
    if sku:
        name = re.sub(rf"[\s,(]*{re.escape(sku)}[\s,)]*", " ", name, flags=re.I)
    name = clean_component(name, 60)
    if name.count("(") > name.count(")"):  # "Alvona (creme, 50 x 80 cm, 218976)" lost its ")" with the SKU
        name += ")"
    return f"{name} ({sku})" if sku else name


def batch_date(batch):
    m = re.match(r"(\d{2})-(\d{2})-(\d{2})", batch.get("id") or "")
    if m:
        return "".join(m.groups())
    c = str(batch.get("createdAt") or "")
    return c[2:4] + c[5:7] + c[8:10] if re.match(r"\d{4}-\d{2}-\d{2}", c) else "000000"


def stars_by_key(ratings):
    """{ 'pid|filename': mean stars over reviewers }; malformed entries are ignored (rating only picks the cover)."""
    acc = {}
    for per in (ratings or {}).values():
        for key, v in (per or {}).items() if isinstance(per, dict) else ():
            s = v.get("stars") if isinstance(v, dict) else v
            if isinstance(s, (int, float)) and not isinstance(s, bool) and 1 <= s <= 5:
                acc.setdefault(key, []).append(float(s))
    return {k: sum(v) / len(v) for k, v in acc.items()}


FOLDER_SKU_RE = re.compile(r"\((IK\d{9}|2\d{5}|26O\d{9})\)", re.I)


def known_folders(latest):
    """{'by_key': {product key: folder_rel}, 'by_src': {source path lower: folder_rel}} from the ACTIVE manifest copies
    (covers excluded). Key = the stored product_key, else the '(SKU)' in the folder name; several folders for one key ->
    the one with most copies. Lets a known product keep its folder when a newer batch titles it differently (4.10 B)."""
    count, by_src = {}, {}
    for e in latest.values():
        if not is_active(e) or e.get("kind") == "cover":
            continue
        folder = os.path.dirname(e["dest_rel"])
        skus = FOLDER_SKU_RE.findall(os.path.basename(folder))
        key = e.get("product_key") or (skus[-1].upper() if skus else "")
        if key:
            c = count.setdefault(key, {})
            c[folder] = c.get(folder, 0) + 1
        by_src[e["src"].lower()] = folder
    by_key = {k: max(sorted(c), key=lambda f: c[f]) for k, c in count.items()}
    return {"by_key": by_key, "by_src": by_src}


def build_plan(records, target, known=None):
    """records: dicts {batch, product, shots:[{src, filename, gold}], stars:{key: mean}}. Returns (items, products,
    brand_alarms). items: {src, dest_rel, batch, kind, product_key}. One folder per product key under <brand>\\<category>; a
    key whose records give no brand or two brands is excluded and reported. Cover = best mean stars, then gold, then newest
    batch. known (known_folders): a product already in the target keeps its folder (same brand + category) instead of the
    newest batch's title, so a re-titled batch (even case only) never renames it into Borttagna (4.10 order B)."""
    products, brand_alarms = {}, []
    for r in sorted(records, key=lambda r: (batch_date(r["batch"]), r["batch"]["id"]), reverse=True):
        p, sku = r["product"], sku_for(r["product"])
        rel = os.path.relpath(r["shots"][0]["src"], r["batch"]["path"]) if r["shots"] and r["batch"].get("path") else ""
        ctx = rel.replace("/", os.sep).split(os.sep)[0] if os.sep in rel.replace("/", os.sep) else ""
        key = sku or clean_component(str(p.get("id") or ""), 80).lower()
        if not sku and ctx and is_generic(str(p.get("id") or "")):
            key = clean_component(f"{ctx} {p.get('id')}", 80).lower()  # "1500px" in two product folders = two products
        e = products.setdefault(key, {"category": category_for(r["batch"], p, ctx), "folder": product_name(p, sku, ctx), "color": p.get("color"),
                                      "files": [], "batches": [], "seen": set(), "brands": set(), "brand_errors": []})
        try:
            e["brands"].add(brand_for(r["batch"], e["category"]))
        except Alarm as exc:
            e["brand_errors"].append(str(exc))
        if r["batch"]["id"] not in e["batches"]:
            e["batches"].append(r["batch"]["id"])
        for s in r["shots"]:
            if s["src"].lower() in e["seen"]:
                continue
            e["seen"].add(s["src"].lower())
            e["files"].append({**s, "batch": r["batch"]["id"], "date": batch_date(r["batch"]),
                               "stars": r["stars"].get(f"{p.get('id')}|{s['filename']}", 0.0)})
    for key in sorted(products):
        e = products[key]
        if e["brand_errors"] or len(e["brands"]) != 1:
            brand_alarms.append({"product_key": key, "batches": e["batches"],
                                 "message": "; ".join(e["brand_errors"]) or f"two brands: {sorted(e['brands'])}"})
            del products[key]
        else:
            e["brand"] = next(iter(e["brands"]))
    items, used = [], set()
    for key, e in sorted(products.items()) if known else ():
        kf = known["by_key"].get(key)
        if not kf:
            hits = [known["by_src"][f["src"].lower()] for f in e["files"] if f["src"].lower() in known["by_src"]]
            kf = max(sorted(set(hits)), key=hits.count) if hits else None
        if kf and os.path.dirname(kf).lower() == os.path.join(e["brand"], e["category"]).lower() and kf.lower() not in used:
            e["folder_rel"] = kf
            used.add(kf.lower())
    name_count = {}
    for e in products.values():
        nk = (e["brand"], e["category"], e["folder"].lower())
        name_count[nk] = name_count.get(nk, 0) + 1
    for key, e in sorted(products.items()):
        if "folder_rel" not in e:
            folder = e["folder"]
            col = clean_component(str(e.get("color") or ""), 30) if e.get("color") else ""
            if name_count[(e["brand"], e["category"], folder.lower())] > 1 and col and col.lower() not in folder.lower():
                folder = f"{folder} {col}"  # same name, different colour: "Black River Legacy Trucker Khaki", not "(2)"
            base = folder_rel = os.path.join(e["brand"], e["category"], folder)
            # Two product keys with the same display folder (same name, no SKU), or a name a known product holds: suffix.
            n = 1
            while folder_rel.lower() in used:
                n += 1
                folder_rel = f"{base} ({n})"
            used.add(folder_rel.lower())
            e["folder_rel"] = folder_rel
        folder_rel = e["folder_rel"]
        names = set()
        room = MAX_PATH - len(os.path.join(target, folder_rel)) - 1
        for f in e["files"]:
            stem, ext = os.path.splitext(os.path.basename(f["src"]))
            base = clean_component(f"{f['date']}_{stem}", max(12, room - len(ext) - 3))
            name, i = base + ext.lower(), 2
            while name.lower() in names:
                name, i = f"{base}_{i}{ext.lower()}", i + 1
            names.add(name.lower())
            items.append({"src": f["src"], "dest_rel": os.path.join(folder_rel, name), "batch": f["batch"], "kind": "image",
                          "product_key": key})
        imgs = [f for f in e["files"] if os.path.splitext(f["src"])[1].lower() in IMAGE_EXTS]
        if imgs:
            best = max(imgs, key=lambda f: (f["stars"], f.get("gold", False), f["date"]))
            items.append({"src": best["src"], "batch": best["batch"], "kind": "cover", "product_key": key,
                          "dest_rel": os.path.join(folder_rel, "00_omslag" + os.path.splitext(best["src"])[1].lower())})
        e["count"] = len(e["files"])
    return items, products, brand_alarms


def load_sync():
    sys.path[:0] = [HUB_DIR, os.path.join(HUB_DIR, "scripts")]
    spec = importlib.util.spec_from_file_location("sharepoint_sync", os.path.join(HUB_DIR, "scripts", "sharepoint_sync.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def inside(path, root):
    a, r = os.path.normcase(os.path.realpath(path)), os.path.normcase(os.path.realpath(root))
    return a == r or a.startswith(r.rstrip(os.sep) + os.sep)


def guard_target(ss, target):
    for b in ss.hub.load_batches():
        src = (b.get("path") or "").strip()
        if src and (inside(src, target) or inside(target, src)):
            raise Alarm(f"hub batch {b.get('id')} overlaps the target {target}: target must never be a hub source")


def collect(ss, pictures, target, reforma=None):
    """Returns (batches, records, alarms, removed, skipped, reforma_hits); reforma_hits = [{batch, product, reason, files,
    srcs}] (product None = whole batch). Reforma content filter on top of the org filter; the list is loaded fail-closed."""
    reforma = reforma if reforma is not None else load_reforma()
    args = type("A", (), {"batch": None, "all": True, "include_drafts": False})()
    batches, skipped = ss.select_batches(args)
    records, alarms, removed, hits = [], [], {}, []
    for b in batches:
        if ss.org_decision(b):  # second, independent org filter: Reforma never leaves the hub
            skipped.append({"batch": b.get("id"), "reason": ss.org_decision(b)})
            continue
        why = reforma_batch_hit(b)
        if why:  # third filter, content based (4.4 order 13)
            skipped.append({"batch": b.get("id"), "reason": f"reforma: {why}"})
            hits.append({"batch": b.get("id"), "product": None, "reason": why, "files": None, "srcs": []})
            continue
        try:
            ss.local_gate(b, pictures)
        except ss.SelectionError as exc:
            alarms.append({"batch": b["id"], "message": str(exc)})
            continue
        root = os.path.normpath(os.path.abspath(b["path"]))
        removed[b["id"]] = ss._count_removed(root)
        stars = stars_by_key(ss.hub.load_ratings(b["id"]))
        for prod in ss.hub.scan_batch_products(b):
            shots, seen = [], set()
            gold_keys = ["Original"] if "Original" in (prod.get("variants") or {}) else list(prod.get("variants") or {})[:1]
            for vk, vshots in (prod.get("variants") or {}).items():
                for s in vshots or []:
                    rel = (s.get("rel_path") or "").replace("\\", "/").lstrip("/")
                    full = os.path.normpath(os.path.join(root, rel)) if rel else ""
                    if not rel or rel in seen or not os.path.isfile(full) or not inside(full, root) or inside(full, target):
                        continue
                    seen.add(rel)
                    shots.append({"src": full, "filename": s.get("filename") or os.path.basename(rel),
                                  "gold": bool(vk in gold_keys and (s.get("is_studio_packshot") or s.get("in_gold_set")))})
            if not shots:
                continue
            rel0 = os.path.relpath(shots[0]["src"], root).replace("/", os.sep)
            why = reforma_product_hit(prod, reforma, rel0.split(os.sep)[0] if os.sep in rel0 else "")
            if why:
                hits.append({"batch": b["id"], "product": prod.get("id"), "reason": why, "files": len(shots),
                             "srcs": [s["src"] for s in shots]})
                continue
            records.append({"batch": b, "product": prod, "shots": shots, "stars": stars})
    return batches, records, alarms, removed, skipped, hits


def sha256(path):
    h = hashlib.sha256()
    with open(lp(path), "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def sig(path):
    try:
        st = os.stat(lp(path))
    except OSError:
        return None
    return [st.st_size, int(st.st_mtime)]


def load_manifest():
    latest = {}
    if os.path.isfile(MANIFEST_FILE):
        with open(MANIFEST_FILE, encoding="utf-8") as f:
            for line in f:
                if line.strip():
                    e = json.loads(line)
                    latest[e["dest_rel"].lower()] = e
    return latest


def append_manifest(latest, entry):
    os.makedirs(os.path.dirname(MANIFEST_FILE), exist_ok=True)
    with open(MANIFEST_FILE, "a", encoding="utf-8", newline="\n") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        f.flush()
        os.fsync(f.fileno())
    latest[entry["dest_rel"].lower()] = entry


def append_move(entry):
    os.makedirs(os.path.dirname(MOVES_FILE), exist_ok=True)
    with open(MOVES_FILE, "a", encoding="utf-8", newline="\n") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        f.flush()
        os.fsync(f.fileno())


def now_iso():
    return datetime.datetime.now().isoformat(timespec="seconds")


def is_active(e):
    return e.get("state", "aktiv") == "aktiv"


def dest_path(target, rel):
    full = os.path.normpath(os.path.join(target, rel))
    if not inside(full, target) or os.path.normcase(full) == os.path.normcase(os.path.normpath(target)):
        raise Alarm(f"path outside target: {full}")
    return full


def sync_item(it, target, latest, dry_run):
    """'unchanged' | 'copied' | 'overwritten' | 'adopted' | 'conflict'; raises OSError on a failed copy."""
    dst = dest_path(target, it["dest_rel"])
    prev, d_sig, s_sig = latest.get(it["dest_rel"].lower()), sig(dst), sig(it["src"])
    if prev and not is_active(prev):
        prev = None  # moved to Borttagna and not moved back (source changed): copy afresh from the source
    if prev and d_sig == prev["dst_sig"] and s_sig == prev["src_sig"] and prev["src"].lower() == it["src"].lower():
        return "unchanged"
    if prev and d_sig not in (None, prev["dst_sig"]):
        return "conflict"  # edited at the target after our copy: left alone, verify reports it, restore fixes it
    if dry_run:
        return "overwritten" if prev else ("conflict" if d_sig and d_sig[0] != s_sig[0] else "copied")
    digest = sha256(it["src"])
    status = "overwritten" if prev else "copied"
    if not prev and d_sig:
        if sha256(dst) != digest:
            return "conflict"  # a file we did not write
        status = "adopted"
    else:
        copy_atomic(it["src"], dst, digest)
    append_manifest(latest, {"ts": now_iso(), "dest_rel": it["dest_rel"],
                             "kind": it["kind"], "batch": it["batch"], "src": it["src"], "sha256": digest,
                             "src_sig": s_sig, "dst_sig": sig(dst), **pkey(it)})
    return status


def copy_atomic(src, dst, digest):
    """Copy to <dst>.godk-tmp in the same folder, check sha256, then os.replace: a run killed mid-copy never leaves a half
    file under the real name (4.11 N2: no manifest entry -> 'conflict' on every later run). clean_tmp removes leftovers."""
    tmp = dst + TMP_SUFFIX
    os.makedirs(lp(os.path.dirname(dst)), exist_ok=True)
    shutil.copy2(lp(src), lp(tmp))
    if sha256(tmp) != digest:
        os.remove(lp(tmp))
        raise OSError(f"hash mismatch after copy: {dst}")
    os.replace(lp(tmp), lp(dst))


def clean_tmp(target, dry_run):
    """Our own *.godk-tmp leftovers of a killed run under target (names only, no file is read); removed unless dry_run."""
    found = []
    for dirpath, _dirs, names in os.walk(lp(target)) if os.path.isdir(lp(target)) else ():
        for n in names:
            if n.lower().endswith(TMP_SUFFIX):
                found.append(os.path.join(dirpath, n))
                if not dry_run:
                    os.remove(os.path.join(dirpath, n))
    return found


def pkey(it):
    return {"product_key": it["product_key"]} if it.get("product_key") else {}


def move_back(items, latest, target, borttagna, dry_run):
    """Planned items whose copy sits in Borttagna: moved back when that file still equals the current source (sha256);
    otherwise 'recopy' (sync_item copies afresh from the source, the stale Borttagna file stays). A freed/cloud-only file
    is never read (4.11 N1: reading hydrates it, or Errno 22 stops the whole sync, also in plan): 'recopy'."""
    out = []
    for it in items:
        prev = latest.get(it["dest_rel"].lower())
        if not prev or is_active(prev):
            continue
        b_path = dest_path(borttagna, prev["borttagen_rel"])
        dst = dest_path(target, it["dest_rel"])
        if (not os.path.isfile(lp(b_path)) or os.path.exists(lp(dst)) or is_cloud(b_path) or is_cloud(it["src"])
                or sha256(b_path) != sha256(it["src"])):
            out.append({"dest_rel": it["dest_rel"], "status": "recopy"})
            continue
        if dry_run:
            out.append({"dest_rel": it["dest_rel"], "status": "would_move_back"})
            continue
        before = sha256(b_path)
        os.makedirs(lp(os.path.dirname(dst)), exist_ok=True)
        shutil.move(lp(b_path), lp(dst))
        after = sha256(dst)
        if after != before:
            raise OSError(f"hash mismatch after move back: {dst}")
        append_manifest(latest, {"ts": now_iso(), "dest_rel": it["dest_rel"], "kind": it["kind"], "batch": it["batch"],
                                 "src": it["src"], "sha256": after, "src_sig": sig(it["src"]), "dst_sig": sig(dst),
                                 **pkey(it)})
        append_move({"ts": now_iso(), "action": "back", "from": b_path, "to": dst, "sha256_before": before,
                     "sha256_after": after, "reason": "tillbaka i hubbens urval"})
        prune_empty_dirs(os.path.dirname(b_path), borttagna)
        out.append({"dest_rel": it["dest_rel"], "status": "moved_back"})
    return out


def prune_empty_dirs(d, root):
    """rmdir d and its parents while empty, never root itself and never outside root (files are never deleted)."""
    d = os.path.normpath(d)
    while inside(d, root) and os.path.normcase(d) != os.path.normcase(os.path.normpath(root)):
        try:
            os.rmdir(lp(d))
        except OSError:
            return
        d = os.path.dirname(d)


def in_removed(src):
    """True when the source left its batch folder and now sits in '<batch folder> Removed' at the same relative path."""
    if os.path.isfile(lp(src)):
        return False
    d = os.path.dirname(os.path.normpath(src))
    while d and os.path.dirname(d) != d:
        if os.path.isfile(lp(os.path.join(d + " Removed", os.path.relpath(src, d)))):
            return True
        d = os.path.dirname(d)
    return False


def mirror_removed(latest, items, target, borttagna, ok_batch_ids, live_batch_ids, dry_run, allow_mass=False,
                   reforma_batches=(), reforma_srcs=()):
    """Active manifest copies no longer planned (covers included) move to Borttagna at the same relative path, when their
    batch was collected without alarm or is no longer live. A live batch that is alarmed/skipped is HELD. Reforma copies are
    held and raise an alarm (they must never sit in a shared folder, nor go to Borttagna). A freed/cloud-only candidate is
    pinned and held to a later run ("pinned"). Never deletes a file."""
    planned = {it["dest_rel"].lower() for it in items}
    ok, live, ref_b = set(ok_batch_ids), set(live_batch_ids), set(reforma_batches)
    ref_s = {s.lower() for s in reforma_srcs}
    res = {"moved": [], "would_move": [], "held": [], "pinned": [], "alarms": []}
    active = [e for e in latest.values() if is_active(e)]
    cands = []
    for e in sorted(active, key=lambda e: e["dest_rel"].lower()):
        if e["dest_rel"].lower() in planned:
            continue
        row = {"dest_rel": e["dest_rel"], "batch": e["batch"]}
        if e["batch"] in ref_b or e["src"].lower() in ref_s:
            res["held"].append({**row, "reason": "reforma"})
            res["alarms"].append({"batch": e["batch"], "message": f"Reforma copy in {TARGET_NAME}: {e['dest_rel']}"})
        elif e["batch"] in live and e["batch"] not in ok:
            res["held"].append({**row, "reason": "batch live men larmad/överhoppad"})
        else:
            reason = ("batch avpublicerad" if e["batch"] not in live else
                      "underkänd (Removed)" if in_removed(e["src"]) else "inte längre i hubbens urval")
            cands.append((e, {**row, "reason": reason}))
    limit = max(MASS_MOVE_MIN, MASS_MOVE_FRACTION * len(active))
    if len(cands) > limit and not allow_mass:
        res["alarms"].append({"batch": None, "message": f"mass move refused: {len(cands)} of {len(active)} (limit {int(limit)}; "
                                                         "--allow-mass-move)"})
        res["would_move"] = [r for _e, r in cands]
        return res
    holds = {} if dry_run else load_pin_holds()
    new_holds = {}
    for e, row in cands:
        if dry_run:
            res["would_move"].append(row)
            continue
        src = dest_path(target, e["dest_rel"])
        dst = dest_path(borttagna, e["dest_rel"])
        if is_cloud(src):
            # Freed (+U) or cloud-only copy: reading it for the sha check fails (Errno 22, INCIDENT 1.6). Pin it (-U +P,
            # stays local) and move it on a later run once hydrated; alarm only after PIN_ALARM_RUNS runs in a row (4.10 A).
            k = e["dest_rel"].lower()
            n = new_holds[k] = holds.get(k, 0) + 1
            try:
                pin_local(src)
                pin = "pinned"
            except (OSError, subprocess.CalledProcessError) as exc:
                pin = f"pin failed: {str(exc)[:120]}"
            res["pinned"].append({**row, "runs": n, "pin": pin})
            if n >= PIN_ALARM_RUNS:
                res["alarms"].append({"batch": e["batch"], "message": f"still cloud-only after {n} runs in a row, not moved to "
                                                                      f"{BORTTAGNA_NAME}: {e['dest_rel']} ({pin})"})
            continue
        if os.path.exists(lp(dst)):  # an older removal of the same path: keep both, the old one untouched
            stem, ext = os.path.splitext(dst)
            dst = f"{stem}_{datetime.datetime.now():%y%m%d%H%M}{ext}"
            i = 2
            while os.path.exists(lp(dst)):
                dst, i = f"{stem}_{datetime.datetime.now():%y%m%d%H%M}_{i}{ext}", i + 1
        status, before = "moved", None
        if os.path.isfile(lp(src)):
            before = sha256(src)
            os.makedirs(lp(os.path.dirname(dst)), exist_ok=True)
            shutil.move(lp(src), lp(dst))
            after = sha256(dst)
            if after != before:
                raise OSError(f"hash mismatch after move to {BORTTAGNA_NAME}: {dst}")
            append_move({"ts": now_iso(), "action": "to_borttagna", "from": src, "to": dst, "sha256_before": before,
                         "sha256_after": after, "reason": row["reason"]})
            prune_empty_dirs(os.path.dirname(src), target)
        else:
            status = "missing"
        append_manifest(latest, {"ts": now_iso(), "dest_rel": e["dest_rel"], "kind": e["kind"], "batch": e["batch"],
                                 "src": e["src"], "state": "borttagen", "borttagen_rel": os.path.relpath(dst, borttagna),
                                 "reason": row["reason"], "sha256": before or e.get("sha256"), "status": status})
        res["moved"].append({**row, "status": status})
    if not dry_run and new_holds != holds:
        save_pin_holds(new_holds)
    return res


def pin_holds_file():
    return os.path.join(os.path.dirname(MANIFEST_FILE), "pin_holds.json")


def load_pin_holds():
    """{dest_rel lower: consecutive runs held as cloud-only}; a copy not held in a run drops out (counter resets)."""
    try:
        with open(pin_holds_file(), encoding="utf-8") as f:
            return {str(k): int(v) for k, v in json.load(f).items()}
    except (OSError, ValueError, AttributeError):
        return {}


def save_pin_holds(holds):
    os.makedirs(os.path.dirname(pin_holds_file()), exist_ok=True)
    with open(pin_holds_file(), "w", encoding="utf-8", newline="\n") as f:
        json.dump(holds, f, ensure_ascii=False, indent=1)


def file_attrs(path):
    """Win32 attributes without opening the file (no hydration); None when missing or not on Windows."""
    if os.name != "nt":
        return None
    a = ctypes.WinDLL("kernel32", use_last_error=True).GetFileAttributesW(lp(path))
    return None if a == 0xFFFFFFFF else a


def is_cloud(path):
    """Freed (+U) or cloud-only: reading it hydrates the file or fails with Errno 22 (INCIDENT 1.6)."""
    a = file_attrs(path)
    return a is not None and bool(a & (ATTR_UNPINNED | ATTR_CLOUD_ONLY))


def pin_local(path):
    """attrib -U +P: OneDrive downloads the file and keeps it local (the inverse of `free`)."""
    subprocess.run(["attrib", "-U", "+P", path], check=True, capture_output=True)


def readme_text(products):
    out = [f"{TARGET_NAME}",
           "Här ligger de godkända bilderna från QA-hubben.",
           "Struktur: Varumärke > Kategori > Produkt (artikelnummer).",
           "I varje produktmapp är 00_omslag den bäst betygsatta bilden. Filnamnen börjar med batchens datum (ÅÅMMDD).",
           "Mappen uppdateras automatiskt från QA-hubben. Bilder som senare blir underkända försvinner automatiskt härifrån.",
           "Frågor om innehållet: Andronik Lindgren.",
           f"Senast uppdaterad: {datetime.datetime.now():%Y-%m-%d %H:%M}", ""]
    tree = {}
    for e in products.values():
        tree.setdefault(e["brand"], {}).setdefault(e["category"], []).append(e)
    for brand in sorted(tree):
        cats = tree[brand]
        out.append(f"{brand} — {sum(len(v) for v in cats.values())} produkter, "
                   f"{sum(e['count'] for v in cats.values() for e in v)} bilder")
        for cat in sorted(cats):
            es = sorted(cats[cat], key=lambda e: e["folder_rel"].lower())
            out.append(f"  {cat} — {len(es)} produkter, {sum(e['count'] for e in es)} bilder")
            out += [f"    {os.path.basename(e['folder_rel'])}: {e['count']} bilder" for e in es]
        out.append("")
    return "\r\n".join(out) + "\r\n"


def links_json(products):
    out = {}
    for key, e in sorted(products.items()):
        url = f"{ONEDRIVE_WEB}/Pictures/{urllib.parse.quote(TARGET_NAME)}/" + "/".join(
            urllib.parse.quote(p) for p in e["folder_rel"].split(os.sep))
        for b in e["batches"]:
            out.setdefault(b, []).append({"product_key": key, "folder": e["folder_rel"], "url": url})
    return out


def write_if_changed(path, text):
    old = open(lp(path), encoding="utf-8").read() if os.path.isfile(lp(path)) else None
    if old != text:
        with open(lp(path), "w", encoding="utf-8", newline="") as f:
            f.write(text)
        return True
    return False


def c_free_gb(path):
    return round(shutil.disk_usage(os.path.splitdrive(os.path.abspath(path))[0] + os.sep).free / 1e9, 1)


def cmd_plan_or_sync(ss, pictures, target, borttagna, dry_run, allow_mass=False):
    guard_target(ss, target)
    guard_target(ss, borttagna)
    batches, records, alarms, removed, skipped, ref_hits = collect(ss, pictures, target)
    latest = load_manifest()
    items, products, brand_alarms = build_plan(records, target, known_folders(latest))
    cats, brands, status_n, errors = {}, {}, {}, []
    tmp_left = clean_tmp(target, dry_run)
    try:
        back = move_back(items, latest, target, borttagna, dry_run)
    except OSError as exc:
        back, errors = [], [{"dest_rel": None, "error": f"move back: {str(exc)[:200]}"}]
    for it in items if not errors else ():
        parts = it["dest_rel"].split(os.sep)
        for agg in (cats.setdefault(os.sep.join(parts[:2]), {"products": set(), "files": 0, "bytes": 0, "covers": 0}),
                    brands.setdefault(parts[0], {"products": set(), "files": 0, "bytes": 0, "covers": 0})):
            agg["products"].add(os.path.dirname(it["dest_rel"]))
            if it["kind"] == "cover":
                agg["covers"] += 1
            else:
                agg["files"] += 1
                agg["bytes"] += os.path.getsize(lp(it["src"]))
        try:
            st = sync_item(it, target, latest, dry_run)
        except (OSError, Alarm) as exc:
            st = "error"
            errors.append({"dest_rel": it["dest_rel"], "error": str(exc)[:200]})
            if not dry_run:
                break
        status_n[st] = status_n.get(st, 0) + 1
    mirror = {"moved": [], "would_move": [], "held": [], "pinned": [], "alarms": []}
    if not errors:
        bad_ids = {a["batch"] for a in alarms} | {s["batch"] for s in skipped}
        ok_ids = {b["id"] for b in batches} - bad_ids
        live_ids = {b.get("id") for b in ss.hub.load_batches() if b.get("published", True)}
        ref_batches = {h["batch"] for h in ref_hits if h["product"] is None} | {
            s["batch"] for s in skipped if "reforma" in str(s.get("reason"))}
        try:
            mirror = mirror_removed(latest, items, target, borttagna, ok_ids, live_ids, dry_run, allow_mass, ref_batches,
                                    [p for h in ref_hits for p in h["srcs"]])
        except (OSError, Alarm) as exc:
            errors.append({"dest_rel": None, "error": f"mirror: {str(exc)[:200]}"})
    n_gate_alarms = len(alarms)
    alarms = alarms +[{"batch": a["batches"], "message": a["message"]} for a in brand_alarms] + mirror["alarms"]

    def row(k, v, key):
        return {key: k, "products": len(v["products"]), "files": v["files"], "gb": round(v["bytes"] / 1e9, 3), "covers": v["covers"]}
    rows = [row(k, v, "category") for k, v in sorted(cats.items())]
    for r in rows + [row(k, v, "brand") for k, v in sorted(brands.items())]:
        print(json.dumps(r, ensure_ascii=False))
    merged = [{"folder": e["folder_rel"], "batches": e["batches"]} for e in products.values() if len(e["batches"]) > 1]
    n_back = {s: sum(1 for b in back if b["status"] == s) for s in ("moved_back", "would_move_back", "recopy")}
    summary = {"mode": "dry-run" if dry_run else "sync", "target": target, "borttagna": borttagna,
               "batches_selected": len(batches), "batches_copied": len(batches) - n_gate_alarms, "products": len(products),
               "files": sum(r["files"] for r in rows), "gb": round(sum(r["gb"] for r in rows), 3),
               "removed_excluded": sum(removed.values()), "status": status_n, "alarms": alarms[:50], "alarm_n": len(alarms),
               "errors": errors, "merged_products": len(merged),
               "max_path_len": max((len(os.path.join(target, it["dest_rel"])) for it in items), default=0),
               "moved": len(mirror["moved"]), "would_move": len(mirror["would_move"]), "held": len(mirror["held"]),
               "pinned": len(mirror["pinned"]), "move_back": n_back, "brand_alarms": len(brand_alarms),
               ("tmp_found" if dry_run else "tmp_removed"): len(tmp_left),
               "reforma_excluded": {"batches": sorted({h["batch"] for h in ref_hits if h["product"] is None}),
                                    "products": sum(1 for h in ref_hits if h["product"] is not None),
                                    "files": sum(h["files"] or 0 for h in ref_hits)},
               "skipped": skipped, "c_free_gb": c_free_gb(target)}
    for e in sorted(products.values(), key=lambda e: e["folder_rel"].lower()):
        print(json.dumps({"folder": e["folder_rel"], "files": e["count"], "batches": e["batches"]}, ensure_ascii=False))
    print(json.dumps({"merged": merged}, ensure_ascii=False))
    print(json.dumps({"removed_per_batch": removed}, ensure_ascii=False))
    print(json.dumps({"reforma_hits": [{k: v for k, v in h.items() if k != "srcs"} for h in ref_hits]}, ensure_ascii=False))
    print(json.dumps({"brand_alarms": brand_alarms[:50]}, ensure_ascii=False))
    for k in ("moved", "would_move", "held", "pinned"):
        print(json.dumps({k: mirror[k][:50]}, ensure_ascii=False))
    print(json.dumps({"move_back": [b for b in back if b["status"] != "recopy"][:50]}, ensure_ascii=False))
    if not dry_run and not errors:
        os.makedirs(lp(target), exist_ok=True)
        summary["readme_written"] = write_if_changed(os.path.join(target, "LÄS MIG.txt"), readme_text(products))
        summary["links_written"] = write_if_changed(LINKS_FILE, json.dumps(links_json(products), ensure_ascii=False, indent=1))
    print(json.dumps({"summary": summary}, ensure_ascii=False))
    return 1 if errors else 3 if alarms or status_n.get("conflict") else 0


def cmd_reforma_scan(ss, pictures, target):
    """Read-only list for the coordinator: per published batch org + Reforma hits; weak model-word matches separately."""
    reforma = load_reforma()
    print(json.dumps({"reforma_list": reforma["file"], "skus": len(reforma["skus"]), "handles": len(reforma["handles"])},
                     ensure_ascii=False))
    _b, records, alarms, _r, skipped, hits = collect(ss, pictures, target, reforma)
    for b in sorted(ss.hub.load_batches(), key=lambda b: b.get("id") or ""):
        if not b.get("published", True):
            continue
        bid = b.get("id")
        print(json.dumps({"batch": bid, "org": b.get("org"), "category": b.get("category"),
                          "strong_hits": sum(1 for h in hits if h["batch"] == bid),
                          "skipped": next((s["reason"] for s in skipped if s["batch"] == bid), None),
                          "alarm": any(a["batch"] == bid for a in alarms)}, ensure_ascii=False))
    for h in hits:
        print(json.dumps({"strong": {k: v for k, v in h.items() if k != "srcs"}}, ensure_ascii=False))
    weak = []
    for r in records:
        rel0 = os.path.relpath(r["shots"][0]["src"], r["batch"]["path"]).replace("/", os.sep)
        w = reforma_weak_hits(r["product"], reforma, rel0.split(os.sep)[0] if os.sep in rel0 else "")
        if w:
            weak.append({"batch": r["batch"]["id"], "product": r["product"].get("id"), "words": w, "files": len(r["shots"])})
    for w in weak:
        print(json.dumps({"weak": w}, ensure_ascii=False))
    print(json.dumps({"summary": {"strong": len(hits), "weak": len(weak), "alarms": alarms}}, ensure_ascii=False))
    return 3 if alarms else 0


def cmd_verify(ss, target, borttagna):
    guard_target(ss, target)
    latest, bad = load_manifest(), []
    for e in latest.values():
        if not is_active(e):
            if e.get("status") != "missing" and not os.path.isfile(lp(dest_path(borttagna, e["borttagen_rel"]))):
                bad.append({"dest_rel": e["dest_rel"], "problem": f"missing in {BORTTAGNA_NAME}"})
            continue
        d = sig(dest_path(target, e["dest_rel"]))
        if d != e["dst_sig"]:
            bad.append({"dest_rel": e["dest_rel"], "problem": "missing" if d is None else "changed at target"})
    free = c_free_gb(target)
    res = {"entries": len(latest), "ok": len(latest) - len(bad), "alarm": bad[:50], "alarm_n": len(bad), "c_free_gb": free,
           "low_disk": free < MIN_FREE_GB}
    print(json.dumps(res, ensure_ascii=False))
    return 3 if bad or free < MIN_FREE_GB else 0


def cmd_restore(ss, target, prefix, dry_run):
    guard_target(ss, target)
    latest, out, code = load_manifest(), [], 0
    clean_tmp(target, dry_run)
    for e in list(latest.values()):
        if not is_active(e) or not e["dest_rel"].lower().startswith(prefix.lower()):
            continue
        dst = dest_path(target, e["dest_rel"])
        if not os.path.isfile(lp(e["src"])) or sha256(e["src"]) != e["sha256"]:
            out.append({"dest_rel": e["dest_rel"], "status": "source changed or gone: run sync"})
            code = 3
            continue
        if not dry_run:
            copy_atomic(e["src"], dst, e["sha256"])
            append_manifest(latest, {**e, "ts": datetime.datetime.now().isoformat(timespec="seconds"), "dst_sig": sig(dst)})
        out.append({"dest_rel": e["dest_rel"], "status": "would_restore" if dry_run else "restored"})
    print(json.dumps({"restore": out}, ensure_ascii=False))
    return code


def cmd_free(ss, target):
    guard_target(ss, target)
    batches, out = ss.hub.load_batches(), []
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    for cat in sorted(os.listdir(lp(target))) if os.path.isdir(lp(target)) else []:
        for prod in sorted(os.listdir(lp(os.path.join(target, cat)))) if os.path.isdir(lp(os.path.join(target, cat))) else []:
            d = os.path.join(target, cat, prod)
            ok, info = ss.folder_synced(d)
            refusal = ss.free_refusal(d, batches) or ("" if inside(d, target) else "refused: outside target")
            if ok and not refusal:
                for dirpath, _dirs, names in os.walk(lp(d)):
                    for n in names:
                        p = os.path.join(dirpath, n)
                        a = k32.GetFileAttributesW(p)
                        if a != 0xFFFFFFFF:
                            k32.SetFileAttributesW(p, (a | ATTR_UNPINNED) & ~ATTR_PINNED)
            out.append({"folder": d, "freed": bool(ok and not refusal), "status": refusal or info})
    print(json.dumps({"free": out}, ensure_ascii=False))
    return 0


def pid_alive(pid):
    """True while a process with this pid runs (Windows: OpenProcess + GetExitCodeProcess == STILL_ACTIVE)."""
    if os.name != "nt":
        try:
            os.kill(pid, 0)
            return True
        except OSError:
            return False
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    h = k32.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
    if not h:
        return False
    try:
        code = ctypes.c_ulong()
        return bool(k32.GetExitCodeProcess(h, ctypes.byref(code))) and code.value == 259
    finally:
        k32.CloseHandle(h)


def read_lock():
    """{pid, started, ts} of the lock file (old format = bare pid, age from mtime); None when unreadable."""
    try:
        with open(LOCK_FILE, encoding="utf-8") as f:
            raw = f.read().strip()
        mtime = os.path.getmtime(LOCK_FILE)
    except OSError:
        return None
    try:
        d = json.loads(raw)
        d = d if isinstance(d, dict) else {"pid": int(d)}
    except ValueError:
        d = {"pid": int(raw)} if raw.isdigit() else {}
    d.setdefault("ts", mtime)
    return d


def acquire_lock():
    """O_EXCL lock with pid + start time. A lock left by a killed run (pid dead, or older than LOCK_STALE_S) is removed
    and logged ("stale lock removed" on stderr) so a timeout/kill/reboot never turns every later run into a skip."""
    for attempt in (1, 2):
        try:
            fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            d = read_lock()
            if attempt == 2 or d is None:
                return False
            age = datetime.datetime.now().timestamp() - float(d.get("ts") or 0)
            if pid_alive(int(d.get("pid") or 0)) and age <= LOCK_STALE_S:
                return False
            try:
                os.remove(LOCK_FILE)
            except OSError:
                return False
            print(json.dumps({"stale_lock_removed": {**d, "age_s": int(age)}}, ensure_ascii=False), file=sys.stderr)
            continue
        os.write(fd, json.dumps({"pid": os.getpid(), "started": now_iso(),
                                 "ts": datetime.datetime.now().timestamp()}).encode())
        os.close(fd)
        return True
    return False


def release_lock():
    """Removes the lock only while it is still ours (a stale-lock removal by a later run must not be undone)."""
    d = read_lock()
    if d is not None and int(d.get("pid") or 0) == os.getpid():
        os.remove(LOCK_FILE)


def main():
    for stream in (sys.stdout, sys.stderr):  # a cp1252 console must never crash on a product name (gate v5 R5)
        stream.reconfigure(errors="backslashreplace")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=("plan", "sync", "verify", "restore", "free", "reforma-scan"))
    ap.add_argument("prefix", nargs="?", default="")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--allow-mass-move", action="store_true", help="allow more than max(50, 10 %%) moves to Borttagna in one sync")
    args = ap.parse_args()
    if args.cmd == "restore" and not args.prefix:
        print("restore needs <dest-prefix>", file=sys.stderr)
        return 2
    locking = args.cmd in ("sync", "restore", "free")
    if locking and not acquire_lock():
        print(f"lock held: {LOCK_FILE}", file=sys.stderr)
        return 4
    try:
        ss = load_sync()
        pictures = os.path.join(ss.DEFAULT_ONEDRIVE, ss.PICTURES_NAME)
        target = os.path.join(pictures, TARGET_NAME)
        borttagna = os.path.join(pictures, BORTTAGNA_NAME)
        if args.cmd in ("plan", "sync"):
            return cmd_plan_or_sync(ss, pictures, target, borttagna, dry_run=args.cmd == "plan",
                                    allow_mass=args.allow_mass_move)
        if args.cmd == "reforma-scan":
            return cmd_reforma_scan(ss, pictures, target)
        if args.cmd == "verify":
            return cmd_verify(ss, target, borttagna)
        if args.cmd == "restore":
            return cmd_restore(ss, target, args.prefix, args.dry_run)
        return cmd_free(ss, target)
    except Alarm as exc:
        print(json.dumps({"ALARM": str(exc)}, ensure_ascii=False))
        return 3
    finally:
        if locking:
            release_lock()

if __name__ == "__main__":
    sys.exit(main())
