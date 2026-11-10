"""Distribution copy of approved hub images in the CEO's OneDrive Pictures (CEO 2026-11-10 "ja kör på pictures-mappen"):
  <OneDrive>\\Pictures\\QA-hubb Godkända\\<Swedish category>\\<Product name (SKU)>\\<yymmdd>_<file>
plus 00_omslag.<ext> (best image by hub rating) per product and LÄS MIG.txt at the root (category -> product -> count + hub).
Selection (CEO-RÄTTELSE 5): every shot the hub shows in a batch folder; only images in "<batch> Removed" are underkända
(Removed is never scanned). Reforma is NEVER copied (org filter twice). A product in several batches gets ONE folder (key =
SKU, else the product id); file names carry the batch date so they never collide.
The target is a copy, never a hub source: the run refuses (exit 3) if any batches_config path lies inside it, and `free`
(+U) refuses hub sources (sharepoint_sync.free_refusal). Sources are read only, never freed. Idempotent: unchanged files are
skipped by stat signature, a changed source overwrites its copy (sha256 checked after copy), a copy edited at the target is
left alone and reported as conflict. Files that left the selection are never deleted, only listed ("underkänd efter
publicering" when the source now sits in Removed). Manifest (JSONL) in ~/qa-hub-backup/pictures-godkanda/.

Commands (hub python, cwd = hub dir):
  plan                dry-run: files + GB per category, Removed excluded, ALARM batches, merged products (read-only)
  sync                copy/overwrite per plan, write LÄS MIG.txt + pictures_godkanda_links.json, append manifest
  verify              target vs manifest by metadata (no download, +U safe) + C: free >= 20 GB; deviation -> exit 3
  restore <dest-prefix> [--dry-run]   copy back from the hub source when its sha256 equals the manifest
  free                +U on product folders OneDrive reports as uploaded (target only)
Exit: 0 OK, 1 copy error, 2 usage, 3 ALARM, 4 lock held."""
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
MANIFEST_FILE = os.path.join(os.path.expanduser("~"), "qa-hub-backup", "pictures-godkanda", "manifest.jsonl")
LINKS_FILE = os.path.join(HUB_DIR, "pictures_godkanda_links.json")
LOCK_FILE = os.path.join(HUB_DIR, ".pictures_godkanda.lock")
HUB_URL = "https://trendcarpet-qa.vercel.app"
ONEDRIVE_WEB = "https://trendcarpet-my.sharepoint.com/personal/andronik_lindgren_trendcarpet_se/Documents"
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".gif", ".avif", ".heic"}
MAX_PATH = 255
MIN_FREE_GB = 20
ATTR_PINNED, ATTR_UNPINNED = 0x80000, 0x100000

# (category, batch-id needles, product needles); first hit wins. Product text = product id + display name + colour.
CATEGORIES = [
    ("Mattor - Trappsteg", ("trappsteg",), ("trappsteg", "halvmåne", "halvmane")),
    ("Rum & Möbler", ("flow-vibes-rum",), ("soffbord", "baddsoffa", "soffa-")),
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


def build_plan(records, target):
    """records: dicts {batch, product, shots:[{src, filename, gold}], stars:{key: mean}}. Returns (items, products).
    items: {src, dest_rel, batch, kind}. One folder per product key; cover = best mean stars, then gold, then newest batch."""
    products = {}
    for r in sorted(records, key=lambda r: (batch_date(r["batch"]), r["batch"]["id"]), reverse=True):
        p, sku = r["product"], sku_for(r["product"])
        rel = os.path.relpath(r["shots"][0]["src"], r["batch"]["path"]) if r["shots"] and r["batch"].get("path") else ""
        ctx = rel.replace("/", os.sep).split(os.sep)[0] if os.sep in rel.replace("/", os.sep) else ""
        key = sku or clean_component(str(p.get("id") or ""), 80).lower()
        if not sku and ctx and is_generic(str(p.get("id") or "")):
            key = clean_component(f"{ctx} {p.get('id')}", 80).lower()  # "1500px" in two product folders = two products
        e = products.setdefault(key, {"category": category_for(r["batch"], p, ctx), "folder": product_name(p, sku, ctx), "color": p.get("color"),
                                      "files": [], "batches": [], "seen": set()})
        if r["batch"]["id"] not in e["batches"]:
            e["batches"].append(r["batch"]["id"])
        for s in r["shots"]:
            if s["src"].lower() in e["seen"]:
                continue
            e["seen"].add(s["src"].lower())
            e["files"].append({**s, "batch": r["batch"]["id"], "date": batch_date(r["batch"]),
                               "stars": r["stars"].get(f"{p.get('id')}|{s['filename']}", 0.0)})
    items, used = [], {}
    name_count = {}
    for e in products.values():
        name_count[(e["category"], e["folder"].lower())] = name_count.get((e["category"], e["folder"].lower()), 0) + 1
    for key, e in sorted(products.items()):
        folder = e["folder"]
        col = clean_component(str(e.get("color") or ""), 30) if e.get("color") else ""
        if name_count[(e["category"], folder.lower())] > 1 and col and col.lower() not in folder.lower():
            folder = f"{folder} {col}"  # same name, different colour: "Black River Legacy Trucker Khaki", not "(2)"
        folder_rel = os.path.join(e["category"], folder)
        # Two product keys with the same display folder (same name, no SKU): suffix the later one.
        n = used.get(folder_rel.lower(), 0)
        used[folder_rel.lower()] = n + 1
        if n:
            folder_rel = f"{folder_rel} ({n + 1})"
        e["folder_rel"] = folder_rel
        names = set()
        room = MAX_PATH - len(os.path.join(target, folder_rel)) - 1
        for f in e["files"]:
            stem, ext = os.path.splitext(os.path.basename(f["src"]))
            base = clean_component(f"{f['date']}_{stem}", max(12, room - len(ext) - 3))
            name, i = base + ext.lower(), 2
            while name.lower() in names:
                name, i = f"{base}_{i}{ext.lower()}", i + 1
            names.add(name.lower())
            items.append({"src": f["src"], "dest_rel": os.path.join(folder_rel, name), "batch": f["batch"], "kind": "image"})
        imgs = [f for f in e["files"] if os.path.splitext(f["src"])[1].lower() in IMAGE_EXTS]
        if imgs:
            best = max(imgs, key=lambda f: (f["stars"], f.get("gold", False), f["date"]))
            items.append({"src": best["src"], "batch": best["batch"], "kind": "cover",
                          "dest_rel": os.path.join(folder_rel, "00_omslag" + os.path.splitext(best["src"])[1].lower())})
        e["count"] = len(e["files"])
    return items, products


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


def collect(ss, pictures, target):
    args = type("A", (), {"batch": None, "all": True, "include_drafts": False})()
    batches, skipped = ss.select_batches(args)
    records, alarms, removed = [], [], {}
    for b in batches:
        if ss.org_decision(b):  # second, independent org filter: Reforma never leaves the hub
            skipped.append({"batch": b.get("id"), "reason": ss.org_decision(b)})
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
            if shots:
                records.append({"batch": b, "product": prod, "shots": shots, "stars": stars})
    return batches, records, alarms, removed, skipped


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


def dest_path(target, rel):
    full = os.path.normpath(os.path.join(target, rel))
    if not inside(full, target) or os.path.normcase(full) == os.path.normcase(os.path.normpath(target)):
        raise Alarm(f"path outside target: {full}")
    return full


def sync_item(it, target, latest, dry_run):
    """'unchanged' | 'copied' | 'overwritten' | 'adopted' | 'conflict'; raises OSError on a failed copy."""
    dst = dest_path(target, it["dest_rel"])
    prev, d_sig, s_sig = latest.get(it["dest_rel"].lower()), sig(dst), sig(it["src"])
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
        os.makedirs(lp(os.path.dirname(dst)), exist_ok=True)
        shutil.copy2(lp(it["src"]), lp(dst))
    if sha256(dst) != digest:
        raise OSError(f"hash mismatch after copy: {dst}")
    append_manifest(latest, {"ts": datetime.datetime.now().isoformat(timespec="seconds"), "dest_rel": it["dest_rel"],
                             "kind": it["kind"], "batch": it["batch"], "src": it["src"], "sha256": digest,
                             "src_sig": s_sig, "dst_sig": sig(dst)})
    return status


def readme_text(products, batches, removed):
    names = {b["id"]: b.get("name") or b["id"] for b in batches}
    out = [f"{TARGET_NAME} — godkända bilder från QA-hubben (allt utom underkända i \"<batch> Removed\")",
           f"Uppdaterad: {datetime.datetime.now():%Y-%m-%d %H:%M}", f"QA-hubben: {HUB_URL}",
           "Varje produktmapp: 00_omslag = bäst betygsatta bilden i hubben; filnamn börjar med batchens datum (ÅÅMMDD).",
           f"Underkända (Removed) som inte kopierats: {sum(removed.values())} bilder.", ""]
    by_cat = {}
    for e in products.values():
        by_cat.setdefault(e["category"], []).append(e)
    for cat in sorted(by_cat):
        es = sorted(by_cat[cat], key=lambda e: e["folder_rel"].lower())
        out.append(f"{cat} — {len(es)} produkter, {sum(e['count'] for e in es)} bilder")
        for e in es:
            out.append(f"  {os.path.basename(e['folder_rel'])}: {e['count']} bilder")
            out += [f"      hubb-batch: {names.get(b, b)}" for b in e["batches"]]
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


def gone_from_selection(latest, items, target, batches):
    """Manifest copies no longer planned: never deleted, listed (with 'underkänd' when the source now sits in Removed)."""
    planned = {it["dest_rel"].lower() for it in items}
    out = []
    for k, e in sorted(latest.items()):
        if k in planned or e["kind"] == "cover":
            continue
        roots = [os.path.normpath(b["path"]) for b in batches if b["id"] == e["batch"]]
        in_removed = not os.path.isfile(lp(e["src"])) and any(
            inside(e["src"], r) and os.path.isfile(lp(os.path.join(r + " Removed", os.path.relpath(e["src"], r)))) for r in roots)
        out.append({"dest_rel": e["dest_rel"], "batch": e["batch"],
                    "reason": "underkänd efter publicering" if in_removed else "inte längre i hubbens urval"})
    return out


def c_free_gb(path):
    return round(shutil.disk_usage(os.path.splitdrive(os.path.abspath(path))[0] + os.sep).free / 1e9, 1)


def cmd_plan_or_sync(ss, pictures, target, dry_run):
    guard_target(ss, target)
    batches, records, alarms, removed, skipped = collect(ss, pictures, target)
    items, products = build_plan(records, target)
    latest = load_manifest()
    cats, status_n, errors = {}, {}, []
    for it in items:
        c = cats.setdefault(it["dest_rel"].split(os.sep)[0], {"products": set(), "files": 0, "bytes": 0, "covers": 0})
        c["products"].add(os.path.dirname(it["dest_rel"]))
        if it["kind"] == "cover":
            c["covers"] += 1
        else:
            c["files"] += 1
            c["bytes"] += os.path.getsize(lp(it["src"]))
        try:
            st = sync_item(it, target, latest, dry_run)
        except (OSError, Alarm) as exc:
            st = "error"
            errors.append({"dest_rel": it["dest_rel"], "error": str(exc)[:200]})
            if not dry_run:
                break
        status_n[st] = status_n.get(st, 0) + 1
    rows = [{"category": k, "products": len(v["products"]), "files": v["files"], "gb": round(v["bytes"] / 1e9, 3),
             "covers": v["covers"]} for k, v in sorted(cats.items())]
    for r in rows:
        print(json.dumps(r, ensure_ascii=False))
    merged = [{"folder": e["folder_rel"], "batches": e["batches"]} for e in products.values() if len(e["batches"]) > 1]
    gone = gone_from_selection(latest, items, target, batches)
    summary = {"mode": "dry-run" if dry_run else "sync", "target": target, "batches_selected": len(batches),
               "batches_copied": len(batches) - len(alarms), "products": len(products),
               "files": sum(r["files"] for r in rows), "gb": round(sum(r["gb"] for r in rows), 3),
               "removed_excluded": sum(removed.values()), "status": status_n, "alarms": alarms, "errors": errors,
               "merged_products": len(merged), "max_path_len": max((len(os.path.join(target, it["dest_rel"])) for it in items),
                                                                   default=0),
               "left_selection": gone, "skipped": skipped, "c_free_gb": c_free_gb(target)}
    for e in sorted(products.values(), key=lambda e: e["folder_rel"].lower()):
        print(json.dumps({"folder": e["folder_rel"], "files": e["count"], "batches": e["batches"]}, ensure_ascii=False))
    print(json.dumps({"merged": merged}, ensure_ascii=False))
    print(json.dumps({"removed_per_batch": removed}, ensure_ascii=False))
    if not dry_run and not errors:
        os.makedirs(lp(target), exist_ok=True)
        summary["readme_written"] = write_if_changed(os.path.join(target, "LÄS MIG.txt"), readme_text(products, batches, removed))
        summary["links_written"] = write_if_changed(LINKS_FILE, json.dumps(links_json(products), ensure_ascii=False, indent=1))
    print(json.dumps({"summary": summary}, ensure_ascii=False))
    return 1 if errors else 3 if alarms or status_n.get("conflict") else 0


def cmd_verify(ss, target):
    guard_target(ss, target)
    latest, bad = load_manifest(), []
    for e in latest.values():
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
    for e in latest.values():
        if not e["dest_rel"].lower().startswith(prefix.lower()):
            continue
        dst = dest_path(target, e["dest_rel"])
        if not os.path.isfile(lp(e["src"])) or sha256(e["src"]) != e["sha256"]:
            out.append({"dest_rel": e["dest_rel"], "status": "source changed or gone: run sync"})
            code = 3
            continue
        if not dry_run:
            os.makedirs(lp(os.path.dirname(dst)), exist_ok=True)
            shutil.copy2(lp(e["src"]), lp(dst))
            if sha256(dst) != e["sha256"]:
                raise OSError(f"hash mismatch after restore: {dst}")
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


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=("plan", "sync", "verify", "restore", "free"))
    ap.add_argument("prefix", nargs="?", default="")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()
    ss = load_sync()
    pictures = os.path.join(ss.DEFAULT_ONEDRIVE, ss.PICTURES_NAME)
    target = os.path.join(pictures, TARGET_NAME)
    if args.cmd == "restore" and not args.prefix:
        print("restore needs <dest-prefix>", file=sys.stderr)
        return 2
    if args.cmd in ("sync", "restore", "free"):
        try:
            fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode())
            os.close(fd)
        except FileExistsError:
            print(f"lock held: {LOCK_FILE}", file=sys.stderr)
            return 4
    try:
        if args.cmd in ("plan", "sync"):
            return cmd_plan_or_sync(ss, pictures, target, dry_run=args.cmd == "plan")
        if args.cmd == "verify":
            return cmd_verify(ss, target)
        if args.cmd == "restore":
            return cmd_restore(ss, target, args.prefix, args.dry_run)
        return cmd_free(ss, target)
    except Alarm as exc:
        print(json.dumps({"ALARM": str(exc)}, ensure_ascii=False))
        return 3
    finally:
        if args.cmd in ("sync", "restore", "free"):
            os.remove(LOCK_FILE)


if __name__ == "__main__":
    sys.exit(main())
