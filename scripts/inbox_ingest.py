"""Inbox for batches built by AG producers: AG never writes OneDrive/Pictures (rule 2026-10-28) and the hub refuses batch
paths outside Pictures (C gate), so AG delivers to <hub dir>\\_inkorg\\<batch-id>\\ (optional "<batch-id> Removed" sibling and
optional <batch-id>.json = {"name", "category", "org", "published"}) and a Claude thread runs this script:
  1. pick the target: an existing hub copy that already holds exactly this delivery (no copy), else
     Pictures\\QA-hubb-batcher\\<batch-id>, else (that folder holds an older delivery) Pictures\\QA-hubb-batcher\\<batch-id>__v<yymmdd-HHMM>
  2. copy (+ Removed), verify file count + sha256 of every file
  3. pin the copy (agent-home pin-batches.py: every image and folder P)
  4. back up batches_config.json, then repoint the entry with that id or register a new one (same fields as /api/batches/add)
  5. live count through the local server before/after, one line appended to _inkorg\\_ingested.jsonl
The inbox and every existing Pictures folder are only read; nothing is moved or deleted. Ratings/reviews are keyed by batch id,
so a new version keeps them.
"Exactly this delivery" = every delivered file's sha256 is in the copy (main or Removed: reviewers move images to Removed) and
the copy's main folder shows no image that is not in the delivery.

Usage (hub python = miniconda3):
  python scripts/inbox_ingest.py list [--inbox DIR]
  python scripts/inbox_ingest.py ingest --batch <id> [--name ..] [--category ..] [--org cama|reforma] [--published] [--dry-run] [--inbox DIR]
Exit: 0 OK, 1 error, 2 usage.
"""
import argparse
import datetime
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import urllib.parse
import urllib.request

HUB_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BATCHES_FILE = os.path.join(HUB_DIR, "batches_config.json")
INBOX = os.path.join(HUB_DIR, "_inkorg")
PICTURES = r"C:\Users\AndronikLindgren\OneDrive - CaMa Gruppen AB\Pictures"
BATCHER_DIR = os.path.join(PICTURES, "QA-hubb-batcher")
CONFIG_BACKUP_DIR = r"C:\Users\AndronikLindgren\agent-home\logs\hub-inbox"
PIN_SCRIPT = r"C:\Users\AndronikLindgren\agent-home\scripts\qa-hub\pictures-1110\pin-batches.py"
HUB_URL = "http://127.0.0.1:8092"
BATCH_ID = re.compile(r"^[a-z0-9][a-z0-9_-]{2,120}$")
IGNORE = {"desktop.ini", "thumbs.db"}
IMAGE_EXT = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".heic", ".gif", ".bmp"}


def norm(p):
    return os.path.normcase(os.path.normpath(p or ""))


def in_pictures(path):
    p, root = norm(path), norm(PICTURES)
    return p.startswith(root + os.sep)


def removed_sibling(path):
    p = os.path.normpath(path)
    return os.path.join(os.path.dirname(p), os.path.basename(p) + " Removed")


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def tree_hashes(root):
    """rel path (forward slashes) -> sha256; {} when root is absent. desktop.ini/Thumbs.db ignored."""
    out = {}
    if not os.path.isdir(root):
        return out
    for dirpath, _dirs, names in os.walk(root):
        for n in names:
            if n.lower() in IGNORE:
                continue
            full = os.path.join(dirpath, n)
            out[os.path.relpath(full, root).replace("\\", "/")] = sha256(full)
    return out


def is_image(rel):
    return os.path.splitext(rel)[1].lower() in IMAGE_EXT


def holds_delivery(src_main, src_removed, cand_main, cand_removed):
    """rel path -> sha256 maps in, True when the candidate copy shows exactly this delivery (see module doc). By path, not by
    content alone: a delivery often holds the same image twice (root + product folder). A reviewer removal moves a file to
    the same rel path under "<name> Removed"."""
    delivered = {**src_removed, **src_main}
    if not delivered:
        return False
    if any(h not in (cand_main.get(rel), cand_removed.get(rel)) for rel, h in delivered.items()):
        return False
    return all(delivered.get(rel) == h for rel, h in cand_main.items() if is_image(rel))


def plan(batch_id, inbox, batches, now=None):
    """Decide the action for one delivery; reads only. Returns a report dict with status:
    error_* | would_reuse (copy already holds it) | would_copy (dst free)."""
    rep = {"batch": batch_id}
    if not BATCH_ID.match(batch_id or ""):
        rep["status"] = "error_bad_batch_id"
        return rep
    src = os.path.join(inbox, batch_id)
    rep["src"] = src
    if not os.path.isdir(src):
        rep["status"] = "error_source_missing"
        return rep
    src_main, src_removed = tree_hashes(src), tree_hashes(removed_sibling(src))
    if not any(is_image(r) for r in src_main):
        rep["status"] = "error_no_images"
        return rep
    rep["files"] = {"main": len(src_main), "removed": len(src_removed)}
    entry = next((b for b in batches if b.get("id") == batch_id), None)
    rep["entry"] = "existing" if entry else "new"
    base = os.path.join(BATCHER_DIR, batch_id)
    candidates = []
    if entry and in_pictures(entry.get("path")):
        candidates.append(os.path.normpath(entry["path"]))
    if norm(base) not in {norm(c) for c in candidates}:
        candidates.append(base)
    for cand in candidates:
        if os.path.isdir(cand) and holds_delivery(src_main, src_removed, tree_hashes(cand), tree_hashes(removed_sibling(cand))):
            rep["dst"] = cand
            rep["status"] = "would_reuse"
            return rep
    dst = base
    if os.path.exists(dst) or os.path.exists(removed_sibling(dst)):
        dst = f"{base}__v{(now or datetime.datetime.now()):%y%m%d-%H%M}"
        if os.path.exists(dst) or os.path.exists(removed_sibling(dst)):
            rep["dst"] = dst
            rep["status"] = "error_collision"
            return rep
    other = next((b for b in batches if norm(b.get("path")) == norm(dst) and b.get("id") != batch_id), None)
    if other:
        rep["dst"] = dst
        rep["status"] = "error_path_owned_by_" + other.get("id", "?")
        return rep
    rep["dst"] = dst
    rep["status"] = "would_copy"
    return rep


def copy_verified(src, dst):
    """copytree src (+ Removed) -> dst, then compare every sha256. Returns (ok, counts|mismatch)."""
    counts = {}
    for s, d in [(src, dst), (removed_sibling(src), removed_sibling(dst))]:
        if not os.path.isdir(s):
            continue
        os.makedirs(os.path.dirname(d), exist_ok=True)
        shutil.copytree(s, d, copy_function=shutil.copy2, ignore=shutil.ignore_patterns("desktop.ini", "Thumbs.db"))
        hs, hd = tree_hashes(s), tree_hashes(d)
        if hs != hd:
            return False, sorted(k for k in set(hs) | set(hd) if hs.get(k) != hd.get(k))[:5]
        counts[os.path.basename(d)] = len(hd)
    return True, counts


def new_entry(batch_id, dst, meta):
    return {
        "id": batch_id,
        "name": meta.get("name") or batch_id,
        "path": dst,
        "category": meta["category"],
        "org": meta.get("org") if meta.get("org") in ("reforma", "cama") else "cama",
        "onedrive_url": "",
        "published": bool(meta.get("published", False)),
        "createdAt": datetime.datetime.now().isoformat(),
    }


def register(batches, batch_id, dst, meta):
    """Mutates batches: repoint the entry with batch_id, or append a new one. Returns the change record."""
    entry = next((b for b in batches if b.get("id") == batch_id), None)
    if entry:
        old = entry.get("path")
        if norm(old) == norm(dst):
            return {"batch": batch_id, "action": "unchanged", "new": dst}
        entry["path"] = dst
        return {"batch": batch_id, "action": "repointed", "old": old, "new": dst}
    batches.append(new_entry(batch_id, dst, meta))
    return {"batch": batch_id, "action": "registered", "new": dst}


def load_config():
    with open(BATCHES_FILE, encoding="utf-8") as f:
        return json.load(f)


def save_config(batches):
    tmp = BATCHES_FILE + ".inbox.tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(batches, f, indent=2, ensure_ascii=False)
    os.replace(tmp, BATCHES_FILE)


def live_count(batch_id):
    """products + shots through the local server (a local request is admin: no token)."""
    try:
        with urllib.request.urlopen(f"{HUB_URL}/api/products?batch={urllib.parse.quote(batch_id)}", timeout=60) as r:
            body = json.loads(r.read().decode("utf-8"))
    except Exception as e:  # noqa: BLE001 - reported, never fatal before the change
        return {"status": getattr(e, "code", type(e).__name__)}
    shots = 0
    stack = [body]
    while stack:
        o = stack.pop()
        if isinstance(o, list):
            stack.extend(o)
        elif isinstance(o, dict):
            if isinstance(o.get("rel_path"), str):
                shots += 1
            stack.extend(o.values())
    prods = body.get("products", []) if isinstance(body, dict) else body
    return {"status": 200, "products": len(prods), "shots": shots}


def read_meta(inbox, batch_id, args):
    meta = {}
    path = os.path.join(inbox, batch_id + ".json")
    if os.path.isfile(path):
        with open(path, encoding="utf-8") as f:
            meta = json.load(f)
    for k in ("name", "category", "org"):
        if getattr(args, k):
            meta[k] = getattr(args, k)
    if args.published:
        meta["published"] = True
    return meta


def cmd_list(inbox):
    if not os.path.isdir(inbox):
        print(json.dumps({"inbox": inbox, "status": "empty_no_inbox"}, ensure_ascii=False))
        return 0
    batches = load_config()
    for name in sorted(os.listdir(inbox)):
        if name.startswith("_") or name.endswith(" Removed") or not os.path.isdir(os.path.join(inbox, name)):
            continue
        print(json.dumps(plan(name, inbox, batches), ensure_ascii=False))
    return 0


def cmd_ingest(args, inbox):
    batches = load_config()
    rep = plan(args.batch, inbox, batches)
    meta = read_meta(inbox, args.batch, args)
    if rep["status"].startswith("error"):
        print(json.dumps(rep, ensure_ascii=False))
        return 1
    if rep["entry"] == "new" and not meta.get("category"):
        rep["status"] = "error_new_batch_needs_category"
        print(json.dumps(rep, ensure_ascii=False))
        return 2
    rep["count_before"] = live_count(args.batch)
    if args.dry_run:
        print(json.dumps(rep, ensure_ascii=False))
        return 0
    if rep["status"] == "would_copy":
        ok, info = copy_verified(rep["src"], rep["dst"])
        if not ok:
            rep["status"] = "error_hash_mismatch"
            rep["mismatch"] = info
            print(json.dumps(rep, ensure_ascii=False))
            return 1
        rep["files_verified"] = info
    os.makedirs(CONFIG_BACKUP_DIR, exist_ok=True)
    backup = os.path.join(CONFIG_BACKUP_DIR, f"batches_config.before-inbox-{datetime.datetime.now():%Y%m%d-%H%M%S}.json")
    shutil.copy2(BATCHES_FILE, backup)
    rep["config_backup"] = backup
    batches = load_config()  # re-read right before writing: the live server may have saved in between
    rep["config"] = register(batches, args.batch, rep["dst"], meta)
    save_config(batches)
    pin = subprocess.run([sys.executable, "-I", PIN_SCRIPT, HUB_DIR, args.batch], capture_output=True, text=True, encoding="utf-8")
    rep["pin"] = {"exit": pin.returncode, "last": (pin.stdout.strip().splitlines() or [""])[-1]}
    rep["count_after"] = live_count(args.batch)
    rep["status"] = "reused" if rep["status"] == "would_reuse" else "copied_verified"
    code = 0 if pin.returncode == 0 and rep["count_after"].get("status") == 200 else 1
    rep["ok"] = code == 0
    with open(os.path.join(inbox, "_ingested.jsonl"), "a", encoding="utf-8") as f:
        f.write(json.dumps({"ts": datetime.datetime.now().isoformat(timespec="seconds"), **rep}, ensure_ascii=False) + "\n")
    print(json.dumps(rep, ensure_ascii=False))
    return code


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=("list", "ingest"))
    ap.add_argument("--batch")
    ap.add_argument("--name")
    ap.add_argument("--category")
    ap.add_argument("--org", choices=("cama", "reforma"))
    ap.add_argument("--published", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--inbox", default=INBOX, help="override for dry-runs/tests")
    args = ap.parse_args()
    if args.cmd == "list":
        return cmd_list(args.inbox)
    if not args.batch:
        ap.error("ingest needs --batch")
    return cmd_ingest(args, args.inbox)


if __name__ == "__main__":
    sys.exit(main())
