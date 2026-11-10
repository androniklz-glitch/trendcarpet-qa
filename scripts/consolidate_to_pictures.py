"""Move a hub batch's SOURCE into the CEO's OneDrive Pictures (the backup) by copying, never moving or deleting.

Per batch whose path is outside Pictures:
  1. copy <path> -> Pictures\\<same folder name> (plus the "<name> Removed" sibling when it exists)
  2. verify file count + sha256 of every file
  3. back up batches_config.json, repoint "path" of every entry with that folder (ratings/reviews are keyed by batch id)
  4. (separate command `free`) attrib +U once OneDrive reports the copy uploaded
A folder name that already exists in Pictures stops that batch (reported). Originals stay untouched.

Usage (hub python = miniconda3):
  python scripts/consolidate_to_pictures.py copy --batch <id> [--batch ...] | --all  [--dry-run]
  python scripts/consolidate_to_pictures.py free --batch <id> [--batch ...] | --all
Exit: 0 OK, 1 error/collision, 2 usage.
"""
import argparse
import datetime
import hashlib
import json
import os
import shutil
import subprocess
import sys

HUB_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BATCHES_FILE = os.path.join(HUB_DIR, "batches_config.json")
PICTURES = r"C:\Users\AndronikLindgren\OneDrive - CaMa Gruppen AB\Pictures"
CONFIG_BACKUP_DIR = r"C:\Users\AndronikLindgren\agent-home\logs\hub-pictures-1110"
SYNC_STATUS_PS1 = r"C:\Users\AndronikLindgren\agent-home\scripts\onedrive-sync-status.ps1"


def norm(p):
    return os.path.normcase(os.path.normpath(p or ""))


def in_pictures(path):
    p, root = norm(path), norm(PICTURES)
    return p == root or p.startswith(root + os.sep)


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
    out = {}
    for dirpath, _dirs, names in os.walk(root):
        for n in names:
            full = os.path.join(dirpath, n)
            out[os.path.relpath(full, root).replace("\\", "/")] = sha256(full)
    return out


def load_config():
    with open(BATCHES_FILE, encoding="utf-8") as f:
        return json.load(f)


def save_config(batches):
    tmp = BATCHES_FILE + ".consolidate.tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(batches, f, indent=2, ensure_ascii=False)
    os.replace(tmp, BATCHES_FILE)


def pick(batches, args):
    if args.all:
        return [b for b in batches if not in_pictures(b.get("path"))]
    wanted = {x.lower() for x in args.batch}
    chosen = [b for b in batches if (b.get("id") or "").lower() in wanted]
    missing = wanted - {(b.get("id") or "").lower() for b in chosen}
    if missing:
        print(f"unknown batch id(s): {sorted(missing)}", file=sys.stderr)
        sys.exit(2)
    return chosen


def copy_batch(batch, dry_run):
    src = os.path.normpath(batch["path"])
    rep = {"batch": batch["id"], "src": src}
    if in_pictures(src):
        rep["status"] = "already_in_pictures"
        return rep
    if not os.path.isdir(src):
        rep["status"] = "error_source_missing"
        return rep
    dst = os.path.join(PICTURES, os.path.basename(src))
    pairs = [(src, dst)]
    if os.path.isdir(removed_sibling(src)):
        pairs.append((removed_sibling(src), removed_sibling(dst)))
    rep["dst"] = dst
    for s, d in pairs:
        if os.path.exists(d):
            rep["status"] = "collision"
            rep["collision"] = d
            return rep
    if dry_run:
        rep["status"] = "would_copy"
        rep["files"] = sum(len(n) for _p, _d, n in os.walk(src))
        return rep
    counts = {}
    for s, d in pairs:
        shutil.copytree(s, d, copy_function=shutil.copy2)
        hs, hd = tree_hashes(s), tree_hashes(d)
        if hs != hd:
            bad = sorted(k for k in set(hs) | set(hd) if hs.get(k) != hd.get(k))[:5]
            rep["status"] = "error_hash_mismatch"
            rep["mismatch"] = bad
            return rep
        counts[os.path.basename(d)] = len(hd)
    rep["files_verified"] = counts
    rep["status"] = "copied_verified"
    return rep


def repoint(reports):
    moved = {norm(r["src"]): r["dst"] for r in reports if r.get("status") == "copied_verified"}
    if not moved:
        return None, []
    os.makedirs(CONFIG_BACKUP_DIR, exist_ok=True)
    backup = os.path.join(CONFIG_BACKUP_DIR, f"batches_config.before-pictures-{datetime.datetime.now():%Y%m%d-%H%M%S}.json")
    shutil.copy2(BATCHES_FILE, backup)
    batches = load_config()  # re-read right before writing: the live server may have saved in between
    changed = []
    for b in batches:
        new = moved.get(norm(b.get("path")))
        if new:
            changed.append({"batch": b["id"], "old": b["path"], "new": new})
            b["path"] = new
    save_config(batches)
    return backup, changed


def free_batch(batch):
    p = os.path.normpath(batch["path"])
    if not in_pictures(p):
        return {"batch": batch["id"], "free": "skipped_not_in_pictures"}
    out = {"batch": batch["id"]}
    for d in (p, removed_sibling(p)):
        if not os.path.isdir(d):
            continue
        st = subprocess.run(["powershell", "-NoProfile", "-File", SYNC_STATUS_PS1, "-Path", d],
                            capture_output=True, text=True, encoding="utf-8", errors="replace")
        if st.returncode == 0:
            subprocess.run(["attrib", "+U", "-P", os.path.join(d, "*"), "/S", "/D"], check=True, capture_output=True)
            out[os.path.basename(d)] = "freed"
        else:
            out[os.path.basename(d)] = "not_synced: " + " | ".join(st.stdout.splitlines()[:2])
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=("copy", "free"))
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--batch", action="append")
    g.add_argument("--all", action="store_true", help="copy: every batch outside Pictures; free: every batch")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()
    batches = load_config()
    if args.cmd == "free":
        targets = batches if args.all else pick(batches, args)
        for b in targets:
            print(json.dumps(free_batch(b), ensure_ascii=False))
        return 0
    reports, seen, code = [], set(), 0
    for b in pick(batches, args):
        if norm(b.get("path")) in seen:
            continue  # duplicate entry for the same folder: repointed together below
        seen.add(norm(b.get("path")))
        rep = copy_batch(b, args.dry_run)
        reports.append(rep)
        print(json.dumps(rep, ensure_ascii=False))
        if rep["status"].startswith("error") or rep["status"] == "collision":
            code = 1
            if rep["status"].startswith("error"):
                break
    if not args.dry_run:
        backup, changed = repoint(reports)
        print(json.dumps({"config_backup": backup, "repointed": changed}, ensure_ascii=False))
    return code


if __name__ == "__main__":
    sys.exit(main())
