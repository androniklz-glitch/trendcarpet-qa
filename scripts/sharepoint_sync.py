"""Copy approved hub images to the CaMa SharePoint library (OneDrive-synced folder, so OneDrive uploads; no Graph API, no keys):
  SharePoint: <lib>\\QA-hubb Godkända\\<batch-id>\\<rel path in batch>
"Godkänt" = everything NOT underkänt (CEO 2026-11-10, see select_not_rejected). The backup is the source itself: every hub
batch lives in the CEO's OneDrive Pictures (pinned, local); there is no separate backup mirror (RÄTTELSE 3).
Gate (fail-closed per batch, ALARM): batch outside Pictures, a source file cloud-only / not pinned, unreadable ratings/reviews
or metadata. Only org "cama" is uploaded. Reforma is a separate company and is NEVER uploaded; unknown org = skip + log.
A file counts as uploaded only when the SharePoint copy's sha256 equals the source. Every upload is appended to the manifest
(JSONL, versioned in ~/qa-hub-backup). Nothing is ever deleted: files that left the selection are only logged, a SharePoint
file changed by someone else is left alone and reported.

Commands (hub python = miniconda3):
  sync    --batch <id> [--batch ...] | --all [--include-drafts]  [--dry-run]
  report  --batch <id> ... | --all   dry-run of the selection: files + GB vs gold, underkänt per rule (read-only)
  free    attrib +U on SharePoint library batch folders OneDrive reports as uploaded (never the hub sources)
  verify  SharePoint vs manifest by metadata (no download) + C: free >= 20 GB; deviation -> ALARM file, exit 3
  restore <batch-id>[/<rel path>] [--dry-run]   copy back from the Pictures source (sha256 must match; manual only)
Exit: 0 OK, 1 copy error / path guard (stop condition), 2 usage, 3 ALARM (verify deviation / gated batch), 4 lock held.
"""
import argparse
import datetime
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time

HUB_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, HUB_DIR)
import hatshop_server as hub  # noqa: E402  (import side effect: makedirs .cache/thumbnails only)

DEFAULT_LIB = r"C:\Users\AndronikLindgren\CaMa Gruppen AB\CaMa Gruppen AB - Dokument"
DEFAULT_ONEDRIVE = r"C:\Users\AndronikLindgren\OneDrive - CaMa Gruppen AB"
SP_ROOT_NAME = "QA-hubb Godkända"
PICTURES_NAME = "Pictures"
MIN_FREE_GB = 20
UPLOAD_ORGS = {"cama"}
EXCLUDED_ORGS = {"reforma"}
MANIFEST_REPO = r"C:\Users\AndronikLindgren\qa-hub-backup"
MANIFEST_REL = "sharepoint/manifest.jsonl"
SYNC_STATUS_PS1 = r"C:\Users\AndronikLindgren\agent-home\scripts\onedrive-sync-status.ps1"
LOG_DIR = os.path.join(HUB_DIR, "_evidence", "sharepoint-sync")
LOCK_FILE = os.path.join(HUB_DIR, ".sharepoint_sync.lock")
LOCK_STALE_S = 2 * 3600
FORBIDDEN_CHARS = '<>:"|?*'


class GuardError(Exception):
    pass


def org_decision(batch):
    org = (batch.get("org") or "").strip().lower()
    if org in UPLOAD_ORGS:
        return None
    if org in EXCLUDED_ORGS:
        return "excluded_org_reforma"
    return f"unknown_org:{org or '-'}"


def safe_component(name):
    cleaned = "".join("_" if c in FORBIDDEN_CHARS or ord(c) < 32 else c for c in name).strip().rstrip(".")
    if not cleaned or cleaned in (".", ".."):
        raise GuardError(f"unsafe path component: {name!r}")
    return cleaned


def rel_parts(rel):
    parts = [p for p in rel.replace("\\", "/").split("/") if p]
    if not parts or any(p in (".", "..") for p in parts):
        raise GuardError(f"unsafe rel path: {rel!r}")
    return [safe_component(p) for p in parts]


def under(root, batch_id, rel):
    """Path <root>\\<batch-id>\\<rel>, guarded to stay strictly inside root."""
    full = os.path.normpath(os.path.join(root, safe_component(batch_id), *rel_parts(rel)))
    root_n = os.path.normcase(os.path.normpath(root))
    if os.path.normcase(os.path.commonpath([root_n, os.path.normcase(full)])) != root_n or os.path.normcase(full) == root_n:
        raise GuardError(f"path outside {root}: {full}")
    return full


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def stat_sig(path):
    try:
        st = os.stat(path)
    except OSError:
        return None
    return st.st_size, int(st.st_mtime)


def versioned_name(path):
    base, ext = os.path.splitext(path)
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M")
    cand, n = f"{base}__{stamp}{ext}", 2
    while os.path.exists(cand):
        cand, n = f"{base}__{stamp}-{n}{ext}", n + 1
    return cand


class Manifest:
    def __init__(self, repo):
        self.repo = repo
        self.path = os.path.join(repo, *MANIFEST_REL.split("/"))
        self.latest = {}
        self.added = 0
        if os.path.isfile(self.path):
            with open(self.path, encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line:
                        e = json.loads(line)
                        self.latest[(e["batch"], e["rel"])] = e

    def append(self, entry):
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        with open(self.path, "a", encoding="utf-8", newline="\n") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
            f.flush()
            os.fsync(f.fileno())
        self.latest[(entry["batch"], entry["rel"])] = entry
        self.added += 1

    def commit(self, message):
        if not self.added or not os.path.isdir(os.path.join(self.repo, ".git")):
            return "nothing to commit" if not self.added else "no git repo"
        for attempt in range(5):
            add = subprocess.run(["git", "-C", self.repo, "add", "--", MANIFEST_REL], capture_output=True, text=True)
            com = subprocess.run(["git", "-C", self.repo, "commit", "-q", "-m", message, "--", MANIFEST_REL],
                                 capture_output=True, text=True)
            if add.returncode == 0 and com.returncode == 0:
                return "committed"
            if "index.lock" not in (add.stderr + com.stderr):
                return f"commit failed: {(add.stderr + com.stderr).strip()[:200]}"
            time.sleep(2 + attempt * 2)
        return "commit failed: index.lock busy"


def select_batches(args):
    batches = hub.load_batches()
    skipped = []
    if args.batch:
        wanted = {b.lower() for b in args.batch}
        chosen = [b for b in batches if (b.get("id") or "").lower() in wanted]
        missing = wanted - {(b.get("id") or "").lower() for b in chosen}
        if missing:
            print(f"unknown batch id(s): {sorted(missing)}", file=sys.stderr)
            sys.exit(2)
    else:
        chosen, seen_paths = [], {}
        # Published entries first so a draft duplicate of the same folder never wins.
        for b in sorted(batches, key=lambda x: not x.get("published", True)):
            if not args.include_drafts and not b.get("published", True):
                skipped.append({"batch": b.get("id"), "reason": "draft"})
                continue
            p = os.path.normcase(os.path.normpath(b.get("path") or ""))
            if p in seen_paths:
                skipped.append({"batch": b.get("id"), "reason": f"same folder as {seen_paths[p]}"})
                continue
            seen_paths[p] = b.get("id")
            chosen.append(b)
    # Hard org filter, also for an explicit --batch (the admin button goes through here).
    keep = []
    for b in chosen:
        reason = org_decision(b)
        if reason:
            skipped.append({"batch": b.get("id"), "reason": reason})
        else:
            keep.append(b)
    return keep, skipped


def sync_file(batch_id, src, rel, sp_root, manifest, dry_run):
    """Returns 'unchanged' | 'copied' | 'conflict'; raises OSError/GuardError on failure (stop condition)."""
    sp = under(sp_root, batch_id, rel)
    src_sig = stat_sig(src)
    sp_sig = stat_sig(sp)
    prev = manifest.latest.get((batch_id, rel))
    if prev and sp_sig == (prev["size"], prev["sp_mtime"]) and src_sig == (prev["src_size"], prev["src_mtime"]):
        return "unchanged"
    if prev and sp_sig != (prev["size"], prev["sp_mtime"]):
        return "conflict"  # edited or removed in SharePoint after our upload: leave it, verify/restore handles it
    if not prev and sp_sig and sp_sig != src_sig:
        return "conflict"  # a file we did not write; an identical pre-manifest copy is adopted after hashing below
    if dry_run:
        return "copied"
    digest = sha256(src)
    if not (sp_sig and sp_sig == src_sig and sha256(sp) == digest):
        os.makedirs(os.path.dirname(sp), exist_ok=True)
        shutil.copy2(src, sp)
    if sha256(sp) != digest:
        raise OSError(f"SharePoint hash mismatch: {sp}")
    size, mtime = stat_sig(sp)
    manifest.append({
        "ts": datetime.datetime.now().isoformat(timespec="seconds"), "batch": batch_id,
        "product": rel.replace("\\", "/").split("/")[0], "rel": rel, "sha256": digest, "size": size,
        "sp_mtime": mtime, "src": src, "src_size": src_sig[0], "src_mtime": src_sig[1],
        "sharepoint": sp,
    })
    return "copied"


REJECT_STATUSES = {"redo", "angle_issue"}
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".gif", ".avif", ".heic"}


class SelectionError(Exception):
    """Ratings/reviews file exists but cannot be read: the batch is skipped (fail-closed), never "all"."""


def _read_json_strict(path):
    """None if the file is missing (= batch not reviewed); raises SelectionError if present but unreadable."""
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as exc:
        raise SelectionError(f"unreadable {os.path.basename(path)}: {exc}")
    if not isinstance(data, dict):
        raise SelectionError(f"unreadable {os.path.basename(path)}: not an object")
    return data


def _low_rated_keys(ratings):
    """'<product>|<filename>' keys any reviewer gave 1-2 stars (undo_remove does not clear ratings)."""
    low = set()
    for reviewer, per in ratings.items():
        if not isinstance(per, dict):
            raise SelectionError(f"ratings: reviewer {reviewer!r} is not an object")
        for key, v in per.items():
            stars = v.get("stars") if isinstance(v, dict) else v
            if isinstance(stars, bool) or not isinstance(stars, (int, float)):
                raise SelectionError(f"ratings: bad stars for {key!r}")
            if 1 <= stars <= 2:
                low.add(key.lower())
    return low


def _count_removed(batch_path):
    root = hub.removed_root_for(batch_path)
    n = 0
    if os.path.isdir(root):
        for _dirpath, _dirs, names in os.walk(root):
            n += sum(1 for x in names if os.path.splitext(x)[1].lower() in IMAGE_EXTS)
    return n


ATTR_OFFLINE, ATTR_PINNED, ATTR_RECALL_ON_DATA_ACCESS = 0x1000, 0x80000, 0x400000
GATE_IGNORE = {"desktop.ini", "thumbs.db"}


def local_gate(batch, pictures_root):
    """C gate (Koordinator 4.2): the source must be in Pictures and every file pinned + local, else the batch is blocked."""
    p = os.path.normcase(os.path.normpath(os.path.abspath(batch.get("path") or "")))
    pics = os.path.normcase(os.path.normpath(pictures_root))
    if not (p == pics or p.startswith(pics + os.sep)):
        raise SelectionError(f"outside Pictures: {batch.get('path')}")
    bad = []
    for dirpath, _dirs, names in os.walk(batch.get("path")):
        if not os.stat(dirpath).st_file_attributes & ATTR_PINNED:  # folder pin = files added later inherit it
            bad.append(os.path.basename(dirpath) + os.sep)
        for n in names:
            if n.lower() in GATE_IGNORE:
                continue
            a = os.stat(os.path.join(dirpath, n)).st_file_attributes
            if a & (ATTR_OFFLINE | ATTR_RECALL_ON_DATA_ACCESS) or not a & ATTR_PINNED:
                bad.append(n)
    if bad:
        raise SelectionError(f"{len(bad)} file(s) cloud-only or not pinned, e.g. {bad[0]}")


def select_not_rejected(batch, pictures_root):
    """CEO 2026-11-10 "i princip allt som inte är underkänt i mapparna" (definition: Koordinator 4.2).
    Underkänt = (1) in "<batch> Removed" (not scanned at all), (2) any reviewer gave 1-2 stars, (3) product status
    redo|angle_issue. Everything else is approved: unrated, pending/approved status, flags only.
    Returns {"files": [(abs, rel)], "gold": [(abs, rel)], "rejected": {"1": n, "2": n, "3": n}}.
    Cloud-only sources are kept (reading them hydrates); size comes from the placeholder."""
    bid = batch["id"]
    local_gate(batch, pictures_root)
    reviews = _read_json_strict(hub.get_review_file(bid))
    ratings = _read_json_strict(hub.get_ratings_file(bid))
    low = _low_rated_keys(ratings) if ratings else set()
    bad_products = set()
    for pid, rv in (reviews or {}).items():
        if not isinstance(rv, dict):
            raise SelectionError(f"reviews: product {pid!r} is not an object")
        if str(rv.get("status") or "").strip().lower() in REJECT_STATUSES:
            bad_products.add(str(pid).lower())

    batch_path = os.path.normpath(os.path.abspath(batch.get("path", "")))
    if not os.path.isdir(batch_path):
        raise SelectionError("batch folder missing")
    # scan_batch_products swallows unreadable metadata (products_manifest.json, upload_order.txt ...) and silently
    # returns another selection; a cloud-only file OneDrive cannot hydrate raises OSError 22 here -> fail-closed.
    for dirpath, _dirs, names in os.walk(batch_path):
        for n in names:
            if os.path.splitext(n)[1].lower() in (".json", ".txt", ".csv"):
                try:
                    with open(os.path.join(dirpath, n), "rb") as f:
                        f.read(1)
                except OSError as exc:
                    raise SelectionError(f"unreadable metadata {os.path.relpath(os.path.join(dirpath, n), batch_path)}: {exc}")
    out ={"files": [], "gold": [], "rejected": {"1": _count_removed(batch_path), "2": 0, "3": 0}}
    seen = set()
    for prod in hub.scan_batch_products(batch):
        pid = str(prod.get("id"))
        variants = prod.get("variants") or {}
        gold_keys = ["Original"] if "Original" in variants else list(variants.keys())[:1]
        for vk, shots in variants.items():
            for s in shots or []:
                rel = (s.get("rel_path") or "").replace("\\", "/").lstrip("/")
                full = os.path.normpath(os.path.join(batch_path, rel)) if rel else ""
                if not rel or rel in seen or not os.path.isfile(full):
                    continue
                if os.path.normcase(os.path.commonpath([batch_path, full])) != os.path.normcase(batch_path):
                    continue
                seen.add(rel)
                if vk in gold_keys and s.get("site_eligible", True) and (s.get("is_studio_packshot") or s.get("in_gold_set")):
                    out["gold"].append((full, rel))
                names = {str(s.get("filename") or ""), os.path.basename(rel)}
                if pid.lower() in bad_products:
                    out["rejected"]["3"] += 1
                elif any(f"{pid}|{n}".lower() in low for n in names if n):
                    out["rejected"]["2"] += 1
                else:
                    out["files"].append((full, rel))
    return out


def cmd_report(args):
    """Dry-run of the not-rejected selection vs gold, per batch: files + GB, rejected per rule. Read-only."""
    batches, skipped = select_batches(args)
    gb = lambda files: round(sum(os.path.getsize(a) for a, _ in files) / 1e9, 3)  # noqa: E731
    rows, tot = [], {"files": 0, "gb": 0.0, "gold_files": 0, "gold_gb": 0.0, "r1": 0, "r2": 0, "r3": 0}
    code = 0
    for b in batches:
        row = {"batch": b["id"]}
        try:
            sel = select_not_rejected(b, os.path.join(args.onedrive, PICTURES_NAME))
        except SelectionError as exc:
            row.update(status="ALARM_skipped", message=str(exc))
            print(f"[ALARM] {b['id']}: {exc}", file=sys.stderr)
            code = 3
            rows.append(row)
            continue
        row.update(status="ok", files=len(sel["files"]), gb=gb(sel["files"]), gold_files=len(sel["gold"]),
                   gold_gb=gb(sel["gold"]), r1=sel["rejected"]["1"], r2=sel["rejected"]["2"], r3=sel["rejected"]["3"])
        for k in tot:
            tot[k] += row[k]
        rows.append(row)
        print(f"{b['id']}: files={row['files']} {row['gb']} GB | gold={row['gold_files']} {row['gold_gb']} GB | "
              f"underkänt r1={row['r1']} r2={row['r2']} r3={row['r3']} | pictures={row['in_pictures']}")
    tot["gb"], tot["gold_gb"] = round(tot["gb"], 3), round(tot["gold_gb"], 3)
    print(f"TOTAL batches={len(rows)}: files={tot['files']} {tot['gb']} GB | gold={tot['gold_files']} {tot['gold_gb']} GB"
          f" | underkänt r1={tot['r1']} r2={tot['r2']} r3={tot['r3']}")
    return code, {"skipped_batches": skipped, "batches": rows, "total": tot}


def sync_batch(batch, sp_root, pictures_root, manifest, dry_run):
    bid = batch["id"]
    rep = {"batch": bid, "org": batch.get("org"), "copied": [], "unchanged": 0, "conflicts": [],
           "deselected": [], "status": "ok"}
    try:
        sel = select_not_rejected(batch, pictures_root)
    except SelectionError as exc:
        rep["status"] = "ALARM_blocked"
        rep["message"] = str(exc)
        return rep
    rep["rejected"] = sel["rejected"]
    wanted = set()
    for src, rel in sel["files"]:
        wanted.add(os.path.normcase(under(sp_root, bid, rel)))
        res = sync_file(bid, src, rel, sp_root, manifest, dry_run)
        if res == "unchanged":
            rep["unchanged"] += 1
        elif res == "conflict":
            rep["conflicts"].append(rel)
        else:
            rep["copied"].append(rel)
    batch_dir = os.path.join(sp_root, safe_component(bid))
    if os.path.isdir(batch_dir):
        for dirpath, _dirs, names in os.walk(batch_dir):
            for n in names:
                full = os.path.join(dirpath, n)
                if os.path.normcase(full) not in wanted:
                    rep["deselected"].append(os.path.relpath(full, batch_dir).replace("\\", "/"))
    return rep


def folder_synced(path):
    st = subprocess.run(["powershell", "-NoProfile", "-File", SYNC_STATUS_PS1, "-Path", path],
                        capture_output=True, text=True, encoding="utf-8", errors="replace")
    return st.returncode == 0, " | ".join(l for l in st.stdout.splitlines()[:3] if l).strip()


def cmd_free(args, roots):
    out = []
    for root in roots:
        if not os.path.isdir(root):
            continue
        for name in sorted(os.listdir(root)):
            d = os.path.join(root, name)
            if not os.path.isdir(d):
                continue
            ok, info = folder_synced(d)
            if ok:
                subprocess.run(["attrib", "+U", "-P", os.path.join(d, "*"), "/S", "/D"], check=True, capture_output=True)
            out.append({"folder": d, "freed": ok, "status": info})
            print(f"{'freed' if ok else 'not synced, kept'}: {d} ({info})")
    return 0, {"free": out}


def cmd_verify(args, manifest):
    missing, changed = [], []
    for (bid, rel), e in sorted(manifest.latest.items()):
        sig = stat_sig(e["sharepoint"])  # metadata only, never hydrates a freed file
        if sig is None:
            missing.append(e["sharepoint"])
        elif sig != (e["size"], e["sp_mtime"]):
            changed.append(e["sharepoint"])
    free_gb = round(shutil.disk_usage("C:\\").free / 1e9, 1)
    low_disk = [f"C: free {free_gb} GB < {MIN_FREE_GB} GB (hub sources are pinned locally)"] if free_gb < MIN_FREE_GB else []
    res = {"entries": len(manifest.latest), "missing": missing, "changed": changed, "c_free_gb": free_gb}
    print(f"verify: entries={len(manifest.latest)} missing={len(missing)} changed={len(changed)} c_free_gb={free_gb}")
    if missing or changed or low_disk:
        alarm = os.path.join(MANIFEST_REPO, "sharepoint",
                             "ALARM-" + datetime.datetime.now().strftime("%Y%m%d-%H%M%S") + ".txt")
        os.makedirs(os.path.dirname(alarm), exist_ok=True)
        with open(alarm, "w", encoding="utf-8") as f:
            f.write("SharePoint verify deviation. Restore: python scripts/sharepoint_sync.py restore <batch-id>[/<rel>]\n")
            for label, items in (("MISSING", missing), ("CHANGED", changed), ("LOW_DISK", low_disk)):
                for p in items:
                    f.write(f"{label}\t{p}\n")
        res["alarm"] = alarm
        print(f"ALARM: {alarm}")
        return 3, res
    return 0, res


def source_for(e):
    """The Pictures source of a manifest entry: the batch's current path + rel, else the recorded src; sha256 must match."""
    cands = [os.path.join(b.get("path") or "", *e["rel"].replace("\\", "/").split("/"))
             for b in hub.load_batches() if (b.get("id") or "") == e["batch"]] + [e.get("src") or ""]
    for c in cands:
        if c and os.path.isfile(c) and sha256(c) == e["sha256"]:
            return c
    raise OSError(f"no source with matching sha256 for {e['batch']}/{e['rel']}")


def cmd_restore(args, manifest):
    target = args.target.replace("\\", "/").strip("/")
    bid, _, rel = target.partition("/")
    picked = [e for (b, r), e in manifest.latest.items() if b == bid and (not rel or r.replace("\\", "/") == rel)]
    if not picked:
        print(f"no manifest entries for {target}", file=sys.stderr)
        return 2, {}
    done = []
    for e in picked:
        if stat_sig(e["sharepoint"]) == (e["size"], e["sp_mtime"]):
            continue
        src = source_for(e)
        under(os.path.join(args.lib, SP_ROOT_NAME), e["batch"], e["rel"])  # path guard
        if not args.dry_run:
            os.makedirs(os.path.dirname(e["sharepoint"]), exist_ok=True)
            shutil.copy2(src, e["sharepoint"])
            if sha256(e["sharepoint"]) != e["sha256"]:
                raise OSError(f"restore hash mismatch: {e['sharepoint']}")
        done.append(e["sharepoint"])
        print(f"{'would restore' if args.dry_run else 'restored'}: {e['sharepoint']}")
    print(f"restore: {len(done)} file(s), {len(picked) - len(done)} already intact")
    return 0, {"restored": done}


def cmd_sync(args, sp_root, pictures_root, manifest):
    batches, skipped = select_batches(args)
    for s in skipped:
        if s["reason"].startswith(("unknown_org", "excluded_org")):
            print(f"[skip] {s['batch']}: {s['reason']}")
    res = {"skipped_batches": skipped, "batches": []}
    code = 0
    for b in batches:
        try:
            rep = sync_batch(b, sp_root, pictures_root, manifest, args.dry_run)
        except (GuardError, OSError) as exc:
            res["batches"].append({"batch": b.get("id"), "status": "error", "message": str(exc)})
            print(f"[STOP] {b.get('id')}: {exc}", file=sys.stderr)
            code = 1
            break
        res["batches"].append(rep)
        if rep["status"] == "ALARM_blocked":
            print(f"[ALARM] {rep['batch']}: blocked: {rep['message']}", file=sys.stderr)
            code = code or 3
            continue
        print(f"{rep['batch']}: {rep['status']} copied={len(rep['copied'])} unchanged={rep['unchanged']} "
              f"conflicts={len(rep['conflicts'])} deselected={len(rep['deselected'])} rejected={rep['rejected']}")
    res["manifest_added"] = manifest.added
    res["manifest_commit"] = manifest.commit(f"sharepoint sync: {manifest.added} file(s)") if not args.dry_run else "dry-run"
    return code, res


def acquire_lock():
    try:
        if time.time() - os.path.getmtime(LOCK_FILE) > LOCK_STALE_S:
            os.remove(LOCK_FILE)
    except OSError:
        pass
    try:
        fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        return False
    os.write(fd, f"{os.getpid()} {datetime.datetime.now().isoformat()}".encode())
    os.close(fd)
    return True


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--lib", default=DEFAULT_LIB, help="local path of the synced SharePoint library")
    ap.add_argument("--onedrive", default=DEFAULT_ONEDRIVE, help="local path of the CEO's OneDrive (Pictures = hub sources)")
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("sync")
    g = s.add_mutually_exclusive_group(required=True)
    g.add_argument("--batch", action="append", help="batch id (repeatable)")
    g.add_argument("--all", action="store_true", help="every published batch (one entry per folder)")
    s.add_argument("--include-drafts", action="store_true", help="with --all: also draft batches")
    s.add_argument("--dry-run", action="store_true")
    rp = sub.add_parser("report")
    gr = rp.add_mutually_exclusive_group(required=True)
    gr.add_argument("--batch", action="append", help="batch id (repeatable)")
    gr.add_argument("--all", action="store_true", help="every published batch (one entry per folder)")
    rp.add_argument("--include-drafts", action="store_true")
    sub.add_parser("free")
    sub.add_parser("verify")
    r = sub.add_parser("restore")
    r.add_argument("target", help="<batch-id> or <batch-id>/<rel path>")
    r.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    for p in (args.lib, args.onedrive):
        if not os.path.isdir(p):
            print(f"folder not found: {p}", file=sys.stderr)
            return 2
    sp_root = os.path.join(args.lib, SP_ROOT_NAME)
    pictures_root = os.path.join(args.onedrive, PICTURES_NAME)
    if not acquire_lock():
        print(f"another run holds {LOCK_FILE}", file=sys.stderr)
        return 4
    started = datetime.datetime.now()
    try:
        manifest = Manifest(MANIFEST_REPO)
        if args.cmd == "sync":
            code, res = cmd_sync(args, sp_root, pictures_root, manifest)
        elif args.cmd == "report":
            code, res = cmd_report(args)
        elif args.cmd == "free":
            code, res = cmd_free(args, (sp_root,))
        elif args.cmd == "verify":
            code, res = cmd_verify(args, manifest)
        else:
            code, res = cmd_restore(args, manifest)
    finally:
        os.remove(LOCK_FILE)
    res.update({"cmd": args.cmd, "started": started.isoformat(timespec="seconds"),
                "finished": datetime.datetime.now().isoformat(timespec="seconds"), "exit_code": code,
                "dry_run": getattr(args, "dry_run", False), "sharepoint_root": sp_root, "pictures_root": pictures_root})
    os.makedirs(LOG_DIR, exist_ok=True)
    log_path = os.path.join(LOG_DIR, f"{started:%Y%m%d-%H%M%S}-{args.cmd}"
                            + ("-dry" if getattr(args, "dry_run", False) else "") + ".json")
    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=1)
    print(f"log: {log_path}")
    return code


if __name__ == "__main__":
    sys.exit(main())
