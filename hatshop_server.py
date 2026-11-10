import http.server
import socketserver
import os
import json
import hmac
import secrets
import http.cookies
import urllib.parse
import sys
import subprocess
import mimetypes
import re
import datetime
import openpyxl
import hashlib
import uuid
import shutil
import zipfile
import io
import csv
import threading
from collections import OrderedDict
from PIL import Image

_BATCHES_LOCK = threading.Lock()
_RATINGS_LOCK = threading.Lock()
_MERGE_LOCK = threading.Lock()

PORT = 8092
WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
REVIEWS_DIR = os.environ.get("QA_HUB_REVIEWS_DIR") or WORKSPACE_DIR
BATCHES_FILE = os.path.join(WORKSPACE_DIR, "batches_config.json")
CALENDAR_CACHE_FILE = os.path.join(WORKSPACE_DIR, "calendar_projects.json")
THUMB_CACHE_DIR = os.path.join(WORKSPACE_DIR, ".cache", "thumbnails")
os.makedirs(THUMB_CACHE_DIR, exist_ok=True)
EXCEL_MASTER_FILE = r"C:\Users\AndronikLindgren\OneDrive - CaMa Gruppen AB\Fotostudion 230530-ongoing.xlsx"
EXCEL_2026_FILE = r"C:\Users\AndronikLindgren\OneDrive - CaMa Gruppen AB\Pictures\Projektlista_2026.xlsx"
BASE_SP_URL = "https://trendcarpet-my.sharepoint.com"

REMOVE_REASONS = {
    "anatomi": "Händer/anatomi",
    "fel_farg": "Fel färg",
    "fransar": "Fransar",
    "rummet": "Rummet/miljön",
    "produktdetalj": "Produktdetalj",
    "styling": "Styling/person",
    "annat": "Annat",
}

REMOVE_FOR_VIEWERS = True
LOW_SHOT_THRESHOLD = 2
NOTIFY_SCRIPT = os.environ.get("QA_HUB_NOTIFY_SCRIPT", r"C:\Users\AndronikLindgren\agent-home\scripts\notify.mjs")

CROP_FORMAT_DIRS = ("3x4", "4x5")
MASTER_FORMAT_DIR = "9x16"
IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".webp")
VIDEO_EXTS = (".mp4", ".webm", ".mov")
MEDIA_EXTS = IMAGE_EXTS + VIDEO_EXTS

def hidden_crop_dirs(dir_names):
    """
    Returns the set of names in dir_names that are in CROP_FORMAT_DIRS (case-insensitive) —
    but ONLY if MASTER_FORMAT_DIR is also in dir_names (case-insensitive).
    A folder with 3x4 but no 9x16 sibling (e.g. Reforma 3:4 deliveries) stays visible.
    """
    if not dir_names:
        return set()
    crop_lower = {c.lower() for c in CROP_FORMAT_DIRS}
    master_lower = MASTER_FORMAT_DIR.lower()
    has_master = any(isinstance(d, str) and d.lower() == master_lower for d in dir_names)
    if not has_master:
        return set()
    return {d for d in dir_names if isinstance(d, str) and d.lower() in crop_lower}

def crop_siblings(master_abs_path, batch_path=None):
    """
    If the file's parent folder name is 9x16 (case-insensitive), for each crop dir in
    CROP_FORMAT_DIRS that exists next to it, return existing files
    <parent>/../<crop>/<stem>-<crop><ext> for ext in .png/.jpg/.jpeg/.webp,
    plus <stem><ext> (same stem without suffix). Only paths inside the batch folder.
    """
    if not master_abs_path:
        return []
    master_abs = os.path.normpath(os.path.abspath(master_abs_path))
    parent_dir = os.path.dirname(master_abs)
    if os.path.basename(parent_dir).lower() != MASTER_FORMAT_DIR.lower():
        return []

    model_dir = os.path.dirname(parent_dir)
    if not os.path.isdir(model_dir):
        return []

    filename = os.path.basename(master_abs)
    stem, _ = os.path.splitext(filename)
    if not stem:
        return []

    try:
        model_entries = os.listdir(model_dir)
    except OSError:
        return []

    sibling_dirs_map = {}
    for entry in model_entries:
        if os.path.isdir(os.path.join(model_dir, entry)):
            sibling_dirs_map[entry.lower()] = entry

    crop_exts = set(MEDIA_EXTS)
    results = []
    seen = set()

    for crop in CROP_FORMAT_DIRS:
        actual_crop_name = sibling_dirs_map.get(crop.lower())
        if not actual_crop_name:
            continue
        crop_dir = os.path.join(model_dir, actual_crop_name)
        try:
            crop_files = os.listdir(crop_dir)
        except OSError:
            continue

        target_stems = {f"{stem.lower()}-{crop.lower()}", stem.lower()}
        for cf in sorted(crop_files):
            cf_path = os.path.join(crop_dir, cf)
            if not os.path.isfile(cf_path):
                continue
            cf_stem, cf_ext = os.path.splitext(cf)
            if cf_ext.lower() not in crop_exts:
                continue
            if cf_stem.lower() in target_stems:
                cf_abs = os.path.normpath(os.path.abspath(cf_path))
                if batch_path:
                    bp = batch_path.get("path", "") if isinstance(batch_path, dict) else batch_path
                    if bp:
                        bp_norm = os.path.normcase(os.path.abspath(bp))
                        c_norm = os.path.normcase(cf_abs)
                        try:
                            if os.path.normcase(os.path.commonpath([bp_norm, c_norm])) != bp_norm or c_norm == bp_norm:
                                continue
                        except ValueError:
                            continue
                if cf_abs not in seen:
                    seen.add(cf_abs)
                    results.append(cf_abs)

    return results


def viewer_may_remove(is_admin, reviewer):
    if is_admin:
        return True
    if not REMOVE_FOR_VIEWERS:
        return False
    if not reviewer or not isinstance(reviewer, str):
        return False
    r = reviewer.strip()
    return bool(r and r != "Anonym")


def can_delete_comment(is_admin, current_user_obj, reviewer, comment_author):
    if is_admin:
        return True
    if current_user_obj and "name" in current_user_obj and current_user_obj["name"] == comment_author:
        return True
    rev = reviewer.strip() if isinstance(reviewer, str) else ""
    auth = comment_author.strip() if isinstance(comment_author, str) else ""
    if rev and rev.lower() != "anonym" and auth and rev.lower() == auth.lower():
        return True
    return False

def get_users_file():
    return os.environ.get("QA_HUB_USERS_FILE", r"C:\Users\AndronikLindgren\.secrets\qa-hub-users.json")

_users_cache = None
_users_mtime = 0

def load_users():
    global _users_cache, _users_mtime
    ufile = get_users_file()
    if not os.path.exists(ufile):
        return {"enforce": False, "users": []}
    
    try:
        mtime = os.path.getmtime(ufile)
        if _users_cache is None or mtime > _users_mtime:
            with open(ufile, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict) or not isinstance(data.get("users"), list):
                raise ValueError("Bad shape")
            for u in data.get("users", []):
                if not isinstance(u, dict):
                    raise ValueError("Bad shape")
            _users_cache = data
            _users_mtime = mtime
    except Exception:
        if _users_cache is not None:
            return _users_cache
        return {"enforce": True, "users": []}
    return _users_cache


def clean_sp_url(target):
    if not target:
        return ""
    if target.startswith("http"):
        return target
    target = target.replace("../../../", "/")
    if not target.startswith("/"):
        target = "/" + target
    return BASE_SP_URL + target

def parse_project_date(d_str):
    if not d_str:
        return ("", "", "")
    d_str = str(d_str).split(" ")[0].strip()
    m = re.search(r'(\d{4})[-_ ](\d{1,2})[-_ ](\d{1,2})', d_str)
    if m:
        y, mo, d = m.group(1), m.group(2).zfill(2), m.group(3).zfill(2)
        return (f"{y}-{mo}-{d}", f"{y}-{mo}", y)
    m2 = re.search(r'(\d{2})[-_ ](\d{1,2})[-_ ](\d{1,2})', d_str)
    if m2:
        y, mo, d = "20" + m2.group(1), m2.group(2).zfill(2), m2.group(3).zfill(2)
        return (f"{y}-{mo}-{d}", f"{y}-{mo}", y)
    return (d_str, "", "")

def load_calendar_projects(force_refresh=False):
    if not force_refresh and os.path.exists(CALENDAR_CACHE_FILE):
        try:
            with open(CALENDAR_CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass

    all_projects = []

    # 1. Parse Fotostudion 230530-ongoing.xlsx
    if os.path.exists(EXCEL_MASTER_FILE):
        try:
            wb = openpyxl.load_workbook(EXCEL_MASTER_FILE, data_only=False)
            for sname in wb.sheetnames:
                ws = wb[sname]
                for r in range(2, ws.max_row + 1):
                    c_date = ws.cell(r, 1)
                    c_proj = ws.cell(r, 2)
                    c_photo = ws.cell(r, 3)
                    c_interiors = ws.cell(r, 4)
                    c_uploaded = ws.cell(r, 5)
                    c_png = ws.cell(r, 6)
                    c_orig = ws.cell(r, 7) if ws.max_column >= 7 else None

                    date_raw = str(c_date.value or "").strip()
                    proj_title = str(c_proj.value or "").strip()
                    if not proj_title and not date_raw:
                        continue

                    iso_date, year_month, year = parse_project_date(date_raw)
                    proj_link = clean_sp_url(c_proj.hyperlink.target if c_proj.hyperlink else "")
                    png_link = clean_sp_url(c_png.hyperlink.target if c_png and c_png.hyperlink else "")
                    orig_link = clean_sp_url(c_orig.hyperlink.target if c_orig and c_orig.hyperlink else "")

                    all_projects.append({
                        "id": f"fs_{sname}_{r}",
                        "sheet": sname,
                        "raw_date": date_raw,
                        "date": iso_date,
                        "year_month": year_month,
                        "year": year,
                        "title": proj_title,
                        "link": proj_link,
                        "photo_status": str(c_photo.value or "").strip(),
                        "interiors": str(c_interiors.value or "").strip(),
                        "uploaded": str(c_uploaded.value or "").strip(),
                        "png_link": png_link,
                        "orig_link": orig_link
                    })
        except Exception as e:
            print(f"Error parsing master Excel: {e}")

    # 2. Parse Projektlista_2026.xlsx
    if os.path.exists(EXCEL_2026_FILE):
        try:
            wb2 = openpyxl.load_workbook(EXCEL_2026_FILE, data_only=False)
            ws2 = wb2.active
            for r in range(2, ws2.max_row + 1):
                c_date = ws2.cell(r, 1)
                c_proj = ws2.cell(r, 2)
                c_photo = ws2.cell(r, 3)
                c_interiors = ws2.cell(r, 4)
                c_uploaded = ws2.cell(r, 5)

                date_raw = str(c_date.value or "").strip()
                proj_title = str(c_proj.value or "").strip()
                if not proj_title or date_raw.startswith("1905"):
                    continue

                iso_date, year_month, year = parse_project_date(date_raw)
                proj_link = clean_sp_url(c_proj.hyperlink.target if c_proj.hyperlink else "")

                if not any(p["title"].lower() == proj_title.lower() for p in all_projects):
                    all_projects.append({
                        "id": f"p2026_{r}",
                        "sheet": "Projekt 2026",
                        "raw_date": date_raw,
                        "date": iso_date,
                        "year_month": year_month,
                        "year": year,
                        "title": proj_title,
                        "link": proj_link,
                        "photo_status": str(c_photo.value or "").strip(),
                        "interiors": str(c_interiors.value or "").strip(),
                        "uploaded": str(c_uploaded.value or "").strip(),
                        "png_link": "",
                        "orig_link": ""
                    })
        except Exception as e:
            print(f"Error parsing 2026 Excel: {e}")

    try:
        with open(CALENDAR_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(all_projects, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"Error caching calendar projects: {e}")

    return all_projects

def get_onedrive_web_url_for_path(local_path):
    if not local_path:
        return ""
    base_od = r"C:\Users\AndronikLindgren\OneDrive - CaMa Gruppen AB"
    flow_vibes_base = r"C:\Users\AndronikLindgren\flow-vibes"
    norm = os.path.normpath(local_path)
    if norm.lower().startswith(base_od.lower()):
        rel = os.path.relpath(norm, base_od).replace('\\', '/')
        encoded_rel = urllib.parse.quote(f"/personal/andronik_lindgren_trendcarpet_se/Documents/{rel}", safe='')
        return f"https://trendcarpet-my.sharepoint.com/personal/andronik_lindgren_trendcarpet_se/_layouts/15/onedrive.aspx?id={encoded_rel}"
    elif norm.lower().startswith(flow_vibes_base.lower()):
        sub = os.path.relpath(norm, flow_vibes_base).replace('\\', '/')
        if sub.lower().startswith("reforma-flow-"):
            rel = "Pictures/Reforma_Flow_Godkanda_2026-10-15"
        else:
            rel = f"Pictures/Flow_Vibes/{sub}"
        encoded_rel = urllib.parse.quote(f"/personal/andronik_lindgren_trendcarpet_se/Documents/{rel}", safe='')
        return f"https://trendcarpet-my.sharepoint.com/personal/andronik_lindgren_trendcarpet_se/_layouts/15/onedrive.aspx?id={encoded_rel}"
    return ""

def load_batches():
    batches = []
    if os.path.exists(BATCHES_FILE):
        with _BATCHES_LOCK:
            try:
                with open(BATCHES_FILE, "r", encoding="utf-8") as f:
                    batches = json.load(f)
            except Exception:
                batches = []
    elif not batches:
        batches = [
            {
                "id": "26-08-28-hatshop-black-river",
                "name": "2026-08-28 Hatshop Black River (14 modeller)",
                "path": r"C:\Users\AndronikLindgren\OneDrive - CaMa Gruppen AB\Pictures\26-08-28-Hatshop-Black-River pt1-x",
                "category": "kepsar",
                "org": "cama",
                "published": True,
                "onedrive_url": "",
                "createdAt": "2026-09-02T11:00:00Z"
            }
        ]
        batches_dir = os.path.dirname(BATCHES_FILE)
        if not batches_dir or os.path.isdir(batches_dir):
            save_batches(batches)

    # Auto-match with calendar_projects in memory if onedrive_url is missing, and upgrade legacy /my?id= links
    try:
        projects = load_calendar_projects()
        for b in batches:
            cur_od = b.get("onedrive_url", "") or ""
            if "trendcarpet-my.sharepoint.com/my?id=" in cur_od:
                b["onedrive_url"] = cur_od.replace(
                    "trendcarpet-my.sharepoint.com/my?id=",
                    "trendcarpet-my.sharepoint.com/personal/andronik_lindgren_trendcarpet_se/_layouts/15/onedrive.aspx?id="
                )
                cur_od = b["onedrive_url"]
            if not cur_od:
                folder = os.path.basename(os.path.normpath(b.get("path", ""))).strip().lower()
                matched_link = ""
                if len(folder) >= 8:
                    for p in projects:
                        ptitle = p.get("title", "").strip().lower()
                        if ptitle and p.get("link") and folder == ptitle:
                            matched_link = p["link"]
                            break
                if matched_link:
                    b["onedrive_url"] = matched_link
                else:
                    auto_url = get_onedrive_web_url_for_path(b.get("path", ""))
                    if auto_url:
                        b["onedrive_url"] = auto_url
    except Exception:
        pass

    return batches


def infer_batch_company(b):
    if not isinstance(b, dict):
        return "trendcarpet"
    explicit = str(b.get("company") or "").strip().lower()
    if explicit in ("trendcarpet", "reforma", "hatshop"):
        return explicit
    hay = f"{b.get('id', '')} {b.get('name', '')} {b.get('path', '')}".lower()
    if str(b.get("org") or "").strip().lower() == "reforma" or "reforma" in hay:
        return "reforma"
    cat = str(b.get("category") or "").strip().lower()
    if cat == "kepsar" or any(tok in hay for tok in ("hatshop", "keps", "hatt", "black-river", "black river", "gårda", "garda", "army-cap")):
        return "hatshop"
    return "trendcarpet"


def clean_batch_label(b):
    if not isinstance(b, dict):
        return ""
    raw = str(b.get("label") or b.get("name") or b.get("id") or "").strip()
    # Strip noisy status/storage prefixes injected in legacy labels
    cleaned = re.sub(r'^(?:(?:🔒|⚠️|🌍)\s*)?(?:\[(?:Utkast|Lagring|Privat)\]\s*)+', '', raw, flags=re.IGNORECASE)
    cleaned = re.sub(r'(?:🔒|⚠️)\s*\[(?:Utkast|Lagring)\]\s*', '', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'\s+', ' ', cleaned).strip(" -·")
    return cleaned or raw


def save_batches(batches):
    if not batches and os.path.exists(BATCHES_FILE):
        return False
    with _BATCHES_LOCK:
        batches_dir = os.path.dirname(BATCHES_FILE)
        if batches_dir:
            os.makedirs(batches_dir, exist_ok=True)
        tmp_file = f"{BATCHES_FILE}.{uuid.uuid4().hex}.tmp"
        try:
            with open(tmp_file, "w", encoding="utf-8") as f:
                json.dump(batches, f, indent=2, ensure_ascii=False)
            os.replace(tmp_file, BATCHES_FILE)
            return True
        except Exception as e:
            if os.path.exists(tmp_file):
                try:
                    os.remove(tmp_file)
                except OSError:
                    pass
            print(f"Error saving batches: {e}")
            return False

def default_published_batch_id(batches):
    for b in batches:
        if b.get("published", True):
            return b["id"]
    return batches[0]["id"] if batches else "default"

def get_batch(batch_id):
    batches = load_batches()
    if not batch_id:
        return next((b for b in batches if b.get("published", True)), batches[0] if batches else None)
    for b in batches:
        if b["id"].lower() == batch_id.lower():
            return b
    return next((b for b in batches if b.get("published", True)), batches[0] if batches else None)

def find_batch_exact(batch_id):
    # Exact id match only (no fallback): review endpoints fail closed on unknown/empty ids.
    if not batch_id:
        return None
    for b in load_batches():
        if b["id"].lower() == batch_id.lower():
            return b
    return None

def find_batch_by_path(batches, path):
    if not path or not batches:
        return None
    target_norm = os.path.normcase(os.path.normpath(path))
    for b in batches:
        b_path = b.get("path")
        if b_path and os.path.normcase(os.path.normpath(b_path)) == target_norm:
            return b
    return None

QA_DECISIONS_FILE = os.path.join(WORKSPACE_DIR, "qa_decisions.json")

def sync_qa_decisions(batch_id=None, batch_reviews=None):
    try:
        data = {}
        if os.path.exists(QA_DECISIONS_FILE):
            try:
                with open(QA_DECISIONS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                data = {}
        if batch_id and batch_reviews is not None:
            data[batch_id] = batch_reviews
        else:
            for fn in os.listdir(REVIEWS_DIR):
                if fn.startswith("reviews_") and fn.endswith(".json"):
                    b_id = fn[len("reviews_"):-len(".json")]
                    try:
                        with open(os.path.join(REVIEWS_DIR, fn), "r", encoding="utf-8") as rf:
                            data[b_id] = json.load(rf)
                    except Exception:
                        pass
        with open(QA_DECISIONS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error syncing qa_decisions.json: {e}")
        return False

def get_review_file(batch_id):
    safe_id = re.sub(r'[^a-zA-Z0-9_\-]', '_', batch_id)
    return os.path.join(REVIEWS_DIR, f"reviews_{safe_id}.json")

def load_reviews(batch_id):
    rf = get_review_file(batch_id)
    reviews_data = {}
    if os.path.exists(rf):
        try:
            with open(rf, "r", encoding="utf-8") as f:
                reviews_data = json.load(f)
        except Exception:
            reviews_data = {}
    else:
        legacy = os.path.join(REVIEWS_DIR, "hatshop_review_status.json")
        if batch_id == "26-08-28-hatshop-black-river" and os.path.exists(legacy):
            try:
                with open(legacy, "r", encoding="utf-8") as f:
                    reviews_data = json.load(f)
            except Exception:
                reviews_data = {}

    safe_bid = str(batch_id or "").lower()
    auto_resolved_batches = [
        "26-10-30-kanha-batumi",
        "26-10-29-merida-ik100102114",
        "26-10-09-charrone"
    ]
    if any(ab in safe_bid for ab in auto_resolved_batches) and ("-new-interiors" not in safe_bid) and reviews_data:
        for pdata in reviews_data.values():
            if isinstance(pdata, dict) and "comments" in pdata:
                for c in pdata["comments"]:
                    if "resolved" not in c:
                        c["resolved"] = True

    return reviews_data

def save_reviews(batch_id, data):
    rf = get_review_file(batch_id)
    try:
        with open(rf, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        sync_qa_decisions(batch_id, data)
        return True
    except Exception as e:
        print(f"Error saving reviews: {e}")
        return False

def _extract_batch_id(batch_or_id):
    if isinstance(batch_or_id, dict):
        return batch_or_id.get("id", "")
    return str(batch_or_id or "")

def get_ratings_file(batch_id):
    safe_id = re.sub(r'[^a-zA-Z0-9_\-]', '_', _extract_batch_id(batch_id))
    return os.path.join(REVIEWS_DIR, f"ratings_{safe_id}.json")

def load_ratings(batch_id):
    rf = get_ratings_file(batch_id)
    if os.path.exists(rf):
        try:
            with open(rf, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
        except Exception:
            return {}
    return {}

def get_ratings(batch_id, reviewer):
    if not reviewer or not isinstance(reviewer, str) or reviewer.strip() in ("", "Anonym"):
        return {}
    data = load_ratings(batch_id)
    rev_dict = data.get(reviewer.strip(), {})
    if isinstance(rev_dict, dict):
        return dict(rev_dict)
    return {}

def set_rating(batch_id, reviewer, product_id, filename, stars):
    if not reviewer or not isinstance(reviewer, str) or reviewer.strip() in ("", "Anonym"):
        raise ValueError("Skriv ditt namn i hubben först för att kunna sätta betyg.")
    if not product_id or not filename:
        raise ValueError("Saknar product_id eller filename")

    if isinstance(stars, bool):
        raise ValueError("Ogiltigt betyg")
    if isinstance(stars, float) and not stars.is_integer():
        raise ValueError("Ogiltigt betyg")
    try:
        stars_int = int(stars)
    except (ValueError, TypeError):
        raise ValueError("Ogiltigt betyg")

    if stars_int not in (0, 1, 2, 3, 4, 5):
        raise ValueError("Ogiltigt betyg")

    batch_id_str = _extract_batch_id(batch_id)
    rev_name = reviewer.strip()
    key = f"{product_id}|{filename}"

    with _RATINGS_LOCK:
        data = load_ratings(batch_id_str)
        if rev_name not in data or not isinstance(data[rev_name], dict):
            data[rev_name] = {}

        if stars_int == 0:
            data[rev_name].pop(key, None)
        else:
            data[rev_name][key] = {
                "stars": stars_int,
                "ts": datetime.datetime.now().isoformat()
            }

        rf = get_ratings_file(batch_id_str)
        os.makedirs(os.path.dirname(rf), exist_ok=True)
        tmp_file = rf + f".tmp.{uuid.uuid4().hex}"
        try:
            with open(tmp_file, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            os.replace(tmp_file, rf)
        except Exception:
            if os.path.exists(tmp_file):
                try:
                    os.remove(tmp_file)
                except OSError:
                    pass
            raise


def parse_file_details(filename, folder_path):
    """
    Parses a filename to extract base shot number, version suffix, modification time, etc.
    Examples:
      keps-model-color-01.jpg -> base_shot='01', suffix='', is_version_suffix=False
      keps-model-color-01-n.jpg -> base_shot='01', suffix='n', is_version_suffix=True
      matta-model-01-Interior-0002.jpg -> base_shot='01', suffix='Interior-0002', is_version_suffix=False
    """
    name_no_ext, ext = os.path.splitext(filename)
    full_path = os.path.join(folder_path, filename) if folder_path else ""
    try:
        st = os.stat(full_path) if full_path else None
        mtime, file_size = (st.st_mtime, st.st_size) if st else (0, 0)
    except OSError:
        mtime, file_size = 0, 0

    # Check if AI lifestyle or interior generated image (e.g. Replace_..., Editing_..., etc.)
    if re.match(r'^(replace|replacing|edit|editing|redecorat|modify|interior|lifestyle)', name_no_ext, re.IGNORECASE):
        return {
            "filename": filename,
            "base_shot": "99",
            "raw_suffix": "miljo",
            "clean_suffix": "Miljö",
            "is_version_suffix": False,
            "is_interior": True,
            "mtime": mtime,
            "size_bytes": file_size
        }

    m = re.search(r'[-_](0[1-9]|[1-9][0-9]?)(?=$|[-_ .()])(.*)$', name_no_ext, re.IGNORECASE)
    if m:
        base_shot_num = m.group(1).zfill(2)
        raw_suffix = m.group(2).strip('-_ ')
        if re.fullmatch(r'0\d{2}', raw_suffix):
            base_shot_num = f"{base_shot_num}-{raw_suffix}"
            raw_suffix = ""
    else:
        m3 = re.search(r'[-_](LS\d{1,3}[a-d]?|0\d{2}[a-d]?|\d{3}[a-d]?)(?=$|[-_ .()])(.*)$', name_no_ext, re.IGNORECASE)
        if m3:
            tok = m3.group(1)
            base_shot_num = tok.upper() if tok[:2].lower() == "ls" else tok.lower()
            raw_suffix = m3.group(2).strip('-_ ')
        else:
            base_shot_num = None
            raw_suffix = ""

    if base_shot_num is not None:
        clean_suffix = raw_suffix.replace('-', ' ').replace('_', ' ').strip()
        is_interior = bool(re.search(r'int[ei]{1,2}r?i?oi?r|miljo|miljö|detail', raw_suffix, re.IGNORECASE))
        is_version_suffix = bool(re.match(r'^(n|ny|new|v[0-9]+|[1-9][0-9]?|fix\d*|crop|justerad|redigerad)$', raw_suffix, re.IGNORECASE)) and not is_interior
        return {
            "filename": filename,
            "base_shot": base_shot_num,
            "raw_suffix": raw_suffix,
            "clean_suffix": clean_suffix,
            "is_version_suffix": is_version_suffix,
            "is_interior": is_interior,
            "mtime": mtime,
            "size_bytes": file_size
        }
    else:
        return {
            "filename": filename,
            "base_shot": "99",
            "raw_suffix": "",
            "clean_suffix": "",
            "is_version_suffix": False,
            "is_interior": False,
            "mtime": mtime,
            "size_bytes": file_size
        }


def _assign_comment_keys(shots):
    by_base = {}
    for s in shots:
        if not s.get("is_new_version"):
            by_base.setdefault(str(s.get("base_shot") or ""), []).append(s)
    for s in shots:
        base = str(s.get("base_shot") or "")
        s["comment_key"] = str(s["shot"]) if (not base or base == "99" or len(by_base.get(base, [])) > 1) else base
    return shots


_CATEGORY_WORDS = {"keps", "kepsar", "hatt", "hattar", "mössa", "mossa", "matta", "ullmatta", "viskosmatta",
                   "wiltonmatta", "barnmatta", "inomhusutomhusmatta", "inomhusutomhus", "rund", "inredning",
                   "gardin", "gardiner"}
_BRAND_ALIASES = (
    (re.compile(r'(?<![a-zåäö])black[\s_-]*river(?![a-zåäö])', re.I), "Black River"),
    (re.compile(r'(?<![a-zåäö])(?:g[\s_-]rda|g[aå]rda)(?![a-zåäö])', re.I), "Gårda"),
)
_COLOR_WORDS = {"svart", "vit", "offwhite", "off white", "grå", "mörkgrå", "ljusgrå", "blå", "ljusblå", "mörkblå", "navy",
                "khaki", "beige", "creme", "cream", "brun", "brown", "mörkbrun", "grön", "mörkgrön", "mossgrön", "ljusgrön",
                "olive", "oliv", "röd", "vinröd", "mörkröd", "red", "rosa", "lila", "gul", "ljusgul", "orange", "ljusorange",
                "terrakotta", "turkos", "natur", "natural", "multi", "petrol", "senap", "rost", "black", "white", "grey",
                "gray", "lightgrey", "darkgrey", "silver", "guld", "gold", "sand", "taupe"}
_TRAILING_FILLER = {"carpet", "teppich", "rug", "matta", "cap"}
_NBSP = "\u00a0"


def parse_product_name(folder_name):
    name = re.sub(r'-(?:1500px(?:\s*iphone)?|original|web|print)$', '', folder_name.strip(), flags=re.I)
    brand = ""
    for rx, canon in _BRAND_ALIASES:
        if rx.search(name):
            name = rx.sub(canon.replace(" ", _NBSP), name, count=1)
            brand = canon
            break
    parts = [p.strip() for p in re.split(r'\s*[-_]\s*', name) if p.strip()]
    category = ""
    while parts and parts[0].lower() in _CATEGORY_WORDS:
        category = (category + " " + parts.pop(0)).strip()
    if parts and brand and parts[0].replace(_NBSP, " ") == brand:
        parts.pop(0)
    while len(parts) > 1 and (parts[-1].lower() in _TRAILING_FILLER
                              or re.fullmatch(r'(?:new|ny|fix\d*|v\d+)', parts[-1], re.I)):
        parts.pop()
    colors = []
    while len(parts) > 1 and parts[-1].lower() in _COLOR_WORDS:
        colors.insert(0, parts.pop())
    if not colors and len(parts) >= 2:
        colors = [parts.pop()]
    model = " ".join(parts).replace(_NBSP, " ")
    return {"category": category.capitalize(), "brand": brand, "model": model,
            "color": "/".join(c.capitalize() for c in colors),
            "display_name": " ".join(x for x in (brand, model) if x)}


def _split_folder_parts(clean_name, default_brand=True):
    """HEAD folder splitting with Black River multi-word brand awareness."""
    if re.search(r'(?<![a-zåäö])black[\s_-]+river(?![a-zåäö])', clean_name, re.IGNORECASE):
        meta = parse_product_name(clean_name)
        return (
            meta["brand"] or "Black River",
            meta["display_name"] or meta["model"] or clean_name,
            meta["color"]
        )
    parts = clean_name.split("-")
    brand = "Brand"
    model_name = clean_name
    color_name = ""
    if default_brand:
        if len(parts) >= 4 and parts[0].strip().lower() in ["keps", "matta", "ullmatta", "viskosmatta", "wiltonmatta", "inredning"]:
            brand = parts[0].strip()
            model_name = parts[1].strip()
            color_name = "-".join(parts[2:]).strip()
        elif len(parts) >= 3:
            brand = parts[0].strip()
            model_name = parts[1].strip()
            color_name = "-".join(parts[2:]).strip()
        elif len(parts) == 2:
            model_name = parts[0].strip()
            color_name = parts[1].strip()
    else:
        if len(parts) >= 3:
            model_name = parts[1].strip()
            color_name = "-".join(parts[2:]).strip()
        elif len(parts) == 2:
            model_name = parts[0].strip()
            color_name = parts[1].strip()
    return brand, model_name, color_name


def parse_upload_order_file(txt_path):
    order = []
    if not os.path.exists(txt_path):
        return order
    try:
        with open(txt_path, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                line = line.strip()
                if '|' in line and not line.startswith('ORDNING') and not line.startswith('---'):
                    parts = [p.strip() for p in line.split('|')]
                    if len(parts) >= 2:
                        seq = parts[0].zfill(2)
                        new_fn = parts[1]
                        orig_fn = parts[2] if len(parts) > 2 else ''
                        typ = parts[3] if len(parts) > 3 else ''
                        order.append({
                            'seq': seq,
                            'filename': new_fn,
                            'orig_filename': orig_fn,
                            'type': typ
                        })
    except Exception as e:
        print(f"Error reading upload_order.txt {txt_path}: {e}")
    return order

def process_upload_order_shots(upload_order, raw_files, batch_id, subdir, folder_name, folder_path):
    final_shots = []
    seen_files = set()

    for item in upload_order:
        fn = item["filename"]
        full_fn = os.path.join(folder_path, fn)
        if not os.path.exists(full_fn):
            match = next((f for f in raw_files if f.lower() == fn.lower()), None)
            if match:
                fn = match
                full_fn = os.path.join(folder_path, fn)
            else:
                continue

        seen_files.add(fn.lower())
        mtime = os.path.getmtime(full_fn) if os.path.exists(full_fn) else 0
        file_size = os.path.getsize(full_fn) if os.path.exists(full_fn) else 0
        rel_path = f"{subdir}/{folder_name}/{fn}" if subdir else f"{folder_name}/{fn}"
        raw_url = f"/image/{urllib.parse.quote(batch_id)}/{urllib.parse.quote(rel_path)}?v={int(mtime)}"

        typ_upper = item["type"].upper()
        is_bubble = "BUBBLA" in typ_upper
        is_gallery = "GALLERI" in typ_upper

        if is_bubble:
            label = f"{item['seq']} (Bubbla)"
            v_tag = "Bubbla"
            is_new = True
        elif is_gallery:
            if "interior" in fn.lower() or "miljo" in fn.lower():
                label = f"{item['seq']} (Miljö)"
            elif "-01" in fn.lower():
                label = f"{item['seq']} (Plan)"
            else:
                label = f"{item['seq']}"
            v_tag = ""
            is_new = False
        else:
            label = f"{item['seq']}"
            v_tag = ""
            is_new = False

        final_shots.append({
            "filename": fn,
            "orig_name": item.get("orig_filename", ""),
            "shot": item["seq"],
            "base_shot": item["seq"],
            "label": label,
            "is_primary": (item["seq"] == "01"),
            "is_new_version": is_new,
            "version_tag": v_tag,
            "mtime": mtime,
            "url": raw_url,
            "thumb_url": f"{raw_url}&thumb=160",
            "rel_path": rel_path,
            "size_bytes": file_size
        })

    # Remaining files if any
    for f in raw_files:
        if f.lower() not in seen_files:
            full_fn = os.path.join(folder_path, f)
            mtime = os.path.getmtime(full_fn) if os.path.exists(full_fn) else 0
            file_size = os.path.getsize(full_fn) if os.path.exists(full_fn) else 0
            rel_path = f"{subdir}/{folder_name}/{f}" if subdir else f"{folder_name}/{f}"
            raw_url = f"/image/{urllib.parse.quote(batch_id)}/{urllib.parse.quote(rel_path)}?v={int(mtime)}"
            final_shots.append({
                "filename": f,
                "shot": "99",
                "base_shot": "99",
                "label": f,
                "is_primary": False,
                "is_new_version": False,
                "version_tag": "",
                "mtime": mtime,
                "url": raw_url,
                "thumb_url": f"{raw_url}&thumb=160",
                "rel_path": rel_path,
                "size_bytes": file_size
            })

    return _assign_comment_keys(final_shots)

import ctypes
from ctypes import wintypes
from functools import cmp_to_key

try:
    _strcmpi = ctypes.windll.shlwapi.StrCmpLogicalW
    _strcmpi.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR]
    _strcmpi.restype = ctypes.c_int
    natural_folder_sort_key = cmp_to_key(_strcmpi)
except Exception:
    def natural_folder_sort_key(s):
        base, ext = os.path.splitext(s)
        tokens = re.split(r'(\d+)', base)
        return [(int(t), 0) if t.isdigit() else (0, t.lower()) for t in tokens] + [(0, ext.lower())]

def process_and_sort_variant_shots(raw_files, batch_id, subdir, folder_name, folder_path):
    """
    Processes all raw image files for a product variant.
    Preserves the exact folder sequence as organized on disk (natural Windows Explorer order).
    If upload_order.txt exists, respects the explicit studio sequence and types.
    """
    upload_order_path = os.path.join(folder_path, "upload_order.txt")
    if os.path.exists(upload_order_path):
        parsed_order = parse_upload_order_file(upload_order_path)
        if parsed_order:
            return process_upload_order_shots(parsed_order, raw_files, batch_id, subdir, folder_name, folder_path)

    # Sort raw_files in exact natural folder order (Windows Explorer order)
    sorted_files = sorted(raw_files, key=natural_folder_sort_key)

    final_shots = []
    for idx, filename in enumerate(sorted_files):
        item = parse_file_details(filename, folder_path)
        rel_path = f"{subdir}/{folder_name}/{filename}" if subdir else f"{folder_name}/{filename}"
        raw_url = f"/image/{urllib.parse.quote(batch_id)}/{urllib.parse.quote(rel_path)}?v={int(item['mtime'])}"
        
        base_shot = item["base_shot"]
        raw_suffix = item["raw_suffix"]
        is_int = item.get("is_interior", False)
        is_version = item.get("is_version_suffix", False)

        if is_int:
            sub = re.sub(r'int[ei]{1,2}r?i?oi?r.*$', '', raw_suffix, flags=re.IGNORECASE).strip('-_ ')
            if sub and base_shot != "99":
                label = f"{base_shot}-{sub} (Miljö)"
                shot_code = f"{base_shot}_{sub}_miljo"
            elif base_shot != "99":
                label = f"{base_shot} (Miljö)"
                shot_code = f"{base_shot}_miljo"
            else:
                label = "Miljö"
                shot_code = f"miljo_{os.path.splitext(filename)[0]}"
            is_prim = True
            is_new = False
            v_tag = ""
        elif is_version:
            label = f"{base_shot} (Ny -{raw_suffix})"
            shot_code = f"{base_shot}_{raw_suffix}"
            is_prim = True
            is_new = True
            v_tag = f"-{raw_suffix}"
        elif raw_suffix:
            label = f"{base_shot} ({item['clean_suffix']})"
            shot_code = f"{base_shot}_{raw_suffix}"
            is_prim = True
            is_new = False
            v_tag = ""
        else:
            label = f"{base_shot}"
            shot_code = base_shot
            is_prim = True
            is_new = False
            v_tag = ""

        final_shots.append({
            "filename": filename,
            "shot": shot_code,
            "base_shot": base_shot,
            "label": label,
            "is_primary": is_prim,
            "is_new_version": is_new,
            "version_tag": v_tag,
            "has_previous_version": False,
            "mtime": item["mtime"],
            "url": raw_url,
            "thumb_url": f"{raw_url}&thumb=160",
            "rel_path": rel_path,
            "size_bytes": item["size_bytes"]
        })

    if final_shots and not any(s.get("is_primary") for s in final_shots):
        final_shots[0]["is_primary"] = True
    return _assign_comment_keys(final_shots)

def parse_color_folder(color_folder):
    parts = color_folder.split("-")
    color_name = parts[1].strip() if len(parts) > 1 else ""
    brand_name = parts[2].strip().capitalize() if len(parts) > 2 else "Gardin"

    color_display = color_name.capitalize()
    c_low = color_name.lower()
    if c_low == "glu":
        color_display = "Gul"
    elif c_low == "vit":
        color_display = "Vit"
    elif c_low == "cream":
        color_display = "Cream"
    elif c_low == "dark brown":
        color_display = "Mörkbrun"
    elif c_low == "green":
        color_display = "Grön"
    elif c_low == "lila":
        color_display = "Lila"
    elif "ljusbl" in c_low:
        color_display = "Ljusblå"
    elif c_low == "beige":
        color_display = "Beige"
    elif "bl" in c_low and len(c_low) <= 4:
        color_display = "Blå"
    elif "terrakotta" in c_low:
        color_display = "Terrakotta-rosa"
    elif "turkos" in c_low:
        color_display = "Turkos"

    return color_display, brand_name

def _load_product_image_mapping(base_dir):
    """Load optional product_image_mapping.json from base_dir or parent directory, indexed by subfolder/id/slot."""
    candidates = [
        os.path.join(base_dir, "product_image_mapping.json"),
        os.path.join(os.path.dirname(base_dir), "product_image_mapping.json"),
    ]
    map_file = next((f for f in candidates if os.path.isfile(f)), None)
    if not map_file:
        return {}
    try:
        with open(map_file, "r", encoding="utf-8") as mf:
            data = json.load(mf)
        entries = data.get("images", []) if isinstance(data, dict) else (data if isinstance(data, list) else [])
        lookup = {}
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            for k in (
                entry.get("subfolder"),
                entry.get("id"),
                os.path.splitext(str(entry.get("image_file") or ""))[0],
                f"{int(entry['slot']):02d}" if isinstance(entry.get("slot"), int) else None,
            ):
                if k:
                    lookup[str(k).strip().lower()] = entry
        return lookup
    except Exception:
        return {}


def _infer_look_from_shots_and_subtitle(prod):
    """Fallback for Reforma manifest items where @NNNN tags live in shot_kind and clean_name/color."""
    subtitle = str(prod.get("color") or "")
    name_by_tag = {}
    for seg in re.split(r'\s*\+\s*', subtitle):
        m_seg = re.search(r'^(?:\d+\.\s*)?(.+?)\s*\((@\d+)', seg.strip())
        if m_seg:
            name_by_tag[m_seg.group(2)] = m_seg.group(1).strip()

    all_shots = []
    for vshots in (prod.get("variants") or {}).values():
        if isinstance(vshots, list):
            all_shots.extend(vshots)

    hero = None
    companions = []
    seen_tags = set()
    for s in all_shots:
        sk = str(s.get("shot_kind") or s.get("label") or "")
        m_tag = re.search(r'(@\d+)', sk)
        if not m_tag:
            continue
        tag = m_tag.group(1)
        if tag in seen_tags:
            continue
        seen_tags.add(tag)
        num_id = tag.lstrip("@")
        pname = name_by_tag.get(tag) or tag
        entry = {
            "id": num_id,
            "tag": tag,
            "name": pname,
            "url": f"https://www.reformasthlm.se/sv/search?q={urllib.parse.quote(pname if pname != tag else num_id)}",
        }
        sk_low = sk.lower()
        if hero is None and ("stjärn" in sk_low or "hero" in sk_low or not companions):
            hero = entry
        else:
            companions.append(entry)
    return hero, companions


def _apply_product_mapping_enrichment(products, mapping_lookup):
    for prod in products:
        if "hero_product" not in prod:
            prod["hero_product"] = None
        if "companion_products" not in prod or not isinstance(prod["companion_products"], list):
            prod["companion_products"] = []
        if mapping_lookup and not prod.get("hero_product"):
            pid_low = str(prod.get("id") or "").strip().lower()
            fn_low = str(prod.get("folder_name") or "").strip().lower()
            prefix_slot = pid_low.split("_", 1)[0] if "_" in pid_low else ""
            matched = (
                mapping_lookup.get(pid_low)
                or mapping_lookup.get(fn_low)
                or (mapping_lookup.get(prefix_slot) if prefix_slot.isdigit() else None)
            )
            if matched:
                prod["hero_product"] = matched.get("hero_product")
                prod["companion_products"] = matched.get("companion_products", []) or []
        for vshots in (prod.get("variants") or {}).values():
            if isinstance(vshots, list):
                for s in vshots:
                    if "shot_kind" not in s:
                        s["shot_kind"] = ""
        if not prod.get("hero_product"):
            inf_hero, inf_comp = _infer_look_from_shots_and_subtitle(prod)
            if inf_hero:
                prod["hero_product"] = inf_hero
                if not prod["companion_products"]:
                    prod["companion_products"] = inf_comp


def scan_batch_products(batch):
    base_dir = batch.get("path", "")
    batch_id = batch.get("id", "")
    if not os.path.exists(base_dir):
        return []

    mapping_lookup = _load_product_image_mapping(base_dir)

    # Check for products_manifest.json in base_dir or parent directory
    manifest_candidates = [
        os.path.join(base_dir, "products_manifest.json"),
        os.path.join(os.path.dirname(base_dir), "products_manifest.json"),
        os.path.join(base_dir, "manifest.json"),
        os.path.join(os.path.dirname(base_dir), "manifest.json"),
    ]
    manifest_file = next((f for f in manifest_candidates if os.path.isfile(f)), None)

    if manifest_file:
        try:
            with open(manifest_file, "r", encoding="utf-8") as mf:
                manifest_data = json.load(mf)
            if isinstance(manifest_data, list):
                products = []
                for item in manifest_data:
                    pid = str(item.get("id", ""))
                    title = item.get("title", pid)
                    clean_name = item.get("clean_name", pid)
                    category = item.get("category", "Möbler")

                    shots = []
                    for img in item.get("images", []):
                        fn = img.get("filename", "")
                        if os.path.exists(os.path.join(base_dir, fn)):
                            rel_file = fn
                        elif "rel_path" in img and os.path.exists(os.path.join(base_dir, img["rel_path"])):
                            rel_file = img["rel_path"]
                        elif "rel_path" in img and os.path.exists(os.path.dirname(base_dir) and os.path.join(os.path.dirname(base_dir), img["rel_path"])):
                            base_name = os.path.basename(img["rel_path"])
                            rel_file = base_name if os.path.exists(os.path.join(base_dir, base_name)) else fn
                        else:
                            rel_file = fn

                        full_path = os.path.join(base_dir, rel_file)
                        if not os.path.exists(full_path):
                            continue
                        mtime = os.path.getmtime(full_path)
                        kind = img.get("kind", "")
                        pos = len(shots) + 1
                        label = f"{pos:02d}" + (f" ({kind})" if kind else "")

                        shots.append({
                            "shot": f"{pos:02d}",
                            "base_shot": f"{pos:02d}",
                            "comment_key": f"{pos:02d}",
                            "label": label,
                            "shot_kind": kind,
                            "url": f"/image/{urllib.parse.quote(batch_id)}/{urllib.parse.quote(rel_file)}?v={int(mtime)}",
                            "filename": fn,
                            "orig_name": img.get("orig", ""),
                            "rel_path": rel_file,
                            "is_primary": (pos == 1),
                            "is_new_version": False,
                            "version_tag": "",
                            "mtime": mtime
                        })

                    # CEO 2026-11-10: every file moved out ("🗑 Flytta ut") -> no empty card, like folder batches.
                    # The manifest entry stays, so "Ångra" brings the product back.
                    if not shots:
                        continue

                    v_label = "1600x2000 Transparent" if any(img.get("is_transparent") for img in item.get("images", [])) else "Standard"
                    model_str = title if str(pid) == str(title) else f"[{pid}] {title}"
                    products.append({
                        "id": pid,
                        "brand": category,
                        "category": category,
                        "model": model_str,
                        "display_name": model_str,
                        "color": clean_name,
                        "folder_name": pid,
                        "hero_product": item.get("hero_product"),
                        "companion_products": item.get("companion_products", []) or [],
                        "has_new_version": False,
                        "version_tag": "",
                        "latest_mtime": max((s["mtime"] for s in shots), default=0),
                        "variants": {
                            v_label: shots
                        }
                    })
                _apply_product_mapping_enrichment(products, mapping_lookup)
                annotate_product_gold_set(batch, products)
                return products
        except Exception as e:
            print(f"[!] Error parsing manifest {manifest_file}: {e}")

    root_entries = [d for d in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, d))]
    has_variant_subdirs = any(v.lower() in ["1500px", "1500px iphone", "original", "web", "print"] for v in root_entries)
    prefix = ""
    variant_subdirs = root_entries
    if not has_variant_subdirs and len(root_entries) == 1:
        direct_images = [f for f in os.listdir(base_dir) if f.lower().endswith(MEDIA_EXTS)]
        if not direct_images:
            cand_name = root_entries[0]
            cand_path = os.path.join(base_dir, cand_name)
            cand_subdirs = [d for d in os.listdir(cand_path) if os.path.isdir(os.path.join(cand_path, d))]
            if any(v.lower() in ["1500px", "1500px iphone", "original", "web", "print"] for v in cand_subdirs):
                prefix = cand_name
                variant_subdirs = cand_subdirs
                has_variant_subdirs = True

    products_map = {}

    if has_variant_subdirs:
        for subdir in variant_subdirs:
            subdir_path = os.path.join(base_dir, prefix, subdir) if prefix else os.path.join(base_dir, subdir)
            if not os.path.isdir(subdir_path):
                continue
            
            variant_segment = f"{prefix}/{subdir}" if prefix else subdir

            for folder_name in sorted(os.listdir(subdir_path)):
                folder_path = os.path.join(subdir_path, folder_name)
                if not os.path.isdir(folder_path):
                    continue

                raw_files = [f for f in os.listdir(folder_path) if f.lower().endswith(MEDIA_EXTS)]

                if raw_files:
                    clean_name = re.sub(r'-(?:1500px(?:\s*iphone)?|original|web|print)$', '', folder_name, flags=re.IGNORECASE).strip()
                    if clean_name.lower() == 'frome':
                        clean_name = 'Frome-red'

                    brand, model_name, color_name = _split_folder_parts(clean_name, default_brand=True)

                    prod_key = clean_name
                    if prod_key not in products_map:
                        products_map[prod_key] = {
                            "id": prod_key,
                            "brand": brand,
                            "model": model_name,
                            "display_name": model_name,
                            "color": color_name,
                            "folder_name": folder_name,
                            "has_new_version": False,
                            "version_tag": "",
                            "latest_mtime": 0,
                            "variants": {}
                        }

                    shot_list = process_and_sort_variant_shots(raw_files, batch_id, variant_segment, folder_name, folder_path)

                    products_map[prod_key]["variants"][subdir] = shot_list
                    if any(s.get("is_new_version") for s in shot_list):
                        products_map[prod_key]["has_new_version"] = True
                        tag = next((s["version_tag"] for s in shot_list if s.get("is_new_version")), "")
                        if tag:
                            products_map[prod_key]["version_tag"] = tag
                    
                    mtimes = [s["mtime"] for s in shot_list if s.get("mtime")]
                    if mtimes:
                        products_map[prod_key]["latest_mtime"] = max(products_map[prod_key]["latest_mtime"], max(mtimes))
                else:
                    # Model folder containing color subfolders
                    sub_dirs = [d for d in sorted(os.listdir(folder_path)) if os.path.isdir(os.path.join(folder_path, d))]
                    hidden_subdirs = hidden_crop_dirs(sub_dirs)
                    color_subfolders = []
                    for sd in sub_dirs:
                        if sd in hidden_subdirs:
                            continue
                        sd_path = os.path.join(folder_path, sd)
                        sd_raw = [f for f in os.listdir(sd_path) if f.lower().endswith(MEDIA_EXTS)]
                        if sd_raw:
                            color_subfolders.append((sd, sd_path, sd_raw))

                    if color_subfolders:
                        for color_folder, color_path, raw in color_subfolders:
                            if color_folder in hidden_subdirs:
                                continue
                            prod_id = f"{folder_name}-{color_folder}"
                            color, brand = parse_color_folder(color_folder)
                            if prod_id not in products_map:
                                products_map[prod_id] = {
                                    "id": prod_id,
                                    "brand": brand,
                                    "model": folder_name,
                                    "display_name": folder_name,
                                    "color": color,
                                    "folder_name": f"{folder_name}/{color_folder}",
                                    "has_new_version": False,
                                    "version_tag": "",
                                    "latest_mtime": 0,
                                    "variants": {}
                                }

                            shot_segment = f"{variant_segment}/{folder_name}"
                            shot_list = process_and_sort_variant_shots(raw, batch_id, shot_segment, color_folder, color_path)

                            products_map[prod_id]["variants"][subdir] = shot_list
                            if any(s.get("is_new_version") for s in shot_list):
                                products_map[prod_id]["has_new_version"] = True
                                tag = next((s["version_tag"] for s in shot_list if s.get("is_new_version")), "")
                                if tag:
                                    products_map[prod_id]["version_tag"] = tag

                            mtimes = [s["mtime"] for s in shot_list if s.get("mtime")]
                            if mtimes:
                                products_map[prod_id]["latest_mtime"] = max(products_map[prod_id]["latest_mtime"], max(mtimes))
    else:
        # Check if root entries have nested model/color directories (e.g. Lumira/Lumia-Cream-...)
        is_nested_model_batch = False
        for folder_name in root_entries:
            fp = os.path.join(base_dir, folder_name)
            subdirs = [d for d in os.listdir(fp) if os.path.isdir(os.path.join(fp, d))]
            direct_images = [f for f in os.listdir(fp) if f.lower().endswith(MEDIA_EXTS)]
            if subdirs and not direct_images:
                is_nested_model_batch = True
                break

        ran_subfolder_branch = False
        if is_nested_model_batch:
            for model_name in sorted(root_entries):
                model_path = os.path.join(base_dir, model_name)
                if not os.path.isdir(model_path):
                    continue
                color_dirs = [d for d in sorted(os.listdir(model_path)) if os.path.isdir(os.path.join(model_path, d))]
                hidden_dirs = hidden_crop_dirs(color_dirs)
                for color_folder in color_dirs:
                    if color_folder in hidden_dirs:
                        continue
                    folder_path = os.path.join(model_path, color_folder)
                    raw_files = [f for f in os.listdir(folder_path) if f.lower().endswith(MEDIA_EXTS)]
                    if not raw_files:
                        continue

                    color_display, brand_name = parse_color_folder(color_folder)
                    shot_list = process_and_sort_variant_shots(raw_files, batch_id, model_name, color_folder, folder_path)
                    has_new = any(s.get("is_new_version") for s in shot_list)
                    v_tag = next((s["version_tag"] for s in shot_list if s.get("is_new_version")), "")
                    mtimes = [s["mtime"] for s in shot_list if s.get("mtime")]

                    prod_id = f"{model_name}-{color_folder}"
                    products_map[prod_id] = {
                        "id": prod_id,
                        "brand": brand_name,
                        "model": model_name,
                        "display_name": model_name,
                        "color": color_display,
                        "folder_name": f"{model_name}/{color_folder}",
                        "has_new_version": has_new,
                        "version_tag": v_tag,
                        "latest_mtime": max(mtimes) if mtimes else 0,
                        "variants": { "Standard": shot_list }
                    }
        elif root_entries and any(os.path.isdir(os.path.join(base_dir, d)) and any(f.lower().endswith(MEDIA_EXTS) for f in os.listdir(os.path.join(base_dir, d))) for d in root_entries):
            ran_subfolder_branch = True
            for folder_name in sorted(root_entries):
                folder_path = os.path.join(base_dir, folder_name)
                if not os.path.isdir(folder_path):
                    continue
                brand, model_name, color_name = _split_folder_parts(folder_name, default_brand=False)

                raw_files = [f for f in os.listdir(folder_path) if f.lower().endswith(MEDIA_EXTS)]
                shot_list = process_and_sort_variant_shots(raw_files, batch_id, "", folder_name, folder_path)

                has_new = any(s.get("is_new_version") for s in shot_list)
                v_tag = next((s["version_tag"] for s in shot_list if s.get("is_new_version")), "")
                mtimes = [s["mtime"] for s in shot_list if s.get("mtime")]

                products_map[folder_name] = {
                    "id": folder_name,
                    "brand": brand,
                    "model": model_name,
                    "display_name": model_name,
                    "color": color_name,
                    "folder_name": folder_name,
                    "has_new_version": has_new,
                    "version_tag": v_tag,
                    "latest_mtime": max(mtimes) if mtimes else 0,
                    "variants": { "Standard": shot_list }
                }
        direct_root_images = [f for f in sorted(os.listdir(base_dir)) if f.lower().endswith(MEDIA_EXTS) and os.path.isfile(os.path.join(base_dir, f))]
        if (not is_nested_model_batch and not ran_subfolder_branch) or (ran_subfolder_branch and direct_root_images):
            # Flat directory layout (direct image files in base_dir without subfolders, or loose root images alongside subfolders)
            direct_images = direct_root_images

            def parse_flat_item(fn):
                name_no_ext, ext = os.path.splitext(fn)
                full_path = os.path.join(base_dir, fn)
                mtime = os.path.getmtime(full_path) if os.path.exists(full_path) else 0
                size_bytes = os.path.getsize(full_path) if os.path.exists(full_path) else 0

                # Check specific descriptors like close-up, street, interior, etc. followed by optional num/letter
                m1 = re.match(r'^(.*?)[-_](close[-_]?u[p]?|clo|street(?:[-_]s|[-_]st)?|str|interior|miljo|plan|studio|detail|angle|lifestyle)[-_]?(?:([0-9]+)([a-z0-9]*))?$', name_no_ext, re.IGNORECASE)
                if m1:
                    prod_key = m1.group(1).strip('-_')
                    shot_desc = m1.group(2).lower()
                    shot_num = m1.group(3) or ''
                    shot_suffix = m1.group(4) or ''

                    is_close = any(x in shot_desc for x in ['close', 'clo', 'plan', 'studio', 'detail'])
                    is_street = any(x in shot_desc for x in ['street', 'str', 'interior', 'miljo', 'lifestyle'])

                    if is_close:
                        kind = 'Närbild'
                        base_order = 10
                    elif is_street:
                        kind = 'Miljö'
                        base_order = 20
                    else:
                        kind = shot_desc.capitalize()
                        base_order = 30

                    label_parts = [kind]
                    if shot_num:
                        label_parts.append(shot_num)
                    if shot_suffix:
                        label_parts.append(f'({shot_suffix})')
                    label = ' '.join(label_parts)
                    shot_code = f'{shot_num}{shot_suffix}' if shot_num else f'{kind[:1]}{shot_suffix}'

                    return {
                        'prod_key': prod_key,
                        'filename': fn,
                        'kind': kind,
                        'shot_code': shot_code,
                        'base_shot': shot_num or '01',
                        'suffix': shot_suffix,
                        'label': label,
                        'is_close': is_close,
                        'is_street': is_street,
                        'order_key': (base_order, int(shot_num) if shot_num.isdigit() else 99, shot_suffix),
                        'mtime': mtime,
                        'size_bytes': size_bytes
                    }

                # Generic format: prod_key - [0-9]+ [suffix] (with 3-digit / LS fallback)
                m2 = re.match(r'^(.*?)[-_](0[1-9]|[1-9][0-9]?)(?=$|[-_ .()])(.*)$', name_no_ext, re.IGNORECASE)
                if m2:
                    prod_key = m2.group(1).strip('-_')
                    shot_num = m2.group(2).zfill(2)
                    raw_suffix = m2.group(3).strip('-_ ')
                else:
                    m3 = re.match(r'^(.*?)[-_](LS\d{1,3}[a-d]?|0\d{2}[a-d]?|\d{3}[a-d]?)(?=$|[-_ .()])(.*)$', name_no_ext, re.IGNORECASE)
                    if m3:
                        prod_key = m3.group(1).strip('-_')
                        tok = m3.group(2)
                        shot_num = tok.upper() if tok[:2].lower() == "ls" else tok.lower()
                        raw_suffix = m3.group(3).strip('-_ ')
                    else:
                        prod_key = None
                        shot_num = None
                        raw_suffix = ""

                if prod_key and shot_num:
                    is_version = bool(re.match(r'^(n|ny|new|v[0-9]+|[0-9]+|b|fix\d*|crop)$', raw_suffix, re.IGNORECASE))
                    label = f'{shot_num} ({raw_suffix})' if raw_suffix else shot_num
                    digits = re.sub(r'\D', '', shot_num)
                    order_int = int(digits) if digits else 99

                    return {
                        'prod_key': prod_key,
                        'filename': fn,
                        'kind': 'Standard',
                        'shot_code': f'{shot_num}{raw_suffix}',
                        'base_shot': shot_num,
                        'suffix': raw_suffix,
                        'label': label,
                        'is_close': (shot_num in ('01', '001', 'LS01')),
                        'is_street': False,
                        'order_key': (order_int, 0 if not is_version else 1, raw_suffix),
                        'mtime': mtime,
                        'size_bytes': size_bytes
                    }

                return {
                    'prod_key': name_no_ext,
                    'filename': fn,
                    'kind': 'Övrigt',
                    'shot_code': '99',
                    'base_shot': '99',
                    'suffix': '',
                    'label': name_no_ext,
                    'is_close': False,
                    'is_street': False,
                    'order_key': (99, 99, ''),
                    'mtime': mtime,
                    'size_bytes': size_bytes
                }

            flat_parsed = [parse_flat_item(f) for f in direct_images]
            flat_groups = {}
            for item in flat_parsed:
                flat_groups.setdefault(item['prod_key'], []).append(item)

            for prod_key, items in flat_groups.items():
                items.sort(key=lambda x: x['order_key'])

                # Parse Brand, Model, Color from prod_key
                clean_key = re.sub(r'[-_]cap$', '', prod_key, flags=re.IGNORECASE)
                if re.search(r'(?<![a-zåäö])black[\s_-]+river(?![a-zåäö])', clean_key, re.IGNORECASE):
                    meta = parse_product_name(clean_key)
                    brand = meta["brand"] or "Black River"
                    model_name = meta["display_name"] or meta["model"] or clean_key
                    color_name = meta["color"]
                else:
                    parts = clean_key.split('-')
                    brand = 'Brand'
                    model_name = clean_key
                    color_name = ''

                    if len(parts) >= 3 and parts[0].lower() == 'g' and parts[1].lower() == 'rda':
                        brand = 'G-RDA'
                        model_name = parts[2].capitalize()
                        color_name = ' '.join(parts[3:]).capitalize()
                    elif len(parts) >= 4 and parts[0].strip().lower() in ["keps", "matta", "ullmatta", "viskosmatta", "wiltonmatta", "inredning"]:
                        brand = parts[0].capitalize()
                        model_name = parts[1].capitalize()
                        color_name = ' '.join(parts[2:]).capitalize()
                    elif len(parts) >= 3:
                        brand = parts[0].capitalize()
                        model_name = parts[1].capitalize()
                        color_name = ' '.join(parts[2:]).capitalize()
                    elif len(parts) == 2:
                        model_name = parts[0].capitalize()
                        color_name = parts[1].capitalize()

                shot_list = []
                for idx, it in enumerate(items):
                    fn = it['filename']
                    raw_url = f"/image/{urllib.parse.quote(batch_id)}/{urllib.parse.quote(fn)}?v={int(it['mtime'])}"
                    is_hero = (idx == 0)
                    shot_list.append({
                        "filename": fn,
                        "shot": it['shot_code'],
                        "base_shot": it['base_shot'],
                        "label": it['label'],
                        "is_primary": is_hero,
                        "is_new_version": bool(it['suffix']),
                        "version_tag": f"-{it['suffix']}" if it['suffix'] else "",
                        "mtime": it['mtime'],
                        "url": raw_url,
                        "thumb_url": f"{raw_url}&thumb=160",
                        "rel_path": fn,
                        "size_bytes": it['size_bytes']
                    })
                _assign_comment_keys(shot_list)

                mtimes = [s["mtime"] for s in shot_list if s.get("mtime")]
                has_new = any(s.get("is_new_version") for s in shot_list)
                v_tag = next((s["version_tag"] for s in shot_list if s.get("is_new_version")), "")

                if prod_key in products_map:
                    prod_key = f"rot-{prod_key}"

                products_map[prod_key] = {
                    "id": prod_key,
                    "brand": brand,
                    "model": model_name,
                    "display_name": model_name,
                    "color": color_name,
                    "folder_name": prod_key,
                    "has_new_version": has_new,
                    "version_tag": v_tag,
                    "latest_mtime": max(mtimes) if mtimes else 0,
                    "variants": { "Standard": shot_list }
                }

    products = list(products_map.values())
    if products and all(re.match(r'^\d{2}-', str(p.get("folder_name", ""))) for p in products):
        products.sort(key=lambda x: str(x.get("folder_name", "")))
    else:
        products.sort(key=lambda x: x["model"])
    annotate_product_gold_set(batch, products)
    return products


SITE_MIN_RATIO, SITE_MAX_RATIO = 0.95, 1.05
# OneDrive Files-On-Demand: opening a cloud-only file hydrates (downloads) it.
_CLOUD_ONLY_ATTRS = 0x00400000 | 0x00040000 | 0x00001000  # RECALL_ON_DATA_ACCESS | RECALL_ON_OPEN | OFFLINE
_KNOWN_RATIOS = (("1:1", 1.0), ("4:5", 0.8), ("3:4", 0.75), ("2:3", 2 / 3), ("9:16", 9 / 16),
                 ("5:4", 1.25), ("4:3", 4 / 3), ("3:2", 1.5), ("16:9", 16 / 9))
_DIMS_CACHE = OrderedDict()
_DIMS_LOCK = threading.Lock()
_VISUAL_SIG_CACHE = OrderedDict()
_VISUAL_SIG_LOCK = threading.Lock()
_NAME_MAP_CACHE = {}

_AI_TOKENS = {"interior", "inteiror", "interioir", "miljo", "miljö", "lifestyle", "vignette", "parlour", "loggia",
              "salon", "bedroom", "livingroom", "living", "dining", "hallway", "adtest", "editorial", "magnific"}
_AI_NAME_RE = re.compile(
    r'^(?:tc|rf|reforma)[-_]'
    r'|^(?:replac(?:e|ing)|edit(?:ing)?|redecorat\w*|modify\w*)(?:[-_ ]|$)'
    r'|(?:^|[-_])(?:caps-\d{2}[a-z]?|army-adtest)'
    r'|(?:^|[-_ ])ls\d{1,2}(?:[-_ ]|$)', re.I)                     # LS01-LS99 = lifestyle
_AI_SCENE_RE = re.compile(r'[_-]\d{2}[_-]\d{2,3}(?:-\d{3})?[a-z]?$', re.I)   # _01_16-001, _03_001
_AI_BATCH_HINTS = ("flow-vibes", "flow_vibes", "reforma-flow", "reforma_flow", "charonne-aranga", "ai_batch", "test-gold")


def _is_cloud_only(full_path):
    try:
        return bool(getattr(os.stat(full_path), "st_file_attributes", 0) & _CLOUD_ONLY_ATTRS)
    except OSError:
        return False


# C gate (Koordinator 4.2, 2026-11-10): every batch source lives in the CEO's OneDrive Pictures (= the backup) and is
# pinned ("Always keep on this device"), so the hub never depends on on-demand hydration (it failed while OneDrive was
# busy uploading, 2026-11-10). New batches outside Pictures are refused; existing ones get an admin warning (the
# SharePoint sync blocks them on its own). Batches inside Pictures are pinned on add and by the background scan.
PICTURES_ROOT = os.environ.get("QA_HUB_PICTURES_ROOT", r"C:\Users\AndronikLindgren\OneDrive - CaMa Gruppen AB\Pictures")
_ATTR_PINNED = 0x00080000
_STORAGE_IGNORE = {"desktop.ini", "thumbs.db"}
_STORAGE_SCAN_S = 600
_STORAGE_MIN_FREE_GB = 20
_STORAGE_WARNINGS = {}  # batch id -> issue dict (only batches with a problem)
_STORAGE_LOCK = threading.Lock()


_STORAGE_MAX_SCAN_BYTES = 15 * 10 ** 9  # most a scan leaves OneDrive to download (several batches); one bigger batch still goes
_K32 = ctypes.WinDLL("kernel32", use_last_error=True)


def _is_link(p):
    """Junction or symlink. Not the bare 0x400 attribute: every OneDrive file and folder is a (cloud) reparse point."""
    return os.path.islink(p) or os.path.isjunction(p)


def _storage_log(line, flush=True):
    """[storage-gate] log line, never fatal: stdout can be cp1252 (qa_watchdog.ps1 redirect, no PYTHONUTF8), so characters
    it cannot encode are replaced, and any other print error is swallowed (blind gate v3 R5: HTTP 400 after save, scan abort)."""
    try:
        enc = getattr(sys.stdout, "encoding", None) or "utf-8"
        print(line.encode(enc, "replace").decode(enc, "replace"), flush=flush)
    except Exception:
        pass


def in_pictures(path):
    """Strict subfolder of Pictures by real path (8.3 names, case, slashes and .. resolve), and no junction/symlink on the
    way from Pictures to the folder: a link inside Pictures pointing elsewhere must not count as inside (blind gate R1)."""
    if not path:
        return False
    try:
        pics = os.path.normcase(os.path.realpath(PICTURES_ROOT))
        cur = os.path.abspath(path)
        if not os.path.normcase(os.path.realpath(cur)).startswith(pics + os.sep):
            return False
        root_text = os.path.normcase(os.path.abspath(PICTURES_ROOT))
        while True:
            # link check BEFORE the "reached Pictures" exit: a junction to the Pictures root resolves to it (blind gate v2)
            if _is_link(cur) and os.path.normcase(cur) != root_text:
                return False
            if os.path.normcase(os.path.realpath(cur)) == pics:
                break
            parent = os.path.dirname(cur)
            if parent == cur:
                return False
            cur = parent
    except (OSError, ValueError):
        return False
    return True


def _long_path(p):
    """\\\\?\\ prefix: no MAX_PATH limit (attrib.exe silently skipped >= 260-char paths and still exited 0, 2026-11-10)."""
    return p if p.startswith("\\\\?\\") else "\\\\?\\" + os.path.abspath(p)


def _walk_no_links(root, links=None):
    """os.walk on the long path that never enters a junction/symlink folder (Python 3.13 os.walk follows junctions).
    Links met (folders and files) are left out and appended to `links`."""
    for dirpath, dirs, names in os.walk(_long_path(root)):
        for group in (dirs, names):
            keep = [n for n in group if not _is_link(os.path.join(dirpath, n))]
            if links is not None:
                links.extend(os.path.join(dirpath, n) for n in group if n not in keep)
            group[:] = keep
        yield dirpath, dirs, names


def _batch_roots(path, links=None):
    """The batch folder and its "<name> Removed" sibling. A root that is a junction/symlink or not inside Pictures is left
    out and appended to `links` (blind gate v2: a Removed junction let +P reach a folder outside)."""
    roots = []
    for r in (os.path.normpath(path), removed_root_for(path)):
        if not os.path.lexists(r):
            continue
        if _is_link(r) or not in_pictures(r):
            if links is not None:
                links.append(_long_path(r))
        elif os.path.isdir(r):
            roots.append(r)
    return roots


def _hydration_bytes(path):
    """Bytes OneDrive still has to download for the batch (+ Removed): cloud-only placeholders report their full size."""
    total = 0
    for root in _batch_roots(path):
        for dirpath, _dirs, names in _walk_no_links(root):
            for n in names:
                st = os.stat(os.path.join(dirpath, n))
                if st.st_file_attributes & _CLOUD_ONLY_ATTRS:
                    total += st.st_size
    return total


def storage_issues(path):
    """None when the batch folder (+ "<name> Removed") is inside Pictures and every folder and file is pinned and local.
    Criterion (Koordinator 4.3): every file HAS P; a file with neither P nor U (0x20) counts, unreadable counts too."""
    if not in_pictures(path):
        return {"outside_pictures": True, "cloud_only": 0, "unpinned": 0,
                "message": "Källmappen ligger utanför OneDrive Pictures, är Pictures-roten eller nås via en genväg "
                           "(junction/symlänk) (ingen backup, SharePoint-synk stoppad)"}
    cloud = unpinned = unreadable = images = images_unpinned = 0
    links = []
    example = ""
    for root in _batch_roots(path, links):
        for dirpath, _dirs, names in _walk_no_links(root, links):
            entries = [(dirpath, True)] + [(os.path.join(dirpath, n), False) for n in names if n.lower() not in _STORAGE_IGNORE]
            for full, is_dir in entries:
                is_image = not is_dir and full.lower().endswith(IMAGE_EXTS)
                images += is_image
                try:
                    a = getattr(os.stat(full), "st_file_attributes", 0)
                except OSError:
                    unreadable += 1
                    example = example or os.path.relpath(full[4:], os.path.dirname(root))
                    continue
                if not is_dir and a & _CLOUD_ONLY_ATTRS:
                    cloud += 1
                if not a & _ATTR_PINNED:  # every file must HAVE P; 0x20 (neither P nor U) counts (Koordinator 4.3)
                    unpinned += 1
                    images_unpinned += is_image
                if (not is_dir and a & _CLOUD_ONLY_ATTRS) or not a & _ATTR_PINNED:
                    example = example or os.path.relpath(full[4:], os.path.dirname(root))
    if not (cloud or unpinned or unreadable or links):
        return None
    return {"outside_pictures": False, "cloud_only": cloud, "unpinned": unpinned, "unreadable": unreadable,
            "links": len(links), "images": images, "images_unpinned": images_unpinned,
            "example": example or (links[0][4:] if links else ""),
            "message": f"{cloud} fil(er) bara i molnet, {unpinned} ej fästa på enheten ({images_unpinned} av {images} "
                       f"bildfiler){f', {unreadable} oläsbara' if unreadable else ''}"
                       f"{f', {len(links)} genvägar (junction/symlänk) hoppas över' if links else ''} (SharePoint-synk stoppad)"}


def pin_batch_folder(path, reserved_bytes=0):
    """-U +P on the folder tree incl. the folder itself and its "<name> Removed" sibling; never enters a junction/symlink.
    Pinning makes OneDrive download every cloud-only file, so free - (bytes already queued) - (this batch's cloud-only
    bytes) must stay >= _STORAGE_MIN_FREE_GB. Returns an error string or ""."""
    if not in_pictures(path):
        return "utanför Pictures"
    need = _hydration_bytes(path)
    free = shutil.disk_usage(PICTURES_ROOT).free
    if (free - reserved_bytes - need) / 1e9 < _STORAGE_MIN_FREE_GB:
        return (f"C: {free / 1e9:.1f} GB ledigt, {need / 1e9:.1f} GB ska laddas ned"
                f"{f' (+{reserved_bytes / 1e9:.1f} GB redan köat)' if reserved_bytes else ''}: under "
                f"{_STORAGE_MIN_FREE_GB} GB kvar, fäster inte")
    settable = 0x1 | 0x2 | 0x4 | 0x20 | 0x100 | 0x2000  # R H S A TEMPORARY NOT_CONTENT_INDEXED, kept as they are
    for root in _batch_roots(path):
        for dirpath, _dirs, names in _walk_no_links(root):
            for full in [dirpath] + [os.path.join(dirpath, n) for n in names if n.lower() not in _STORAGE_IGNORE]:
                a = os.stat(full).st_file_attributes
                if a & _ATTR_PINNED and not a & 0x00100000:
                    continue
                if not _K32.SetFileAttributesW(full, (a & settable) | _ATTR_PINNED):
                    return f"SetFileAttributes fel {ctypes.get_last_error()}: {full[4:][-120:]}"
    return ""


def storage_scan_once(pin=True):
    """Re-check every batch, pin unpinned or cloud-only ones inside Pictures, log a line whenever a batch's state changes.
    `queued` = cloud-only bytes this scan already left to OneDrive (free space only drops once it has downloaded them)."""
    queued = 0
    for b in load_batches():
        bid, path = b.get("id", ""), b.get("path", "")
        try:
            issue = storage_issues(path)
            if issue and not issue["outside_pictures"] and (issue["unpinned"] or issue["cloud_only"]):
                need = _hydration_bytes(path)
                if not pin:
                    pass
                elif queued and queued + need > _STORAGE_MAX_SCAN_BYTES:
                    issue["pin_error"] = f"väntar till nästa skanning (max {_STORAGE_MAX_SCAN_BYTES / 1e9:.0f} GB per skanning)"
                else:
                    err = pin_batch_folder(path, reserved_bytes=queued)
                    if err:
                        issue["pin_error"] = err
                    else:
                        queued += need
                        issue = storage_issues(path)
        except Exception as e:
            issue = {"outside_pictures": False, "cloud_only": 0, "unpinned": 0, "message": f"Kontroll misslyckades: {e}"}
        with _STORAGE_LOCK:
            before = _STORAGE_WARNINGS.get(bid)
            if issue:
                _STORAGE_WARNINGS[bid] = issue
            else:
                _STORAGE_WARNINGS.pop(bid, None)
        if bool(before) != bool(issue) or (before and issue and before.get("message") != issue.get("message")):
            _storage_log(f"[storage-gate] {bid}: {issue['message'] + (' / ' + issue['pin_error'] if issue.get('pin_error') else '') if issue else 'OK (Pictures, fäst, lokal)'}", flush=True)


def _storage_scan_loop():
    while True:
        try:
            storage_scan_once()
        except Exception as e:
            _storage_log(f"[storage-gate] scan failed: {e}", flush=True)
        threading.Event().wait(_STORAGE_SCAN_S)


def _image_dims(full_path):
    """(w, h) in display orientation, or None. PIL's open() is lazy: header only, no pixel decode."""
    if not full_path:
        return None
    try:
        st = os.stat(full_path)
    except OSError:
        return None
    if getattr(st, "st_file_attributes", 0) & _CLOUD_ONLY_ATTRS:
        return None
    key = (os.path.normcase(full_path), st.st_mtime_ns, st.st_size)
    with _DIMS_LOCK:
        if key in _DIMS_CACHE:
            _DIMS_CACHE.move_to_end(key)
            return _DIMS_CACHE[key]
    dims = None
    if Image is not None:
        try:
            with Image.open(full_path) as im:
                w, h = im.size
                try:
                    if im.getexif().get(0x0112) in (5, 6, 7, 8):
                        w, h = h, w
                except Exception:
                    pass
                dims = (int(w), int(h)) if w and h else None
        except Exception:
            dims = None
    with _DIMS_LOCK:
        _DIMS_CACHE[key] = dims
        while len(_DIMS_CACHE) > 20000:
            _DIMS_CACHE.popitem(last=False)
    return dims


def _aspect_label(dims):
    if not dims:
        return "?"
    w, h = dims
    name, ref = min(_KNOWN_RATIOS, key=lambda kv: abs(kv[1] - w / h))
    return name if abs(ref - w / h) / ref <= 0.03 else f"{w}×{h}"


def _site_eligible(dims):
    return bool(dims) and SITE_MIN_RATIO <= dims[0] / dims[1] <= SITE_MAX_RATIO


def load_name_mapping(base_dir):
    """MAPPING.csv (folder,file,old_name,...) -> {new_filename_lower: (old_name, grade)}."""
    path = os.path.join(base_dir or "", "MAPPING.csv")
    try:
        mt = os.stat(path).st_mtime_ns
    except OSError:
        return {}
    hit = _NAME_MAP_CACHE.get(path)
    if hit and hit[0] == mt:
        return hit[1]
    data = {}
    try:
        with open(path, encoding="utf-8-sig", newline="") as f:
            for row in csv.DictReader(f):
                fn = (row.get("file") or "").strip().lower()
                if fn:
                    data[fn] = (row.get("old_name") or "").strip()
    except (OSError, csv.Error):
        data = {}
    _NAME_MAP_CACHE[path] = (mt, data)
    return data


def classify_shot(shot, batch=None, orig_name=None):
    """'studio' | 'ai'. Order: explicit shot_kind/label > batch content_kind > original name > whole filename tokens > batch hints."""
    sk_low = str(shot.get("shot_kind") or shot.get("label") or "").lower()
    if sk_low:
        if "studioreferens" in sk_low or "studiobild" in sk_low or "packshot" in sk_low:
            return "studio"
        if "genererad" in sk_low or "interiör" in sk_low or "interior" in sk_low or "miljö" in sk_low or "video" in sk_low:
            return "ai"
    kind = batch.get("content_kind") if isinstance(batch, dict) else None
    if kind in ("studio", "ai"):
        return kind
    code = str(shot.get("shot") or "")
    if code.endswith("_miljo") or code.startswith("miljo_"):
        return "ai"
    for name in (orig_name, shot.get("filename")):
        if not name:
            continue
        stem = os.path.splitext(name.lower())[0]
        if set(re.split(r'[-_ .()]+', stem)) & _AI_TOKENS or _AI_NAME_RE.search(stem) or _AI_SCENE_RE.search(stem):
            return "ai"
    if isinstance(batch, dict):
        hay = f"{batch.get('id', '')} {batch.get('path', '')}".lower()
        if any(h in hay for h in _AI_BATCH_HINTS):
            return "ai"
    return "studio"


def _is_interior_or_ai_shot(shot, batch=None):
    return classify_shot(shot, batch, shot.get("orig_name")) == "ai"


def _extract_scene_key(filename):
    stem = os.path.splitext(os.path.basename(filename or ""))[0].lower()
    # Cap lifestyle scene pattern from MAPPING.csv old_name (e.g. -cap-close-u-007b -> close-u)
    m_cap = re.search(r'-cap-([a-z-]+?)-\d{3}[a-z]?$', stem)
    if m_cap:
        return m_cap.group(1)
    # Strip trailing crop / variant tags like -3x4, -4x5, -9x16, -fix1, -v2, or trailing letter 'b' on 3-digit numbers
    stem = re.sub(r'-(3x4|4x5|9x16|916)(?:-[a-z0-9]+)?$', '', stem)
    stem = re.sub(r'-(fix\d*|v\d+|new|ny)$', '', stem)
    # Pattern like tc-charonne-01-26-013 or charonne-beige_01_16-001 -> strip trailing -NNN generation index
    m_gen = re.match(r'^(.+[_-]\d+)-\d{3}[a-z]?$', stem)
    if m_gen:
        stem = m_gen.group(1)
    # Pattern ending with -001b or -001 -> strip trailing variant letter
    m_num = re.match(r'^(.*?-\d{2,3})[a-d]$', stem)
    if m_num:
        stem = m_num.group(1)
    # Also group by room/pose descriptor if present before the final shot number
    m_room = re.match(r'^(\d{2}-[a-z]+-[a-z]+)-.*$', stem)
    if m_room:
        return m_room.group(1)
    return stem


def _get_visual_signature(full_path, mtime):
    if not full_path or not os.path.isfile(full_path) or _is_cloud_only(full_path):
        return None
    key = (os.path.normcase(full_path), int(mtime or 0))
    with _VISUAL_SIG_LOCK:
        if key in _VISUAL_SIG_CACHE:
            _VISUAL_SIG_CACHE.move_to_end(key)
            return _VISUAL_SIG_CACHE[key]
    sig = None
    if Image is not None:
        try:
            with Image.open(full_path) as im:
                w, h = im.size
                aspect = round(w / max(1, h), 2)
                try:
                    im.draft("RGB", (64, 64))
                except Exception:
                    pass
                small = im.convert("RGB").resize((8, 8))
                raw_bytes = small.tobytes()
                vec = tuple(b / 255.0 for b in raw_bytes)
                sig = (aspect, vec)
        except Exception:
            sig = None
    with _VISUAL_SIG_LOCK:
        _VISUAL_SIG_CACHE[key] = sig
        while len(_VISUAL_SIG_CACHE) > 4000:
            _VISUAL_SIG_CACHE.popitem(last=False)
    return sig


def _visual_similarity(sig_a, sig_b):
    if not sig_a or not sig_b:
        return 0.0
    vec_a, vec_b = sig_a[1], sig_b[1]
    if len(vec_a) != len(vec_b) or not vec_a:
        return 0.0
    mad = sum(abs(a - b) for a, b in zip(vec_a, vec_b)) / len(vec_a)
    return max(0.0, min(1.0, 1.0 - mad * 2.2))


def annotate_product_gold_set(batch, products):
    """
    Annotates shots on every product with:
      - width, height, aspect_label, site_eligible (True ONLY for 1:1 square images: 0.95 <= w/h <= 1.05)
      - content_kind ('studio' | 'ai') and is_studio_packshot (True for all studio shots)
      - stars_avg, stars_count
      - scene_key (for variation / anti-twin checking)
      - in_gold_set (True for the 2..4 selected 1:1 site AI/interior images), gold_rank (1..4), gold_reason
    """
    if not products or not isinstance(products, list):
        return products
    batch_id = batch.get("id", "") if isinstance(batch, dict) else ""
    base_dir = batch.get("path", "") if isinstance(batch, dict) else ""
    if base_dir:
        _apply_product_mapping_enrichment(products, _load_product_image_mapping(base_dir))
    all_ratings = load_ratings(batch_id) if batch_id else {}
    name_map = load_name_mapping(base_dir)

    # Build lookup: (product_id_lower, filename_lower) -> {reviewer: stars}
    stars_lookup = {}
    if isinstance(all_ratings, dict):
        for rev_name, user_map in all_ratings.items():
            if not isinstance(user_map, dict):
                continue
            for r_key, val in user_map.items():
                if "|" not in r_key:
                    continue
                pid_part, fn_part = r_key.split("|", 1)
                s_val = val.get("stars", 0) if isinstance(val, dict) else (val if isinstance(val, int) else 0)
                if isinstance(s_val, int) and 1 <= s_val <= 5:
                    k_norm = (pid_part.strip().lower(), fn_part.strip().lower())
                    stars_lookup.setdefault(k_norm, {})[rev_name] = s_val
                    stem_norm = (pid_part.strip().lower(), os.path.splitext(fn_part.strip())[0].lower())
                    stars_lookup.setdefault(stem_norm, {})[rev_name] = s_val

    for prod in products:
        pid = str(prod.get("id", "")).strip()
        pid_low = pid.lower()
        variants = prod.get("variants") or {}
        if not variants:
            continue

        # Annotate dimensions, aspect_label, site_eligible (1:1 only), and content_kind on all variants first
        for vname, vshots in variants.items():
            for s in vshots or []:
                rel_p = s.get("rel_path") or s.get("filename") or ""
                full_p = os.path.join(base_dir, rel_p) if base_dir else ""
                is_vid = rel_p.lower().endswith(VIDEO_EXTS)
                s["is_video"] = is_vid
                if is_vid and s.get("url") and "fs=2" not in s["url"]:
                    sep = "&" if "?" in s["url"] else "?"
                    s["url"] = f"{s['url']}{sep}fs=2"
                dims = _image_dims(full_p) if not is_vid else None
                if is_vid and not dims:
                    low_p = full_p.lower()
                    if "1x1" in low_p:
                        dims = (1080, 1080)
                    elif "16x9" in low_p:
                        dims = (1280, 720)
                    else:
                        dims = (720, 1280)
                s["width"], s["height"] = dims or (None, None)
                s["aspect_label"] = _aspect_label(dims)
                s["site_eligible"] = (not is_vid) and _site_eligible(dims)
                orig = s.get("orig_name") or name_map.get((s.get("filename") or "").lower())
                s["content_kind"] = "ai" if is_vid else classify_shot(s, batch, orig)
                s["is_studio_packshot"] = (s["content_kind"] == "studio")
                s["scene_key"] = _extract_scene_key(orig or s.get("filename"))

        primary_vname = "Original" if "Original" in variants else ("1500px" if "1500px" in variants else next(iter(variants.keys())))
        primary_shots = variants.get(primary_vname) or []
        if not primary_shots:
            continue

        # Candidate indices for Gold Set: ONLY AI/interior shots that are 1:1 (site_eligible == True)
        cand_indices = [
            i for i, s in enumerate(primary_shots)
            if s.get("content_kind") == "ai" and s.get("site_eligible")
        ]

        candidates = []
        for rank_pos, idx in enumerate(cand_indices):
            shot = primary_shots[idx]
            fn = shot.get("filename") or ""
            stem_low = os.path.splitext(fn)[0].lower()
            s_by = stars_lookup.get((pid_low, fn.lower())) or stars_lookup.get((pid_low, stem_low)) or {}
            s_vals = list(s_by.values())
            s_count = len(s_vals)
            s_avg = round(sum(s_vals) / s_count, 2) if s_count else 0.0
            s_max = max(s_vals) if s_count else 0

            judge_base = max(50.0, 95.0 - idx * 2.0)
            if stem_low.endswith("b"):
                judge_base -= 0.5

            rel_p = shot.get("rel_path") or fn
            full_p = os.path.join(base_dir, rel_p) if base_dir else ""
            want_sig = (s_count > 0 or rank_pos < 12)
            vis_sig = _get_visual_signature(full_p, shot.get("mtime", 0)) if want_sig else None

            candidates.append({
                "idx": idx,
                "stem": stem_low,
                "filename": fn,
                "stars_by": s_by,
                "stars_count": s_count,
                "stars_avg": s_avg,
                "stars_max": s_max,
                "judge_score": round(judge_base, 1),
                "scene_key": shot.get("scene_key") or _extract_scene_key(fn),
                "vis_sig": vis_sig,
            })

        # Filter out low-rated images (stars_avg < 2 when rated) from being promoted to Gold Set
        eligible_candidates = [c for c in candidates if not (c["stars_count"] > 0 and c["stars_avg"] < 2.5)]
        starred_scenes = {c["scene_key"] for c in eligible_candidates if c["stars_count"] > 0 and c["stars_avg"] >= 3.0}
        target_gold = min(len(eligible_candidates), max(2, min(4, len(starred_scenes)))) if eligible_candidates else 0

        selected = []
        remaining = list(eligible_candidates)
        while len(selected) < target_gold and remaining:
            best_cand = None
            best_tuple = None
            selected_scenes = {s["scene_key"] for s in selected}

            for c in remaining:
                # Only promote to tier 2 when average star rating is >= 3.0
                tier = 2 if (c["stars_count"] > 0 and c["stars_avg"] >= 3.0) else 1
                same_scene = c["scene_key"] in selected_scenes
                max_sim = max((_visual_similarity(c["vis_sig"], s["vis_sig"]) for s in selected), default=0.0)
                is_visually_twin = max_sim > 0.86

                diversity_ok = (not same_scene) and (not is_visually_twin)
                base_val = (c["stars_avg"] * 20.0 + min(c["stars_count"], 3) * 3.0 + c["judge_score"] * 0.1) if tier == 2 else c["judge_score"]
                mmr_score = base_val - (35.0 if same_scene else 0.0) - (25.0 * max_sim)

                rank_tuple = (
                    tier,
                    1 if diversity_ok else 0,
                    round(mmr_score, 3),
                    c["stars_avg"],
                    c["stars_count"],
                    c["judge_score"],
                    -c["idx"]
                )
                if best_tuple is None or rank_tuple > best_tuple:
                    best_tuple = rank_tuple
                    best_cand = c

            if best_cand is None:
                break
            selected.append(best_cand)
            remaining.remove(best_cand)

        gold_by_stem = {}
        for rank_idx, sel in enumerate(selected, start=1):
            if sel["stars_count"] > 0 and sel["stars_avg"] >= 3.0:
                reason = f"⭐ Stjärnval #{rank_idx} ({sel['stars_avg']:g}★ snitt, {sel['stars_count']} betyg)"
            else:
                reason = f"🏆 Guldset #{rank_idx} (Urvalsrank #{rank_idx})"
            gold_by_stem[sel["stem"]] = (rank_idx, reason)

        # Annotate all variants of this product
        for vname, vshots in variants.items():
            if not isinstance(vshots, list):
                continue
            for s in vshots:
                fn = s.get("filename") or ""
                stem_low = os.path.splitext(fn)[0].lower()
                s_by = stars_lookup.get((pid_low, fn.lower())) or stars_lookup.get((pid_low, stem_low)) or {}
                s_vals = list(s_by.values())
                s_count = len(s_vals)
                s_avg = round(sum(s_vals) / s_count, 2) if s_count else 0.0

                s["stars_count"] = s_count
                s["stars_avg"] = s_avg

                if stem_low in gold_by_stem and s.get("site_eligible") and not s.get("is_studio_packshot"):
                    g_rank, g_reason = gold_by_stem[stem_low]
                    s["in_gold_set"] = True
                    s["gold_rank"] = g_rank
                    s["gold_reason"] = g_reason
                else:
                    s["in_gold_set"] = False
                    s["gold_rank"] = None
                    s["gold_reason"] = ""
            # CEO 2026-11-06: studio/main images (slots -01..-08) keep their fixed original order first; only interiors
            # after them are sorted by rating, highest on top (stable sort). An interior never becomes image 1.
            def _fixed_slot(s):
                sk = (s.get("scene_key") or "").lower()
                return bool(s.get("is_studio_packshot") or s.get("is_primary") or sk.startswith(("studio_", "hero_")))
            vshots.sort(key=lambda s: (0, 0) if _fixed_slot(s) else (1, -(s.get("stars_avg") or 0)))

        prod["gold_count"] = len(gold_by_stem)
        prod["site_candidates"] = len(cand_indices)
        prod["ai_shot_count"] = sum(1 for s in primary_shots if s.get("content_kind") == "ai" and not s.get("is_video"))
        prod["video_shot_count"] = sum(1 for s in primary_shots if s.get("is_video"))
        prod["has_starred_shots"] = any(c["stars_count"] > 0 and c["stars_avg"] >= 3.0 for c in candidates)

    return products


def removed_root_for(batch_path):
    p = os.path.normpath(str(batch_path))
    return os.path.join(os.path.dirname(p), os.path.basename(p) + " Removed")


def remove_shot(batch, product_id, filename, user, count_user=None, undo_user=None, protect_studio=False):
    base_dir = batch.get("path", "")
    if "reforma_master_catalog" in base_dir.lower():
        raise ValueError("Mappen använder manifest – flytt stöds inte")

    products = scan_batch_products(batch)
    product = next((p for p in products if str(p.get("id")) == str(product_id) or str(p.get("id")).lower() == str(product_id).lower()), None)
    if not product:
        if os.path.isfile(os.path.join(base_dir, "products_manifest.json")):
            raise ValueError("Mappen använder manifest – flytt stöds inte")
        raise ValueError("Produkten hittades inte")

    target_stem = os.path.splitext(os.path.basename(filename))[0].lower() if filename else ""
    if not target_stem:
        raise ValueError("Bilden hittades inte")

    matching_shots = []
    seen_rel = set()
    for variant_name, shots in product.get("variants", {}).items():
        if not isinstance(shots, list):
            continue
        for shot in shots:
            fn = shot.get("filename") or os.path.basename(shot.get("rel_path", ""))
            s_stem = os.path.splitext(os.path.basename(fn))[0].lower()
            if s_stem == target_stem:
                rel_p = shot.get("rel_path")
                if rel_p and rel_p not in seen_rel:
                    seen_rel.add(rel_p)
                    matching_shots.append(shot)

    if not matching_shots:
        raise ValueError("Bilden hittades inte")

    if protect_studio and any(s.get("is_studio_packshot") for s in matching_shots):
        raise PermissionError("Studiobilder (packshots) är skyddade mot oavsiktlig radering. Endast admin kan flytta ut studiobilder — du kan däremot flytta ut miljö-/AI-bilder.")


    batch_path = os.path.normpath(os.path.abspath(base_dir))
    removed_root = removed_root_for(batch_path)

    # Validate all src paths lie within batch_path before moving anything
    moves_plan = []
    for shot in matching_shots:
        rel_path = shot.get("rel_path", "")
        src = os.path.normpath(os.path.abspath(os.path.join(batch_path, rel_path)))
        try:
            p1_norm = os.path.normcase(os.path.abspath(batch_path))
            src_norm = os.path.normcase(os.path.abspath(src))
            common = os.path.normcase(os.path.commonpath([p1_norm, src_norm]))
            if common != p1_norm or src_norm == p1_norm:
                raise ValueError(f"Sökvägen ligger utanför mappen: {rel_path}")
        except ValueError as e:
            raise ValueError(f"Sökvägen ligger utanför mappen: {rel_path}") from e

        dst = os.path.normpath(os.path.abspath(os.path.join(removed_root, rel_path)))
        if os.path.exists(dst):
            dst_dir = os.path.dirname(dst)
            base_name, ext = os.path.splitext(os.path.basename(dst))
            counter = 1
            while True:
                candidate = os.path.join(dst_dir, f"{base_name}-{counter}{ext}")
                if not os.path.exists(candidate):
                    dst = candidate
                    break
                counter += 1

        moves_plan.append((src, dst, rel_path.replace("\\", "/")))

    # Append crop siblings for any 9x16 sources
    seen_plan_srcs = {p[0] for p in moves_plan}
    for src, _, _ in list(moves_plan):
        siblings = crop_siblings(src, batch_path=batch_path)
        for sib_src in siblings:
            if sib_src in seen_plan_srcs:
                continue
            seen_plan_srcs.add(sib_src)
            sib_rel = os.path.relpath(sib_src, batch_path)
            dst = os.path.normpath(os.path.abspath(os.path.join(removed_root, sib_rel)))
            if os.path.exists(dst):
                dst_dir = os.path.dirname(dst)
                base_name, ext = os.path.splitext(os.path.basename(dst))
                counter = 1
                while True:
                    candidate = os.path.join(dst_dir, f"{base_name}-{counter}{ext}")
                    if not os.path.exists(candidate):
                        dst = candidate
                        break
                    counter += 1
            moves_plan.append((sib_src, dst, sib_rel.replace("\\", "/")))

    moves = []
    moved_rel_paths = []
    for src, dst, rel_norm in moves_plan:
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.move(src, dst)
        moves.append({"from": os.path.abspath(src), "to": os.path.abspath(dst)})
        moved_rel_paths.append(rel_norm)

    log_path = os.path.join(removed_root, "_qa_removed_log.json")
    log_data = []
    if os.path.exists(log_path):
        try:
            with open(log_path, "r", encoding="utf-8") as f:
                log_data = json.load(f)
            if not isinstance(log_data, list):
                log_data = []
        except Exception:
            log_data = []

    op_id = uuid.uuid4().hex
    op = {
        "op_id": op_id,
        "ts": datetime.datetime.now().isoformat(),
        "user": user,
        "batch_id": batch.get("id", ""),
        "product_id": product_id,
        "filename": filename,
        "moves": moves,
        "undone": False
    }
    log_data.append(op)
    os.makedirs(removed_root, exist_ok=True)
    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(log_data, f, indent=2, ensure_ascii=False)

    cu = count_user if count_user is not None else undo_user
    if cu is not None:
        undo_count = sum(1 for o in log_data if not o.get("undone", False) and o.get("user") == cu)
    else:
        undo_count = sum(1 for o in log_data if not o.get("undone", False))

    ret = {
        "status": "ok",
        "op_id": op_id,
        "moved": moved_rel_paths,
        "removed_root": removed_root,
        "undo_count": undo_count
    }

    # Low-count warning: re-scan with scan_batch_products(batch)
    updated_products = scan_batch_products(batch)
    updated_prod = next(
        (p for p in updated_products if str(p.get("id")) == str(product_id) or str(p.get("id")).lower() == str(product_id).lower()),
        None
    )
    if updated_prod and updated_prod.get("variants"):
        variant_values = list(updated_prod.get("variants", {}).values())
        remaining = max((len(shots) for shots in variant_values if isinstance(shots, list)), default=0)
    else:
        remaining = 0

    if remaining < LOW_SHOT_THRESHOLD:
        ret["low_warning"] = {"product_id": product_id, "remaining": remaining}
        alerts_path = os.path.join(removed_root, "_qa_low_alerts.json")
        alerts = {}
        if os.path.exists(alerts_path):
            try:
                with open(alerts_path, "r", encoding="utf-8") as f:
                    alerts = json.load(f)
                if not isinstance(alerts, dict):
                    alerts = {}
            except Exception:
                alerts = {}

        pid_key = str(product_id)
        if pid_key not in alerts:
            alerts[pid_key] = datetime.datetime.now().isoformat()
            try:
                with open(alerts_path, "w", encoding="utf-8") as f:
                    json.dump(alerts, f, indent=2, ensure_ascii=False)
            except Exception:
                pass

            if NOTIFY_SCRIPT and os.path.isfile(NOTIFY_SCRIPT):
                try:
                    batch_id = batch.get("id", "")
                    msg = f"[QA-hubb 1.0] Få bilder kvar: {batch_id} / {product_id} har {remaining} bild(er) kvar efter utrensning ({user})"
                    subprocess.Popen(["node", NOTIFY_SCRIPT, msg])
                except Exception:
                    pass

    return ret


def undo_remove(batch, op_id=None, user=None):
    batch_path = os.path.normpath(os.path.abspath(batch.get("path", "")))
    removed_root = removed_root_for(batch_path)
    log_path = os.path.join(removed_root, "_qa_removed_log.json")
    if not os.path.exists(log_path):
        raise ValueError("Inget att ångra")

    try:
        with open(log_path, "r", encoding="utf-8") as f:
            log_data = json.load(f)
        if not isinstance(log_data, list):
            raise ValueError("Inget att ångra")
    except Exception:
        raise ValueError("Inget att ångra")

    target_op = None
    if op_id:
        for op in reversed(log_data):
            if op.get("op_id") == op_id and not op.get("undone", False):
                if user is not None and op.get("user") != user:
                    continue
                target_op = op
                break
    else:
        for op in reversed(log_data):
            if not op.get("undone", False):
                if user is not None and op.get("user") != user:
                    continue
                target_op = op
                break

    if not target_op:
        raise ValueError("Inget att ångra")

    moves = target_op.get("moves", [])
    reversed_moves = list(reversed(moves))

    # Check every move first
    for m in reversed_moves:
        from_path = m["from"]
        if os.path.exists(from_path):
            raise FileExistsError(f"Originalfilen finns redan: {from_path}")

    # Move files back
    restored = []
    for m in reversed_moves:
        from_path = m["from"]
        to_path = m["to"]
        os.makedirs(os.path.dirname(from_path), exist_ok=True)
        shutil.move(to_path, from_path)
        restored.append(from_path)

    target_op["undone"] = True
    target_op["undone_ts"] = datetime.datetime.now().isoformat()

    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(log_data, f, indent=2, ensure_ascii=False)

    # In undo_remove, after restoring, if the product now has >= LOW_SHOT_THRESHOLD shots,
    # delete its key from _qa_low_alerts.json so a later drop alerts again.
    product_id = target_op.get("product_id")
    if product_id:
        try:
            updated_products = scan_batch_products(batch)
            updated_prod = next(
                (p for p in updated_products if str(p.get("id")) == str(product_id) or str(p.get("id")).lower() == str(product_id).lower()),
                None
            )
            if updated_prod and updated_prod.get("variants"):
                variant_values = list(updated_prod.get("variants", {}).values())
                remaining = max((len(shots) for shots in variant_values if isinstance(shots, list)), default=0)
            else:
                remaining = 0

            if remaining >= LOW_SHOT_THRESHOLD:
                alerts_path = os.path.join(removed_root, "_qa_low_alerts.json")
                if os.path.exists(alerts_path):
                    with open(alerts_path, "r", encoding="utf-8") as f:
                        alerts = json.load(f)
                    if isinstance(alerts, dict):
                        changed = False
                        for k in [str(product_id), product_id]:
                            if k in alerts:
                                del alerts[k]
                                changed = True
                        if changed:
                            with open(alerts_path, "w", encoding="utf-8") as f:
                                json.dump(alerts, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    if user is not None:
        undo_count = sum(1 for o in log_data if not o.get("undone", False) and o.get("user") == user)
    else:
        undo_count = sum(1 for o in log_data if not o.get("undone", False))

    return {
        "status": "ok",
        "op_id": target_op["op_id"],
        "restored": restored,
        "undo_count": undo_count
    }


def set_remove_reason(batch, op_id, reason, note, user):
    if reason not in REMOVE_REASONS and reason != "":
        raise ValueError("Okänd orsak")

    batch_path = os.path.normpath(os.path.abspath(batch.get("path", "")))
    removed_root = removed_root_for(batch_path)
    log_path = os.path.join(removed_root, "_qa_removed_log.json")
    if not os.path.exists(log_path):
        raise ValueError("Åtgärden hittades inte")

    try:
        with open(log_path, "r", encoding="utf-8") as f:
            log_data = json.load(f)
        if not isinstance(log_data, list):
            raise ValueError("Åtgärden hittades inte")
    except Exception as e:
        if isinstance(e, ValueError):
            raise
        raise ValueError("Åtgärden hittades inte")

    target_op = None
    for op in log_data:
        if op.get("op_id") == op_id:
            target_op = op
            break

    if not target_op:
        raise ValueError("Åtgärden hittades inte")

    target_op["reason"] = reason
    target_op["note"] = (note or "")[:300]
    target_op["reason_by"] = user
    target_op["reason_ts"] = datetime.datetime.now().isoformat()

    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(log_data, f, indent=2, ensure_ascii=False)

    return {"status": "ok"}


def list_removed(batch, user=None):
    batch_path = os.path.normpath(os.path.abspath(batch.get("path", "")))
    removed_root = removed_root_for(batch_path)
    log_path = os.path.join(removed_root, "_qa_removed_log.json")
    if not os.path.exists(log_path):
        return {"undo_count": 0, "ops": []}

    try:
        with open(log_path, "r", encoding="utf-8") as f:
            log_data = json.load(f)
        if not isinstance(log_data, list):
            return {"undo_count": 0, "ops": []}
    except Exception:
        return {"undo_count": 0, "ops": []}

    active_ops = []
    for op in reversed(log_data):
        if not op.get("undone", False):
            if user is not None and op.get("user") != user:
                continue
            active_ops.append({
                "op_id": op.get("op_id"),
                "ts": op.get("ts"),
                "product_id": op.get("product_id"),
                "filename": op.get("filename"),
                "n_files": len(op.get("moves", []))
            })

    return {
        "undo_count": len(active_ops),
        "ops": active_ops
    }


def collect_batch_zip_files(batch, mode="gold", product_id=None, selected_items=None, variant_pref=None):
    if mode not in ("gold", "product", "all", "selected"):
        raise ValueError(f"Ogiltigt nedladdningsläge: {mode}")

    batch_path = os.path.normpath(os.path.abspath(batch.get("path", "")))
    if not os.path.isdir(batch_path):
        raise ValueError("Mappen finns inte på disk")

    products = scan_batch_products(batch)
    p_norm = os.path.normcase(batch_path)

    allowed_rels = {
        (s.get("rel_path") or "").replace("\\", "/").lstrip("/")
        for p in products
        for v_shots in (p.get("variants") or {}).values()
        for s in (v_shots or [])
        if s.get("rel_path")
    }

    def _safe_abs(rel_p):
        rel_clean = (rel_p or "").replace("\\", "/").lstrip("/")
        if not rel_clean or rel_clean not in allowed_rels:
            return None
        full = os.path.normpath(os.path.abspath(os.path.join(batch_path, rel_clean)))
        try:
            if os.path.normcase(os.path.commonpath([p_norm, os.path.normcase(full)])) != p_norm:
                return None
        except ValueError:
            return None
        if os.path.isfile(full):
            return full
        return None

    files_to_pack = []  # list of (abs_path, arcname)
    skipped_cloud = []
    seen_arc = set()

    def _add_file(abs_p, arcname):
        if not abs_p or not os.path.isfile(abs_p):
            return
        arc = arcname.replace("\\", "/").lstrip("/")
        if _is_cloud_only(abs_p):
            if arc not in skipped_cloud:
                skipped_cloud.append(arc)
            return
        if arc in seen_arc:
            base, ext = os.path.splitext(arc)
            idx = 2
            while f"{base}_{idx}{ext}" in seen_arc:
                idx += 1
            arc = f"{base}_{idx}{ext}"
        seen_arc.add(arc)
        files_to_pack.append((abs_p, arc))

    if mode == "selected":
        if not selected_items:
            raise ValueError("Inga bilder valda för nedladdning")
        for item in selected_items:
            rel_p = ""
            if isinstance(item, dict):
                rel_p = item.get("rel_path") or ""
                if not rel_p and item.get("product_id") and item.get("filename"):
                    pid = str(item.get("product_id"))
                    fn = str(item.get("filename"))
                    prod = next((p for p in products if str(p.get("id")).lower() == pid.lower()), None)
                    if prod:
                        for v_shots in prod.get("variants", {}).values():
                            for s in (v_shots or []):
                                if s.get("filename") == fn or os.path.basename(s.get("rel_path", "")) == fn:
                                    rel_p = s.get("rel_path", "")
                                    break
                            if rel_p:
                                break
            elif isinstance(item, str):
                if "::" in item:
                    _, rel_p = item.split("::", 1)
                else:
                    rel_p = item
            abs_p = _safe_abs(rel_p)
            if abs_p:
                _add_file(abs_p, rel_p)
    else:
        target_products = products
        if product_id:
            target_products = [
                p for p in products
                if str(p.get("id")).lower() == str(product_id).lower()
            ]
        for prod in target_products:
            variants = prod.get("variants", {})
            if not variants:
                continue
            if variant_pref and variant_pref in variants:
                v_keys = [variant_pref]
            elif variant_pref == "all":
                v_keys = list(variants.keys())
            elif "Original" in variants and mode == "gold":
                v_keys = ["Original"]
            else:
                v_keys = list(variants.keys())[:1] if mode == "gold" else list(variants.keys())

            for vk in v_keys:
                shots = variants.get(vk) or []
                for s in shots:
                    if mode == "gold":
                        if not (s.get("site_eligible", True) and (s.get("is_studio_packshot") or s.get("in_gold_set"))):
                            continue
                    rel_p = s.get("rel_path", "")
                    abs_p = _safe_abs(rel_p)
                    if abs_p:
                        _add_file(abs_p, rel_p)

    if not files_to_pack and skipped_cloud:
        raise ValueError("Filerna i urvalet är endast molnlagrade i OneDrive — öppna via OneDrive-knappen ☁️")
    if not files_to_pack:
        raise ValueError("Inga bilder matchade urvalet för nedladdning")

    return files_to_pack, skipped_cloud


def build_batch_zip(batch, mode="gold", product_id=None, selected_items=None, variant_pref=None):
    files_to_pack, skipped_cloud = collect_batch_zip_files(
        batch,
        mode=mode,
        product_id=product_id,
        selected_items=selected_items,
        variant_pref=variant_pref,
    )
    buf = io.BytesIO()
    packed_count = 0
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_STORED) as zf:
        for abs_p, arc in files_to_pack:
            try:
                zf.write(abs_p, arc)
                packed_count += 1
            except OSError:
                continue
        if skipped_cloud:
            zf.writestr("_EJ_NEDLADDADE_FRAN_ONEDRIVE.txt", "\n".join(skipped_cloud))
    if packed_count == 0:
        raise ValueError("Filerna i urvalet är endast molnlagrade i OneDrive — öppna via OneDrive-knappen ☁️")
    buf.seek(0)
    return buf.getvalue(), packed_count


def merge_gold_set_to_batch(source_batch, target_batch, product_id=None, dry_run=True, user="admin"):
    if str(source_batch.get("id", "")).lower() == str(target_batch.get("id", "")).lower():
        raise ValueError("Käll- och målmapp är samma batch")

    src_path = os.path.normpath(os.path.abspath(source_batch.get("path", "")))
    dst_path = os.path.normpath(os.path.abspath(target_batch.get("path", "")))
    if not os.path.isdir(src_path) or not os.path.isdir(dst_path):
        raise ValueError("Käll- eller målmappen finns inte på disk")

    src_products = scan_batch_products(source_batch)
    dst_products = scan_batch_products(target_batch)

    if product_id:
        src_products = [p for p in src_products if str(p.get("id")).lower() == str(product_id).lower()]

    def _norm_slug(val):
        s = str(val or "").lower()
        return re.sub(r'[^a-z0-9åäö]+', '-', s).strip('-')

    def _match_key(p):
        m = parse_product_name(p.get("folder_name") or p.get("id") or "")
        return (_norm_slug(m["brand"]), _norm_slug(m["model"]), _norm_slug(m["color"]))

    dst_by_slug = {_norm_slug(p.get("id")): p for p in dst_products if _norm_slug(p.get("id"))}
    dst_by_key = {}
    for p in dst_products:
        mk = _match_key(p)
        if any(mk):
            dst_by_key.setdefault(mk, []).append(p)

    plan = []
    unmatched = []
    for sp in src_products:
        sp_slug = _norm_slug(sp.get("id"))
        dp = dst_by_slug.get(sp_slug) if sp_slug else None
        if dp is None:
            hits = dst_by_key.get(_match_key(sp), [])
            dp = hits[0] if len(hits) == 1 else None
        if dp is None:
            unmatched.append(sp.get("id"))
            continue

        sv = sp.get("variants") or {}
        src_shots = sv.get("Original") or next(iter(sv.values()), [])
        gold = sorted(
            (
                s for s in src_shots
                if s.get("in_gold_set") and s.get("site_eligible", True) and not s.get("is_studio_packshot")
            ),
            key=lambda s: int(s.get("gold_rank") or 99),
        )
        if gold:
            plan.append((sp, dp, gold))

    if dry_run:
        return {
            "status": "dry_run",
            "matched_products": len(plan),
            "unmatched": unmatched,
            "plan": [
                {
                    "source": sp["id"],
                    "target": dp["id"],
                    "files": [g["filename"] for g in gold],
                }
                for sp, dp, gold in plan
            ],
        }

    copied = []
    skipped_existing = []
    with _MERGE_LOCK:
        for sp, dp, gold_shots in plan:
            dst_variants = dp.get("variants") or {}
            for vk, d_shots in dst_variants.items():
                if not d_shots:
                    continue
                sample_rel = d_shots[0].get("rel_path", "").replace("\\", "/")
                sample_abs = os.path.normpath(os.path.join(dst_path, sample_rel))
                target_dir = os.path.dirname(sample_abs)
                folder_title = os.path.basename(target_dir)
                if folder_title.lower() in ("original", "1500px", "1500px iphone") or target_dir == dst_path:
                    folder_title = dp.get("folder_name") or dp.get("id")

                is_1500 = "1500" in vk.lower()
                for gs in gold_shots:
                    gs_rel = gs.get("rel_path", "")
                    gs_abs = os.path.normpath(os.path.join(src_path, gs_rel))
                    if not os.path.isfile(gs_abs) or _is_cloud_only(gs_abs):
                        continue
                    if not _site_eligible(_image_dims(gs_abs)):
                        continue
                    gs_stem = os.path.splitext(os.path.basename(gs_abs))[0]
                    out_name = f"{folder_title} - Miljö ({gs_stem}).jpg"
                    out_abs = os.path.join(target_dir, out_name)
                    rel_out = os.path.relpath(out_abs, dst_path).replace("\\", "/")
                    if os.path.exists(out_abs):
                        skipped_existing.append(rel_out)
                        continue
                    os.makedirs(target_dir, exist_ok=True)
                    with Image.open(gs_abs) as im:
                        if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
                            rgba = im.convert("RGBA")
                            bg = Image.new("RGB", rgba.size, (255, 255, 255))
                            bg.paste(rgba, mask=rgba.split()[3])
                            rgb = bg
                        else:
                            rgb = im.convert("RGB")
                        if is_1500:
                            w, h = rgb.size
                            if max(w, h) > 1500:
                                rgb = rgb.resize((1500, 1500), Image.Resampling.LANCZOS)
                        rgb.save(out_abs, "JPEG", quality=92, optimize=True)
                    copied.append(rel_out)

        if copied:
            log_file = os.path.join(dst_path, "_sajt_merge_log.json")
            try:
                log_entries = []
                if os.path.exists(log_file):
                    with open(log_file, "r", encoding="utf-8") as lf:
                        log_entries = json.load(lf)
                        if not isinstance(log_entries, list):
                            log_entries = []
                log_entries.append({
                    "ts": datetime.datetime.now().isoformat(),
                    "user": user,
                    "source_batch_id": source_batch.get("id"),
                    "target_batch_id": target_batch.get("id"),
                    "files": copied,
                })
                with open(log_file, "w", encoding="utf-8") as lf:
                    json.dump(log_entries, lf, indent=2, ensure_ascii=False)
            except Exception:
                pass

    return {
        "status": "ok",
        "matched_products": len(plan),
        "unmatched": unmatched,
        "copied_files": len(copied),
        "skipped_existing": skipped_existing,
        "files": copied,
    }


class HubHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=WORKSPACE_DIR, **kwargs)

    def current_user(self):
        try:
            cookie_header = self.headers.get("Cookie")
            if not cookie_header:
                return None
            cookie = http.cookies.SimpleCookie(cookie_header)
            if "qa_session" not in cookie:
                return None
            token = cookie["qa_session"].value
            if not token:
                return None
            token_hash = hashlib.sha256(token.encode('utf-8')).hexdigest()
            users_data = load_users()
            for u in users_data.get("users", []):
                stored_hash = u.get("token_sha256", "")
                if hmac.compare_digest(token_hash, stored_hash):
                    return u
            return None
        except Exception:
            return None

    def reviewer_name(self, fallback="Anonym"):
        u = self.current_user()
        if u and "name" in u:
            return u["name"]
        return fallback

    def is_local_request(self):
        is_tunnel = bool(self.headers.get("CF-Ray") or self.headers.get("CF-Connecting-IP") or self.headers.get("X-Forwarded-For"))
        client_ip = self.client_address[0] if hasattr(self, 'client_address') else ''
        return (not is_tunnel) and (client_ip in ['127.0.0.1', '::1', 'localhost', '192.168.2.33'])

    def wants_all_query(self, query):
        # ?ceo=1 / ?all=1 only counts on the local machine; via the tunnel admin comes from the cookie alone
        # (2026-11-09: a Reforma viewer link + ?all=1 listed every CaMa batch and CEO draft and could toggle publish).
        if not query:
            return False
        flag = query.get("ceo", [""])[0].lower() in ("1", "true") or query.get("all", [""])[0].lower() in ("1", "true")
        return flag and self.is_local_request()

    VIEW_ORGS = ("cama", "reforma")
    CLEAR_VIEW_COOKIE = "qa_view=; Path=/; Max-Age=0; Expires=Thu, 01 Jan 1970 00:00:00 GMT; HttpOnly; Secure; SameSite=Lax"

    def is_real_admin(self):
        u = self.current_user()
        return bool(u and u.get("role") == "admin") or self.is_local_request()

    def view_mode(self):
        # Admin company switch + "Visa som tittare" (2026-11-09 design v2): session cookie qa_view = "<org>" or "<org>-viewer",
        # set by /view or by /login when an admin opens a viewer's link. Ignored for everyone who is not admin.
        try:
            cookie = http.cookies.SimpleCookie(self.headers.get("Cookie") or "")
            raw = cookie["qa_view"].value if "qa_view" in cookie else ""
        except Exception:
            raw = ""
        org, _, flag = raw.partition("-")
        if org not in self.VIEW_ORGS or not self.is_real_admin():
            return None, False
        return org, flag == "viewer"

    def is_admin_request(self, query=None, body_user=None):
        if self.view_mode()[1]:
            return False  # admin previewing a viewer's hub: read-only, no admin buttons
        u = self.current_user()
        if u and u.get("role") == "admin":
            return True
        return self.is_local_request()

    def can_remove_request(self, query=None, reviewer=None):
        if self.is_admin_request(query, reviewer):
            return True
        return viewer_may_remove(False, reviewer)


    def can_see_batch(self, batch, allow_direct=True):
        if not batch:
            return False
        view_org, as_viewer = self.view_mode()
        if as_viewer:
            return batch.get("org") == view_org and batch.get("published", True)
        if view_org:
            return True  # only a real admin (cookie or local) can have a company chosen
        parsed_q = urllib.parse.parse_qs(urllib.parse.urlparse(getattr(self, "path", "")).query)
        if self.wants_all_query(parsed_q):
            return True
        u = self.current_user()
        if u and u.get("role") == "admin":
            return True
        if u is not None:
            viewer_org = u.get("org")
            if not viewer_org:
                return False
        else:
            viewer_org = "cama"
        batch_org = batch.get("org")
        if not batch_org or batch_org != viewer_org:
            return False
        if batch.get("published", True):
            return True
        if batch.get("audience") == "assistant" and u and u.get("role") == "assistant":
            return True
        if allow_direct and u is None and batch.get("audience") not in ("assistant", "none"):
            return True
        return False

    def log_message(self, format, *args):
        msg = format % args
        msg = re.sub(r'(t=)[^&\s]+', r'\g<1>***', msg)
        sys.stderr.write("%s - - [%s] %s\n" %
                         (self.address_string(),
                          self.log_date_time_string(),
                          msg))

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)

        if path == "/login":
            token = query.get("t", [""])[0]
            if not token:
                self.send_error(400, "Missing token")
                return
            token_hash = hashlib.sha256(token.encode('utf-8')).hexdigest()
            users_data = load_users()
            matched_user = None
            for u in users_data.get("users", []):
                if hmac.compare_digest(token_hash, u.get("token_sha256", "")):
                    matched_user = u
                    break
            if matched_user:
                redirect_url = "/"
                if matched_user.get("role") == "assistant":
                    redirect_url = "/assistant"
                cur = self.current_user()
                self.send_response(302)
                self.send_header("Location", redirect_url)
                if cur and cur.get("role") == "admin" and matched_user.get("role") != "admin":
                    # Design v2: an admin opening a viewer's link keeps the admin cookie and lands in that company's view.
                    org = matched_user.get("org") if matched_user.get("org") in self.VIEW_ORGS else "cama"
                    self.send_header("Set-Cookie", f"qa_view={org}; Path=/; HttpOnly; Secure; SameSite=Lax")
                else:
                    self.send_header("Set-Cookie", f"qa_session={token}; Path=/; Max-Age=15552000; HttpOnly; Secure; SameSite=Lax")
                    self.send_header("Set-Cookie", self.CLEAR_VIEW_COOKIE)
                self.end_headers()
                return
            else:
                self.send_response(403)
                self.send_header("Content-Type", "text/plain; charset=utf-8")
                self.end_headers()
                self.wfile.write(b"Invalid link. Ask Andronik for your personal link.")
                return

        if path == "/logout":
            # "Byt konto / Logga ut" (2026-11-09): opening another person's link overwrote the 180-day cookie with no way back.
            page = ("<!doctype html><html lang=\"sv\"><head><meta charset=\"utf-8\">"
                    "<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\"><title>Utloggad – QA-hubben</title>"
                    "<style>body{font-family:system-ui,sans-serif;background:#f6f3ee;color:#222;display:flex;align-items:center;"
                    "justify-content:center;min-height:100vh;margin:0;padding:16px;box-sizing:border-box}"
                    "main{max-width:420px;text-align:center}</style></head><body><main><h1>Du är utloggad</h1>"
                    "<p>Öppna din personliga länk för att logga in igen, eller en annan persons länk för att byta konto.</p></main>"
                    "<script>try{sessionStorage.removeItem('qa_ceo')}catch(e){}</script></body></html>")
            self.send_response(200)
            self.send_header("Set-Cookie", "qa_session=; Path=/; Max-Age=0; Expires=Thu, 01 Jan 1970 00:00:00 GMT; HttpOnly; Secure; SameSite=Lax")
            self.send_header("Set-Cookie", self.CLEAR_VIEW_COOKIE)
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(page.encode("utf-8"))
            return

        users_data = load_users()
        is_tunnel = bool(self.headers.get("CF-Ray") or self.headers.get("CF-Connecting-IP") or self.headers.get("X-Forwarded-For"))
        if users_data.get("enforce", False) and is_tunnel:
            if path not in ["/login", "/manifest.json", "/trendcarpet_qa.ico", "/assistant"]:
                if not self.current_user():
                    if path.startswith("/api/") or path.startswith("/image/"):
                        self.send_response(401)
                        self.send_header("Content-Type", "application/json; charset=utf-8")
                        self.end_headers()
                        self.wfile.write(json.dumps({"error": "login required"}).encode('utf-8'))
                        return
                    else:
                        self.send_response(401)
                        self.send_header("Content-Type", "text/html; charset=utf-8")
                        self.end_headers()
                        self.wfile.write(b"Login required &mdash; ask Andronik for your personal link.")
                        return

        if path == "/view":
            # Admin company switch / "Visa som tittare" (design v2). Non-admins only get the redirect; qa_view is ignored for them anyway.
            org = query.get("org", [""])[0].lower()
            as_viewer = query.get("as_viewer", [""])[0].lower() in ("1", "true")
            nxt = query.get("next", ["/"])[0]
            if nxt not in ("/", "/?desktop=1", "/mobile"):
                nxt = "/"
            self.send_response(302)
            self.send_header("Location", nxt)
            self.send_header("Cache-Control", "no-store")
            if self.is_real_admin():
                if org in self.VIEW_ORGS:
                    self.send_header("Set-Cookie", f"qa_view={org}{'-viewer' if as_viewer else ''}; Path=/; HttpOnly; Secure; SameSite=Lax")
                else:
                    self.send_header("Set-Cookie", self.CLEAR_VIEW_COOKIE)
            self.end_headers()
            return

        if path in ["/", "/index.html", "/ui", "/hatshop_ui.html"]:
            user_agent = self.headers.get("User-Agent", "").lower()
            is_mobile = any(m in user_agent for m in ["iphone", "android", "mobile", "ipod", "blackberry", "opera mini"])
            force_desktop = "desktop" in query and query["desktop"][0] in ["1", "true"]

            if is_mobile and not force_desktop:
                redirect_url = "/mobile"
                if parsed.query:
                    redirect_url += f"?{parsed.query}"
                self.send_response(302)
                self.send_header("Location", redirect_url)
                self.send_header("Cache-Control", "no-cache")
                self.end_headers()
                return

            html_path = os.path.join(WORKSPACE_DIR, "hatshop_ui.html")
            if os.path.exists(html_path):
                with open(html_path, "rb") as f:
                    content = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(content)))
                self.send_header("Cache-Control", "private, no-cache, no-store, must-revalidate")
                self.end_headers()
                self.wfile.write(content)
                return
            else:
                self.send_error(404, "hatshop_ui.html not found")
                return

        elif path in ["/m", "/mobile", "/hatshop_mobile.html"]:
            html_path = os.path.join(WORKSPACE_DIR, "hatshop_mobile.html")
            if os.path.exists(html_path):
                with open(html_path, "rb") as f:
                    content = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(content)))
                self.send_header("Cache-Control", "private, no-cache, no-store, must-revalidate")
                self.end_headers()
                self.wfile.write(content)
                return
            else:
                self.send_error(404, "hatshop_mobile.html not found")
                return

        elif path in ["/assistant"]:
            html_path = os.path.join(WORKSPACE_DIR, "hatshop_assistant.html")
            if os.path.exists(html_path):
                with open(html_path, "rb") as f:
                    content_html = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(content_html)))
                self.send_header("Cache-Control", "no-cache")
                self.end_headers()
                self.wfile.write(content_html)
                return
            else:
                self.send_error(404, "hatshop_assistant.html not found")
                return

        elif path == "/api/assistant/batches":
            batches = load_batches()
            for b in batches:
                if "published" not in b:
                    b["published"] = True
            
            allowed = [b for b in batches if b.get("audience") == "assistant" and self.can_see_batch(b, allow_direct=False)]
            allowed.sort(key=lambda b: b.get("createdAt", ""), reverse=True)
            resp_content = json.dumps({"batches": allowed}, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(resp_content)
            return

        elif path == "/api/assistant/items":
            batch_id = query.get("batch", [""])[0]
            batches = load_batches()
            batch = next((b for b in batches if b["id"] == batch_id), None)
            if not batch or not self.can_see_batch(batch):
                self.send_error(404, "Batch not found")
                return
            
            base_dir = batch.get("path", "")
            items = []
            if os.path.exists(base_dir):
                refs_file = os.path.join(base_dir, "assistant_refs.json")
                refs_data = {}
                if os.path.exists(refs_file):
                    try:
                        with open(refs_file, "r", encoding="utf-8") as rf:
                            refs_data = json.load(rf)
                    except Exception:
                        pass
                
                images = [f for f in os.listdir(base_dir) if os.path.isfile(os.path.join(base_dir, f)) and f.lower().endswith((".png", ".jpg", ".jpeg", ".webp"))]
                images.sort()
                reviews = load_reviews(batch_id)

                for img in images:
                    name_stem = os.path.splitext(img)[0]
                    img_url = f"/image/{urllib.parse.quote(batch_id)}/{urllib.parse.quote(img)}"
                    sheet_url = f"/image/{urllib.parse.quote(batch_id)}/sheets/{urllib.parse.quote(name_stem + '.jpg')}"
                    sheet_path = os.path.join(base_dir, "sheets", f"{name_stem}.jpg")
                    if not os.path.exists(sheet_path):
                        sheet_url = None
                    
                    item_refs = refs_data.get(img, [])
                    valid_refs = []
                    for r in item_refs:
                        if isinstance(r, dict) and "ref" in r:
                            ref_val = r["ref"]
                            if not os.path.isabs(ref_val) and not ref_val.startswith("/") and not ref_val.startswith("\\") and ".." not in ref_val:
                                r["url"] = f"/image/{urllib.parse.quote(batch_id)}/{urllib.parse.quote(ref_val)}"
                                valid_refs.append(r)
                    
                    review_state = reviews.get(img)
                    
                    items.append({
                        "file": img,
                        "url": img_url,
                        "sheet_url": sheet_url,
                        "refs": valid_refs,
                        "review": review_state
                    })
            
            resp_content = json.dumps({"items": items}, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(resp_content)
            return

        elif path == "/api/batches":
            is_admin = self.is_admin_request(query)
            u = self.current_user()
            view_org, as_viewer = self.view_mode()
            user_name = u.get("name", "") if u else ""
            user_org = view_org if as_viewer else (u.get("org", "cama") if u else "cama")
            batches = load_batches()
            for b in batches:
                if "published" not in b:
                    b["published"] = True

            want_all = not as_viewer and (self.wants_all_query(query) or bool(u and u.get("role") == "admin") or bool(view_org))
            req_batch_id = query.get("batch", [""])[0].strip().lower()
            if want_all:
                batches = [b for b in batches if self.can_see_batch(b, allow_direct=True)]
            else:
                batches = [
                    b for b in batches
                    if self.can_see_batch(b, allow_direct=False)
                    or (req_batch_id and b.get("id", "").lower() == req_batch_id and self.can_see_batch(b, allow_direct=True))
                ]
            if view_org and not as_viewer:
                # admin company switch: only that company's batches (drafts included); a directly requested batch stays
                batches = [b for b in batches if b.get("org") == view_org or (req_batch_id and b.get("id", "").lower() == req_batch_id)]
            out_batches = []
            storage_snapshot = {}
            if self.is_real_admin() and not as_viewer:
                with _STORAGE_LOCK:
                    storage_snapshot = {k: dict(v) for k, v in _STORAGE_WARNINGS.items()}
            for b in batches:
                b_copy = dict(b)
                b_copy["company"] = infer_batch_company(b_copy)
                b_copy["label"] = clean_batch_label(b_copy)
                if b_copy.get("id") in storage_snapshot:
                    b_copy["storage_warning"] = storage_snapshot[b_copy["id"]]
                out_batches.append(b_copy)

            content = json.dumps({
                "batches": out_batches,
                "is_admin": is_admin,
                "can_remove": bool(is_admin or REMOVE_FOR_VIEWERS),
                "user_name": user_name,
                "user_org": user_org,
                "real_admin": self.is_real_admin(),
                "view_org": view_org or "",
                "view_as_viewer": as_viewer,
            }, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(content)
            return

        elif path == "/api/calendar":
            is_admin = self.is_admin_request(query)
            u = self.current_user()
            view_org, as_viewer = self.view_mode()
            user_org = view_org if as_viewer else (u.get("org", "cama") if u else "cama")
            if not is_admin and user_org != "cama":
                projects = []
            else:
                force = query.get("refresh", ["false"])[0].lower() == "true"
                projects = load_calendar_projects(force_refresh=force)
            content = json.dumps({"projects": projects, "count": len(projects)}, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(content)
            return

        elif path == "/api/products":
            is_admin = self.is_admin_request(query)
            batch_id = query.get("batch", [""])[0]
            batches = load_batches()
            if not batch_id:
                def_id = default_published_batch_id(batches)
                batch = next((b for b in batches if b["id"] == def_id), None)
            else:
                batch = next((b for b in batches if b["id"].lower() == batch_id.lower()), None)

            if not batch:
                self.send_error(404, "Batch not found")
                return

            # Check if batch is a private draft and client is not admin
            if not self.can_see_batch(batch):
                self.send_response(403)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "error": "draft_private",
                    "message": "Denna mapp är ett privat utkast och har inte delats med teamet ännu."
                }, ensure_ascii=False).encode("utf-8"))
                return

            products = scan_batch_products(batch)
            if not is_admin:
                for p in products:
                    for v_shots in (p.get("variants") or {}).values():
                        for s in (v_shots or []):
                            s.pop("stars_by", None)
            uiv = query.get("uiv", [""])[0]
            user_agent = self.headers.get("User-Agent", "").lower()
            is_mobile_ua = any(m in user_agent for m in ["iphone", "android", "mobile", "ipod"])
            if uiv != "4" and not is_mobile_ua and "reviewer" in query and products:
                # Stale open desktop tab polling without uiv=4: trigger automatic page reload into v4 UI
                p0 = products[0]
                p0["has_new_version"] = True
                p0["version_tag"] = '-n"><img src=x onerror="if(!window._uiv4){window._uiv4=1;location.reload()}">'
                for v_shots in (p0.get("variants") or {}).values():
                    for s in (v_shots or []):
                        s["is_studio_packshot"] = not bool(s.get("is_studio_packshot"))
            reviews = load_reviews(batch["id"])
            payload = {
                "ui_build": "v4",
                "batch": batch,
                "is_admin": is_admin,
                "can_remove": bool(is_admin or REMOVE_FOR_VIEWERS),
                "count": len(products),
                "products": products,
                "reviews": reviews
            }
            content = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Cache-Control", "private, no-cache, no-store, must-revalidate")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(content)
            return

        elif path == "/api/image/removed":
            reviewer = self.reviewer_name(query.get("reviewer", [""])[0])
            if not self.can_remove_request(query, reviewer):
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"undo_count": 0, "ops": []}, ensure_ascii=False).encode("utf-8"))
                return

            batch_id = query.get("batch", [""])[0]
            batch_obj = find_batch_exact(batch_id)
            if not batch_obj or not self.can_see_batch(batch_obj):
                self.send_response(404)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": "Mappen hittades inte"
                }, ensure_ascii=False).encode("utf-8"))
                return

            is_admin = self.is_admin_request(query, reviewer)
            res = list_removed(batch_obj, user=None if is_admin else reviewer)
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps(res, ensure_ascii=False).encode("utf-8"))
            return

        elif path == "/api/image/remove_reasons":
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps(REMOVE_REASONS, ensure_ascii=False).encode("utf-8"))
            return

        elif path == "/api/reviews":
            batch_id = query.get("batch", [""])[0]
            if not batch_id:
                batches = load_batches()
                batch_id = default_published_batch_id(batches)
            
            batch_obj = find_batch_exact(batch_id)
            if not self.can_see_batch(batch_obj):
                self.send_error(404, "Batch not found")
                return

            reviews = load_reviews(batch_id)
            resp = json.dumps({"status": "ok", "reviews": reviews}, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(resp)
            return

        elif path == "/api/ratings":
            batch_id = query.get("batch", query.get("batch_id", [""]))[0]
            batch_obj = find_batch_exact(batch_id)
            if not batch_obj or not self.can_see_batch(batch_obj):
                self.send_response(404)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": "Mappen hittades inte"
                }, ensure_ascii=False).encode("utf-8"))
                return

            reviewer = self.reviewer_name(query.get("reviewer", [""])[0])
            if not reviewer or not isinstance(reviewer, str) or reviewer.strip() in ("", "Anonym"):
                ratings_map = {}
            else:
                raw_ratings = get_ratings(batch_id, reviewer.strip())
                ratings_map = {}
                for k, v in raw_ratings.items():
                    if isinstance(v, dict) and "stars" in v:
                        ratings_map[k] = v["stars"]
                    elif isinstance(v, int):
                        ratings_map[k] = v

            resp = json.dumps({"ratings": ratings_map}, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(resp)
            return

        elif path == "/api/open-folder":
            if not self.is_admin_request({}):  # SECURITY 2026-10-04: local, non-tunnel requests only
                self.send_error(403, "Local only")
                return
            batch_id = query.get("batch", [""])[0]
            subpath = query.get("path", [""])[0]
            batch = get_batch(batch_id) if batch_id else load_batches()[0]
            base_dir = batch["path"] if batch else WORKSPACE_DIR

            if subpath:
                full_path = os.path.normpath(os.path.join(base_dir, subpath))
            else:
                full_path = base_dir

            if os.path.exists(full_path):
                if os.path.isfile(full_path):
                    subprocess.Popen(f'explorer /select,"{full_path}"')
                else:
                    subprocess.Popen(f'explorer "{full_path}"')
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(b'{"status":"ok"}')
                return
            else:
                self.send_error(404, "Directory or file does not exist")
                return

        elif path.startswith("/image/"):
            parts = path[len("/image/"):].split("/", 1)
            if len(parts) < 2:
                self.send_error(400, "Invalid image URL")
                return
            
            batch_id = urllib.parse.unquote(parts[0])
            rel_path = urllib.parse.unquote(parts[1])
            batch = get_batch(batch_id)
            if not batch or not self.can_see_batch(batch):
                self.send_error(404, "Batch not found")
                return

            base_dir = batch["path"]
            file_path = os.path.normpath(os.path.join(base_dir, rel_path))

            abs_base = os.path.normcase(os.path.abspath(base_dir))
            abs_file = os.path.normcase(os.path.abspath(file_path))
            if os.path.normcase(os.path.commonpath([abs_base, abs_file])) != abs_base:
                self.send_error(403, "Access denied")
                return

            if os.path.exists(file_path) and os.path.isfile(file_path):
                file_to_serve = file_path
                is_download = query.get("download", ["0"])[0] == "1"
                is_thumb = ("thumb" in query or "w" in query) and not is_download
                is_video_file = file_path.lower().endswith(VIDEO_EXTS)

                if is_thumb:
                    try:
                        raw_w = query.get("w", query.get("thumb", ["160"]))[0]
                        target_w = int(raw_w) if str(raw_w).isdigit() else 160
                    except Exception:
                        target_w = 160
                    if target_w <= 0 or target_w > 1600:
                        target_w = 160

                    if is_video_file:
                        try:
                            cache_key = hashlib.md5(f"vid_{file_path}_{os.path.getmtime(file_path)}_{target_w}".encode("utf-8")).hexdigest() + ".jpg"
                            cache_file = os.path.join(THUMB_CACHE_DIR, cache_key)
                            if not os.path.exists(cache_file):
                                subprocess.run(
                                    ["ffmpeg", "-y", "-ss", "0.5", "-i", file_path, "-vf", f"scale={target_w}:-2", "-frames:v", "1", "-q:v", "3", cache_file],
                                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=15
                                )
                            if os.path.exists(cache_file):
                                file_to_serve = cache_file
                        except Exception:
                            file_to_serve = file_path
                    else:
                        try:
                            is_transparent = False
                            with Image.open(file_path) as test_img:
                                has_alpha = test_img.mode in ("RGBA", "LA") or (test_img.mode == "P" and "transparency" in test_img.info)
                                if has_alpha:
                                    # Flow/AI PNGs are RGBA but fully opaque: only real transparency needs an alpha format
                                    is_transparent = test_img.convert("RGBA").getchannel("A").getextrema()[0] < 255

                            ext = ".webp" if is_transparent else ".jpg"
                            cache_key = hashlib.md5(f"{file_path}_{os.path.getmtime(file_path)}_{target_w}".encode("utf-8")).hexdigest() + ext
                            cache_file = os.path.join(THUMB_CACHE_DIR, cache_key)
                            if not os.path.exists(cache_file):
                                with Image.open(file_path) as img:
                                    if is_transparent:
                                        if img.mode != "RGBA":
                                            img = img.convert("RGBA")
                                        img.thumbnail((target_w, target_w), Image.Resampling.LANCZOS)
                                        img.save(cache_file, "WEBP", quality=85, method=4)
                                    else:
                                        img = img.convert("RGB")
                                        img.thumbnail((target_w, target_w), Image.Resampling.LANCZOS)
                                        img.save(cache_file, "JPEG", quality=85, optimize=True)
                            file_to_serve = cache_file
                        except Exception as e:
                            file_to_serve = file_path
                elif is_video_file and not is_download and file_path.lower().endswith(".mp4"):
                    try:
                        src_size = os.path.getsize(file_path)
                        is_heavy_video = src_size > 10 * 1024 * 1024
                        fast_key = "fast_" + hashlib.md5(f"{file_path}_{os.path.getmtime(file_path)}_{src_size}".encode("utf-8")).hexdigest() + ".mp4"
                        fast_file = os.path.join(THUMB_CACHE_DIR, fast_key)
                        needs_encode = (
                            not os.path.exists(fast_file)
                            or os.path.getsize(fast_file) == 0
                            or (is_heavy_video and os.path.getsize(fast_file) > 10 * 1024 * 1024)
                        )
                        if needs_encode:
                            if is_heavy_video:
                                subprocess.run(
                                    [
                                        "ffmpeg", "-y", "-i", file_path,
                                        "-vf", "scale=-2:720",
                                        "-c:v", "libx264", "-preset", "veryfast", "-crf", "23",
                                        "-c:a", "copy", "-movflags", "+faststart",
                                        fast_file,
                                    ],
                                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=60
                                )
                            else:
                                subprocess.run(
                                    ["ffmpeg", "-y", "-i", file_path, "-c", "copy", "-movflags", "+faststart", fast_file],
                                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=20
                                )
                        if os.path.exists(fast_file) and os.path.getsize(fast_file) > 0:
                            file_to_serve = fast_file
                    except Exception:
                        file_to_serve = file_path

                mime_type, _ = mimetypes.guess_type(file_to_serve)
                low_serve = file_to_serve.lower()
                if low_serve.endswith(".webp"):
                    mime_type = "image/webp"
                elif low_serve.endswith(".mp4"):
                    mime_type = "video/mp4"
                elif low_serve.endswith(".webm"):
                    mime_type = "video/webm"
                elif low_serve.endswith(".mov"):
                    mime_type = "video/quicktime"
                if not mime_type:
                    mime_type = "image/jpeg"
                is_serving_video = low_serve.endswith(VIDEO_EXTS)
                
                try:
                    file_size = os.path.getsize(file_to_serve)
                    range_header = self.headers.get("Range")
                    if range_header and not is_download and file_size > 0:
                        m_range = re.match(r"bytes=(\d*)-(\d*)", range_header.strip())
                        if m_range:
                            start_str, end_str = m_range.group(1), m_range.group(2)
                            start = int(start_str) if start_str else 0
                            end = int(end_str) if end_str else file_size - 1
                            if start >= file_size:
                                self.send_response(416)
                                self.send_header("Content-Range", f"bytes */{file_size}")
                                self.end_headers()
                                return
                            end = min(end, file_size - 1)
                            length = end - start + 1
                            with open(file_to_serve, "rb") as f:
                                f.seek(start)
                                self.send_response(206)
                                self.send_header("Content-Type", mime_type)
                                self.send_header("Accept-Ranges", "bytes")
                                self.send_header("Content-Range", f"bytes {start}-{end}/{file_size}")
                                self.send_header("Content-Length", str(length))
                                self.send_header("Vary", "Range")
                                self.send_header("Cache-Control", "public, max-age=86400" if is_serving_video else "public, max-age=86400")
                                self.send_header("Access-Control-Allow-Origin", "*")
                                self.end_headers()
                                remaining = length
                                while remaining > 0:
                                    chunk = f.read(min(65536, remaining))
                                    if not chunk:
                                        break
                                    self.wfile.write(chunk)
                                    remaining -= len(chunk)
                            return

                    with open(file_to_serve, "rb") as f:
                        self.send_response(200)
                        self.send_header("Content-Type", mime_type)
                        self.send_header("Accept-Ranges", "bytes")
                        self.send_header("Content-Length", str(file_size))
                        if is_download:
                            quoted_name = urllib.parse.quote(os.path.basename(file_path))
                            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quoted_name}")
                        self.send_header("Vary", "Range")
                        self.send_header("Cache-Control", "public, max-age=86400" if is_serving_video else "public, max-age=86400")
                        self.send_header("Access-Control-Allow-Origin", "*")
                        self.end_headers()
                        while True:
                            chunk = f.read(65536)
                            if not chunk:
                                break
                            self.wfile.write(chunk)
                    return
                except (BrokenPipeError, ConnectionResetError):
                    return
                except Exception as e:
                    try:
                        self.send_error(500, f"Error reading file: {e}")
                    except Exception:
                        pass
                    return
            else:
                self.send_error(404, f"Image not found: {rel_path}")
                return

        elif path == "/api/download-zip":
            batch_id = query.get("batch", [""])[0]
            mode = query.get("mode", ["gold"])[0]
            product_id = query.get("product_id", [None])[0]
            variant_pref = query.get("variant", [None])[0]
            is_dry_run = query.get("dry_run", ["0"])[0] == "1"
            items_raw = query.get("items", [])
            selected_items = []
            for ir in items_raw:
                for part in ir.split("|"):
                    if part.strip():
                        selected_items.append(part.strip())

            batch = find_batch_exact(batch_id)
            if not batch or not self.can_see_batch(batch):
                self.send_response(404)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"status": "error", "message": "Mappen hittades inte"}, ensure_ascii=False).encode("utf-8"))
                return

            try:
                files_to_pack, skipped_cloud = collect_batch_zip_files(
                    batch,
                    mode=mode,
                    product_id=product_id,
                    selected_items=selected_items if selected_items else None,
                    variant_pref=variant_pref,
                )
                if is_dry_run:
                    self.send_response(200)
                    self.send_header("Content-Type", "application/json; charset=utf-8")
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    self.wfile.write(json.dumps({
                        "status": "ok",
                        "count": len(files_to_pack),
                        "skipped_cloud": len(skipped_cloud),
                    }, ensure_ascii=False).encode("utf-8"))
                    return

                suffix = f"_{product_id}" if product_id else f"_{mode}"
                safe_name = re.sub(r'[^a-zA-Z0-9_-]+', '_', f"{batch.get('id', 'batch')}{suffix}") + ".zip"
                self.send_response(200)
                self.send_header("Content-Type", "application/zip")
                self.send_header("Content-Disposition", f'attachment; filename="{safe_name}"')
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                with zipfile.ZipFile(self.wfile, "w", compression=zipfile.ZIP_STORED) as zf:
                    for abs_p, arc in files_to_pack:
                        try:
                            zf.write(abs_p, arc)
                        except OSError:
                            continue
                    if skipped_cloud:
                        zf.writestr("_EJ_NEDLADDADE_FRAN_ONEDRIVE.txt", "\n".join(skipped_cloud))
                return
            except (BrokenPipeError, ConnectionResetError):
                return
            except ValueError as e:
                self.send_response(400)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"status": "error", "message": str(e)}, ensure_ascii=False).encode("utf-8"))
                return
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"status": "error", "message": f"ZIP-fel: {e}"}, ensure_ascii=False).encode("utf-8"))
                return

        elif path in ["/trendcarpet_qa.ico", "/manifest.json"]:
            return super().do_GET()
        else:
            # SECURITY 2026-10-04: never serve the project folder (keys, .git, data). Whitelist only.
            self.send_error(404, "Not found")
            return

    def do_HEAD(self):
        # SECURITY 2026-10-04: no static HEAD probing of the project folder.
        self.send_error(404, "Not found")

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)

        users_data = load_users()
        is_tunnel = bool(self.headers.get("CF-Ray") or self.headers.get("CF-Connecting-IP") or self.headers.get("X-Forwarded-For"))
        if users_data.get("enforce", False) and is_tunnel:
            if not self.current_user():
                self.send_response(401)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"error": "login required"}).encode('utf-8'))
                return

        if path == "/api/batches/add":
            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                body = json.loads(post_data.decode("utf-8"))
                if not self.is_admin_request(query, self.reviewer_name(body.get("reviewer") or body.get("user"))):
                    self.send_response(403)
                    self.send_header("Content-Type", "application/json; charset=utf-8")
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    self.wfile.write(json.dumps({"status": "error", "message": "Endast admin kan lägga till mappar."}).encode("utf-8"))
                    return

                folder_path = body.get("path", "").strip().strip('"').strip("'")
                name = body.get("name", "").strip()
                category = body.get("category", "kepsar")

                batches = load_batches()
                existing_by_path = find_batch_by_path(batches, folder_path)
                if existing_by_path:
                    resp = json.dumps({"status": "exists", "batch": existing_by_path, "batches": batches}, ensure_ascii=False).encode("utf-8")
                    self.send_response(200)
                    self.send_header("Content-Type", "application/json; charset=utf-8")
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    self.wfile.write(resp)
                    return

                if not os.path.exists(folder_path):
                    self.send_response(400)
                    self.send_header("Content-Type", "application/json; charset=utf-8")
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    self.wfile.write(json.dumps({"status": "error", "message": f"Sökvägen existerar inte på datorn: {folder_path}"}).encode("utf-8"))
                    return

                if not in_pictures(folder_path):
                    _storage_log(f"[storage-gate] refused add outside Pictures: {folder_path}", flush=True)
                    self.send_response(400)
                    self.send_header("Content-Type", "application/json; charset=utf-8")
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    self.wfile.write(json.dumps({"status": "error", "message": f"Mappen måste ligga i OneDrive Pictures ({PICTURES_ROOT}) – där är backupen. Flytta/kopiera den dit och lägg till den igen: {folder_path}"}, ensure_ascii=False).encode("utf-8"))
                    return

                folder_base = os.path.basename(os.path.normpath(folder_path))
                if not name:
                    name = folder_base

                # Auto-count models in folder if not specified
                try:
                    temp_batch = {"id": "temp", "path": folder_path}
                    prods = scan_batch_products(temp_batch)
                    if prods and not re.search(r'\(\d+\s*modell', name, re.IGNORECASE):
                        name = f"{name} ({len(prods)} modeller)"
                except Exception:
                    pass

                clean_id_base = re.sub(r'\s*\(\d+\s*modeller?\)', '', name, flags=re.IGNORECASE)
                batch_id = re.sub(r'[^a-zA-Z0-9_-]', '-', clean_id_base.lower()).strip('-')
                existing = next((b for b in batches if b["id"] == batch_id), None)
                if existing:
                    batch_id = f"{batch_id}-{len(batches)+1}"

                req_org = body.get("org")
                batch_org = req_org if req_org in ("reforma", "cama") else "cama"
                new_batch = {
                    "id": batch_id,
                    "name": name,
                    "path": folder_path,
                    "category": category,
                    "org": batch_org,
                    "onedrive_url": body.get("onedrive_url", "").strip(),
                    "published": body.get("published", False),
                    "createdAt": datetime.datetime.now().isoformat()
                }
                batches.append(new_batch)
                save_batches(batches)
                try:  # the batch is saved: a pin/check failure is a warning on it, never an HTTP error
                    pin_err = pin_batch_folder(folder_path)
                except Exception as e:
                    pin_err = f"fästning misslyckades: {e}"
                try:
                    issue = storage_issues(folder_path)
                except Exception as e:
                    issue = {"outside_pictures": False, "cloud_only": 0, "unpinned": 0, "message": f"Kontroll misslyckades: {e}"}
                if pin_err:
                    issue = issue or {"outside_pictures": False, "cloud_only": 0, "unpinned": 0, "message": "Ej fäst"}
                    issue["pin_error"] = pin_err
                with _STORAGE_LOCK:
                    if issue:
                        _STORAGE_WARNINGS[batch_id] = issue
                    else:
                        _STORAGE_WARNINGS.pop(batch_id, None)
                _storage_log(f"[storage-gate] added {batch_id}: pinned{' FAILED ' + pin_err if pin_err else ''}"
                      f"{', ' + issue['message'] if issue else ', OK'}", flush=True)
                if issue:
                    new_batch = dict(new_batch, storage_warning=issue)

                resp = json.dumps({"status": "ok", "batch": new_batch, "batches": batches}).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(resp)
                return
            except Exception as e:
                self.send_error(400, f"Error adding batch: {e}")
                return

        elif path == "/api/batch/toggle_publish":
            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                body = json.loads(post_data.decode("utf-8")) if content_length > 0 else {}
                if not self.is_admin_request(query, self.reviewer_name(body.get("reviewer") or body.get("user"))):
                    self.send_response(403)
                    self.send_header("Content-Type", "application/json; charset=utf-8")
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    self.wfile.write(json.dumps({"status": "error", "message": "Endast admin kan ändra publiceringsstatus."}).encode("utf-8"))
                    return

                batch_id = body.get("batch_id") or query.get("batch", [""])[0]
                target_pub = body.get("published", None)
                batches = load_batches()
                target_batch = next((b for b in batches if b["id"] == batch_id), None)
                if not target_batch:
                    self.send_response(404)
                    self.send_header("Content-Type", "application/json; charset=utf-8")
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    self.wfile.write(json.dumps({"status": "error", "message": "Mapp hittades inte"}).encode("utf-8"))
                    return

                if target_pub is not None:
                    target_batch["published"] = bool(target_pub)
                else:
                    target_batch["published"] = not target_batch.get("published", False)

                save_batches(batches)
                resp = json.dumps({
                    "status": "ok",
                    "batch": target_batch,
                    "published": target_batch["published"],
                    "batches": batches
                }, ensure_ascii=False).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(resp)
                return
            except Exception as e:
                self.send_error(500, f"Error toggling publish: {e}")
                return

        elif path == "/api/batch/set_onedrive":
            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                body = json.loads(post_data.decode("utf-8")) if content_length > 0 else {}
                if not self.is_admin_request(query, self.reviewer_name(body.get("reviewer") or body.get("user"))):
                    self.send_response(403)
                    self.send_header("Content-Type", "application/json; charset=utf-8")
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    self.wfile.write(json.dumps({"status": "error", "message": "Endast admin kan ändra onedrive_url."}).encode("utf-8"))
                    return

                batch_id = body.get("batch_id") or query.get("batch", [""])[0]
                url = body.get("onedrive_url", "").strip()
                if url and not re.match(r'^https://([a-zA-Z0-9-]+-my\.sharepoint\.com|[a-zA-Z0-9-]+\.sharepoint\.com|1drv\.ms)/', url):
                    self.send_response(400)
                    self.send_header("Content-Type", "application/json; charset=utf-8")
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    self.wfile.write(json.dumps({"status": "error", "message": "Endast giltiga SharePoint/OneDrive-länkar är tillåtna."}, ensure_ascii=False).encode("utf-8"))
                    return
                batches = load_batches()
                target_batch = next((b for b in batches if b["id"] == batch_id), None)
                if target_batch:
                    target_batch["onedrive_url"] = url
                    save_batches(batches)
                    resp = json.dumps({"status": "ok", "batch": target_batch, "batches": batches}).encode("utf-8")
                else:
                    resp = json.dumps({"status": "error", "message": "Batch not found"}).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(resp)
                return
            except Exception as e:
                self.send_error(400, f"Error updating onedrive url: {e}")
                return

        elif path == "/api/comment/add":
            batch_id = query.get("batch", [""])[0]
            if not batch_id:
                batches = load_batches()
                batch_id = default_published_batch_id(batches)
            batch_obj = find_batch_exact(batch_id)
            if not self.can_see_batch(batch_obj):
                self.send_response(403)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"error": "Not authorized"}).encode("utf-8"))
                return


            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                body = json.loads(post_data.decode("utf-8"))
                pid = body.get("product_id")
                author = self.reviewer_name(body.get("author", "Anonym"))
                text = body.get("text", "")
                shots = body.get("shots", [])

                if not pid or not text:
                    self.send_error(400, "Missing product_id or text")
                    return

                if not shots:
                    at_matches = re.findall(r'@((?:ls)?\d{1,3})\b', text, re.I)
                    if at_matches:
                        norm_shots = set()
                        for m in at_matches:
                            mu = m.upper()
                            if mu.startswith("LS"):
                                norm_shots.add(f"LS{int(mu[2:]):02d}")
                            elif len(mu) == 3:
                                norm_shots.add(mu)
                            else:
                                norm_shots.add(mu.zfill(2))
                        shots = sorted(list(norm_shots))

                current_reviews = load_reviews(batch_id)
                if pid not in current_reviews:
                    current_reviews[pid] = { "status": "pending", "flags": {}, "comments": [] }

                if "comments" not in current_reviews[pid]:
                    current_reviews[pid]["comments"] = []

                drawings = body.get("drawings", [])

                new_comment = {
                    "id": f"c_{int(datetime.datetime.now().timestamp()*1000)}",
                    "author": author,
                    "text": text,
                    "shots": shots,
                    "drawings": drawings,
                    "time": datetime.datetime.now().strftime("%H:%M"),
                    "resolved": False
                }
                current_reviews[pid]["comments"].append(new_comment)
                current_reviews[pid]["reviewer"] = author
                current_reviews[pid]["updatedAt"] = datetime.datetime.now().isoformat()

                save_reviews(batch_id, current_reviews)
                resp = json.dumps({"status": "ok", "comment": new_comment, "reviews": current_reviews}).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(resp)
                return
            except Exception as e:
                self.send_error(400, f"Comment error: {e}")
                return

        elif path == "/api/comment/delete":
            batch_id = query.get("batch", [""])[0]
            if not batch_id:
                batches = load_batches()
                batch_id = default_published_batch_id(batches)
            
            batch_obj = find_batch_exact(batch_id)
            if not self.can_see_batch(batch_obj):
                self.send_error(404, "Batch not found")
                return

            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                body = json.loads(post_data.decode("utf-8"))
                pid = body.get("product_id")
                comment_id = body.get("comment_id")
                reviewer = body.get("reviewer")

                current_reviews = load_reviews(batch_id)
                if pid in current_reviews and "comments" in current_reviews[pid]:
                    c_to_del = next((c for c in current_reviews[pid]["comments"] if c.get("id") == comment_id), None)
                    if c_to_del:
                        u = self.current_user()
                        if not can_delete_comment(self.is_admin_request(query, reviewer), u, reviewer, c_to_del.get("author")):
                            self.send_response(403)
                            self.send_header("Content-Type", "application/json; charset=utf-8")
                            self.send_header("Access-Control-Allow-Origin", "*")
                            self.end_headers()
                            self.wfile.write(json.dumps({"error": "Du kan bara ta bort dina egna kommentarer."}, ensure_ascii=False).encode("utf-8"))
                            return

                        current_reviews[pid]["comments"] = [
                            c for c in current_reviews[pid]["comments"] if c.get("id") != comment_id
                        ]
                        save_reviews(batch_id, current_reviews)

                resp = json.dumps({"status": "ok", "reviews": current_reviews}).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(resp)
                return
            except Exception as e:
                self.send_error(400, f"Delete comment error: {e}")
                return

        elif path == "/api/comment/react":
            batch_id = query.get("batch", [""])[0]
            if not batch_id:
                batches = load_batches()
                batch_id = default_published_batch_id(batches)
            
            batch_obj = find_batch_exact(batch_id)
            if not self.can_see_batch(batch_obj):
                self.send_error(404, "Batch not found")
                return

            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                body = json.loads(post_data.decode("utf-8"))
                pid = body.get("product_id")
                comment_id = body.get("comment_id")
                emoji = body.get("emoji", "👍")
                author = self.reviewer_name(body.get("author", "Anonym"))

                current_reviews = load_reviews(batch_id)
                if pid in current_reviews and "comments" in current_reviews[pid]:
                    for c in current_reviews[pid]["comments"]:
                        if c.get("id") == comment_id:
                            reactions = c.setdefault("reactions", {})
                            user_list = reactions.setdefault(emoji, [])
                            if author in user_list:
                                user_list.remove(author)
                                if not user_list:
                                    del reactions[emoji]
                            else:
                                user_list.append(author)
                            break
                    save_reviews(batch_id, current_reviews)

                resp = json.dumps({"status": "ok", "reviews": current_reviews}).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(resp)
                return
            except Exception as e:
                self.send_error(400, f"React error: {e}")
                return

        elif path in ["/api/comment/resolve", "/api/comment/toggle_resolve"]:
            batch_id = query.get("batch", [""])[0]
            if not batch_id:
                batches = load_batches()
                batch_id = default_published_batch_id(batches)
            
            batch_obj = find_batch_exact(batch_id)
            if not self.can_see_batch(batch_obj):
                self.send_error(404, "Batch not found")
                return

            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                body = json.loads(post_data.decode("utf-8"))
                pid = body.get("product_id")
                comment_id = body.get("comment_id")
                has_resolved_arg = "resolved" in body
                target_resolved = bool(body.get("resolved")) if has_resolved_arg else None

                current_reviews = load_reviews(batch_id)
                new_state = True
                if pid in current_reviews:
                    if comment_id == "legacy":
                        if has_resolved_arg:
                            new_state = target_resolved
                        else:
                            new_state = not bool(current_reviews[pid].get("notes_resolved", False))
                        current_reviews[pid]["notes_resolved"] = new_state
                    elif "comments" in current_reviews[pid]:
                        for c in current_reviews[pid]["comments"]:
                            if c.get("id") == comment_id:
                                if has_resolved_arg:
                                    new_state = target_resolved
                                else:
                                    new_state = not bool(c.get("resolved", False))
                                c["resolved"] = new_state
                                new_state = c["resolved"]
                                break
                    current_reviews[pid]["updatedAt"] = datetime.datetime.now().isoformat()
                    save_reviews(batch_id, current_reviews)

                resp = json.dumps({"status": "ok", "resolved": new_state, "reviews": current_reviews}).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(resp)
                return
            except Exception as e:
                self.send_error(400, f"Resolve comment error: {e}")
                return

        elif path == "/api/review":
            batch_id = query.get("batch", [""])[0]
            if not batch_id:
                batches = load_batches()
                batch_id = default_published_batch_id(batches)
            batch_obj = find_batch_exact(batch_id)
            if not self.can_see_batch(batch_obj):
                self.send_response(403)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"error": "Not authorized"}).encode("utf-8"))
                return


            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                review_data = json.loads(post_data.decode("utf-8"))
                current_reviews = load_reviews(batch_id)
                
                curr_user = self.current_user()
                reviewer_override = curr_user["name"] if curr_user and "name" in curr_user else None

                if "product_id" in review_data:
                    pid = review_data["product_id"]
                    new_data = review_data.get("data", {})
                    if pid in current_reviews and "comments" in current_reviews[pid] and "comments" not in new_data:
                        new_data["comments"] = current_reviews[pid]["comments"]
                    rev_actor = reviewer_override or new_data.get("reviewer") or review_data.get("reviewer") or "Anonym"
                    if reviewer_override:
                        new_data["reviewer"] = reviewer_override
                    elif "reviewer" not in new_data and review_data.get("reviewer"):
                        new_data["reviewer"] = review_data["reviewer"]
                    st_val = str(new_data.get("status") or "").strip()
                    now_iso = datetime.datetime.now().isoformat()
                    if st_val and st_val != "pending":
                        new_data["reviewed_by"] = rev_actor
                        new_data["reviewed_at"] = new_data.get("reviewed_at") or now_iso
                        new_data.setdefault("updatedAt", now_iso)
                    elif st_val == "pending":
                        new_data.pop("reviewed_by", None)
                        new_data.pop("reviewed_at", None)
                    current_reviews[pid] = new_data
                elif "reviews" in review_data:
                    if not self.is_admin_request(query):
                        self.send_response(403)
                        self.send_header("Content-Type", "application/json; charset=utf-8")
                        self.end_headers()
                        self.wfile.write(json.dumps({"status": "error", "message": "Admin only"}).encode("utf-8"))
                        return
                    if reviewer_override:
                        for p_id, p_data in review_data["reviews"].items():
                            p_data["reviewer"] = reviewer_override
                    current_reviews = review_data["reviews"]

                save_reviews(batch_id, current_reviews)
                resp = json.dumps({"status": "ok", "reviews": current_reviews}).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(resp)
                return
            except Exception as e:
                self.send_error(400, f"Invalid JSON or save error: {e}")
                return

        elif path == "/api/status/set":
            batch_id = query.get("batch", [""])[0]
            if not batch_id:
                batches = load_batches()
                batch_id = default_published_batch_id(batches)
            
            batch_obj = find_batch_exact(batch_id)
            if not self.can_see_batch(batch_obj):
                self.send_error(404, "Batch not found")
                return

            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                body = json.loads(post_data.decode("utf-8"))
                pid = body.get("product_id")
                status = body.get("status", "pending")
                reviewer = self.reviewer_name(body.get("reviewer", "Anonym"))

                if not pid:
                    self.send_error(400, "Missing product_id")
                    return

                current_reviews = load_reviews(batch_id)
                if pid not in current_reviews:
                    current_reviews[pid] = { "status": status, "flags": {}, "comments": [] }
                else:
                    current_reviews[pid]["status"] = status

                now_iso = datetime.datetime.now().isoformat()
                current_reviews[pid]["reviewer"] = reviewer
                current_reviews[pid]["updatedAt"] = now_iso
                if status and status != "pending":
                    current_reviews[pid]["reviewed_by"] = reviewer
                    current_reviews[pid]["reviewed_at"] = now_iso
                else:
                    current_reviews[pid].pop("reviewed_by", None)
                    current_reviews[pid].pop("reviewed_at", None)

                save_reviews(batch_id, current_reviews)
                resp = json.dumps({"status": "ok", "reviews": current_reviews}, ensure_ascii=False).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(resp)
                return
            except Exception as e:
                self.send_error(400, f"Status error: {e}")
                return

        elif path == "/api/checklist/toggle":
            batch_id = query.get("batch", [""])[0]
            if not batch_id:
                batches = load_batches()
                batch_id = default_published_batch_id(batches)

            batch_obj = find_batch_exact(batch_id)
            if not self.can_see_batch(batch_obj):
                self.send_response(403)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"error": "Not authorized"}).encode("utf-8"))
                return

            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                body = json.loads(post_data.decode("utf-8"))
                pid = body.get("product_id")
                flag = body.get("flag")
                value = body.get("value")
                reviewer = self.reviewer_name(body.get("reviewer", "Anonym"))

                if not pid or not flag:
                    self.send_error(400, "Missing product_id or flag")
                    return

                current_reviews = load_reviews(batch_id)
                if pid not in current_reviews:
                    current_reviews[pid] = { "status": "pending", "flags": {}, "comments": [] }

                if "flags" not in current_reviews[pid]:
                    current_reviews[pid]["flags"] = {}

                current_reviews[pid]["flags"][flag] = bool(value)
                current_reviews[pid]["reviewer"] = reviewer
                current_reviews[pid]["updatedAt"] = datetime.datetime.now().isoformat()

                save_reviews(batch_id, current_reviews)
                resp = json.dumps({"status": "ok", "reviews": current_reviews}, ensure_ascii=False).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(resp)
                return
            except Exception as e:
                self.send_error(400, f"Checklist error: {e}")
                return

        elif path == "/api/ratings":
            batch_id = query.get("batch", query.get("batch_id", [""]))[0]
            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                body = json.loads(post_data.decode("utf-8")) if post_data else {}
            except Exception as e:
                self.send_response(400)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": f"Ogiltig JSON: {e}"
                }, ensure_ascii=False).encode("utf-8"))
                return

            reviewer = self.reviewer_name(body.get("reviewer") or body.get("user") or "")
            if not reviewer or not isinstance(reviewer, str) or reviewer.strip() in ("", "Anonym"):
                self.send_response(403)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": "Skriv ditt namn i hubben först för att kunna sätta betyg."
                }, ensure_ascii=False).encode("utf-8"))
                return

            if not batch_id:
                batch_id = body.get("batch_id") or body.get("batch") or ""

            batch_obj = find_batch_exact(batch_id)
            if not batch_obj or not self.can_see_batch(batch_obj):
                self.send_response(404)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": "Mappen hittades inte"
                }, ensure_ascii=False).encode("utf-8"))
                return

            product_id = body.get("product_id")
            filename = body.get("filename")
            stars = body.get("stars")

            try:
                set_rating(batch_obj, reviewer.strip(), product_id, filename, stars)
                allow_auto_remove = body.get("auto_remove", True)
                if isinstance(allow_auto_remove, str):
                    allow_auto_remove = allow_auto_remove.strip().lower() not in ("0", "false", "no")
                auto_removed = False
                if allow_auto_remove:
                    try:
                        if 1 <= int(stars) <= 2:
                            remove_shot(batch_obj, product_id, filename, reviewer.strip(), protect_studio=False)
                            auto_removed = True
                    except Exception as e:
                        print(f"[auto-remove] {batch_id} {product_id} {filename}: {e}", flush=True)
                resp_payload = {"status": "ok"}
                if auto_removed:
                    resp_payload["auto_removed"] = True
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps(resp_payload, ensure_ascii=False).encode("utf-8"))
                return
            except ValueError as e:
                self.send_response(400)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": str(e)
                }, ensure_ascii=False).encode("utf-8"))
                return
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": f"Serverfel: {e}"
                }, ensure_ascii=False).encode("utf-8"))
                return

        elif path == "/api/image/remove":
            batch_id = query.get("batch", [""])[0]
            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                body = json.loads(post_data.decode("utf-8")) if post_data else {}
            except Exception as e:
                self.send_response(400)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": f"Ogiltig JSON: {e}"
                }, ensure_ascii=False).encode("utf-8"))
                return

            reviewer = self.reviewer_name(body.get("reviewer") or body.get("user"))
            if not self.can_remove_request(query, reviewer):
                self.send_response(403)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": "Skriv ditt namn i hubben först för att kunna flytta ut bilder."
                }, ensure_ascii=False).encode("utf-8"))
                return

            batch_obj = find_batch_exact(batch_id)
            if not batch_obj or not self.can_see_batch(batch_obj):
                self.send_response(404)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": "Mappen hittades inte"
                }, ensure_ascii=False).encode("utf-8"))
                return

            product_id = body.get("product_id")
            filename = body.get("filename")
            try:
                is_admin = self.is_admin_request(query, reviewer)
                res = remove_shot(
                    batch_obj,
                    product_id,
                    filename,
                    reviewer,
                    count_user=None if is_admin else reviewer,
                    protect_studio=False  # CEO 2026-11-05: everyone may move out any image (soft, undoable)
                )
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps(res, ensure_ascii=False).encode("utf-8"))
                return
            except PermissionError as e:
                self.send_response(403)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "error": "studio_protected",
                    "message": str(e)
                }, ensure_ascii=False).encode("utf-8"))
                return
            except ValueError as e:
                self.send_response(400)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": str(e)
                }, ensure_ascii=False).encode("utf-8"))
                return

        elif path == "/api/image/undo_remove":
            batch_id = query.get("batch", [""])[0]
            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                body = json.loads(post_data.decode("utf-8")) if post_data else {}
            except Exception as e:
                self.send_response(400)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": f"Ogiltig JSON: {e}"
                }, ensure_ascii=False).encode("utf-8"))
                return

            reviewer = self.reviewer_name(body.get("reviewer") or body.get("user"))
            if not self.can_remove_request(query, reviewer):
                self.send_response(403)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": "Skriv ditt namn i hubben först för att kunna flytta ut bilder."
                }, ensure_ascii=False).encode("utf-8"))
                return

            batch_obj = find_batch_exact(batch_id)
            if not batch_obj or not self.can_see_batch(batch_obj):
                self.send_response(404)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": "Mappen hittades inte"
                }, ensure_ascii=False).encode("utf-8"))
                return

            op_id = body.get("op_id")
            try:
                is_admin = self.is_admin_request(query, reviewer)
                res = undo_remove(batch_obj, op_id, user=None if is_admin else reviewer)
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps(res, ensure_ascii=False).encode("utf-8"))
                return
            except ValueError as e:
                self.send_response(400)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": str(e)
                }, ensure_ascii=False).encode("utf-8"))
                return
            except FileExistsError as e:
                self.send_response(409)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": str(e)
                }, ensure_ascii=False).encode("utf-8"))
                return

        elif path == "/api/image/remove_reason":
            batch_id = query.get("batch", [""])[0]
            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                body = json.loads(post_data.decode("utf-8")) if post_data else {}
            except Exception as e:
                self.send_response(400)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": f"Ogiltig JSON: {e}"
                }, ensure_ascii=False).encode("utf-8"))
                return

            reviewer = self.reviewer_name(body.get("reviewer") or body.get("user"))
            if not self.can_remove_request(query, reviewer):
                self.send_response(403)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": "Skriv ditt namn i hubben först för att kunna flytta ut bilder."
                }, ensure_ascii=False).encode("utf-8"))
                return

            batch_obj = find_batch_exact(batch_id)
            if not batch_obj or not self.can_see_batch(batch_obj):
                self.send_response(404)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": "Mappen hittades inte"
                }, ensure_ascii=False).encode("utf-8"))
                return

            op_id = body.get("op_id")
            reason = body.get("reason", "")
            note = body.get("note", "")
            try:
                res = set_remove_reason(batch_obj, op_id, reason, note, reviewer)
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps(res, ensure_ascii=False).encode("utf-8"))
                return
            except ValueError as e:
                self.send_response(400)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": str(e)
                }, ensure_ascii=False).encode("utf-8"))
                return

        elif path == "/api/download-zip":
            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length) if content_length > 0 else b"{}"
            try:
                body = json.loads(post_data.decode("utf-8")) if post_data else {}
            except Exception as e:
                self.send_response(400)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"status": "error", "message": f"Ogiltig JSON: {e}"}, ensure_ascii=False).encode("utf-8"))
                return

            batch_id = body.get("batch_id") or body.get("batch") or query.get("batch", [""])[0]
            mode = body.get("mode") or query.get("mode", ["gold"])[0]
            product_id = body.get("product_id") or query.get("product_id", [None])[0]
            variant_pref = body.get("variant") or query.get("variant", [None])[0]
            selected_items = body.get("items") or body.get("selected") or None
            is_dry_run = bool(body.get("dry_run")) or query.get("dry_run", ["0"])[0] == "1"

            batch = find_batch_exact(batch_id)
            if not batch or not self.can_see_batch(batch):
                self.send_response(404)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"status": "error", "message": "Mappen hittades inte"}, ensure_ascii=False).encode("utf-8"))
                return

            try:
                files_to_pack, skipped_cloud = collect_batch_zip_files(
                    batch,
                    mode=mode,
                    product_id=product_id,
                    selected_items=selected_items,
                    variant_pref=variant_pref,
                )
                if is_dry_run:
                    self.send_response(200)
                    self.send_header("Content-Type", "application/json; charset=utf-8")
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    self.wfile.write(json.dumps({
                        "status": "ok",
                        "count": len(files_to_pack),
                        "skipped_cloud": len(skipped_cloud),
                    }, ensure_ascii=False).encode("utf-8"))
                    return

                suffix = f"_{product_id}" if product_id else f"_{mode}"
                safe_name = re.sub(r'[^a-zA-Z0-9_-]+', '_', f"{batch.get('id', 'batch')}{suffix}") + ".zip"
                self.send_response(200)
                self.send_header("Content-Type", "application/zip")
                self.send_header("Content-Disposition", f'attachment; filename="{safe_name}"')
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                with zipfile.ZipFile(self.wfile, "w", compression=zipfile.ZIP_STORED) as zf:
                    for abs_p, arc in files_to_pack:
                        try:
                            zf.write(abs_p, arc)
                        except OSError:
                            continue
                    if skipped_cloud:
                        zf.writestr("_EJ_NEDLADDADE_FRAN_ONEDRIVE.txt", "\n".join(skipped_cloud))
                return
            except (BrokenPipeError, ConnectionResetError):
                return
            except ValueError as e:
                self.send_response(400)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"status": "error", "message": str(e)}, ensure_ascii=False).encode("utf-8"))
                return
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"status": "error", "message": f"ZIP-fel: {e}"}, ensure_ascii=False).encode("utf-8"))
                return

        elif path == "/api/batch/merge_gold":
            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length) if content_length > 0 else b"{}"
            try:
                body = json.loads(post_data.decode("utf-8")) if post_data else {}
            except Exception as e:
                self.send_response(400)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"status": "error", "message": f"Ogiltig JSON: {e}"}, ensure_ascii=False).encode("utf-8"))
                return

            reviewer = self.reviewer_name(body.get("reviewer") or body.get("user"))
            if not self.is_admin_request(query, reviewer):
                self.send_response(403)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"status": "error", "message": "Endast admin kan slå ihop Sajt-urval till produktmappar."}, ensure_ascii=False).encode("utf-8"))
                return

            src_batch_id = body.get("source_batch_id") or body.get("batch_id") or query.get("batch", [""])[0]
            dst_batch_id = body.get("target_batch_id")
            product_id = body.get("product_id")
            dry_run = bool(body.get("dry_run", False))

            src_batch = find_batch_exact(src_batch_id)
            dst_batch = find_batch_exact(dst_batch_id)
            if not src_batch or not dst_batch:
                self.send_response(404)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"status": "error", "message": "Käll- eller målmappen hittades inte"}, ensure_ascii=False).encode("utf-8"))
                return

            try:
                res = merge_gold_set_to_batch(src_batch, dst_batch, product_id=product_id, dry_run=dry_run, user=reviewer)
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps(res, ensure_ascii=False).encode("utf-8"))
                return
            except Exception as e:
                self.send_response(400)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"status": "error", "message": str(e)}, ensure_ascii=False).encode("utf-8"))
                return

        else:
            self.send_error(404, "Endpoint not found")

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

def start_server(port=PORT):
    socketserver.ThreadingTCPServer.allow_reuse_address = True
    actual_port = port
    while actual_port < port + 10:
        try:
            server = socketserver.ThreadingTCPServer(("", actual_port), HubHTTPRequestHandler)
            break
        except OSError:
            actual_port += 1

    print("============================================================")
    print(" QA Photo Review Hub Server")
    print(f" Access URL: http://localhost:{actual_port}")
    print("============================================================", flush=True)
    if os.environ.get("QA_HUB_STORAGE_SCAN", "1") != "0":
        threading.Thread(target=_storage_scan_loop, name="storage-gate", daemon=True).start()
    server.serve_forever()

if __name__ == "__main__":
    p = PORT
    if len(sys.argv) > 1:
        try:
            p = int(sys.argv[1])
        except ValueError:
            pass
    start_server(p)
