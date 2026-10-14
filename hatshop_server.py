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
from PIL import Image

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
    norm = os.path.normpath(local_path)
    if norm.lower().startswith(base_od.lower()):
        rel = os.path.relpath(norm, base_od).replace('\\', '/')
        encoded_rel = urllib.parse.quote(f"/personal/andronik_lindgren_trendcarpet_se/Documents/{rel}", safe='')
        return f"https://trendcarpet-my.sharepoint.com/my?id={encoded_rel}"
    return ""

def load_batches():
    batches = []
    if os.path.exists(BATCHES_FILE):
        try:
            with open(BATCHES_FILE, "r", encoding="utf-8") as f:
                batches = json.load(f)
        except Exception:
            pass
    if not batches:
        batches = [
            {
                "id": "26-08-28-hatshop-black-river",
                "name": "2026-08-28 Hatshop Black River (14 modeller)",
                "path": r"C:\Users\AndronikLindgren\OneDrive - CaMa Gruppen AB\Pictures\26-08-28-Hatshop-Black-River pt1-x",
                "category": "kepsar",
                "onedrive_url": "",
                "createdAt": "2026-09-02T11:00:00Z"
            }
        ]
        save_batches(batches)

    # Auto-match with calendar_projects if onedrive_url is missing
    changed = False
    try:
        projects = load_calendar_projects()
        for b in batches:
            if not b.get("onedrive_url"):
                folder = os.path.basename(os.path.normpath(b.get("path", ""))).strip().lower()
                bname = b.get("name", "").lower()
                matched_link = ""
                for p in projects:
                    ptitle = p.get("title", "").strip().lower()
                    if ptitle and p.get("link"):
                        if folder == ptitle or folder in ptitle or ptitle in folder:
                            matched_link = p["link"]
                            break
                if matched_link:
                    b["onedrive_url"] = matched_link
                    changed = True
                else:
                    auto_url = get_onedrive_web_url_for_path(b.get("path", ""))
                    if auto_url:
                        b["onedrive_url"] = auto_url
                        changed = True
        if changed:
            save_batches(batches)
    except Exception:
        pass

    return batches

def save_batches(batches):
    try:
        with open(BATCHES_FILE, "w", encoding="utf-8") as f:
            json.dump(batches, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error saving batches: {e}")
        return False

def get_batch(batch_id):
    batches = load_batches()
    if not batch_id:
        return batches[0] if batches else None
    for b in batches:
        if b["id"].lower() == batch_id.lower():
            return b
    return batches[0] if batches else None

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

def get_review_file(batch_id):
    safe_id = re.sub(r'[^a-zA-Z0-9_\-]', '_', batch_id)
    return os.path.join(REVIEWS_DIR, f"reviews_{safe_id}.json")

def load_reviews(batch_id):
    rf = get_review_file(batch_id)
    if os.path.exists(rf):
        try:
            with open(rf, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    legacy = os.path.join(REVIEWS_DIR, "hatshop_review_status.json")
    if batch_id == "26-08-28-hatshop-black-river" and os.path.exists(legacy):
        try:
            with open(legacy, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_reviews(batch_id, data):
    rf = get_review_file(batch_id)
    try:
        with open(rf, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error saving reviews: {e}")
        return False

def parse_file_details(filename, folder_path):
    """
    Parses a filename to extract base shot number, version suffix, modification time, etc.
    Examples:
      keps-model-color-01.jpg -> base_shot='01', suffix='', is_version_suffix=False
      keps-model-color-01-n.jpg -> base_shot='01', suffix='n', is_version_suffix=True
      matta-model-01-Interior-0002.jpg -> base_shot='01', suffix='Interior-0002', is_version_suffix=False
    """
    name_no_ext, ext = os.path.splitext(filename)
    full_path = os.path.join(folder_path, filename)
    mtime = os.path.getmtime(full_path) if os.path.exists(full_path) else 0
    file_size = os.path.getsize(full_path) if os.path.exists(full_path) else 0

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

    m = re.search(r'[-_](0[1-9]|[1-9][0-9]?)(.*)$', name_no_ext, re.IGNORECASE)
    if m:
        base_shot_num = m.group(1).zfill(2)
        raw_suffix = m.group(2).strip('-_ ')
        clean_suffix = raw_suffix.replace('-', ' ').replace('_', ' ').strip()
        
        # Check if true replacement version (e.g. -n, -ny, -v2, -fix, -2)
        is_interior = bool(re.search(r'interioi?r|miljo|miljö|detail', raw_suffix, re.IGNORECASE))
        is_version_suffix = bool(re.match(r'^(n|ny|new|v[0-9]+|[0-9]+|fix|crop|justerad|redigerad)$', raw_suffix, re.IGNORECASE)) and not is_interior
        
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

    return final_shots

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
            sub = re.sub(r'interioi?r.*$', '', raw_suffix, flags=re.IGNORECASE).strip('-_ ')
            if sub and base_shot != "99":
                label = f"{base_shot}-{sub} (Miljö)"
                shot_code = f"{base_shot}_{sub}_miljo"
            elif base_shot != "99":
                label = f"{base_shot} (Miljö)"
                shot_code = f"{base_shot}_miljo"
            else:
                label = f"Miljö"
                shot_code = f"miljo_{idx+1}"
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
    return final_shots

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

def scan_batch_products(batch):
    base_dir = batch.get("path", "")
    batch_id = batch.get("id", "")
    if not os.path.exists(base_dir):
        return []

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
                        elif "rel_path" in img and os.path.exists(os.path.join(os.path.dirname(base_dir), img["rel_path"])):
                            base_name = os.path.basename(img["rel_path"])
                            rel_file = base_name if os.path.exists(os.path.join(base_dir, base_name)) else fn
                        else:
                            rel_file = fn

                        full_path = os.path.join(base_dir, rel_file)
                        mtime = os.path.getmtime(full_path) if os.path.exists(full_path) else 0
                        kind = img.get("kind", "")
                        pos = img.get("pos", len(shots) + 1)
                        label = f"{pos:02d}" + (f" ({kind})" if kind else "")

                        shots.append({
                            "shot": pos,
                            "label": label,
                            "url": f"/image/{urllib.parse.quote(batch_id)}/{urllib.parse.quote(rel_file)}",
                            "filename": fn,
                            "rel_path": rel_file,
                            "is_primary": (pos == 1),
                            "is_new_version": False,
                            "version_tag": "",
                            "mtime": mtime
                        })

                    v_label = "1600x2000 Transparent" if any(img.get("is_transparent") for img in item.get("images", [])) else "Standard"
                    products.append({
                        "id": pid,
                        "brand": category,
                        "model": f"[{pid}] {title}",
                        "color": clean_name,
                        "folder_name": pid,
                        "has_new_version": False,
                        "version_tag": "",
                        "latest_mtime": max((s["mtime"] for s in shots), default=0),
                        "variants": {
                            v_label: shots
                        }
                    })
                return products
        except Exception as e:
            print(f"[!] Error parsing manifest {manifest_file}: {e}")

    root_entries = [d for d in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, d))]
    has_variant_subdirs = any(v.lower() in ["1500px", "1500px iphone", "original", "web", "print"] for v in root_entries)
    prefix = ""
    variant_subdirs = root_entries
    if not has_variant_subdirs and len(root_entries) == 1:
        direct_images = [f for f in os.listdir(base_dir) if f.lower().endswith((".jpg", ".jpeg", ".png", ".webp"))]
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

                raw_files = [f for f in os.listdir(folder_path) if f.lower().endswith((".jpg", ".jpeg", ".png", ".webp"))]

                if raw_files:
                    # Clean variant suffixes from folder name if present
                    clean_name = re.sub(r'-(?:1500px(?:\s*iphone)?|original|web|print)$', '', folder_name, flags=re.IGNORECASE).strip()
                    if clean_name.lower() == 'frome':
                        clean_name = 'Frome-red'

                    parts = clean_name.split("-")
                    brand = "Brand"
                    model_name = clean_name
                    color_name = ""

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

                    prod_key = clean_name
                    if prod_key not in products_map:
                        products_map[prod_key] = {
                            "id": prod_key,
                            "brand": brand,
                            "model": model_name,
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
                    color_subfolders = []
                    for sd in sub_dirs:
                        sd_path = os.path.join(folder_path, sd)
                        sd_raw = [f for f in os.listdir(sd_path) if f.lower().endswith((".jpg", ".jpeg", ".png", ".webp"))]
                        if sd_raw:
                            color_subfolders.append((sd, sd_path, sd_raw))

                    if color_subfolders:
                        for color_folder, color_path, raw in color_subfolders:
                            prod_id = f"{folder_name}-{color_folder}"
                            color, brand = parse_color_folder(color_folder)
                            if prod_id not in products_map:
                                products_map[prod_id] = {
                                    "id": prod_id,
                                    "brand": brand,
                                    "model": folder_name,
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
        # instead of containing image files directly
        is_nested_model_batch = False
        for folder_name in root_entries:
            fp = os.path.join(base_dir, folder_name)
            subdirs = [d for d in os.listdir(fp) if os.path.isdir(os.path.join(fp, d))]
            direct_images = [f for f in os.listdir(fp) if f.lower().endswith((".jpg", ".jpeg", ".png", ".webp"))]
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
                for color_folder in color_dirs:
                    folder_path = os.path.join(model_path, color_folder)
                    raw_files = [f for f in os.listdir(folder_path) if f.lower().endswith((".jpg", ".jpeg", ".png", ".webp"))]
                    if not raw_files:
                        continue

                    # Parse color and brand nicely
                    color_display, brand_name = parse_color_folder(color_folder)

                    # Use subdir=model_name so rel_path becomes model_name/color_folder/file.jpg
                    shot_list = process_and_sort_variant_shots(raw_files, batch_id, model_name, color_folder, folder_path)
                    has_new = any(s.get("is_new_version") for s in shot_list)
                    v_tag = next((s["version_tag"] for s in shot_list if s.get("is_new_version")), "")
                    mtimes = [s["mtime"] for s in shot_list if s.get("mtime")]

                    prod_id = f"{model_name}-{color_folder}"
                    products_map[prod_id] = {
                        "id": prod_id,
                        "brand": brand_name,
                        "model": model_name,
                        "color": color_display,
                        "folder_name": f"{model_name}/{color_folder}",
                        "has_new_version": has_new,
                        "version_tag": v_tag,
                        "latest_mtime": max(mtimes) if mtimes else 0,
                        "variants": { "Standard": shot_list }
                    }
        elif root_entries and any(os.path.isdir(os.path.join(base_dir, d)) and any(f.lower().endswith((".jpg", ".jpeg", ".png", ".webp")) for f in os.listdir(os.path.join(base_dir, d))) for d in root_entries):
            ran_subfolder_branch = True
            for folder_name in sorted(root_entries):
                folder_path = os.path.join(base_dir, folder_name)
                if not os.path.isdir(folder_path):
                    continue
                parts = folder_name.split("-")
                brand = "Brand"
                model_name = folder_name
                color_name = ""

                if len(parts) >= 3:
                    model_name = parts[1].strip()
                    color_name = "-".join(parts[2:]).strip()
                elif len(parts) == 2:
                    model_name = parts[0].strip()
                    color_name = parts[1].strip()

                raw_files = [f for f in os.listdir(folder_path) if f.lower().endswith((".jpg", ".jpeg", ".png", ".webp"))]
                shot_list = process_and_sort_variant_shots(raw_files, batch_id, "", folder_name, folder_path)

                has_new = any(s.get("is_new_version") for s in shot_list)
                v_tag = next((s["version_tag"] for s in shot_list if s.get("is_new_version")), "")
                mtimes = [s["mtime"] for s in shot_list if s.get("mtime")]

                products_map[folder_name] = {
                    "id": folder_name,
                    "brand": brand,
                    "model": model_name,
                    "color": color_name,
                    "folder_name": folder_name,
                    "has_new_version": has_new,
                    "version_tag": v_tag,
                    "latest_mtime": max(mtimes) if mtimes else 0,
                    "variants": { "Standard": shot_list }
                }
        direct_root_images = [f for f in sorted(os.listdir(base_dir)) if f.lower().endswith((".jpg", ".jpeg", ".png", ".webp")) and os.path.isfile(os.path.join(base_dir, f))]
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

                # Generic format: prod_key - [0-9]+ [suffix]
                m2 = re.match(r'^(.*?)[-_](0[1-9]|[1-9][0-9]?)(.*)$', name_no_ext, re.IGNORECASE)
                if m2:
                    prod_key = m2.group(1).strip('-_')
                    shot_num = m2.group(2).zfill(2)
                    raw_suffix = m2.group(3).strip('-_ ')
                    is_version = bool(re.match(r'^(n|ny|new|v[0-9]+|[0-9]+|b|fix|crop)$', raw_suffix, re.IGNORECASE))
                    label = f'{shot_num} ({raw_suffix})' if raw_suffix else shot_num

                    return {
                        'prod_key': prod_key,
                        'filename': fn,
                        'kind': 'Standard',
                        'shot_code': f'{shot_num}{raw_suffix}',
                        'base_shot': shot_num,
                        'suffix': raw_suffix,
                        'label': label,
                        'is_close': (shot_num == '01'),
                        'is_street': False,
                        'order_key': (int(shot_num), 0 if not is_version else 1, raw_suffix),
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

                mtimes = [s["mtime"] for s in shot_list if s.get("mtime")]
                has_new = any(s.get("is_new_version") for s in shot_list)
                v_tag = next((s["version_tag"] for s in shot_list if s.get("is_new_version")), "")

                if prod_key in products_map:
                    prod_key = f"rot-{prod_key}"

                products_map[prod_key] = {
                    "id": prod_key,
                    "brand": brand,
                    "model": model_name,
                    "color": color_name,
                    "folder_name": prod_key,
                    "has_new_version": has_new,
                    "version_tag": v_tag,
                    "latest_mtime": max(mtimes) if mtimes else 0,
                    "variants": { "Standard": shot_list }
                }

    products = list(products_map.values())
    products.sort(key=lambda x: x["model"])
    return products


def removed_root_for(batch_path):
    p = os.path.normpath(str(batch_path))
    return os.path.join(os.path.dirname(p), os.path.basename(p) + " Removed")


def remove_shot(batch, product_id, filename, user):
    base_dir = batch.get("path", "")
    manifest_candidates = [
        os.path.join(base_dir, "products_manifest.json"),
        os.path.join(os.path.dirname(base_dir), "products_manifest.json"),
        os.path.join(base_dir, "manifest.json"),
        os.path.join(os.path.dirname(base_dir), "manifest.json"),
    ]
    if any(os.path.isfile(f) for f in manifest_candidates):
        raise ValueError("Mappen använder manifest – flytt stöds inte")

    products = scan_batch_products(batch)
    product = next((p for p in products if str(p.get("id")) == str(product_id) or str(p.get("id")).lower() == str(product_id).lower()), None)
    if not product:
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

    undo_count = sum(1 for o in log_data if not o.get("undone", False))
    return {
        "status": "ok",
        "op_id": op_id,
        "moved": moved_rel_paths,
        "removed_root": removed_root,
        "undo_count": undo_count
    }


def undo_remove(batch, op_id=None):
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
                target_op = op
                break
    else:
        for op in reversed(log_data):
            if not op.get("undone", False):
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


def list_removed(batch):
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

    def is_admin_request(self, query=None, body_user=None):
        u = self.current_user()
        if u and u.get("role") == "admin":
            return True
        is_tunnel = bool(self.headers.get("CF-Ray") or self.headers.get("CF-Connecting-IP") or self.headers.get("X-Forwarded-For"))
        client_ip = self.client_address[0] if hasattr(self, 'client_address') else ''
        is_local = (not is_tunnel) and (client_ip in ['127.0.0.1', '::1', 'localhost', '192.168.2.33'])
        return is_local


    def can_see_batch(self, batch):
        if not batch:
            return False
        if self.is_admin_request({}):
            return True
        u = self.current_user()
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
                self.send_response(302)
                self.send_header("Location", redirect_url)
                self.send_header("Set-Cookie", f"qa_session={token}; Path=/; Max-Age=15552000; HttpOnly; Secure; SameSite=Lax")
                self.end_headers()
                return
            else:
                self.send_response(403)
                self.send_header("Content-Type", "text/plain; charset=utf-8")
                self.end_headers()
                self.wfile.write(b"Invalid link. Ask Andronik for your personal link.")
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
                self.send_header("Cache-Control", "no-cache")
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
                self.send_header("Cache-Control", "no-cache")
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
            
            allowed = [b for b in batches if b.get("audience") == "assistant" and self.can_see_batch(b)]
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
            batches = load_batches()
            for b in batches:
                if "published" not in b:
                    b["published"] = True

            batches = [b for b in batches if self.can_see_batch(b)]

            content = json.dumps({"batches": batches, "is_admin": is_admin}, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(content)
            return

        elif path == "/api/calendar":
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
            batch = next((b for b in batches if b["id"] == batch_id), batches[0] if batches else None)

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
            reviews = load_reviews(batch["id"])
            payload = {
                "batch": batch,
                "is_admin": is_admin,
                "count": len(products),
                "products": products,
                "reviews": reviews
            }
            content = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(content)
            return

        elif path == "/api/image/removed":
            if not self.is_admin_request(query):
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

            res = list_removed(batch_obj)
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
                batch_id = batches[0]["id"] if batches else "default"
            
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

            abs_base = os.path.abspath(base_dir)
            abs_file = os.path.abspath(file_path)
            if os.path.commonpath([abs_base, abs_file]) != abs_base:
                self.send_error(403, "Access denied")
                return

            if os.path.exists(file_path) and os.path.isfile(file_path):
                file_to_serve = file_path
                is_thumb = "thumb" in query or "w" in query

                if is_thumb:
                    try:
                        raw_w = query.get("w", query.get("thumb", ["160"]))[0]
                        target_w = int(raw_w) if str(raw_w).isdigit() else 160
                    except Exception:
                        target_w = 160
                    if target_w <= 0 or target_w > 1600:
                        target_w = 160

                    try:
                        is_transparent = False
                        with Image.open(file_path) as test_img:
                            is_transparent = test_img.mode in ("RGBA", "LA") or (test_img.mode == "P" and "transparency" in test_img.info)

                        ext = ".png" if is_transparent else ".jpg"
                        cache_key = hashlib.md5(f"{file_path}_{os.path.getmtime(file_path)}_{target_w}".encode("utf-8")).hexdigest() + ext
                        cache_file = os.path.join(THUMB_CACHE_DIR, cache_key)
                        if not os.path.exists(cache_file):
                            with Image.open(file_path) as img:
                                if is_transparent:
                                    if img.mode != "RGBA":
                                        img = img.convert("RGBA")
                                    img.thumbnail((target_w, target_w), Image.Resampling.LANCZOS)
                                    img.save(cache_file, "PNG", optimize=True)
                                else:
                                    img = img.convert("RGB")
                                    img.thumbnail((target_w, target_w), Image.Resampling.LANCZOS)
                                    img.save(cache_file, "JPEG", quality=85, optimize=True)
                        file_to_serve = cache_file
                    except Exception as e:
                        file_to_serve = file_path

                mime_type, _ = mimetypes.guess_type(file_to_serve)
                if not mime_type:
                    mime_type = "image/jpeg"
                
                try:
                    file_size = os.path.getsize(file_to_serve)
                    with open(file_to_serve, "rb") as f:
                        self.send_response(200)
                        self.send_header("Content-Type", mime_type)
                        self.send_header("Content-Length", str(file_size))
                        self.send_header("Cache-Control", "public, max-age=86400")
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
                    self.end_headers()
                    self.wfile.write(json.dumps({"status": "error", "message": "Endast admin kan ändra onedrive_url."}).encode("utf-8"))
                    return

                batch_id = body.get("batch_id") or query.get("batch", [""])[0]
                url = body.get("onedrive_url", "").strip()
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
                batch_id = batches[0]["id"] if batches else "default"
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
                    at_matches = re.findall(r'@(0[1-9]|[1-9][0-9]?)', text)
                    if at_matches:
                        shots = sorted(list(set(m.zfill(2) for m in at_matches)))

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
                    "time": datetime.datetime.now().strftime("%H:%M")
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
                batch_id = batches[0]["id"] if batches else "default"
            
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

                current_reviews = load_reviews(batch_id)
                if pid in current_reviews and "comments" in current_reviews[pid]:
                    c_to_del = next((c for c in current_reviews[pid]["comments"] if c.get("id") == comment_id), None)
                    if c_to_del:
                        u = self.current_user()
                        is_owner = u and "name" in u and u["name"] == c_to_del.get("author")
                        if not (self.is_admin_request({}) or is_owner):
                            self.send_response(403)
                            self.send_header("Content-Type", "application/json; charset=utf-8")
                            self.end_headers()
                            self.wfile.write(json.dumps({"error": "Not authorized"}).encode("utf-8"))
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
                batch_id = batches[0]["id"] if batches else "default"
            
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

        elif path == "/api/review":
            batch_id = query.get("batch", [""])[0]
            if not batch_id:
                batches = load_batches()
                batch_id = batches[0]["id"] if batches else "default"
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
                    if reviewer_override:
                        new_data["reviewer"] = reviewer_override
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
                batch_id = batches[0]["id"] if batches else "default"
            
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
                self.send_error(400, f"Status error: {e}")
                return

        elif path == "/api/checklist/toggle":
            batch_id = query.get("batch", [""])[0]
            if not batch_id:
                batches = load_batches()
                batch_id = batches[0]["id"] if batches else "default"

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
            if not self.is_admin_request(query, reviewer):
                self.send_response(403)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": "Endast admin kan flytta ut bilder."
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
                res = remove_shot(batch_obj, product_id, filename, reviewer)
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
            if not self.is_admin_request(query, reviewer):
                self.send_response(403)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": "Endast admin kan flytta ut bilder."
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
                res = undo_remove(batch_obj, op_id)
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
            if not self.is_admin_request(query, reviewer):
                self.send_response(403)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "status": "error",
                    "message": "Endast admin kan flytta ut bilder."
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
    server.serve_forever()

if __name__ == "__main__":
    p = PORT
    if len(sys.argv) > 1:
        try:
            p = int(sys.argv[1])
        except ValueError:
            pass
    start_server(p)
