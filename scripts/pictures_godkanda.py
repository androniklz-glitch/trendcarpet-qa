"""Distribution copy of approved hub environment/interior images in the CEO's OneDrive Pictures:
  <OneDrive>\\Pictures\\QA-hubb Godkända\\<Brand>\\<Swedish category>\\<modell-farg-sku-01>.<ext>
Brand = Hatshop | Trendcarpet from the batch metadata (`category` kepsar | mattor_inredning); unknown metadata or a
contradiction with the product category = ALARM, the product is never placed. Flat category folders (no product subfolders,
no 00_omslag duplicates): files are numbered -01, -02, ... per product with the best-rated image at -01, plus LÄS MIG.txt
at the root (descriptive only: brand -> category -> product prefix -> count).
White-background / cutout studio packshots are permanently excluded (CEO 2026-11-11: "Ta bort studiobilderna, de ska inte
vara där!"): only approved environment/interior/lifestyle shots and videos enter the target.
Selection (CEO-RÄTTELSE 5): every non-studio shot the hub shows in a batch folder; only images in "<batch> Removed" are
underkända (Removed is never scanned). Reforma is NEVER copied: org filter twice plus a content filter (batch id/name
"reforma", the flow-vibes-rum/mattor batches, furniture products, product id/SKU/handle in REFORMA_PRODUCTS_FILE; an
unreadable list = ALARM).
The target is a copy, never a hub source: the run refuses (exit 3) if any batches_config path lies inside it or inside
Pictures\\QA-hubb Borttagna, and `free` (+U) refuses hub sources (sharepoint_sync.free_refusal). Sources are read only,
never freed. Idempotent: unchanged files are skipped by stat signature, a changed source overwrites its copy (temp name +
sha256 check + os.replace; a killed run's *.godk-tmp is removed next sync), a copy edited at the target is left alone and
reported as conflict.
Mirror: a copy whose image left the selection (Removed, or batch unpublished) is MOVED, never deleted, to
Pictures\\QA-hubb Borttagna\\<same relative path> (sha256 before == after), and moved back when the image returns; empty
folders are removed (rmdir only). Copies of a live batch that failed its gate are held, never moved. More than max(50, 10 %)
moves at once is refused unless --allow-mass-move. Every move is logged in moves.jsonl, every copy/move in manifest.jsonl.

Commands (hub python, cwd = hub dir):
  plan                dry-run: products + files + GB per brand and category, Removed + studio excluded, ALARM batches,
                      Reforma exclusions, what would move to / back from Borttagna (read-only)
  sync [--allow-mass-move]   move back, copy/overwrite per plan, move left-selection copies to Borttagna, write LÄS MIG.txt
                      + pictures_godkanda_links.json, append manifest + moves log
  verify              target + Borttagna vs manifest by metadata (no download, +U safe) + C: free >= 20 GB; deviation -> exit 3
  restore <dest-prefix> [--dry-run]   copy back from the hub source when its sha256 equals the manifest
  free                +U on category folders OneDrive reports as uploaded (target only)
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
import unicodedata
import urllib.parse

from PIL import Image

HUB_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TARGET_NAME = "QA-hubb Godkända"
BORTTAGNA_NAME = "QA-hubb Borttagna"
MANIFEST_FILE = os.path.join(os.path.expanduser("~"), "qa-hub-backup", "pictures-godkanda", "manifest.jsonl")
MOVES_FILE = os.path.join(os.path.expanduser("~"), "qa-hub-backup", "pictures-godkanda", "moves.jsonl")
STUDIO_CACHE_FILE = os.path.join(os.path.expanduser("~"), "qa-hub-backup", "pictures-godkanda", "studio_bg_cache.json")
MASS_MOVE_MIN, MASS_MOVE_FRACTION = 50, 0.10
BRANDS = {"kepsar": "Hatshop", "mattor_inredning": "Trendcarpet"}
HATSHOP_CATEGORIES = {"Kepsar", "Hattar"}
# Reforma shop list (reformahome.myshopify.com products.json, 2026-11-03). Read only; never copied into the hub.
REFORMA_PRODUCTS_FILE = os.path.join(os.path.expanduser("~"), "scratch", "Projects", "Reforma_FTP_gamla_sajten", "live_audit",
                                     "26-11-03-pos1", "products-2026-11-03.json")
REFORMA_BATCH_NEEDLES = ("reforma", "flow-vibes-rum", "flow-vibes-mattor")
FURNITURE_NEEDLES = ("soffbord", "baddsoffa", "soffa-", "sangbord", "sängbord", "matbord", "byra-", "tv-bank", "skank",
                     "fåtölj", "fatolj", "hylla", "stol-")  # Trendcarpet/Hatshop sell no furniture: suspected Reforma
LINKS_FILE = os.path.join(HUB_DIR, "pictures_godkanda_links.json")
LOCK_FILE = os.path.join(HUB_DIR, ".pictures_godkanda.lock")
HUB_URL = "https://trendcarpet-qa.vercel.app"
ONEDRIVE_WEB = "https://trendcarpet-my.sharepoint.com/personal/andronik_lindgren_trendcarpet_se/Documents"
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".gif", ".avif", ".heic"}
MAX_PATH = 200
MIN_FREE_GB = 20
ATTR_PINNED, ATTR_UNPINNED = 0x80000, 0x100000
ATTR_CLOUD_ONLY = 0x1000 | 0x40000 | 0x400000  # OFFLINE | RECALL_ON_OPEN | RECALL_ON_DATA_ACCESS: data not on disk
PIN_ALARM_RUNS = 3
TMP_SUFFIX = ".godk-tmp"
LOCK_STALE_S = 30 * 60

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
    ("Mattor - Rya & Shaggy", (), ("rya", "shaggy", "merida", "bianca", "amalia", "kanha", "batumi", "aranga", "charonne",
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

CHAR_MAP = str.maketrans({
    "å": "a", "ä": "a", "ö": "o", "ø": "o", "æ": "ae", "ß": "ss", "ł": "l",
    "Å": "a", "Ä": "a", "Ö": "o", "Ø": "o", "Æ": "ae", "Ł": "l",
})
CATEGORY_STRIP_WORDS = {
    "ullmatta", "ullmattor", "jutematta", "jutemattor", "rollakansmatta", "rollakansmattor",
    "trappstegsmatta", "trappstegsmattor", "badrumsmatta", "badrumsmattor",
    "bomullsmatta", "bomullsmattor", "bomullsgardin", "bomullsgardiner",
    "linnegardin", "linnegardiner", "sammetsgardin", "veckbandsgardin",
    "trasmatta", "trasmattor", "ryamatta", "ryamattor", "wiltonmatta", "wiltonmattor",
    "viskosmatta", "viskosmattor", "barnmatta", "barnmattor",
    "inomhusutomhusmatta", "inomhusutomhus", "rundmattor", "rundmatta",
    "matta", "mattor", "kuddfodral", "kudde", "kuddar",
    "gardin", "gardiner", "keps", "kepsar", "mossa", "mossor", "hatt", "hattar",
}
MARKETING_FILLER_PHRASES = (
    r"\bhand[- ]woven[- ]wool[- ]premium\b",
    r"\bflat[- ]handloom[- ]wool[- ]premium\b",
    r"\bhand[- ]tufted[- ]wool[- ]premium\b",
    r"\bhandtufted[- ]wool[- ]premium\b",
    r"\bhandloom[- ]wool[- ]premium\b",
    r"\bflatweave[- ]tufted[- ]premium\b",
    r"\btufted[- ]wool[- ]premium\b",
    r"\bloop[- ]handtufted[- ]wool\b",
    r"\bspecial[- ]luxury[- ]edition\b",
    r"\bnatural[- ]cotton[- ]shaggy\b",
    r"\bcheck[- ]cotton[- ]shaggy\b",
    r"\bstripe[- ]cotton[- ]shaggy\b",
    r"\bsolid[- ]cotton[- ]shaggy\b",
    r"\bcotton[- ]shaggy\b",
    r"\bsuper[- ]soft[- ]fur\b",
    r"\bsuper[- ]soft\b",
    r"\bsoft[- ]shaggy\b",
    r"\brecycled[- ]pet\b",
    r"\bbraided[- ]jute\b",
    r"\bbraided[- ]wool\b",
    r"\bjutewool[- ]bubble\b",
    r"\bwool[- ]bubble\b",
    r"\bwool[- ]look\b",
    r"\bfedora[- ]wool[- ]hat\b",
    r"\barmy[- ]cap\b",
    r"\borganiskt[- ]formad\b",
)
NOISE_WORDS = {"carpet", "teppich", "wollteppich", "ullteppe", "vorhang", "curtain", "pillow", "rug", "rugs", "ny", "nya", "bild", "bilder"}
COLOR_EQUIV_GROUPS = (
    {"morkbrun", "dark-brown", "darkbrown"},
    {"ljusbla", "light-blue", "lightblue"},
    {"morkgron", "dark-green", "darkgreen"},
    {"morkgra", "dark-grey", "darkgrey"},
    {"ljusgra", "light-grey", "lightgrey"},
    {"rod", "red"},
    {"bla", "blue"},
    {"gron", "green"},
    {"svart", "black"},
    {"vit", "white"},
    {"gra", "grey", "gray"},
    {"gul", "yellow"},
    {"brun", "brown"},
    {"creme", "cream"},
)
FLAT_NAME_RE = re.compile(r"^([a-z0-9][a-z0-9-]*)-(\d{2,3})(\.[a-z0-9]+)$", re.I)


class Alarm(Exception):
    pass


def lp(path):
    """Long-path form for Win32 file APIs."""
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
    if hit and category_for({"id": ""}, product, context) == "Mattor - Övriga":
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


def load_studio_cache():
    try:
        with open(STUDIO_CACHE_FILE, encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_studio_cache(cache):
    try:
        os.makedirs(os.path.dirname(STUDIO_CACHE_FILE), exist_ok=True)
        tmp = STUDIO_CACHE_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8", newline="\n") as f:
            json.dump(cache, f, ensure_ascii=False)
        os.replace(tmp, STUDIO_CACHE_FILE)
    except OSError:
        pass


def is_studio_white_bg(path, cache=None, dirty=None):
    """True when an image is a white-background or transparent cutout studio packshot (12 border/corner sample points).
    Non-image files (e.g. .mp4 videos) return False. Results are cached by [size, mtime] for sub-second repeat runs."""
    ext = os.path.splitext(path)[1].lower()
    if ext not in IMAGE_EXTS:
        return False
    s_sig = sig(path)
    key = os.path.normcase(os.path.abspath(path))
    if cache is not None and s_sig is not None:
        hit = cache.get(key)
        if isinstance(hit, list) and len(hit) == 3 and hit[0] == s_sig[0] and hit[1] == s_sig[1]:
            return bool(hit[2])
    res = False
    try:
        with Image.open(lp(path)) as im:
            w, h = im.size
            if w >= 20 and h >= 20:
                if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
                    rgba = im.convert("RGBA")
                    apts = ((2, 2), (w - 3, 2), (2, h - 3), (w - 3, h - 3),
                            (w // 2, 2), (w // 2, h - 3), (2, h // 2), (w - 3, h // 2))
                    if sum(1 for pt in apts if rgba.getpixel(pt)[3] < 15) >= 6:
                        res = True
                    else:
                        rgb = Image.new("RGB", rgba.size, (255, 255, 255))
                        rgb.paste(rgba, mask=rgba.split()[3])
                else:
                    rgb = im.convert("RGB")
                if not res:
                    pts = (
                        (4, 4), (w - 5, 4), (4, h - 5), (w - 5, h - 5),
                        (w // 2, 4), (w // 2, h - 5), (4, h // 2), (w - 5, h // 2),
                        (w // 4, 4), (3 * w // 4, 4), (4, h // 4), (w - 5, h // 4),
                    )
                    white_pts = pure_white_pts = 0
                    for pt in pts:
                        r_c, g_c, b_c = rgb.getpixel(pt)
                        if r_c >= 246 and g_c >= 246 and b_c >= 246 and max(r_c, g_c, b_c) - min(r_c, g_c, b_c) <= 6:
                            white_pts += 1
                        if r_c >= 252 and g_c >= 252 and b_c >= 252:
                            pure_white_pts += 1
                    res = pure_white_pts >= 10 or white_pts >= 11
    except Exception:  # noqa: BLE001 - unreadable/corrupt image falls back to scene handling
        res = False
    if cache is not None and s_sig is not None:
        cache[key] = [s_sig[0], s_sig[1], res]
        if dirty is not None:
            dirty.append(1)
    return res


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
    """Readable product name before SKU/length handling."""
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
    for sep in (" (", " — ") if len(name) > 50 else ():
        if sep in name[8:]:
            name = name[:8 + name[8:].index(sep)]
    if sku:
        name = re.sub(rf"[\s,(]*{re.escape(sku)}[\s,)]*", " ", name, flags=re.I)
    name = clean_component(name, 60)
    if name.count("(") > name.count(")"):
        name += ")"
    return f"{name} ({sku})" if sku else name


def ascii_slug(s):
    s = str(s or "").translate(CHAR_MAP)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[^a-z0-9]+", "-", s.lower())
    return re.sub(r"-+", "-", s).strip("-")


def _has_color_equiv(existing_tokens, cand_token):
    joined = "-".join(existing_tokens)
    for grp in COLOR_EQUIV_GROUPS:
        if cand_token in grp and (grp & set(existing_tokens) or any(g in joined for g in grp if "-" in g)):
            return True
    return False


def _clean_color_hint(col_raw, sku=""):
    if not col_raw:
        return []
    c = re.sub(r"^\d{1,2}\.\s*", "", str(col_raw))
    m_paren = re.match(r"^[^()]+?\(([^()0-9]{2,25})\)\s*$", c.strip())
    c = m_paren.group(1) if m_paren else re.split(r"\s*[—–\[\(]\s*", c)[0]
    if sku:
        c = re.sub(rf"\b{re.escape(sku)}\b", " ", c, flags=re.I)
    c = NOISE_RE.sub(" ", c)
    for pat in MARKETING_FILLER_PHRASES:
        c = re.sub(pat, " ", c, flags=re.I)
    return [t for t in ascii_slug(c).split("-") if t and t not in CATEGORY_STRIP_WORDS and t not in NOISE_WORDS][:4]


def product_slug_base(product, sku, context=""):
    """Concise lowercase ASCII <modell>-<farg>[-<sku>] prefix (~12-32 chars before -01.jpg)."""
    s = base_name(product, context)
    s = re.sub(r"\b\d+\s*x\s*\d+\s*(?:cm|mm|m)?\b", " ", s, flags=re.I)
    s = re.sub(r"\b\d+\s*(?:cm|mm)\b", " ", s, flags=re.I)

    def _clean_paren(m):
        inner = m.group(1)
        if re.search(r"\b(bild|bilder|ny|cm|ullpals|ullpäls|\d{4,})\b", inner, flags=re.I):
            first = inner.split(",")[0].strip()
            return f" {first} " if first and not re.search(r"\d", first) else " "
        return f" {inner} "

    s = re.sub(r"\(([^()]*)\)", _clean_paren, s)
    if sku:
        s = re.sub(rf"\b{re.escape(sku)}\b", " ", s, flags=re.I)
    s = NOISE_RE.sub(" ", s)
    for pat in MARKETING_FILLER_PHRASES:
        s = re.sub(pat, " ", s, flags=re.I)

    tokens = [t for t in ascii_slug(s).split("-") if t and t not in NOISE_WORDS]
    if len(tokens) >= 2 and re.fullmatch(r"0\d|[12]\d", tokens[0]) and any(t.isalpha() for t in tokens[1:]):
        tokens = tokens[1:]
    filtered = [t for t in tokens if t not in CATEGORY_STRIP_WORDS]
    if filtered and any(len(t) >= 2 for t in filtered):
        tokens = filtered

    dedup = []
    for t in tokens:
        if t not in dedup and not _has_color_equiv(dedup, t):
            dedup.append(t)
    tokens = dedup

    col_toks = [ct for ct in _clean_color_hint(product.get("color"), sku)
                if ct not in tokens and not _has_color_equiv(tokens, ct)]
    if col_toks:
        if len(tokens) >= 4:
            tokens = tokens[:3] + col_toks[:2]
        else:
            for ct in col_toks:
                if len(tokens) < 5:
                    tokens.append(ct)
    elif len(tokens) > 5:
        tokens = tokens[:3] + tokens[-2:]

    base = "-".join(tokens) or "produkt"
    if len(base) > 28:
        acc = []
        for t in tokens:
            cand = "-".join(acc + [t])
            if len(cand) <= 28 or not acc:
                acc.append(t)
            else:
                break
        base = "-".join(acc)[:28].strip("-")

    if sku:
        sku_clean = ascii_slug(sku)
        if sku_clean and sku_clean not in base.split("-"):
            base = f"{base}-{sku_clean}"
    return base


def batch_date(batch):
    m = re.match(r"(\d{2})-(\d{2})-(\d{2})", batch.get("id") or "")
    if m:
        return "".join(m.groups())
    c = str(batch.get("createdAt") or "")
    return c[2:4] + c[5:7] + c[8:10] if re.match(r"\d{4}-\d{2}-\d{2}", c) else "000000"


def stars_by_key(ratings):
    """{'pid|filename': mean stars over reviewers}; malformed entries are ignored."""
    acc = {}
    for per in (ratings or {}).values():
        for key, v in (per or {}).items() if isinstance(per, dict) else ():
            s = v.get("stars") if isinstance(v, dict) else v
            if isinstance(s, (int, float)) and not isinstance(s, bool) and 1 <= s <= 5:
                acc.setdefault(key, []).append(float(s))
    return {k: sum(v) / len(v) for k, v in acc.items()}


def known_prefixes(latest):
    """{'by_key': {product_key: (cat_rel, slug_base)}, 'by_src': {src_lower: (cat_rel, slug_base)}} from ACTIVE flat
    manifest entries (<Brand>\\<Category>\\<slug_base>-<NN>.<ext>), so a re-titled batch keeps its established prefix."""
    count, by_src = {}, {}
    for e in latest.values():
        if not is_active(e) or e.get("kind") == "cover":
            continue
        parts = e["dest_rel"].replace("/", os.sep).split(os.sep)
        if len(parts) != 3:
            continue
        m = FLAT_NAME_RE.match(parts[2])
        if not m:
            continue
        pair = (os.path.join(parts[0], parts[1]), m.group(1).lower())
        key = e.get("product_key") or ""
        if key:
            c = count.setdefault(key, {})
            c[pair] = c.get(pair, 0) + 1
        by_src[e["src"].lower()] = pair
    by_key = {k: max(sorted(c), key=lambda p: c[p]) for k, c in count.items()}
    return {"by_key": by_key, "by_src": by_src}


def build_plan(records, target, known=None):
    """records: dicts {batch, product, shots:[{src, filename, gold}], stars:{key: mean}}. Returns (items, products,
    brand_alarms). Flat layout: <Brand>\\<Category>\\<slug_base>-<01..99>.<ext> with best-rated image at -01 and zero
    00_omslag duplicates."""
    products, brand_alarms = {}, []
    for r in sorted(records, key=lambda r: (batch_date(r["batch"]), r["batch"]["id"]), reverse=True):
        p, sku = r["product"], sku_for(r["product"])
        rel = os.path.relpath(r["shots"][0]["src"], r["batch"]["path"]) if r["shots"] and r["batch"].get("path") else ""
        ctx = rel.replace("/", os.sep).split(os.sep)[0] if os.sep in rel.replace("/", os.sep) else ""
        key = sku or clean_component(str(p.get("id") or ""), 80).lower()
        if not sku and ctx and is_generic(str(p.get("id") or "")):
            key = clean_component(f"{ctx} {p.get('id')}", 80).lower()
        e = products.setdefault(key, {
            "category": category_for(r["batch"], p, ctx),
            "title": product_name(p, sku, ctx),
            "raw_slug": product_slug_base(p, sku, ctx),
            "color": p.get("color"),
            "sku": sku,
            "files": [],
            "batches": [],
            "seen": set(),
            "brands": set(),
            "brand_errors": [],
        })
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
            e["cat_rel"] = os.path.join(e["brand"], e["category"])

    used_by_cat = {}
    for key, e in sorted(products.items()) if known else ():
        kp = known["by_key"].get(key)
        if not kp:
            hits = [known["by_src"][f["src"].lower()] for f in e["files"] if f["src"].lower() in known["by_src"]]
            kp = max(sorted(set(hits)), key=hits.count) if hits else None
        used = used_by_cat.setdefault(e["cat_rel"].lower(), set())
        if kp and kp[0].lower() == e["cat_rel"].lower() and kp[1].lower() not in used:
            e["slug_base"] = kp[1].lower()
            used.add(e["slug_base"])

    items = []
    for key, e in sorted(products.items()):
        used = used_by_cat.setdefault(e["cat_rel"].lower(), set())
        if "slug_base" not in e:
            base = e["raw_slug"]
            if base in used and e.get("color"):
                col_t = [t for t in _clean_color_hint(e["color"], e["sku"]) if t not in base.split("-")]
                if col_t and f"{base}-{col_t[0]}" not in used and len(f"{base}-{col_t[0]}") <= 34:
                    base = f"{base}-{col_t[0]}"
            if base in used:
                key_toks = [t for t in ascii_slug(key).split("-")
                            if t not in base.split("-") and t not in CATEGORY_STRIP_WORDS and t not in NOISE_WORDS]
                cand = f"{base}-{key_toks[-1]}" if key_toks else ""
                if cand and cand not in used and len(cand) <= 34:
                    base = cand
                else:
                    n = 2
                    while f"{base}-v{n}" in used:
                        n += 1
                    base = f"{base}-v{n}"
            used.add(base)
            e["slug_base"] = base

        cat_rel = e["cat_rel"]
        imgs = [f for f in e["files"] if os.path.splitext(f["src"])[1].lower() in IMAGE_EXTS]
        best_src = max(imgs, key=lambda f: (f["stars"], f.get("gold", False), f["date"]))["src"].lower() if imgs else None
        ordered = sorted(
            e["files"],
            key=lambda f: (0 if f["src"].lower() == best_src else 1, -f["stars"], f["date"], os.path.basename(f["src"]).lower()),
        )
        width = max(2, len(str(len(ordered))))
        room = MAX_PATH - len(os.path.join(target, cat_rel)) - 1
        hero_rel = None
        for idx, f in enumerate(ordered, 1):
            ext = os.path.splitext(f["src"])[1].lower()
            suffix = f"-{idx:0{width}d}{ext}"
            max_stem = max(8, room - len(suffix))
            stem = e["slug_base"][:max_stem].rstrip("-")
            rel_dst = os.path.join(cat_rel, f"{stem}{suffix}")
            if hero_rel is None and ext in IMAGE_EXTS:
                hero_rel = rel_dst
            items.append({"src": f["src"], "dest_rel": rel_dst, "batch": f["batch"], "kind": "image", "product_key": key})
        e["hero_rel"] = hero_rel or (os.path.join(cat_rel, f"{e['slug_base']}-{1:0{width}d}{os.path.splitext(ordered[0]['src'])[1].lower()}") if ordered else "")
        e["folder_rel"] = os.path.join(cat_rel, e["slug_base"])
        e["count"] = len(ordered)
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


def collect(ss, pictures, target, reforma=None, studio_excluded_out=None):
    """Returns (batches, records, alarms, removed, skipped, reforma_hits); reforma_hits = [{batch, product, reason, files,
    srcs}] (product None = whole batch). Reforma content filter on top of the org filter; white-background studio packshots
    are permanently excluded."""
    reforma = reforma if reforma is not None else load_reforma()
    studio_cache, cache_dirty = load_studio_cache(), []
    args = type("A", (), {"batch": None, "all": True, "include_drafts": False})()
    batches, skipped = ss.select_batches(args)
    records, alarms, removed, hits = [], [], {}, []
    studio_excl = studio_excluded_out if studio_excluded_out is not None else {}
    for b in batches:
        if ss.org_decision(b):  # second, independent org filter: Reforma never leaves the hub
            skipped.append({"batch": b.get("id"), "reason": ss.org_decision(b)})
            continue
        why = reforma_batch_hit(b)
        if why:  # third filter, content based
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
        mixed_mapped = b.get("content_kind") == "mixed" and os.path.isfile(os.path.join(root, "MAPPING.csv"))
        for prod in ss.hub.scan_batch_products(b):
            shots, seen = [], set()
            variants_map = prod.get("variants") or {}
            gold_keys = ["Original"] if "Original" in variants_map else list(variants_map)[:1]
            scan_keys = gold_keys if (mixed_mapped and "Original" in variants_map) else list(variants_map)
            for vk in scan_keys:
                for s in variants_map.get(vk) or []:
                    rel = (s.get("rel_path") or "").replace("\\", "/").lstrip("/")
                    full = os.path.normpath(os.path.join(root, rel)) if rel else ""
                    if not rel or rel in seen or not os.path.isfile(full) or not inside(full, root) or inside(full, target):
                        continue
                    seen.add(rel)
                    shots.append({"src": full, "filename": s.get("filename") or os.path.basename(rel),
                                  "is_studio_packshot": bool(s.get("is_studio_packshot")),
                                  "gold": bool(vk in gold_keys and (s.get("is_studio_packshot") or s.get("in_gold_set")))})
            if not shots:
                continue
            rel0 = os.path.relpath(shots[0]["src"], root).replace("/", os.sep)
            why = reforma_product_hit(prod, reforma, rel0.split(os.sep)[0] if os.sep in rel0 else "")
            if why:
                hits.append({"batch": b["id"], "product": prod.get("id"), "reason": why, "files": len(shots),
                             "srcs": [s["src"] for s in shots]})
                continue
            env_shots = []
            for s in shots:
                if (mixed_mapped and s.get("is_studio_packshot")) or is_studio_white_bg(s["src"], studio_cache, cache_dirty):
                    studio_excl[b["id"]] = studio_excl.get(b["id"], 0) + 1
                else:
                    env_shots.append(s)
            if not env_shots:
                continue
            records.append({"batch": b, "product": prod, "shots": env_shots, "stars": stars})
    if cache_dirty:
        save_studio_cache(studio_cache)
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
        prev = None
    if prev and d_sig == prev["dst_sig"] and s_sig == prev["src_sig"] and prev["src"].lower() == it["src"].lower():
        return "unchanged"
    if prev and d_sig not in (None, prev["dst_sig"]):
        return "conflict"
    if dry_run:
        return "overwritten" if prev else ("conflict" if d_sig and d_sig[0] != s_sig[0] else "copied")
    digest = sha256(it["src"])
    status = "overwritten" if prev else "copied"
    if not prev and d_sig:
        if sha256(dst) != digest:
            return "conflict"
        status = "adopted"
    else:
        copy_atomic(it["src"], dst, digest)
    append_manifest(latest, {"ts": now_iso(), "dest_rel": it["dest_rel"],
                             "kind": it["kind"], "batch": it["batch"], "src": it["src"], "sha256": digest,
                             "src_sig": s_sig, "dst_sig": sig(dst), **pkey(it)})
    return status


def copy_atomic(src, dst, digest):
    tmp = dst + TMP_SUFFIX
    os.makedirs(lp(os.path.dirname(dst)), exist_ok=True)
    shutil.copy2(lp(src), lp(tmp))
    if sha256(tmp) != digest:
        os.remove(lp(tmp))
        raise OSError(f"hash mismatch after copy: {dst}")
    os.replace(lp(tmp), lp(dst))


def clean_tmp(target, dry_run):
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
        if os.path.exists(lp(dst)):
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
    if os.name != "nt":
        return None
    a = ctypes.WinDLL("kernel32", use_last_error=True).GetFileAttributesW(lp(path))
    return None if a == 0xFFFFFFFF else a


def is_cloud(path):
    a = file_attrs(path)
    return a is not None and bool(a & (ATTR_UNPINNED | ATTR_CLOUD_ONLY))


_SUBPROCESS_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0


def pin_local(path):
    subprocess.run(["attrib", "-U", "+P", path], check=True, capture_output=True, creationflags=_SUBPROCESS_NO_WINDOW)


def readme_text(products):
    out = [f"{TARGET_NAME}",
           "Här ligger de godkända miljö- och interiörbilderna från QA-hubben (studiobilder med vit bakgrund ingår inte).",
           "Struktur: Varumärke > Kategori > korta bildfilnamn direkt i kategorimappen.",
           "Filnamn: modell-farg-artikelnummer-01.jpg, där -01 är den bäst betygsatta bilden för produkten.",
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
            es = sorted(cats[cat], key=lambda e: e["slug_base"])
            out.append(f"  {cat} — {len(es)} produkter, {sum(e['count'] for e in es)} bilder")
            out += [f"    {e['slug_base']}-* ({e['title']}): {e['count']} bilder" for e in es]
        out.append("")
    return "\r\n".join(out) + "\r\n"


def links_json(products):
    out = {}
    for key, e in sorted(products.items()):
        rel_parts = e["hero_rel"].split(os.sep) if e.get("hero_rel") else e["cat_rel"].split(os.sep)
        url = f"{ONEDRIVE_WEB}/Pictures/{urllib.parse.quote(TARGET_NAME)}/" + "/".join(
            urllib.parse.quote(p) for p in rel_parts)
        for b in e["batches"]:
            out.setdefault(b, []).append({"product_key": key, "folder": e["cat_rel"], "prefix": e["slug_base"],
                                          "hero": e.get("hero_rel", ""), "url": url})
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
    studio_excl = {}
    batches, records, alarms, removed, skipped, ref_hits = collect(ss, pictures, target, studio_excluded_out=studio_excl)
    latest = load_manifest()
    items, products, brand_alarms = build_plan(records, target, known_prefixes(latest))
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
            agg["products"].add(it.get("product_key") or os.path.dirname(it["dest_rel"]))
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
    alarms = alarms + [{"batch": a["batches"], "message": a["message"]} for a in brand_alarms] + mirror["alarms"]

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
               "removed_excluded": sum(removed.values()), "studio_excluded": sum(studio_excl.values()),
               "status": status_n, "alarms": alarms[:50], "alarm_n": len(alarms),
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
    print(json.dumps({"studio_excluded_per_batch": studio_excl}, ensure_ascii=False))
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
    for brand in sorted(os.listdir(lp(target))) if os.path.isdir(lp(target)) else []:
        bdir = os.path.join(target, brand)
        if not os.path.isdir(lp(bdir)):
            continue
        for cat in sorted(os.listdir(lp(bdir))):
            d = os.path.join(bdir, cat)
            if not os.path.isdir(lp(d)):
                continue
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
    if os.name != "nt":
        try:
            os.kill(pid, 0)
            return True
        except OSError:
            return False
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    h = k32.OpenProcess(0x1000, False, pid)
    if not h:
        return False
    try:
        code = ctypes.c_ulong()
        return bool(k32.GetExitCodeProcess(h, ctypes.byref(code))) and code.value == 259
    finally:
        k32.CloseHandle(h)


def read_lock():
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
    d = read_lock()
    if d is not None and int(d.get("pid") or 0) == os.getpid():
        os.remove(LOCK_FILE)


def main():
    for stream in (sys.stdout, sys.stderr):
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
