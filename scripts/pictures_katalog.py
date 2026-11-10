"""Katalog-PDF:er för Pictures\\QA-hubb Godkända (order 9, QA-hubb 1.15; v2 links 1.16).

Läser den färdiga målmappen (<Varumärke>\\<Kategori>\\<Produkt>\\00_omslag.* + bilder) och skriver en PDF per varumärke
i målmappens rot: "Katalog Hatshop.pdf", "Katalog Trendcarpet.pdf". En PDF per varumärke (inte en stor) = mindre filer
som förhandsgranskas snabbt i OneDrive mobil. Sida 1 = omslag med innehåll (kategori, produkter, bilder, sida), sedan
en sektion per kategori: 3 x 4 produktkort (omslagsbild, produktnamn, antal bilder).
Länkar (v2): ett kort öppnar produktmappen i OneDrive webb, en innehållsrad hoppar till kategorins första sida.

Rör ALDRIG något annat än de två PDF-filerna: läser bilder, skriver via .tmp + os.replace. Samma lås som
pictures_godkanda.py (sync och katalog kör aldrig samtidigt). Kör efter `pictures_godkanda.py sync`.

  python -I -B -X utf8 scripts/pictures_katalog.py plan                 # dry-run: räkna, inga filer skrivs
  python -I -B -X utf8 scripts/pictures_katalog.py build [--out DIR]    # skriv PDF:erna (default: målmappens rot)
"""
import argparse
import datetime
import io
import json
import math
import os
import sys
import urllib.parse

from PIL import Image, ImageDraw, ImageFont, ImageOps

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pictures_godkanda as pg  # noqa: E402  (read-only reuse: TARGET_NAME, lp, guard_target, LOCK_FILE, IMAGE_EXTS)

DPI = 150
PAGE_W, PAGE_H = 1240, 1754  # A4 @ 150 dpi
MARGIN = 80
COLS, ROWS = 3, 4
PER_PAGE = COLS * ROWS
COVER_STEM = "00_omslag"
PDF_NAME = "Katalog {brand}.pdf"
FONT_DIR = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
INK, MUTED, RULE = (28, 28, 28), (110, 110, 110), (225, 225, 225)


def nb(n, one, many):
    return f"{n} {one if n == 1 else many}"


def font(size, bold=False):
    try:
        return ImageFont.truetype(os.path.join(FONT_DIR, "segoeuib.ttf" if bold else "segoeui.ttf"), size)
    except OSError:
        return ImageFont.load_default(size)


def scan(target):
    """{brand: {category: [{name, cover, count}]}} from the live target tree; count = files except 00_omslag.*."""
    tree = {}
    for brand in sorted(os.listdir(pg.lp(target)), key=str.lower):
        bdir = os.path.join(target, brand)
        if not os.path.isdir(pg.lp(bdir)):
            continue
        for cat in sorted(os.listdir(pg.lp(bdir)), key=str.lower):
            cdir = os.path.join(bdir, cat)
            if not os.path.isdir(pg.lp(cdir)):
                continue
            for prod in sorted(os.listdir(pg.lp(cdir)), key=str.lower):
                pdir = os.path.join(cdir, prod)
                if not os.path.isdir(pg.lp(pdir)):
                    continue
                names = [n for n in os.listdir(pg.lp(pdir)) if os.path.isfile(pg.lp(os.path.join(pdir, n)))]
                covers = [n for n in names if os.path.splitext(n)[0].lower() == COVER_STEM
                          and os.path.splitext(n)[1].lower() in pg.IMAGE_EXTS]
                tree.setdefault(brand, {}).setdefault(cat, []).append({
                    "name": prod, "cover": os.path.join(pdir, covers[0]) if covers else None,
                    "count": sum(1 for n in names if os.path.splitext(n)[0].lower() != COVER_STEM)})
    return tree


def layout(cats):
    """[(category, first_page, n_pages)] with page 1 = cover page."""
    out, page = [], 2
    for cat in sorted(cats, key=str.lower):
        n = max(1, math.ceil(len(cats[cat]) / PER_PAGE))
        out.append((cat, page, n))
        page += n
    return out


def text_fit(draw, text, fnt, width, lines=2):
    """Wrap on words into at most `lines` lines; the last one gets an ellipsis if cut. A word wider than the line is cut
    on characters."""
    out, cur = [], ""
    for w in text.split():
        while len(w) > 1 and draw.textlength(w, font=fnt) > width:
            i = len(w) - 1
            while i > 1 and draw.textlength(w[:i], font=fnt) > width:
                i -= 1
            if cur:
                out.append(cur)
                cur = ""
            out.append(w[:i])
            w = w[i:]
        t = f"{cur} {w}".strip()
        if draw.textlength(t, font=fnt) <= width:
            cur = t
            continue
        if cur:
            out.append(cur)
        cur = w
    if cur:
        out.append(cur)
    if len(out) > lines:
        out = out[:lines]
        last = out[-1]
        while last and draw.textlength(last + "…", font=fnt) > width:
            last = last[:-1]
        out[-1] = last.rstrip() + "…"
    return out


def thumb(path, box):
    """Cover image fitted into box (w, h) on white; None if missing or unreadable."""
    if not path:
        return None
    try:
        with Image.open(pg.lp(path)) as im:
            im = ImageOps.exif_transpose(im)
            im.draft("RGB", (box[0] * 2, box[1] * 2))
            if im.mode in ("RGBA", "LA", "P"):
                im = im.convert("RGBA")
                bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
                im = Image.alpha_composite(bg, im)
            im = im.convert("RGB")
            im.thumbnail(box, Image.LANCZOS)
            return im.copy()
    except (OSError, ValueError, Image.DecompressionBombError):
        return None


def footer(draw, brand, label, page, total):
    f = font(20)
    draw.line((MARGIN, PAGE_H - 70, PAGE_W - MARGIN, PAGE_H - 70), fill=RULE, width=2)
    draw.text((MARGIN, PAGE_H - 58), f"{brand} · {label}", font=f, fill=MUTED)
    s = f"Sida {page} av {total}"
    draw.text((PAGE_W - MARGIN - draw.textlength(s, font=f), PAGE_H - 58), s, font=f, fill=MUTED)


def cover_page(brand, cats, plan, total, now):
    pg_img = Image.new("RGB", (PAGE_W, PAGE_H), "white")
    d = ImageDraw.Draw(pg_img)
    n_prod = sum(len(v) for v in cats.values())
    n_img = sum(p["count"] for v in cats.values() for p in v)
    y = 150
    d.text((MARGIN, y), brand, font=font(96, True), fill=INK)
    y += 130
    d.text((MARGIN, y), "Godkända bilder — katalog", font=font(40), fill=INK)
    y += 70
    d.text((MARGIN, y), f"{nb(n_prod, 'produkt', 'produkter')} · {nb(n_img, 'bild', 'bilder')} · {now:%Y-%m-%d}", font=font(28), fill=MUTED)
    y += 90
    # Mosaic of covers (first product per category, then the rest) to make page 1 look like the brand; one row when
    # the table of contents is long (Trendcarpet: 14 categories) so everything fits on page 1.
    picks, seen, n_mosaic = [], set(), 6 if len(plan) <= 6 else 3
    for rnd in range(PER_PAGE):
        for cat in sorted(cats, key=str.lower):
            if len(picks) < n_mosaic and rnd < len(cats[cat]) and cats[cat][rnd]["cover"] and cats[cat][rnd]["name"] not in seen:
                picks.append(cats[cat][rnd]["cover"])
                seen.add(cats[cat][rnd]["name"])
    gap, cw = 16, (PAGE_W - 2 * MARGIN - 2 * 16) // 3
    for i, p in enumerate(picks):
        x0, y0 = MARGIN + (i % 3) * (cw + gap), y + (i // 3) * (cw + gap)
        d.rectangle((x0, y0, x0 + cw, y0 + cw), fill="white", outline=RULE, width=2)
        t = thumb(p, (cw - 6, cw - 6))  # inset: keep the 2 px frame visible
        if t:
            pg_img.paste(t, (x0 + (cw - t.width) // 2, y0 + (cw - t.height) // 2))
    y += (cw + gap) * math.ceil(len(picks) / 3) + 40 if picks else 0
    d.text((MARGIN, y), "Innehåll", font=font(34, True), fill=INK)
    y += 60
    fr, fb = font(26), font(26, True)
    links = []
    for cat, first, _n in plan:
        es = cats[cat]
        links.append(("goto", (MARGIN, y, PAGE_W - MARGIN, y + 40), first))  # tap a contents row -> its first page
        d.text((MARGIN, y), cat, font=fb, fill=INK)
        meta = f"{nb(len(es), 'produkt', 'produkter')} · {nb(sum(p['count'] for p in es), 'bild', 'bilder')}"
        d.text((MARGIN + 420, y), meta, font=fr, fill=MUTED)
        s = f"sida {first}"
        d.text((PAGE_W - MARGIN - d.textlength(s, font=fr), y), s, font=fr, fill=MUTED)
        y += 44
    y += 20
    for line in ("Tryck på ett kort för att öppna produktmappen i OneDrive (Pictures > QA-hubb Godkända > "
                 f"{brand} > Kategori > Produkt). Tryck på en rad ovan för att hoppa till kategorin.",
                 "Bilden är produktens omslag (00_omslag, bäst betygsatt)."):
        for ln in text_fit(d, line, font(22), PAGE_W - 2 * MARGIN, 3):
            d.text((MARGIN, y), ln, font=font(22), fill=MUTED)
            y += 32
    footer(d, brand, "Innehåll", 1, total)
    return pg_img, links


def folder_url(*parts):
    """OneDrive web URL of a folder under Pictures\\QA-hubb Godkända (same scheme as pictures_godkanda.links_json)."""
    return f"{pg.ONEDRIVE_WEB}/Pictures/{urllib.parse.quote(pg.TARGET_NAME)}/" + "/".join(urllib.parse.quote(p) for p in parts)


def section_pages(brand, cat, products, first, total):
    pages, fh, fn, fc = [], font(44, True), font(22, True), font(20)
    gap = 28
    cw = (PAGE_W - 2 * MARGIN - (COLS - 1) * gap) // COLS
    top = MARGIN + 110
    ch = (PAGE_H - 90 - top - (ROWS - 1) * gap) // ROWS
    img_h = ch - 94  # 2 name lines (28 px) + count (26 px) + gaps
    n_pages = max(1, math.ceil(len(products) / PER_PAGE))
    for i in range(n_pages):
        im = Image.new("RGB", (PAGE_W, PAGE_H), "white")
        d = ImageDraw.Draw(im)
        head = cat if i == 0 else f"{cat} (forts.)"
        d.text((MARGIN, MARGIN), head, font=fh, fill=INK)
        meta = f"{nb(len(products), 'produkt', 'produkter')} · {nb(sum(p['count'] for p in products), 'bild', 'bilder')}"
        d.text((PAGE_W - MARGIN - d.textlength(meta, font=fc), MARGIN + 20), meta, font=fc, fill=MUTED)
        links = []
        for j, p in enumerate(products[i * PER_PAGE:(i + 1) * PER_PAGE]):
            x0, y0 = MARGIN + (j % COLS) * (cw + gap), top + (j // COLS) * (ch + gap)
            links.append(("uri", (x0, y0, x0 + cw, y0 + ch), folder_url(brand, cat, p["name"])))
            d.rectangle((x0, y0, x0 + cw, y0 + img_h), fill="white", outline=RULE, width=2)
            t = thumb(p["cover"], (cw - 6, img_h - 6))  # inset: keep the 2 px frame visible
            if t:
                im.paste(t, (x0 + (cw - t.width) // 2, y0 + (img_h - t.height) // 2))
            else:
                s = "ingen omslagsbild"
                d.text((x0 + (cw - d.textlength(s, font=fc)) // 2, y0 + img_h // 2 - 12), s, font=fc, fill=MUTED)
            ty = y0 + img_h + 8
            for ln in text_fit(d, p["name"], fn, cw, 2):
                d.text((x0, ty), ln, font=fn, fill=INK)
                ty += 28
            d.text((x0, ty), nb(p['count'], 'bild', 'bilder'), font=fc, fill=MUTED)
        footer(d, brand, cat, first + i, total)
        pages.append((im, links))
    return pages


def pdf_text(s):
    """PDF text string: UTF-16BE hex with BOM (handles å/ä/ö and any URL-unsafe char)."""
    return "<FEFF" + s.encode("utf-16-be").hex().upper() + ">"


def pdf_lit(s):
    """PDF literal string body (URIs are 7-bit ASCII): escape backslash and parentheses."""
    return s.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def write_pdf(path, pages, title):
    """Minimal PDF 1.4 writer: one JPEG (DCTDecode) per A4 page + link annotations (PIL's PDF writer has no links).
    pages: [(PIL RGB image, [("uri", (x0, y0, x1, y1) px, url) | ("goto", rect px, page_no 1-based)])]."""
    pw, ph, k = 595.28, 841.89, 72 / DPI
    n = len(pages)
    page_id = [3 + 4 * i for i in range(n)]  # per page: page, content, image, annots array
    objs = {1: b"<< /Type /Catalog /Pages 2 0 R >>",
            2: f"<< /Type /Pages /Count {n} /Kids [{' '.join(f'{i} 0 R' for i in page_id)}] >>".encode()}
    info_id = 3 + 4 * n
    objs[info_id] = f"<< /Title {pdf_text(title)} /Producer (QA-hubben pictures_katalog) >>".encode()
    for i, (im, links) in enumerate(pages):
        pid, cid, iid, aid = page_id[i], page_id[i] + 1, page_id[i] + 2, page_id[i] + 3
        buf = io.BytesIO()
        im.save(buf, "JPEG", quality=82, optimize=True)
        jpg = buf.getvalue()
        objs[iid] = (f"<< /Type /XObject /Subtype /Image /Width {im.width} /Height {im.height} /ColorSpace /DeviceRGB "
                     f"/BitsPerComponent 8 /Filter /DCTDecode /Length {len(jpg)} >>\nstream\n").encode() + jpg + b"\nendstream"
        content = f"q {pw:.2f} 0 0 {ph:.2f} 0 0 cm /Im0 Do Q".encode()
        objs[cid] = f"<< /Length {len(content)} >>\nstream\n".encode() + content + b"\nendstream"
        annots = []
        for kind, (x0, y0, x1, y1), target in links:
            rect = f"[{x0 * k:.2f} {ph - y1 * k:.2f} {x1 * k:.2f} {ph - y0 * k:.2f}]"
            act = (f"/A << /S /URI /URI ({pdf_lit(target)}) >>" if kind == "uri" else f"/Dest [{page_id[target - 1]} 0 R /Fit]")
            annots.append(f"<< /Type /Annot /Subtype /Link /Rect {rect} /Border [0 0 0] {act} >>")
        objs[aid] = ("[" + " ".join(annots) + "]").encode()
        objs[pid] = (f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {pw:.2f} {ph:.2f}] /Contents {cid} 0 R "
                     f"/Resources << /XObject << /Im0 {iid} 0 R >> >> /Annots {aid} 0 R >>").encode()
    out, offs = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n"), {}
    for oid in sorted(objs):
        offs[oid] = len(out)
        out += f"{oid} 0 obj\n".encode() + objs[oid] + b"\nendobj\n"
    xref = len(out)
    out += f"xref\n0 {info_id + 1}\n0000000000 65535 f \n".encode()
    out += b"".join(f"{offs[o]:010d} 00000 n \n".encode() for o in range(1, info_id + 1))
    out += f"trailer\n<< /Size {info_id + 1} /Root 1 0 R /Info {info_id} 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    with open(pg.lp(path), "wb") as f:
        f.write(out)


def render_brand(brand, cats, out_path, now=None):
    """Write one PDF via <out>.tmp + os.replace; returns (page count, link count)."""
    now = now or datetime.datetime.now()
    plan = layout(cats)
    total = 1 + sum(n for _c, _f, n in plan)
    pages = [cover_page(brand, cats, plan, total, now)]
    for cat, first, _n in plan:
        pages += section_pages(brand, cat, cats[cat], first, total)
    tmp = out_path + ".tmp"
    try:
        write_pdf(tmp, pages, f"{brand} — Godkända bilder")
        os.replace(pg.lp(tmp), pg.lp(out_path))
    except BaseException:
        if os.path.exists(pg.lp(tmp)):
            os.remove(pg.lp(tmp))
        raise
    return len(pages), sum(len(lk) for _im, lk in pages)


def summary(tree):
    out = {}
    for brand, cats in tree.items():
        plan = layout(cats)
        out[brand] = {"products": sum(len(v) for v in cats.values()),
                      "images": sum(p["count"] for v in cats.values() for p in v),
                      "pages": 1 + sum(n for _c, _f, n in plan),
                      "categories": {c: {"products": len(cats[c]), "first_page": f, "pages": n} for c, f, n in plan},
                      "missing_cover": [f"{c}\\{p['name']}" for c in cats for p in cats[c] if not p["cover"]]}
    return out


def main():
    for stream in (sys.stdout, sys.stderr):
        stream.reconfigure(errors="backslashreplace")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=("plan", "build"))
    ap.add_argument("--target", help="read this tree instead of Pictures\\QA-hubb Godkända (tests/preview)")
    ap.add_argument("--out", help="write the PDFs here instead of the target root (preview)")
    args = ap.parse_args()
    if args.target:
        target = os.path.abspath(args.target)
    else:
        ss = pg.load_sync()
        target = os.path.join(ss.DEFAULT_ONEDRIVE, ss.PICTURES_NAME, pg.TARGET_NAME)
        pg.guard_target(ss, target)
    if not os.path.isdir(pg.lp(target)):
        print(json.dumps({"ALARM": f"target missing: {target}"}, ensure_ascii=False))
        return 3
    out_dir = os.path.abspath(args.out) if args.out else target
    tree = scan(target)
    res = {"target": target, "out_dir": out_dir, "brands": summary(tree),
           "would_write" if args.cmd == "plan" else "written": [os.path.join(out_dir, PDF_NAME.format(brand=b)) for b in tree]}
    if args.cmd == "build":
        lock = None if args.target else pg.LOCK_FILE
        if lock:
            try:
                fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.write(fd, str(os.getpid()).encode())
                os.close(fd)
            except FileExistsError:
                print(f"lock held: {lock}", file=sys.stderr)
                return 4
        try:
            os.makedirs(pg.lp(out_dir), exist_ok=True)
            for brand, cats in tree.items():
                p = os.path.join(out_dir, PDF_NAME.format(brand=brand))
                n, n_links = render_brand(brand, cats, p)
                res["brands"][brand]["pdf_pages"] = n
                res["brands"][brand]["pdf_links"] = n_links
                res["brands"][brand]["pdf_bytes"] = os.path.getsize(pg.lp(p))
        finally:
            if lock:
                os.remove(lock)
    print(json.dumps(res, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
