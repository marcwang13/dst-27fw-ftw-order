#!/usr/bin/env python3
"""從 DESCENTE 27FW 鞋類原廠訂單本重產 data.js、imgs.js、big/。

用法：python3 tools/gen_ftw.py /path/to/order-sheet.xlsx

只讀工作表 27FW。尺寸以表頭 220–300 欄的灰格當作沒開的尺寸。
售價只用 TW-SRP（台幣建議售價）。原廠下單日、PRE/REGULAR 欄不寫進商品資料。
"""
import hashlib
import json
import sys
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
SIZES = list(range(220, 301, 5))

COLOR_ZH = {
    "WHT0": "白色",
    "BLK0": "黑色",
    "BKWH": "黑白",
    "BLCK": "黑色",
    "BKBE": "黑米",
    "PKWH": "粉白",
    "GNWH": "綠白",
    "BEWH": "米白",
    "GRY0": "灰色",
    "SAND": "沙色",
    "DTP0": "印花",
}
ORIGIN_ZH = {
    "Vietnam": "越南",
    "South Korea": "韓國",
    "China": "中國",
    "Korea": "韓國",
}
DEV_ZH = {"C/O": "延續款", "NU": "新款", "CU": "新配色"}
LINE_ZH = {"PERFORMANCE": "性能", "MOVEWEAR": "生活"}

# 英文品名是型號，中文只補品類，不硬翻專有名詞。
NAME = {
    "ENERZITE SUPER V6": ("ENERZITE SUPER V6", "ENERZITE SUPER V6 跑鞋", "跑鞋", "跑鞋"),
    "EZ-SUPER ZEROMOTION": ("EZ-SUPER ZEROMOTION", "EASY SUPER ZEROMOTION 跑鞋", "跑鞋", "跑鞋"),
    "RUN STEP V3": ("RUN STEP V3", "RUN STEP V3 跑鞋", "跑鞋", "跑鞋"),
    "CHRON RACER": ("CHRON RACER", "CHRON RACER 競速鞋", "休閒鞋", "競速鞋"),
    "CHRON TEMPO SLIP-ON": ("CHRON TEMPO SLIP-ON", "CHRON TEMPO 懶人鞋", "休閒鞋", "懶人鞋"),
    "DRIBBLE NEO": ("DRIBBLE NEO", "DRIBBLE NEO 休閒鞋", "休閒鞋", "休閒鞋"),
    "KINETIC MJ": ("KINETIC MJ", "KINETIC MJ 休閒鞋", "休閒鞋", "休閒鞋"),
    "DUOREST SLIDE": ("DUOREST SLIDE", "DUOREST 拖鞋", "拖鞋", "拖鞋"),
    "ICINGNAS+": ("ICINGNAS+", "ICINGNAS+ 拖鞋", "拖鞋", "拖鞋"),
    "TEMPERPRO SLIDE": ("TEMPERPRO SLIDE", "TEMPERPRO 拖鞋", "拖鞋", "拖鞋"),
}


def en_key(style_name):
    s = str(style_name or "")
    if "(" in s:
        s = s.split("(", 1)[0]
    return " ".join(s.replace("＋", "+").split()).strip()


def ymd(v):
    if isinstance(v, datetime):
        return v
    return None


def fmt_date(d):
    return f"{d.year}/{d.month:02d}/{d.day:02d}"


def mon_label(d):
    return f"{d.month}月"


def launch_label(d):
    return f"{d.year}年{d.month}月"


def blocked_sizes(ws, row):
    blocked = []
    for i, sz in enumerate(SIZES):
        cell = ws.cell(row, 20 + i)  # T=220 … AJ=300
        if cell.fill and cell.fill.patternType == "solid":
            blocked.append(sz)
    return blocked


def flatten(im, max_w, quality):
    im = im.convert("RGBA")
    bg = Image.new("RGB", im.size, (255, 255, 255))
    bg.paste(im, mask=im.split()[-1])
    if bg.width > max_w:
        h = max(1, round(bg.height * max_w / bg.width))
        bg = bg.resize((max_w, h), Image.Resampling.LANCZOS)
    buf = BytesIO()
    bg.save(buf, format="JPEG", quality=quality, optimize=True)
    return buf.getvalue()


def main():
    if len(sys.argv) != 2:
        print("用法：python3 tools/gen_ftw.py order-sheet.xlsx", file=sys.stderr)
        sys.exit(2)
    src = Path(sys.argv[1])
    wb = load_workbook(src, data_only=False)
    ws = wb["27FW"]

    by_row = {}
    for im in ws._images:
        er = im.anchor._from.row + 1
        data = im._data()
        by_row.setdefault(er, []).append(data)

    big_dir = ROOT / "big"
    big_dir.mkdir(exist_ok=True)
    for old in big_dir.glob("*.jpg"):
        old.unlink()

    items = []
    img_map = {}
    seq = {}
    unused = []
    for row in range(8, ws.max_row + 1):
        style = ws.cell(row, 2).value
        if not style:
            continue
        key = en_key(ws.cell(row, 11).value)
        if key not in NAME:
            raise SystemExit(f"第 {row} 列品名沒有對照：{ws.cell(row, 11).value!r}")
        en, zh, sheet, cat = NAME[key]
        cc = str(ws.cell(row, 4).value or "").strip()
        if cc not in COLOR_ZH:
            raise SystemExit(f"第 {row} 列色號沒有中文：{cc}")
        gender = str(ws.cell(row, 13).value or "").strip()
        gen = {"U": "中性", "W": "女", "M": "男"}.get(gender)
        if not gen:
            raise SystemExit(f"第 {row} 列性別沒有對照：{gender}")
        launch = ymd(ws.cell(row, 18).value)
        if not launch:
            raise SystemExit(f"第 {row} 列沒有上市日")
        price = ws.cell(row, 39).value
        if not isinstance(price, (int, float)):
            raise SystemExit(f"第 {row} 列沒有台幣建議售價")
        blocked = set(blocked_sizes(ws, row))
        sizes = [str(sz) for sz in SIZES if sz not in blocked]
        if not sizes:
            raise SystemExit(f"第 {row} 列沒有可下尺寸")
        seq[sheet] = seq.get(sheet, 0) + 1
        n = seq[sheet]
        blobs = sorted(by_row.get(row, []), key=len, reverse=True)
        doll = ""
        if blobs:
            big = flatten(Image.open(BytesIO(blobs[0])), 1000, 82)
            doll = hashlib.md5(big).hexdigest()[:12]
            (big_dir / f"{doll}.jpg").write_bytes(big)
            thumb = flatten(Image.open(BytesIO(blobs[0])), 240, 70)
            img_map[doll] = "data:image/jpeg;base64," + __import__("base64").b64encode(thumb).decode()
            if len(blobs) > 1:
                unused.append((row, style, cc, len(blobs) - 1))
        line = str(ws.cell(row, 6).value or "").strip()
        silo = str(ws.cell(row, 7).value or "").strip()
        dev = DEV_ZH.get(str(ws.cell(row, 8).value or "").strip(), str(ws.cell(row, 8).value or "").strip())
        origin = ORIGIN_ZH.get(str(ws.cell(row, 15).value or "").strip(), str(ws.cell(row, 15).value or "").strip())
        run = str(ws.cell(row, 12).value or "").strip()
        etd = launch + timedelta(days=14)
        items.append({
            "id": f"{sheet}-{n}",
            "sheet": sheet,
            "no": str(n),
            "mon": mon_label(launch),
            "theme": silo,
            "cat": cat,
            "gen": gen,
            "style": str(style).strip(),
            "zh": zh,
            "en": en,
            "color": COLOR_ZH[cc],
            "cc": cc,
            "sku": f"{str(style).strip()}{cc}",
            "fab": "",
            "tech": "、".join(x for x in [dev, f"產地{origin}" if origin else ""] if x),
            "series": f"{LINE_ZH.get(line, line)}｜{silo}",
            "launch": launch_label(launch),
            "eta": fmt_date(launch),
            "etd": fmt_date(etd),
            "price": int(price),
            "sizes": sizes,
            "doll": doll,
            "run": "FTW",
            "div": "FTW",
            "ov": f"尺寸 {run}" if run else "",
        })

    # 分頁順序：跑鞋、休閒鞋、拖鞋，同頁依表上的列序（已是）
    order = {"跑鞋": 0, "休閒鞋": 1, "拖鞋": 2}
    items.sort(key=lambda it: (order[it["sheet"]], int(it["no"])))
    # 重新編號，避免排序後 no 跳號；目前各頁本身已依列序遞增，排序不改組內順序
    (ROOT / "data.js").write_text(
        "window.SDR_ITEMS=" + json.dumps(items, ensure_ascii=False, separators=(", ", ": ")) + ";\n",
        encoding="utf-8",
    )
    (ROOT / "imgs.js").write_text(
        "window.SDR_IMG=" + json.dumps(img_map, ensure_ascii=False, separators=(",", ":")) + ";\n",
        encoding="utf-8",
    )
    print(f"款色 {len(items)}　圖 {len(img_map)}")
    from collections import Counter
    print("分頁", Counter(i["sheet"] for i in items))
    if unused:
        print("同列多出來、沒採用的圖（已用該列最大的一張）:")
        for u in unused:
            print(" ", u)


if __name__ == "__main__":
    main()
