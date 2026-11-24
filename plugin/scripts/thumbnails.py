"""컷 번호 라벨이 있는 칸(= 컷의 첫 칸)을 잘라 썸네일로 저장하고, 정답 시트의 썸네일과 대조한다.

컷 번호는 label_sheets.py가 만든 접촉 시트를 사람이(또는 Claude가) 읽어 labels 인자로 준다.
EP12 라벨 규칙(접촉 시트 168칸을 육안으로 확인): 순번 k의 컷 번호는
  k<=11 -> C{k}, k==12 -> C11B, 13<=k<=139 -> C{k-1}, k==140 -> C0(번호 없는 도입 컷), k>=141 -> C{k-2}
사용: python -I thumbnails.py <pdf> <xlsx> <out_dir>
"""
import sys, io, json, pathlib
import numpy as np
import pymupdf
import openpyxl
from PIL import Image

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from label_sheets import label_slots, Y0, PITCH, DPI as LDPI

HI_DPI = 150
PIC_X = (0.115, 0.538)


def ep12_label(k):
    if k <= 11:
        return f"C{k:02d}"
    if k == 12:
        return "C11B"
    if k <= 139:
        return f"C{k - 1:02d}" if k - 1 < 100 else f"C{k - 1}"
    if k == 140:
        return "C0"
    n = k - 2
    return f"C{n:02d}" if n < 100 else f"C{n}"


def crop_first_panels(doc, slots, label=ep12_label):
    scale = HI_DPI / LDPI
    out = {}
    cache = {}
    for k, s in enumerate(slots, start=1):
        pn = s["page"] - 1
        if pn not in cache:
            pix = doc[pn].get_pixmap(dpi=HI_DPI, colorspace=pymupdf.csGRAY)
            cache[pn] = np.frombuffer(pix.samples, np.uint8).reshape(pix.h, pix.w)
        g = cache[pn]
        w = g.shape[1]
        a, b = int((s["y"] + 5) * scale), int((s["y"] + PITCH - 5) * scale)
        x0, x1 = int(w * PIC_X[0]) + 6, int(w * PIC_X[1]) - 6
        out[label(k)] = Image.fromarray(g[a:b, x0:x1]).resize((480, 270))
    return out


def tiny(im):
    a = np.asarray(im.convert("L").resize((48, 27)), float)
    a = (a - a.mean()) / (a.std() + 1e-6)
    return a.ravel()


def golden_images(xlsx):
    ws = openpyxl.load_workbook(xlsx).worksheets[0]
    res = []
    for im in ws._images:
        row = im.anchor._from.row + 1
        cid = ws.cell(row, 1).value
        res.append((cid, Image.open(io.BytesIO(im._data()))))
    return res


def main(pdf, xlsx, out_dir):
    out_dir = pathlib.Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open(pdf)
    slots = label_slots(doc)
    thumbs = crop_first_panels(doc, slots)
    for cid, im in thumbs.items():
        im.save(out_dir / f"{cid}.png")
    json.dump(list(thumbs), open(out_dir / "labels.json", "w", encoding="utf-8"), ensure_ascii=False)
    print("thumbnails:", len(thumbs))
    mine = {cid: tiny(im) for cid, im in thumbs.items()}
    gold = golden_images(xlsx)
    ok = 0
    bad = []
    base = lambda c: str(c).lower().rstrip("abcd_ch").upper().replace("_", "") if c else c
    for gid, gim in gold:
        gv = tiny(gim)
        best = max(mine, key=lambda c: float(np.dot(mine[c], gv)))
        gb = str(gid)[:4]  # C001 -> 'C001', 비교용 4자
        bb = best if best == "C0" else "C%03d" % int(best[1:].rstrip("B"))
        hit = bb == gb
        ok += hit
        if not hit:
            bad.append((gid, best))
    print(f"golden thumbnails={len(gold)} matched to same cut number: {ok}/{len(gold)} ({ok / len(gold):.0%})")
    print("mismatch:", bad[:30])


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(*sys.argv[1:4])
