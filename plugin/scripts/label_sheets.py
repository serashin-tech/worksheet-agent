"""콘티 PDF의 고정 5칸 격자에서 컷 번호 라벨이 있는 칸을 찾아 라벨 크롭 접촉 시트를 만든다.

OCR이 없어 컷 번호는 접촉 시트를 Claude가 읽어 labels.json에 옮긴다(쪽·칸 번호가 캡션).
사용: python -I label_sheets.py <pdf> <out_dir>
출력: slots.json(라벨이 있는 칸 목록), sheet_N.png
"""
import sys, json, pathlib
import numpy as np
import pymupdf
from PIL import Image, ImageDraw

DPI, Y0, PITCH = 110, 331, 432.5  # 110dpi 기준 첫 칸 상단과 칸 높이(5칸 고정 격자)
LABEL_X = (0.06, 0.115)
THRESH = 0.01
COLS, ROWS, SCALE = 6, 7, 2


def label_slots(doc):
    out = []
    for pn in range(len(doc)):
        pix = doc[pn].get_pixmap(dpi=DPI, colorspace=pymupdf.csGRAY)
        g = np.frombuffer(pix.samples, np.uint8).reshape(pix.h, pix.w)
        w = g.shape[1]
        x0, x1 = int(w * LABEL_X[0]), int(w * LABEL_X[1])
        for i in range(5):
            a = int(Y0 + PITCH * i)
            tag = g[a + 6:a + 70, int(w * 0.075):int(w * 0.108)]
            if (tag < 110).mean() > THRESH:
                out.append({"page": pn + 1, "slot": i + 1, "y": a, "crop": g[a + 4:a + 60, x0:x1]})
    return out


def main(pdf, out_dir):
    out_dir = pathlib.Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    slots = label_slots(pymupdf.open(pdf))
    meta = [{k: v for k, v in s.items() if k != "crop"} for s in slots]
    (out_dir / "slots.json").write_text(json.dumps(meta, ensure_ascii=False, indent=1))
    tw, th = (slots[0]["crop"].shape[1]) * SCALE, (slots[0]["crop"].shape[0]) * SCALE
    cap = 22
    per = COLS * ROWS
    for n in range(0, len(slots), per):
        sheet = Image.new("L", (COLS * tw, ROWS * (th + cap)), 255)
        dr = ImageDraw.Draw(sheet)
        for k, s in enumerate(slots[n:n + per]):
            r, c = divmod(k, COLS)
            im = Image.fromarray(s["crop"]).resize((tw, th))
            sheet.paste(im, (c * tw, r * (th + cap) + cap))
            dr.text((c * tw + 4, r * (th + cap) + 4), f"#{n + k + 1} p{s['page']}s{s['slot']}", fill=0)
        sheet.save(out_dir / f"sheet_{n // per + 1}.png")
    print("label slots:", len(slots), "sheets:", (len(slots) + per - 1) // per)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv[1], sys.argv[2])
