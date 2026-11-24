"""스토리보드 PDF의 모든 칸(PICTURE)을 컷별로 묶어 panels.json으로 저장한다.

PDF는 쪽마다 5칸 고정 격자다. CUT 열에 번호가 있는 칸이 새 컷의 첫 칸이고,
번호가 없는 칸은 앞 컷의 이어지는 칸이다. 컷은 칸 순서대로 a, b, c로 부른다.
컷 번호는 thumbnails.ep12_label 규칙(EP12 육안 확인)을 따른다.
사용: python -I storyboard_panels.py <pdf> <out_json>
출력: {"C001": [jpeg base64, ...], "C11B": [...], "C0": [...]}  (키는 콘솔의 컷 ID, 추가 컷은 C11B·C0)
"""
import sys, io, json, base64, pathlib
import numpy as np
import pymupdf
from PIL import Image

HERE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from label_sheets import DPI as LDPI, Y0, PITCH, LABEL_X, THRESH
from thumbnails import ep12_label, PIC_X

HI_DPI = 150
W, H = 240, 135


def key(label):
    """thumbnails 라벨(C01, C100)을 콘솔 ID(C001, C100)로. C11B와 C0은 그대로 둔다."""
    if label in ("C11B", "C0"):
        return label
    return "C%03d" % int(label[1:])


def walk(doc):
    """그림이 있는 칸마다 (쪽 번호, 칸 번호, 컷 ID, 150dpi 그림 crop)을 PDF 순서대로 낸다."""
    scale = HI_DPI / LDPI
    k, cur = 0, None
    for pn in range(len(doc)):
        lo = np.frombuffer((p := doc[pn].get_pixmap(dpi=LDPI, colorspace=pymupdf.csGRAY)).samples, np.uint8).reshape(p.h, p.w)
        hi = np.frombuffer((q := doc[pn].get_pixmap(dpi=HI_DPI, colorspace=pymupdf.csGRAY)).samples, np.uint8).reshape(q.h, q.w)
        w = lo.shape[1]
        for i in range(5):
            a = int(Y0 + PITCH * i)
            if a + PITCH > lo.shape[0] + 40:
                break
            tag = lo[a + 6:a + 70, int(w * 0.075):int(w * 0.108)]
            start = (tag < 110).mean() > THRESH
            ya, yb = int((a + 5) * scale), int((a + PITCH - 5) * scale)
            x0, x1 = int(hi.shape[1] * PIC_X[0]) + 6, int(hi.shape[1] * PIC_X[1]) - 6
            crop = hi[ya:yb, x0:x1]
            if not start and (crop < 110).mean() < 0.002:  # 그림이 없는 빈 칸
                continue
            if start:
                k += 1
                cur = key(ep12_label(k))
            if cur is not None:
                yield pn, i, cur, crop


def main(pdf, out):
    groups, order = {}, []
    for _, _, cur, crop in walk(pymupdf.open(pdf)):
        if cur not in groups:
            groups[cur] = []
            order.append(cur)
        buf = io.BytesIO()
        Image.fromarray(crop).resize((W, H)).save(buf, "JPEG", quality=70)
        groups[cur].append(base64.b64encode(buf.getvalue()).decode())
    json.dump({"order": order, "panels": groups}, open(out, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    multi = {c: len(v) for c, v in groups.items() if len(v) > 1}
    print(f"cuts={len(groups)} panels={sum(len(v) for v in groups.values())} multi={len(multi)}")
    print("C001", len(groups.get("C001", [])), "C003", len(groups.get("C003", [])), "C004", len(groups.get("C004", [])), "C005", len(groups.get("C005", [])))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv[1], sys.argv[2])
