"""PoC: 콘티 PDF(쪽마다 전체가 래스터 한 장)에서 콘티 칸과 컷 시작을 찾는다.

텍스트·벡터가 없어서 이미지 분석만 쓴다.
- 칸 경계: 그림 열에서 가로로 길게 이어진 검은 선
- 컷 시작: 왼쪽 CUT 열의 칸 윗부분에 글자(잉크)가 있으면 새 컷
사용: python -I panel_split.py <pdf> <out_dir> [첫쪽 끝쪽]
"""
import sys, json, pathlib
import numpy as np
import pymupdf
from PIL import Image

DPI = 110
# 그림 열(PDF pt 기준 비율). 쪽 크기 1190x1684pt, 그림 열은 약 x 136~640pt
PIC_X = (0.115, 0.538)
CUT_X = (0.07, 0.113)  # CUT 열(컷 번호 글자)


def page_gray(page):
    pix = page.get_pixmap(dpi=DPI, colorspace=pymupdf.csGRAY)
    return np.frombuffer(pix.samples, np.uint8).reshape(pix.h, pix.w), pix


def find_lines(gray):
    h, w = gray.shape
    x0, x1 = int(w * (PIC_X[0] + 0.01)), int(w * (PIC_X[1] - 0.01))
    dark = (gray[:, x0:x1] < 90).mean(axis=1)
    rows = np.where(dark > 0.85)[0]
    # 굵은 선 하나가 여러 행이므로 묶는다
    groups, cur = [], [rows[0]] if len(rows) else []
    for r in rows[1:]:
        if r - cur[-1] <= 3:
            cur.append(r)
        else:
            groups.append(cur); cur = [r]
    if cur:
        groups.append(cur)
    return [int(np.mean(g)) for g in groups]


def split_page(gray):
    h, w = gray.shape
    ys = [y for y in find_lines(gray) if y > h * 0.12]  # 머리글 아래만
    panels = []
    for a, b in zip(ys, ys[1:]):
        if b - a < h * 0.1:  # 너무 얇으면 칸이 아님
            continue
        panels.append((a, b))
    cx0, cx1 = int(w * CUT_X[0]), int(w * CUT_X[1])
    out = []
    for a, b in panels:
        top = gray[a + 3:a + int((b - a) * 0.18), cx0:cx1]
        ink = (top < 110).mean()
        sub = gray[a + 4:b - 4, int(w * PIC_X[0]) + 4:int(w * PIC_X[1]) - 4]
        # 빗금(미사용 칸) 판정은 후속 과제. 지금은 잉크량만 기록
        out.append({"y0": a, "y1": b, "cut_start": bool(ink > 0.01), "ink": round(float((sub < 110).mean()), 4)})
    return out


def main():
    pdf, out_dir = sys.argv[1], pathlib.Path(sys.argv[2])
    out_dir.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open(pdf)
    first = int(sys.argv[3]) if len(sys.argv) > 3 else 0
    last = int(sys.argv[4]) if len(sys.argv) > 4 else len(doc)
    report = []
    for pn in range(first, last):
        gray, pix = page_gray(doc[pn])
        panels = split_page(gray)
        w = gray.shape[1]
        for i, p in enumerate(panels):
            crop = gray[p["y0"] + 4:p["y1"] - 4, int(w * PIC_X[0]) + 4:int(w * PIC_X[1]) - 4]
            Image.fromarray(crop).save(out_dir / f"p{pn + 1:02d}_{i + 1}.png")
        report.append({"page": pn + 1, "panels": panels})
    (out_dir / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")
    n_pan = sum(len(r["panels"]) for r in report)
    n_cut = sum(p["cut_start"] for r in report for p in r["panels"])
    print(f"pages={len(report)} panels={n_pan} cut_starts={n_cut}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
