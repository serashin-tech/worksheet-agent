"""스토리보드 PDF의 DIALOGUE·NOTE 열을 OCR로 읽어 칸(SHOT)마다 붙인다.

PDF는 글자 층이 없는 그림이라 Windows 내장 OCR(ocr_win.ps1, 한국어)로 읽는다.
한 문단은 그 문단의 첫 줄이 놓인 칸에 붙인다(긴 NOTE가 다음 칸까지 내려와도 시작한 칸의 값).
칸 순서와 컷 ID는 storyboard_panels.walk와 같아서 panels.json의 그림과 짝이 맞는다.
OCR은 오타가 남는다(예: '응?'→'으7'). 콘솔에서는 PD가 확인할 초안으로만 쓴다.
사용: python -I storyboard_text.py <pdf> <out_json>
      python -I storyboard_text.py --attach <out_json> <storyboard_panels 출력> <콘솔 panels.json>
        콘솔 panels.json(C11B·C0을 앞 컷에 합친 판)에 칸 그림이 같은 자리끼리 짝지어 "text"를 넣는다.
출력: {"C001": [{"dia": "...", "note": "..."}, ...], ...}  (panels.json의 panels와 같은 키·길이)
"""
import sys, re, json, pathlib, subprocess, tempfile
import pymupdf

HERE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from label_sheets import DPI as LDPI, Y0, PITCH
from storyboard_panels import walk

DPI = 300
COLS = {"dia": (0.536, 0.680), "note": (0.681, 0.885)}  # 쪽 너비 대비 DIALOGUE·NOTE 열(EP12 육안 확인)
TOP, BOTTOM = 0.125, 0.97  # 머리글 아래부터 꼬리말 위까지


def keep(text, col):
    """그림 속 손글씨(PAN, c31A, 숫자 조각)를 거른다. DIALOGUE는 한글이 있는 줄만, NOTE는 한글이나 영문 3자 이상((WIPE))이 있는 줄만."""
    return bool(re.search("[가-힣]", text)) or (col == "note" and len(re.findall("[A-Za-z]", text)) >= 3)


def paragraphs(lines):
    """OCR 줄 -> [(첫 줄 y, 문단 글)]. 같은 높이의 조각은 한 줄로, 줄 간격이 크게 벌어지면 새 문단."""
    rows = []
    for l in sorted(lines, key=lambda l: (l["y"], l["x"])):
        if rows and abs(l["y"] - rows[-1][0]) < max(l["h"], 10) / 2:
            rows[-1][1].append(l)
        else:
            rows.append([l["y"], [l]])
    out, prev = [], None
    for y, parts in rows:
        text = " ".join(p["text"] for p in sorted(parts, key=lambda p: p["x"]))
        new_cut = re.match(r"\d{1,3}\.", text)  # NOTE의 '32. F.S. …'는 컷 첫 줄이라 줄 간격이 좁아도 새 문단(EP12 C31 '(WIPE)' 아래 C32)
        if prev is not None and y - prev < DPI * 0.3 and not new_cut:  # 줄 간격(약 0.2인치)보다 조금 넓게
            out[-1][1] += "\n" + text
        else:
            out.append([y, text])
        prev = y
    return out


def sentences(text):
    """NOTE 문단: OCR 줄바꿈을 지우고 한 문장씩 한 줄로 둔다. '#장소' 줄은 따로 둔다.
    한글 뒤의 . ! ? 다음에서 끊되, '에디. (뒷모습)'처럼 괄호가 이어지면 붙여 둔다. 'B.S.' 같은 영문 약어에서는 끊지 않는다."""
    out, buf = [], []
    for line in text.split("\n") + ["#"]:
        if line.startswith("#"):
            joined = " ".join(" ".join(buf).split())
            out += [x for x in re.split(r"(?<=[가-힣][.!?])\s+(?!\()", joined) if x]
            out.append(" ".join(line.split())); buf = []
        else:
            buf.append(line)
    return "\n".join(out[:-1])


def band(y_px, page_h):
    """crop 안의 y(px) -> 칸 번호(0~4).
    칸의 글은 칸 윗선에 거의 붙어 시작해서, OCR 상자가 조금만 위로 잡혀도 앞 칸으로 넘어간다(EP12 C17·C21).
    그래서 칸 높이의 6%만큼 아래로 당겨 정한다. 칸 맨 아래 6%에서 시작하는 글은 없다고 본다."""
    y_pt = y_px / DPI * 72 + page_h * TOP
    return int((y_pt * LDPI / 72 - Y0) / PITCH + 0.06)


def ocr(pdf):
    """-> {(쪽, 칸): {"dia": 글, "note": 글}}"""
    doc, tmp = pymupdf.open(pdf), pathlib.Path(tempfile.mkdtemp(prefix="sbocr_"))
    paths = []
    for pn, page in enumerate(doc):
        w, h = page.rect.width, page.rect.height
        for col, (a, b) in COLS.items():
            f = tmp / f"{pn:03d}_{col}.png"
            page.get_pixmap(dpi=DPI, clip=pymupdf.Rect(w * a, h * TOP, w * b, h * BOTTOM), colorspace=pymupdf.csGRAY).save(f)
            paths.append(str(f))
    (tmp / "list.txt").write_text("\n".join(paths), encoding="utf-8")
    subprocess.run(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(HERE / "ocr_win.ps1"),
                    str(tmp / "list.txt"), str(tmp / "out.json")], check=True)
    cells = {}
    for r in json.loads((tmp / "out.json").read_text(encoding="utf-8-sig")):
        name = pathlib.Path(r["path"]).stem
        pn, col = int(name[:3]), name[4:]
        for y, text in paragraphs([l for l in r["lines"] or [] if keep(l["text"], col)]):
            c = cells.setdefault((pn, band(y, doc[pn].rect.height)), {"dia": "", "note": ""})
            c[col] = (c[col] + "\n" + (sentences(text) if col == "note" else text)).strip()
    return cells


def main(pdf, out, label=None):
    cells = ocr(pdf)
    res = {}
    for pn, i, cut, _ in walk(pymupdf.open(pdf), *([label] if label else [])):
        res.setdefault(cut, []).append(cells.pop((pn, i), {"dia": "", "note": ""}))
    print("칸에 붙지 않은 글(빈 칸 높이에서 시작):", {k: v for k, v in cells.items() if any(v.values())})
    json.dump(res, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    n = sum(len(v) for v in res.values())
    print(f"cuts={len(res)} panels={n} dia={sum(bool(p['dia']) for v in res.values() for p in v)} note={sum(bool(p['note']) for v in res.values() for p in v)}")


def attach(text, raw, target):
    """raw(storyboard_panels 출력)와 text는 같은 키·순서. target의 칸 그림마다 같은 그림의 글을 붙인다."""
    by_img = {img: t for k, imgs in raw["panels"].items() for img, t in zip(imgs, text[k])}
    target["text"] = {k: [by_img.get(img, {"dia": "", "note": ""}) for img in imgs] for k, imgs in target["panels"].items()}
    return sum(img in by_img for imgs in target["panels"].values() for img in imgs)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    if sys.argv[1:2] == ["--test"]:
        ps = paragraphs([{"text": "에디:", "x": 8, "y": 3770, "h": 30}, {"text": "응?", "x": 152, "y": 3772, "h": 30},
                         {"text": "에디: 성공이야~!", "x": 8, "y": 3650, "h": 30}, {"text": "뽀로로: 그게 뭐야?", "x": 6, "y": 3710, "h": 30},
                         {"text": "SE: 띠리링~", "x": 5, "y": 76, "h": 30}])
        assert ps == [[76, "SE: 띠리링~"], [3650, "에디: 성공이야~!\n뽀로로: 그게 뭐야?\n에디: 응?"]], ps
        assert paragraphs([{"text": "(WIPE)", "x": 0, "y": 100, "h": 30}, {"text": "32. F.S. 약부감.", "x": 0, "y": 160, "h": 30}]) == [[100, "(WIPE)"], [160, "32. F.S. 약부감."]]
        assert band(76, 1683.84) == 0 and band(3650, 1683.84) == 3, (band(76, 1683.84), band(3650, 1683.84))
        assert [keep(t, c) for t, c in [("PAN", "dia"), ("c31A", "dia"), ("SE: 띠리링~", "dia"), ("(WIPE)", "note"), ("AN", "note"), ("6", "note")]] == [False, False, True, True, False, False]
        assert sentences("#에디 집 내부\n계기판에 연결된 네모모양의 절제\n상자가 보인다. 계기판 앞에서 버튼을 늘러\n조작하는 에디. (뒷모습)") == \
            "#에디 집 내부\n계기판에 연결된 네모모양의 절제 상자가 보인다.\n계기판 앞에서 버튼을 늘러 조작하는 에디. (뒷모습)"
        assert sentences("3. 계기판 B.S. 에디가 버튼을\n누른다. 고른다.") == "3. 계기판 B.S. 에디가 버튼을 누른다.\n고른다."
        print("ok")
        t = attach({"C1": [{"dia": "a", "note": ""}], "C1B": [{"dia": "b", "note": "n"}]}, {"panels": {"C1": ["x"], "C1B": ["y"]}}, tg := {"panels": {"C1": ["x", "y"]}})
        assert t == 2 and tg["text"]["C1"][1] == {"dia": "b", "note": "n"}, tg
        print("ok attach")
    elif sys.argv[1:2] == ["--attach"]:
        text, raw, tp = json.load(open(sys.argv[2], encoding="utf-8")), json.load(open(sys.argv[3], encoding="utf-8")), sys.argv[4]
        target = json.load(open(tp, encoding="utf-8"))
        n = attach(text, raw, target)
        json.dump(target, open(tp, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
        print(f"짝지은 칸 {n}/{sum(len(v) for v in target['panels'].values())}")
    else:
        main(sys.argv[1], sys.argv[2])
