"""Phase 2: 컷 목록(대본 + 규칙 + PD 결정)으로 Worksheet xlsx를 조립한다. Drive에 올리면 Google Sheet로 변환된다.

계산은 모두 이 스크립트가 한다(AI가 숫자를 쓰지 않는다).
- Start Frame(기본 101)과 End Frame은 값이고, Frame = End - Start + 1, Sec = ROUND(Frame/FPS,1)은 수식이다(EP12 시트와 같음). FPS는 overrides의 fps, 없으면 29.97
- Frame 추정값은 End Frame에 들어가며 노란색으로 표시한다. PD가 End를 고치면 일반색으로 바꿔도 된다.
- 요약은 별도 '요약' 탭에 둔다(2026-11-23 협의). 시트 맨 아래 요약 행은 만들지 않는다. 범위는 마지막 행까지 동적이다.
- 방영 길이(분)와 목표 컷 수(EP12는 166)는 시나리오와 함께 받는 고정 입력이다. 행 수(분할 포함)는 PD가 썸네일을 보고 분할을 정해야 확정되는 결과라 입력이 아니다. 요약 탭이 총 러닝타임과 컷 수를 이 값과 견준다.
- 공정 열(LO·AN·Render·Comp)은 드롭다운만 만들고 작업자는 비워 둔다.
사용: python -I build_worksheet.py <docx> <thumbs_dir> <asset_xlsx> <out.xlsx> [--ep 12] [--overrides o.json] [--fit golden.xlsx] [--runtime-min 10] [--target-cuts 166] [--tol-min 0.5]
overrides.json: {"splits": {"11": 2}, "frames": {"C001": 120}, "estimated": ["C001"], "starts": {"C001b": 151}, "fps": 24, "workers": {...}, "lists": {"an": ["송민주", ...]}}  (게이트 1·2에서 PD가 정한 값)
  lists: 작업현황 드롭다운 목록(status·rstatus·ver·an·render·comp). 콘솔의 '작업현황 드롭다운 목록'에서 PD가 고친 값이고, 없는 키는 LISTS 기본값을 쓴다. cam(Cam Move)은 콘솔 선택지용이라 시트에 드롭다운을 걸지 않는다(한 칸에 여러 값).
"""
import re, sys, os, json, argparse, pathlib
import numpy as np
import openpyxl
from openpyxl.drawing.image import Image as XlImage
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

HERE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from parse_docx import parse, DEFAULT_SPEAKERS
from extract_meta import cam_moves, SAME as SAME_CAM
from extract_split import split_count, same_camera
from extract_assets import load_assets, bg_codes, ch_codes, pr_codes, SKY

FPS, START = 29.97, 101
MODEL = HERE / "frame_model.json"
YELLOW = PatternFill("solid", fgColor="FFF2CC")
GRAY = PatternFill("solid", fgColor="CCCCCC")
STATUS = ["수정", "완료", "진행중", "x", "101f 완료"]
RSTATUS = ["수정", "완료", "진행중", "완료/FX대기", "X"]
WORKERS = {"an": ["송민주", "김은엽", "임선형", "윤지민", "탁진영"], "render": ["김지희", "심미림"], "comp": ["정승환", "김유진"]}
VERS = ["V01", "V02", "V03", "V04", "V05"]

# 열 배치: (제목, 종류, 목록). 제목이 빈 칸이면 앞 그룹의 보조 열.
FIXED = ["", "Thumbnail", "Cam\nMove", "Same\nCamera", "Sec", "Frame", "Start\nFrame", "End Frame", "BG", "CH", "PR", "CH_Count", "Dialogue", "Description"]
PROC = [("LO", "status"), ("작업자", "an"), ("요청사항", None)]
for k in range(1, 5):
    PROC += [(f"AN TK0{k}", "status"), ("작업자", "an"), ("요청사항", None)]
PROC += [("AN 완료", "ver"), ("Render TK01", "status"), ("작업자", "render"), ("요청사항", None)]
for k in range(2, 5):  # EP12 시트: Render TK01·TK02는 LO·AN 상태 목록, TK03부터 Render·Comp 상태 목록
    PROC += [(f"Render TK0{k}", "status" if k == 2 else "rstatus"), ("요청사항", None)]
for k in range(1, 6):
    PROC += [(f"Comp TK0{k}", "rstatus"), ("작업자", "comp")]
LISTS = {"status": STATUS, "rstatus": RSTATUS, "ver": VERS, "an": WORKERS["an"], "render": WORKERS["render"], "comp": WORKERS["comp"]}


def picklists(overrides):
    """기본 목록에 PD가 고친 목록을 덮는다. 엑셀 목록 드롭다운은 '항목,항목' 문자열이 255자를 넘거나 항목에 쉼표가 있으면 깨진다."""
    lists = {**LISTS, **overrides.get("lists", {})}
    for k, v in lists.items():
        s = ",".join(v)
        if not v or any("," in x for x in v) or len(s) > 255:
            raise ValueError(f"드롭다운 목록 {k}: 비어 있거나 쉼표가 들었거나 255자를 넘습니다 ({s[:40]})")
    return lists


def ranges(nums):
    nums, out = sorted(nums), []
    for n in nums:
        if out and n == out[-1][1] + 1:
            out[-1][1] = n
        else:
            out.append([n, n])
    return ", ".join(f"{a:02d}" if a == b else f"{a:02d}-{b:02d}" for a, b in out)


def resolve(code, table):
    """규칙이 만든 코드를 확정된 Asset 코드로 맞춘다. 못 찾으면 None."""
    codes = list(table.values())
    for pred in (lambda c: c == code, lambda c: c.lower() == code.lower(), lambda c: c.lower().endswith(code.lower()),
                 lambda c: c.lower().startswith(code.lower())):
        hit = [c for c in codes if pred(c)]
        if hit:
            return hit[0]
    return None


def features(c, ch_count):
    t = c["dialogue"]
    return [1, len(t), len(c["description"]), int(bool(cam_moves(c["description"]))), ch_count, int("SE" in t)]


def derive(docx_path, asset_xlsx, overrides):
    parsed = parse(docx_path, DEFAULT_SPEAKERS + ["노래", "다같이"])
    assets = load_assets(asset_xlsx)
    cuts = parsed["cuts"]
    rows, unknown, prev_place = [], set(), None
    for c in cuts:
        n = c["cut"]
        k = int(overrides.get("splits", {}).get(str(n), split_count(c["description"])))
        text = c["description"] + " " + c["dialogue"]
        mine, babies = ch_codes(text)
        ch = [code for _, code in [("", x) for x in mine]]
        if babies:
            ch.insert(0, "BabyPororo_Christmas_" + ranges(babies))
        bg = []
        for code in bg_codes(c["place"], text):
            hit = resolve(code, assets["BG"]) or resolve(code, assets["Matte"])
            if hit:
                bg.append(hit)
            else:
                unknown.add(code)
        pr = pr_codes(c["description"], assets["PR"])
        cam = cam_moves(c["description"] + " " + c["transition"])
        ref = SAME_CAM.search(c["description"])
        if ref and not cam and int(ref.group(1)) <= len(cuts):
            p = cuts[int(ref.group(1)) - 1]
            cam = cam_moves(p["description"] + " " + p["transition"])
        place = c["place"] if c["place"] != prev_place else ""
        prev_place = c["place"]
        desc = (place + "\n" if place else "") + c["description"] + (("\n" + c["transition"]) if c["transition"] else "")
        for i in range(k):
            cid = "C%03d" % n + ("abcdefghij"[i] if k > 1 else "")  # 콘솔은 컷 하나를 10행까지 나눈다
            rows.append({"cut": n, "id": cid, "first": i == 0, "cam": cam if i == 0 else [], "same": same_camera(c["description"]) if i == 0 else "",
                         "bg": bg, "ch": ch, "pr": pr, "count": len(mine) + len(babies), "dialogue": c["dialogue"] if i == 0 else "",
                         "description": desc if i == 0 else "", "feat": features(c, len(mine) + len(babies)) if i == 0 else features({"dialogue": "", "description": ""}, 0),
                         "split": k > 1, "dup": bool(ref)})
    return parsed, rows, sorted(unknown)


def fit(golden_xlsx, docx_path, asset_xlsx):
    """정답 시트의 일반 행(분할 제외)으로 Frame 선형 모델 계수를 맞춰 저장한다."""
    _, rows, _ = derive(docx_path, asset_xlsx, {})
    gold = {}
    ws = openpyxl.load_workbook(golden_xlsx).worksheets[0]
    for r in range(2, 181):
        gold[str(ws.cell(r, 1).value)] = ws.cell(r, 8).value - ws.cell(r, 7).value + 1 if isinstance(ws.cell(r, 8).value, (int, float)) else None
    X, y = [], []
    for r in rows:
        g = gold.get(r["id"])
        if g and not r["split"]:
            X.append(r["feat"]); y.append(g)
    w, *_ = np.linalg.lstsq(np.array(X, float), np.array(y, float), rcond=None)
    MODEL.write_text(json.dumps({"coef": list(map(float, w)), "n": len(y), "note": "EP12 일반 행 선형 회귀. 초안 값으로만 쓴다."}, indent=1))
    return w


def estimate(feat):
    w = json.loads(MODEL.read_text())["coef"]
    return int(min(300, max(24, round(float(np.dot(w, feat))))))


def shrink(data, px=480):
    """렌더 이미지(EP12는 1920x1080 PNG 63장, 29MB)를 셀 크기에 맞게 줄인다. 크기는 anchor가 정하므로 화면 모양은 같다.
    시트가 Drive 커넥터 한도(요청 본문 10MB, 파일 약 7MB) 안에 들어와야 01_산출에 Google Sheet로 올릴 수 있다."""
    from io import BytesIO
    from PIL import Image as PILImage
    im, out = PILImage.open(BytesIO(data)), BytesIO()
    im.thumbnail((px, px))
    if im.mode in ("RGBA", "LA", "P"):
        im.save(out, "PNG", optimize=True)
    else:
        im.convert("RGB").save(out, "JPEG", quality=85)
    out.seek(0)
    return out


def copy_asset_sheet(asset_xlsx, wb):
    """에셋 목록 파일의 Asset 시트(둘째 탭)를 값·서식·열 폭·행 높이·숨긴 열·렌더 이미지째 옮긴다.
    EP12 시트의 Asset 탭과 같은 모양(제목, 머리글 9열, BG·CH·PR·Matte·2D 분류 행, 리깅 범례 K2:N3)을 그대로 쓰기 위해서다.
    ponytail: 입력 파일이 이 형식이라는 전제(load_assets와 같음). 폴더에서만 에셋을 받는 경우는 그때 서식을 새로 그린다."""
    from copy import copy
    from io import BytesIO
    src = openpyxl.load_workbook(asset_xlsx).worksheets[1]
    wa = wb.create_sheet("Asset")
    for row in src.iter_rows():
        for s in row:
            if s.value is None and not s.has_style:
                continue
            d = wa.cell(s.row, s.column, s.value)
            if s.has_style:
                d.font, d.fill, d.border, d.alignment = copy(s.font), copy(s.fill), copy(s.border), copy(s.alignment)
                d.number_format, d.protection = s.number_format, copy(s.protection)
            if s.hyperlink:
                d.hyperlink = copy(s.hyperlink)
    for k, dim in src.column_dimensions.items():
        cd = wa.column_dimensions[k]
        cd.width, cd.hidden, cd.min, cd.max = dim.width, dim.hidden, dim.min, dim.max
    for r, dim in src.row_dimensions.items():
        if dim.height:
            wa.row_dimensions[r].height = dim.height
    for im in src._images:
        x = XlImage(shrink(im._data()))
        x.anchor = copy(im.anchor)
        wa.add_image(x)
    return wa


def build(docx_path, thumbs, asset_xlsx, out, ep, overrides, runtime_min=10, target_cuts=None, tol_min=0.5):
    parsed, rows, unknown = derive(docx_path, asset_xlsx, overrides)
    lists = picklists(overrides)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"EP{ep}"
    heads = FIXED + [h for h, _ in PROC]
    heads[0] = f"EP{ep}"
    ws.append(heads)
    for c in ws[1]:
        c.fill, c.font, c.alignment = GRAY, Font(name="Arial", bold=True), Alignment(wrap_text=True, horizontal="center", vertical="center")
    last = len(rows) + 1
    thumbs = pathlib.Path(thumbs)
    est_rows = 0
    for i, r in enumerate(rows, start=2):
        frame = overrides.get("frames", {}).get(r["id"])
        est = frame is None or r["id"] in overrides.get("estimated", [])  # 콘솔이 넘긴 추정값도 노란색
        frame = estimate(r["feat"]) if est else int(frame)
        est_rows += est
        start = int(overrides.get("starts", {}).get(r["id"], START))
        ws.append([r["id"], None, "\n".join(r["cam"]), r["same"], f"=ROUND(F{i}/{overrides.get('fps', FPS)},1)", f"=H{i}-G{i}+1", start, start + frame - 1,
                   "\n".join(r["bg"]), "\n".join(r["ch"]), "\n".join(r["pr"]), r["count"], r["dialogue"], r["description"]])
        if est:
            ws.cell(i, 8).fill = YELLOW
        ws.row_dimensions[i].height = 64.5
        for col in range(1, len(heads) + 1):
            ws.cell(i, col).alignment = Alignment(wrap_text=True, vertical="center")
        name = None
        if r["first"]:
            name = "C%02d" % r["cut"] if r["cut"] < 100 else "C%d" % r["cut"]
        elif r["id"].endswith("b") and (thumbs / ("C%02dB.png" % r["cut"])).exists():
            name = "C%02dB" % r["cut"]
        if name and (thumbs / f"{name}.png").exists():
            im = XlImage(shrink((thumbs / f"{name}.png").read_bytes(), 240))  # 셀에는 114x64로 보이므로 240px JPEG면 충분하고, 시트가 커넥터 한도 안에 든다
            im.width, im.height = 114, 64
            ws.add_image(im, f"B{i}")
    # 드롭다운
    for idx, (_, kind) in enumerate(PROC, start=len(FIXED) + 1):
        if kind:
            dv = DataValidation(type="list", formula1='"%s"' % ",".join(lists[kind]), allow_blank=True)
            ws.add_data_validation(dv)
            dv.add(f"{L(idx)}2:{L(idx)}{last}")
    widths = {"A": 9, "B": 17, "C": 8, "D": 8, "E": 6, "F": 8, "G": 8, "H": 9, "I": 20, "J": 30, "K": 22, "L": 9, "M": 40, "N": 42}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    ws.freeze_panes = "C2"
    copy_asset_sheet(asset_xlsx, wb)
    # 요약 탭: 고정 입력(방영 길이, 목표 행 수), 합계 검증, 작업자별 집계
    sm = wb.create_sheet("요약")
    n_cuts = len(parsed["cuts"])
    dup = sum(r["dup"] for r in rows if r["first"])
    q = f"'{ws.title}'"
    sm.append(["항목", "값"])
    sm.append(["방영 길이(분) - 시나리오와 함께 받은 값", runtime_min])
    sm.append(["목표 컷 수 - 시나리오와 함께 받은 값", target_cuts])
    sm.append(["허용 오차(분) - 토스트 확인 전 가정", tol_min])
    sm.append(["시나리오 컷 수", n_cuts])
    sm.append(["컷 수 판정", '=IF(B3="","미입력",IF(B5=B3,"일치","차이 "&(B5-B3)))'])
    sm.append(["총 Frame", f"=SUM({q}!F2:F{last})"])
    sm.append(["총 러닝타임(초)", f"=SUM({q}!E2:E{last})"])
    sm.append(["총 러닝타임(분)", "=ROUND(B8/60,2)"])
    sm.append(["방영 길이와의 차이(분)", '=IF(B2="","미입력",ROUND(B9-B2,2))'])
    sm.append(["러닝타임 판정", '=IF(B2="","미입력",IF(ABS(B10)<=B4,"일치","확인 필요"))'])
    for a_ in [("시트 행 수(분할 포함)", f"=COUNTA({q}!A2:A{last})"), ("분할로 늘어난 행", len(rows) - n_cuts), ("중복 연출(동일 연출) 컷", dup),
               ("Cam Move가 있는 행", f"=COUNTIF({q}!C2:C{last},\"?*\")"),
               ("Frame 추정(노란색) 행 - 빌드 시점", est_rows), ("확정되지 않은 에셋 코드", ", ".join(unknown) or "없음")]:
        sm.append(list(a_))
    sm.append([]); sm.append(["작업자", "맡은 칸 수", "러닝타임(초)"])
    head_row = sm.max_row
    wcols = {k: [L(i) for i, (h, kind) in enumerate(PROC, start=len(FIXED) + 1) if h == "작업자" and kind == k] for k in WORKERS}
    for kind in WORKERS:
        for n in lists[kind]:
            cnt = "+".join(f'COUNTIF({q}!{c}2:{c}{last},"{n}")' for c in wcols[kind])
            sec = "+".join(f'SUMIF({q}!{c}2:{c}{last},"{n}",{q}!E2:E{last})' for c in wcols[kind])
            sm.append([n, "=" + cnt, "=ROUND(" + sec + ",1)"])
    for r_ in (1, head_row):
        for c in sm[r_]:
            c.fill, c.font = GRAY, Font(name="Arial", bold=True)
    sm.column_dimensions["A"].width, sm.column_dimensions["B"].width, sm.column_dimensions["C"].width = 40, 28, 16
    wb.save(out)
    summary = {"ep": ep, "cuts": n_cuts, "rows": len(rows), "last_row": last, "dup_cuts": dup, "estimated_frame_rows": est_rows,
               "unknown_asset_codes": unknown, "runtime_min": runtime_min, "target_cuts": target_cuts, "tol_min": tol_min, "frames_total": sum(estimate(r["feat"]) if overrides.get("frames", {}).get(r["id"]) is None else int(overrides["frames"][r["id"]]) for r in rows)}
    return summary


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("docx"); ap.add_argument("thumbs"); ap.add_argument("assets"); ap.add_argument("out")
    ap.add_argument("--ep", default="12"); ap.add_argument("--overrides"); ap.add_argument("--fit"); ap.add_argument("--runtime-min", type=float, default=10); ap.add_argument("--target-cuts", type=int); ap.add_argument("--tol-min", type=float, default=0.5)
    a = ap.parse_args()
    if a.fit:
        print("fit coef:", [round(x, 3) for x in fit(a.fit, a.docx, a.assets)])
    ov = json.loads(pathlib.Path(a.overrides).read_text(encoding="utf-8")) if a.overrides else {}
    s = build(a.docx, a.thumbs, a.assets, a.out, a.ep, ov, a.runtime_min, a.target_cuts, a.tol_min)
    pathlib.Path(a.out).with_suffix(".summary.json").write_text(json.dumps(s, ensure_ascii=False, indent=1), encoding="utf-8")
    print(s)
