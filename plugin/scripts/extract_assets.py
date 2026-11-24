"""BG·CH·CH_Count·PR 도출 규칙을 정답 시트(EP12) 열 I·J·K·L과 컷 번호로 대조한다.
에셋 이름→코드 표는 Asset 시트(정답)에서 읽는다. 실제 운영에서는 PD가 게이트 2에서 확정한 Asset 목록이 이 표를 대신한다.
사용: python -I extract_assets.py <docx> <xlsx>
"""
import re, sys, os
import openpyxl
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from parse_docx import parse, DEFAULT_SPEAKERS

HOUSE = {"에디": "Eddy", "뽀로로": "Pororo", "루피": "Loopy"}
SKY = [("낮", "SkyDay"), ("노을", "SkySunsetYellow"), ("밤", "SkyNight"), ("아침", "SkyMorning")]
MAIN = [("뽀로로", "Pororo"), ("크롱", "Crong"), ("루피", "Loopy"), ("에디", "Eddy"), ("패티", "Petty")]
BABY = re.compile(r"아기\s?뽀로로\s?([0-9][0-9~,\s/부터까지\-]*)?|아기\s?([0-9]+)")


def load_assets(xlsx):
    a = openpyxl.load_workbook(xlsx).worksheets[1]
    cat, out = None, {"BG": {}, "CH": {}, "PR": {}, "Matte": {}}
    for r in range(3, 60):
        n, code = a.cell(r, 1).value, a.cell(r, 2).value
        if n in out and code is None:
            cat = n
        elif cat and n and code:
            out[cat][str(n).strip()] = str(code).strip()
    return out


def bg_codes(place, text):
    code = []
    m = re.search(r"(에디|뽀로로|루피)\s*집\s*[/\-]?\s*(내부|외부)?", place)
    if m and m.group(2):
        code.append(HOUSE[m.group(1)] + ("In" if m.group(2) == "내부" else "Out"))
    if "눈길" in place:
        code.append("SnowRoad")
    if code and code[0].endswith("Out"):
        for k, v in SKY:
            if k in place:
                code.append(v)
                break
    return code


def baby_numbers(text):
    nums = set()
    for m in BABY.finditer(text):
        s = (m.group(1) or m.group(2) or "").strip()
        for a, b in re.findall(r"(\d+)\s*(?:부터|~|-)\s*(\d+)", s):
            nums.update(range(int(a), int(b) + 1))
        s2 = re.sub(r"\d+\s*(?:부터|~|-)\s*\d+", " ", s)
        nums.update(int(x) for x in re.findall(r"\d+", s2))
    return nums


def ch_codes(text):
    # 장소 이름('에디 집', '루피네 집')의 인물 이름은 등장인물이 아니다(EP12 정답 대조 CH주요 120→123/159).
    # 대사 화자는 빼지 않는다: 빼면 C001처럼 화면 밖 목소리는 맞지만 전체가 121/159로 떨어진다.
    text = re.sub(r"(에디|뽀로로|루피)\s*(네\s*)?집", " ", text)
    t = re.sub(r"아기\s?뽀로로", "아기", text)
    out = [code for kor, code in MAIN if re.search(kor, t)]
    babies = baby_numbers(text)
    return out, babies


def pr_codes(text, props):
    return [code for name, code in props.items() if re.search(re.escape(re.sub(r"\s*\(.*\)", "", name)), text)]


def main(docx_path, xlsx_path):
    assets = load_assets(xlsx_path)
    parsed = {c["cut"]: c for c in parse(docx_path, DEFAULT_SPEAKERS + ["노래", "다같이"])["cuts"]}
    ws = openpyxl.load_workbook(xlsx_path).worksheets[0]
    S = {k: [0, 0] for k in ("BG장소", "CH주요", "CH아기유무", "CH_Count")}
    pr_tp = pr_fp = pr_fn = 0
    bad = {k: [] for k in S}
    sp = lambda v: {x.strip() for x in str(v or "").split("\n") if x.strip() and x != "None"}
    for r in range(2, 181):
        cid = str(ws.cell(r, 1).value)
        m = re.match(r"C(\d+)", cid)
        if not m or int(m.group(1)) not in parsed or re.search(r"[a-d]$|_", cid):
            continue
        c = parsed[int(m.group(1))]
        text = c["description"] + " " + c["dialogue"]
        gbg = {x.replace("EP12", "") for x in sp(ws.cell(r, 9).value) if not x.startswith("\\")}
        mybg = set(bg_codes(c["place"], text))
        ok = {x for x in gbg if not x.startswith("Sky")} == {x for x in mybg if not x.startswith("Sky")}
        S["BG장소"][1] += 1; S["BG장소"][0] += ok
        if not ok: bad["BG장소"].append((cid, sorted(gbg), sorted(mybg), c["place"]))
        gch = sp(ws.cell(r, 10).value)
        gmain = {x for x in gch if not x.startswith("BabyPororo") and x in {v for _, v in MAIN}}
        gbaby = any(x.startswith("BabyPororo") for x in gch)
        mmain, mb = ch_codes(text)
        okm = set(mmain) == gmain
        S["CH주요"][1] += 1; S["CH주요"][0] += okm
        if not okm: bad["CH주요"].append((cid, sorted(gmain), sorted(mmain)))
        okb = bool(mb) == gbaby
        S["CH아기유무"][1] += 1; S["CH아기유무"][0] += okb
        if not okb: bad["CH아기유무"].append((cid, gbaby, sorted(mb)))
        try:
            gcount = int(float(ws.cell(r, 12).value))
        except (TypeError, ValueError):
            gcount = 0
        mc = len(mmain) + len(mb)
        S["CH_Count"][1] += 1; S["CH_Count"][0] += mc == gcount
        if mc != gcount: bad["CH_Count"].append((cid, gcount, mc, sorted(mmain), sorted(mb)))
        gpr = {x for x in sp(ws.cell(r, 11).value)}
        mpr = set(pr_codes(c["description"], assets["PR"]))
        pr_tp += len(gpr & mpr); pr_fp += len(mpr - gpr); pr_fn += len(gpr - mpr)
    for k, (a, n) in S.items():
        print(f"{k}: {a}/{n} ({a / max(n, 1):.0%})")
    print(f"PR: 재현율 {pr_tp / max(pr_tp + pr_fn, 1):.0%}, 정밀도 {pr_tp / max(pr_tp + pr_fp, 1):.0%} (tp {pr_tp}, fp {pr_fp}, fn {pr_fn})")
    for k, v in bad.items():
        print(k, "불일치 예시", v[:8])


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv[1], sys.argv[2])
