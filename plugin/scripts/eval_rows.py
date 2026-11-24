"""정답 시트(EP12)와 docx 파싱 결과를 컷 번호로 대조한다. Sec 열은 제외(프레임 기준이 다름).

정답 행의 번호(C011a -> 11)로 파싱된 같은 번호 컷을 찾아 Description·Dialogue 유사도(0.8 이상 일치)를 센다.
분할 행(a/b/c)은 일부만 담으므로 별도로 센다. 화자 목록은 정답이 아니라 기본값만 쓴다.
사용: python -I eval_rows.py <docx> <xlsx>
"""
import re, sys, os, difflib
import openpyxl
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from parse_docx import parse

norm = lambda s: re.sub(r"[\s#\d\.\-~!?,…()]", "", s or "")
sim = lambda a, b: difflib.SequenceMatcher(None, a, b).ratio()
EXTRA_SPEAKERS = ["노래", "다같이"]  # 정답이 아니라 docx를 보고 PD가 입력했을 법한 화자


def main(docx_path, xlsx_path):
    from parse_docx import DEFAULT_SPEAKERS
    parsed = {c["cut"]: c for c in parse(docx_path, DEFAULT_SPEAKERS + EXTRA_SPEAKERS)["cuts"]}
    ws = openpyxl.load_workbook(xlsx_path).worksheets[0]
    stat = {"plain": [0, 0, 0, 0], "split": [0, 0, 0, 0]}  # desc ok, desc n, dia ok, dia n
    weak = []
    for r in range(2, 181):
        cid = str(ws.cell(r, 1).value)
        m = re.match(r"C(\d+)", cid)
        if not m or int(m.group(1)) not in parsed:
            continue
        key = "split" if re.search(r"[a-d]$|_", cid) else "plain"
        c = parsed[int(m.group(1))]
        desc, dia = norm(ws.cell(r, 14).value), norm(ws.cell(r, 13).value)
        if desc:
            stat[key][1] += 1
            s = sim(desc, norm(c["description"]))
            stat[key][0] += s >= 0.8
            if s < 0.8 and key == "plain":
                weak.append((cid, "desc", round(s, 2)))
        if dia:
            stat[key][3] += 1
            s = sim(dia, norm(c["dialogue"]))
            stat[key][2] += s >= 0.8
            if s < 0.8 and key == "plain":
                weak.append((cid, "dialogue", round(s, 2)))
    print("parsed cuts:", len(parsed))
    for k, (a, n, b, m) in stat.items():
        print(f"{k}: Description {a}/{n} ({a / max(n, 1):.0%}), Dialogue {b}/{m} ({b / max(m, 1):.0%})")
    print("weak(plain):", weak[:30])


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv[1], sys.argv[2])
