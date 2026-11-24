"""Cam Move 추출 규칙을 정답 시트(EP12) 열 C와 컷 번호로 대조한다. 분할 행(a/b/c)은 따로 센다.
사용: python -I extract_meta.py <docx> <xlsx>
"""
import re, sys, os
import openpyxl
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from parse_docx import parse, DEFAULT_SPEAKERS

# 긴 토큰이 먼저 걸리도록 정렬. 정규형은 정답 시트의 표기.
CAM = [("Q.Z.I", r"Q\.?\s?Z\.?\s?I"), ("S.Z.O", r"S\.?\s?Z\.?\s?O"), ("Q-PAN", r"Q[-.\s]?PAN"),
       ("Z.I", r"Z\.\s?I(?![A-Za-z])"), ("Z.O", r"Z\.\s?O(?![A-Za-z])"), ("T.D", r"T\.\s?D(?![A-Za-z])"), ("T.U", r"T\.\s?U(?![A-Za-z])"),
       ("PAN", r"(?<![A-Za-z-])PAN(?![A-Za-z])"), ("CAM Follow", r"CAM\s?Follow"), ("CAM Shake", r"CAM\s?Shake"),
       ("CAM Move", r"CAM\s?Move")]


SAME = re.compile(r"(\d+)\s?번\s?컷과?\s?(?:동일|같은)\s?연출")


def cam_moves(text):
    found, t = [], text
    for name, pat in CAM:
        if re.search(pat, t, re.I if name.startswith("CAM") else 0):
            found.append(name)
            t = re.sub(pat, " ", t)  # 긴 토큰이 먹은 자리는 지운다(Q.Z.I 안의 Z.I 방지)
    return found


def norm_gold(v):
    s = re.sub(r"\s+", " ", str(v or "")).strip()
    return set(cam_moves(s)) if s != "None" else set()


def main(docx_path, xlsx_path):
    parsed = {c["cut"]: c for c in parse(docx_path, DEFAULT_SPEAKERS + ["노래", "다같이"])["cuts"]}
    ws = openpyxl.load_workbook(xlsx_path).worksheets[0]
    st = {"plain": [0, 0, 0, 0], "split": [0, 0, 0, 0]}  # exact, n, gold-with-cam hit, gold-with-cam n
    bad = []
    for r in range(2, 181):
        cid = str(ws.cell(r, 1).value)
        m = re.match(r"C(\d+)", cid)
        if not m or int(m.group(1)) not in parsed:
            continue
        key = "split" if re.search(r"[a-d]$|_", cid) else "plain"
        c = parsed[int(m.group(1))]
        mine = set(cam_moves(c["description"] + " " + c["transition"]))
        ref = SAME.search(c["description"])
        if ref and not mine and int(ref.group(1)) in parsed:  # '34번컷과 동일 연출'은 그 컷의 Cam Move를 따른다
            p = parsed[int(ref.group(1))]
            mine = set(cam_moves(p["description"] + " " + p["transition"]))
        gold = norm_gold(ws.cell(r, 3).value)
        st[key][1] += 1
        st[key][0] += mine == gold
        if gold:
            st[key][3] += 1
            st[key][2] += bool(mine & gold)
        if mine != gold and key == "plain":
            bad.append((cid, sorted(gold), sorted(mine), str(ws.cell(r, 3).value).replace("\n", "/")))
    for k, (a, n, h, hn) in st.items():
        print(f"{k}: 일치 {a}/{n} ({a / max(n, 1):.0%}), 정답에 Cam이 있는 행 중 하나 이상 적중 {h}/{hn}")
    for b in bad:
        print(b)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv[1], sys.argv[2])
