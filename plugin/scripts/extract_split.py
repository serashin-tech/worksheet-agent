"""분할 행(a/b/c)과 Same Camera 후보 규칙을 정답 시트(EP12)와 대조한다.
- 분할: 지문 앞의 '(4분할 컷)', '(일자 2분할)', '(사선 3분할 컷)' 표기에서 칸 수 N을 읽어 행 N개를 제안한다.
- Same Camera: 'N번 컷과 동일 연출' 문구의 N을 C{N}으로 제안한다. 분할 컷의 b·c·d 칸이 a와 같은 카메라인지는 PD가 정한다.
사용: python -I extract_split.py <docx> <xlsx>
"""
import re, sys, os
import openpyxl
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from parse_docx import parse, DEFAULT_SPEAKERS

SPLIT = re.compile(r"\(\s*(?:[가-힣]+\s)?(\d)\s?분할")
SAME = re.compile(r"(\d+)\s?번\s?컷과?\s?(?:동일|같은)\s?연출")


def split_count(desc):
    m = SPLIT.search(desc)
    return int(m.group(1)) if m else 1


def same_camera(desc):
    m = SAME.search(desc)
    return "C%03d" % int(m.group(1)) if m else ""


def main(docx_path, xlsx_path):
    parsed = parse(docx_path, DEFAULT_SPEAKERS + ["노래", "다같이"])
    cuts = {c["cut"]: c for c in parsed["cuts"]}
    ws = openpyxl.load_workbook(xlsx_path).worksheets[0]
    rows, gold_same = {}, {}
    for r in range(2, 181):
        cid = str(ws.cell(r, 1).value)
        m = re.match(r"C(\d+)(.*)", cid)
        if m:
            rows.setdefault(int(m.group(1)), []).append(m.group(2))
        v = ws.cell(r, 4).value
        if v:
            gold_same[cid] = str(v)
    gold_split = {n: len(s) for n, s in rows.items() if len(s) > 1}
    mine_split = {n: split_count(c["description"]) for n, c in cuts.items() if split_count(c["description"]) > 1}
    print(f"시트 행 수: 컷 {len(cuts)}개 -> 정답 {sum(len(s) for s in rows.values())}행, 규칙 제안 {len(cuts) + sum(v - 1 for v in mine_split.values())}행")
    print("정답 분할:", gold_split)
    print("규칙 분할:", mine_split)
    miss = {n: g for n, g in gold_split.items() if mine_split.get(n) != g}
    print("놓치거나 칸 수가 다른 컷(PD 확인 필요):", miss, "| 잘못 제안:", {n: v for n, v in mine_split.items() if n not in gold_split})
    tp = fp = 0
    wrong = []
    for n, c in cuts.items():
        s = same_camera(c["description"])
        if not s:
            continue
        g = gold_same.get("C%03d" % n)
        if g == s:
            tp += 1
        else:
            fp += 1
            wrong.append(("C%03d" % n, s, g))
    fn = [k for k in gold_same if not same_camera(cuts.get(int(re.match(r"C(\d+)", k).group(1)), {}).get("description", ""))]
    print(f"Same Camera: 제안 {tp + fp}건 중 정답 일치 {tp}, 불일치 {fp} {wrong} / 정답 {len(gold_same)}건 중 문구로 못 찾는 {len(fn)}건 {fn}")
    n_dup = sum(bool(same_camera(c["description"])) for c in cuts.values())
    print(f"'동일 연출' 컷: {n_dup}개 -> 총 {len(cuts)}컷 - {n_dup} = {len(cuts) - n_dup}컷 (대본 머리말: {parsed['summary'][:60]})")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv[1], sys.argv[2])
