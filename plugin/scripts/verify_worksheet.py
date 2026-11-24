"""조립한 Worksheet xlsx를 만든 코드와 따로 다시 열어 검사한다. 하나라도 FAIL이면 종료 코드 1.
사용: python -I verify_worksheet.py <xlsx> [<summary.json>]
"""
import re, sys, json, pathlib
import openpyxl

PROC_START = 15  # O열부터 공정 열


def verify(path, summary=None):
    wb = openpyxl.load_workbook(path)
    ws = wb.worksheets[0]
    last = ws.max_row
    while ws.cell(last, 1).value in (None, "") or not re.match(r"C\d+", str(ws.cell(last, 1).value)):
        last -= 1
    ids = [str(ws.cell(r, 1).value) for r in range(2, last + 1)]
    codes = {str(r[1].value) for r in wb["Asset"].iter_rows(min_row=2) if r[1].value}
    res = []
    add = lambda name, ok, detail="": res.append((name, bool(ok), detail))

    add("행 ID 중복 없음", len(ids) == len(set(ids)), f"{len(ids)}행")
    add("컷 번호 연속(C001~)", sorted({int(re.match(r"C(\d+)", i).group(1)) for i in ids}) == list(range(1, max(int(re.match(r"C(\d+)", i).group(1)) for i in ids) + 1)))
    bad = [r for r in range(2, last + 1) if not (isinstance(ws.cell(r, 6).value, int) and ws.cell(r, 6).value > 0)]
    add("Frame이 모두 양의 정수", not bad, f"문제 행 {bad[:5]}")
    add("Sec 수식이 29.97", all(ws.cell(r, 5).value == f"=ROUND(F{r}/29.97,1)" for r in range(2, last + 1)))
    add("Start=101, End=Start+Frame-1", all(ws.cell(r, 7).value == 101 and ws.cell(r, 8).value == f"=G{r}+F{r}-1" for r in range(2, last + 1)))
    unknown = set()
    for r in range(2, last + 1):
        for c in (9, 10, 11):
            for code in str(ws.cell(r, c).value or "").split("\n"):
                code = code.strip()
                if code and code not in codes and not code.startswith("BabyPororo_Christmas_"):
                    unknown.add(code)
    add("BG·CH·PR 코드가 모두 Asset에 있음", not unknown, f"없는 코드 {sorted(unknown)}")
    cnt_bad = [r for r in range(2, last + 1) if ws.cell(r, 12).value is None]
    add("CH_Count 비어 있지 않음", not cnt_bad)
    sm = wb["요약"]
    tail = "".join(str(c.value) for row in sm.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("="))
    ends = set(re.findall(r"[A-Z]+\$?2:[A-Z]+\$?(\d+)", tail))
    add("요약 탭 범위가 마지막 행까지", ends == {str(last)}, f"끝 행 {sorted(ends)}, 마지막 {last}")
    add("시트 아래에 요약 행이 없음", ws.max_row == last, f"max_row {ws.max_row}, 마지막 {last}")
    cover = {}
    for dv in ws.data_validations.dataValidation:
        for rng in str(dv.sqref).split():
            m = re.match(r"([A-Z]+)2:([A-Z]+)(\d+)", rng)
            cover[m.group(1)] = int(m.group(3))
    add("드롭다운이 모든 행에 걸림", cover and all(v == last for v in cover.values()), f"{len(cover)}열")
    worker_cols = [c for c in range(PROC_START, ws.max_column + 1) if ws.cell(1, c).value == "작업자"]
    filled = [(r, c) for c in worker_cols for r in range(2, last + 1) if ws.cell(r, c).value]
    add("작업자 칸이 비어 있음", not filled, f"{len(worker_cols)}열")
    firsts = [r for r in range(2, last + 1) if re.fullmatch(r"C\d+", ids[r - 2])]
    imgs = {im.anchor._from.row + 1 for im in ws._images}
    add("컷마다 썸네일이 있음", set(firsts) <= imgs, f"썸네일 {len(imgs)}개, 없는 행 {sorted(set(firsts) - imgs)[:5]}")
    yellow = sum(1 for r in range(2, last + 1) if ws.cell(r, 6).fill.fgColor.rgb.endswith("FFF2CC"))
    if summary:
        s = json.loads(pathlib.Path(summary).read_text(encoding="utf-8"))
        add("노란색 Frame 행 수 = 요약값", yellow == s["estimated_frame_rows"], f"{yellow}")
        add("시트 행 수 = 요약값", len(ids) == s["rows"])
        if s.get("target_cuts") is not None:
            add("시나리오 컷 수 = 목표 컷 수", s["cuts"] == s["target_cuts"], f"{s['cuts']} / 목표 {s['target_cuts']}")
        if s.get("runtime_min") is not None:
            sec = sum(round(ws.cell(r, 6).value / 29.97, 1) for r in range(2, last + 1))
            mins = round(sec / 60, 2)
            add("총 러닝타임이 방영 길이 허용 오차 안", abs(mins - s["runtime_min"]) <= s["tol_min"], f"{mins}분 / 방영 {s['runtime_min']}분, 허용 {s['tol_min']}분")
    return res


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    out = verify(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
    for n, ok, d in out:
        print(("PASS" if ok else "FAIL"), n, d)
    sys.exit(0 if all(ok for _, ok, _ in out) else 1)
