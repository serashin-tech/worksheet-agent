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
    codes = {str(r[1].value).strip() for r in wb["Asset"].iter_rows(min_row=2) if r[1].value}  # EP12 Asset 탭에는 " LoopyHanky"처럼 앞 공백이 든 코드가 있다
    res = []
    add = lambda name, ok, detail="": res.append((name, bool(ok), detail))

    head = [wb["Asset"].cell(2, c).value for c in range(1, 10)]
    add("Asset 탭이 EP12 형식(머리글 9열)", head == ["Asset Name(국문)", "Asset Code(영문)", "경로", "제작자", "설명", "렌더 이미지", "비고", "디자인 필요 여부", "리깅"], str(head))
    wa = wb["Asset"]
    rig = [dv for dv in wa.data_validations.dataValidation if str(dv.sqref).startswith("I3:")]
    add("Asset 리깅 열에 ○·△·□ 드롭다운", rig and all(x in rig[0].formula1 for x in "○△□"), str(rig[0].formula1) if rig else "없음")
    colpx = lambda c: (wa.column_dimensions[openpyxl.utils.get_column_letter(c)].width or 13) * 7 + 5
    rowpx = lambda r: (wa.row_dimensions[r].height or 15) * 96 / 72
    over = [im.anchor._from.row + 1 for im in wa._images if im.anchor.ext.width / 9525 + im.anchor._from.colOff / 9525 > colpx(im.anchor._from.col + 1) + 1 or im.anchor.ext.height / 9525 + im.anchor._from.rowOff / 9525 > rowpx(im.anchor._from.row + 1) + 1]
    add("Asset 렌더 이미지가 셀 안에 듦", not over, f"넘치는 행 {over[:5]}")
    add("행 ID 중복 없음", len(ids) == len(set(ids)), f"{len(ids)}행")
    add("컷 번호 연속(C001~)", sorted({int(re.match(r"C(\d+)", i).group(1)) for i in ids}) == list(range(1, max(int(re.match(r"C(\d+)", i).group(1)) for i in ids) + 1)))
    # EP12 시트와 같이 Start·End는 값, Frame = End - Start + 1, Sec = ROUND(Frame/FPS,1). FPS는 첫 행 수식에서 읽는다
    frame = lambda r: ws.cell(r, 8).value - ws.cell(r, 7).value + 1
    bad = [r for r in range(2, last + 1) if not all(isinstance(ws.cell(r, c).value, (int, float)) and float(ws.cell(r, c).value).is_integer() for c in (7, 8)) or frame(r) <= 0]
    add("Start·End Frame이 정수이고 Frame이 양수", not bad, f"문제 행 {bad[:5]}")
    add("Frame = End - Start + 1", all(ws.cell(r, 6).value == f"=H{r}-G{r}+1" for r in range(2, last + 1)))
    m = re.fullmatch(r"=ROUND\(F2/([\d.]+),1\)", str(ws.cell(2, 5).value))
    fps = float(m.group(1)) if m else None
    add("Sec 수식이 모든 행에서 같은 FPS", fps and all(ws.cell(r, 5).value == f"=ROUND(F{r}/{m.group(1)},1)" for r in range(2, last + 1)), f"FPS {fps}")
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
    yellow = sum(1 for r in range(2, last + 1) if ws.cell(r, 8).fill.fgColor.rgb.endswith("FFF2CC"))  # 추정 Frame은 End Frame 칸에 노란색
    if summary:
        s = json.loads(pathlib.Path(summary).read_text(encoding="utf-8"))
        add("노란색 Frame 행 수 = 요약값", yellow == s["estimated_frame_rows"], f"{yellow}")
        add("시트 행 수 = 요약값", len(ids) == s["rows"])
        if s.get("fps") is not None:
            add("시트 FPS = 요약값(PD가 정한 FPS)", fps == float(s["fps"]), f"시트 {fps} / 요약 {s['fps']}")
        if s.get("target_cuts") is not None:
            add("시나리오 컷 수 = 목표 컷 수", s["cuts"] == s["target_cuts"], f"{s['cuts']} / 목표 {s['target_cuts']}")
        if s.get("runtime_min") is not None:
            sec = sum(round(frame(r) / fps, 1) for r in range(2, last + 1)) if fps else 0
            mins = round(sec / 60, 2)
            add("총 러닝타임이 방영 길이 허용 오차 안", abs(mins - s["runtime_min"]) <= s["tol_min"], f"{mins}분 / 방영 {s['runtime_min']}분, 허용 {s['tol_min']}분")
    return res


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    out = verify(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
    for n, ok, d in out:
        print(("PASS" if ok else "FAIL"), n, d)
    sys.exit(0 if all(ok for _, ok, _ in out) else 1)
