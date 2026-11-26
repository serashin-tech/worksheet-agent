"""콘솔이 작업 폴더 02_작업에 남긴 cuts.json을 build_worksheet.py의 overrides.json으로 바꾼다.

콘솔에서 PD가 정한 값(분할 수, 행별 Frame과 Start, FPS, 워크시트 드롭다운 목록)을 그대로 옮긴다.
PD가 확인하지 않은 Frame(콘솔의 추정값)은 값은 넘기되 estimated에 넣어 시트에서 노란색으로 남긴다.
3단계(에셋)에서 확정한 에셋 목록(assets)과 시트 행별 BG·CH·PR(rowassets)도 넘긴다. build는 이 값으로 Asset 탭과 BG·CH·PR 열을 쓴다.
2단계(스토리보드)에서 PD가 확인한 칸(sbp)은 시트 행별 Dialogue·Description(rowtext)과 THUMBNAIL(rowimgs, JPEG base64)로 넘긴다. TIME은 시트에 열이 없어 넘기지 않는다.
Same Camera·Cam Move는 아직 넘기지 않는다(build가 시나리오에서 다시 도출).
사용: python -I console_overrides.py <cuts.json> <overrides.json>
표준 출력: build_worksheet.py에 붙일 인자(--ep, --runtime-min, --target-cuts)
"""
import sys, json, pathlib

SUF = "abcdefghij"


def convert(src):
    ov = {"splits": {}, "frames": {}, "estimated": [], "starts": {}, "fps": src["fps"], "lists": {k: v for k, v in src.get("lists", {}).items() if k != "cam"},
          "title": src.get("title", ""), "assets": src.get("assets") or {}, "rowassets": {}, "rowtext": {}, "rowimgs": {}}
    for c in src["cuts"]:
        n, k = int(c["cut"]), int(c.get("split") or 1)
        ov["splits"][str(n)] = k
        rf = c.get("rf") or [c["frame"] // k + (i < c["frame"] % k) for i in range(k)]
        rfc, rs = c.get("rfc") or [False] * k, c.get("rs") or []
        for i in range(k):
            rid = c["id"] + (SUF[i] if k > 1 else "")
            ov["frames"][rid] = int(rf[i])
            if not (c.get("fc") or (i < len(rfc) and rfc[i])):
                ov["estimated"].append(rid)
            if i < len(rs) and rs[i]:
                ov["starts"][rid] = int(rs[i])
            ra = c.get("ra") or []
            if i < len(ra) and ra[i]:
                ov["rowassets"][rid] = {k: ra[i].get(k, [] if k != "count" else 0) for k in ("bg", "ch", "pr", "count")}
            sbp = c.get("sbp") or []
            if i < len(sbp) and sbp[i]:
                ov["rowtext"][rid] = {"dialogue": sbp[i].get("dia", ""), "description": sbp[i].get("note", "")}
                if sbp[i].get("img"):
                    ov["rowimgs"][rid] = sbp[i]["img"]
    return ov


if __name__ == "__main__":
    src = json.loads(pathlib.Path(sys.argv[1]).read_text(encoding="utf-8"))
    pathlib.Path(sys.argv[2]).write_text(json.dumps(convert(src), ensure_ascii=False, indent=1), encoding="utf-8")
    args = ["--ep", str(src["ep"])] + (["--runtime-min", str(src["runtime"])] if src.get("runtime") else []) + (["--target-cuts", str(src["target"])] if src.get("target") else [])
    print(" ".join(args))
    # 자체 점검: 2행 분할, 확인한 행·추정 행, Start
    t = convert({"fps": 24, "lists": {"cam": ["PAN"], "an": ["a"]}, "cuts": [{"id": "C001", "cut": "1", "split": 2, "frame": 9, "rf": [5, 4], "rfc": [True, False], "rs": [None, 120]}, {"id": "C002", "cut": "2", "split": 1, "frame": 30, "fc": True, "ra": [{"bg": ["EddyIn", "SkyNight"], "ch": ["Eddy"], "pr": [], "count": 1}]}], "assets": {"Matte": [{"name": "밤", "code": "SkyNight"}]}})
    assert t["frames"] == {"C001a": 5, "C001b": 4, "C002": 30} and t["estimated"] == ["C001b"] and t["starts"] == {"C001b": 120} and "cam" not in t["lists"] and t["splits"] == {"1": 2, "2": 1}, t
    assert t["rowassets"] == {"C002": {"bg": ["EddyIn", "SkyNight"], "ch": ["Eddy"], "pr": [], "count": 1}} and t["assets"]["Matte"][0]["code"] == "SkyNight", t
    s = convert({"fps": 24, "cuts": [{"id": "C003", "cut": "3", "split": 2, "frame": 20, "sbp": [{"dia": "에디: 응?", "note": "F.S.", "time": "", "img": "QUJD"}, {"dia": "", "note": "", "time": "", "img": ""}]}, {"id": "C004", "cut": "4", "frame": 10}]})
    assert s["rowtext"] == {"C003a": {"dialogue": "에디: 응?", "description": "F.S."}, "C003b": {"dialogue": "", "description": ""}} and s["rowimgs"] == {"C003a": "QUJD"}, s  # 스토리보드가 없는 컷(C004)은 시나리오 값을 쓴다
    assert t["fps"] == 24 and convert({"fps": 29.97, "cuts": []})["fps"] == 29.97  # PD가 정한 FPS를 그대로 넘긴다
    try: convert({"cuts": []}); raise SystemExit("fps 없이 변환됨")
    except KeyError: pass  # 숨은 기본값 없음
