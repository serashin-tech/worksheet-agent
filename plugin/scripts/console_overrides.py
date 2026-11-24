"""콘솔이 작업 폴더 02_작업에 남긴 cuts.json을 build_worksheet.py의 overrides.json으로 바꾼다.

콘솔에서 PD가 정한 값(분할 수, 행별 Frame과 Start, FPS, 워크시트 드롭다운 목록)을 그대로 옮긴다.
PD가 확인하지 않은 Frame(콘솔의 추정값)은 값은 넘기되 estimated에 넣어 시트에서 노란색으로 남긴다.
Same Camera·Cam Move·BG·CH·PR은 아직 넘기지 않는다(build가 시나리오에서 다시 도출).
사용: python -I console_overrides.py <cuts.json> <overrides.json>
표준 출력: build_worksheet.py에 붙일 인자(--ep, --runtime-min, --target-cuts)
"""
import sys, json, pathlib

SUF = "abcdefghij"


def convert(src):
    ov = {"splits": {}, "frames": {}, "estimated": [], "starts": {}, "fps": src.get("fps", 29.97), "lists": {k: v for k, v in src.get("lists", {}).items() if k != "cam"}}
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
    return ov


if __name__ == "__main__":
    src = json.loads(pathlib.Path(sys.argv[1]).read_text(encoding="utf-8"))
    pathlib.Path(sys.argv[2]).write_text(json.dumps(convert(src), ensure_ascii=False, indent=1), encoding="utf-8")
    args = ["--ep", str(src["ep"])] + (["--runtime-min", str(src["runtime"])] if src.get("runtime") else []) + (["--target-cuts", str(src["target"])] if src.get("target") else [])
    print(" ".join(args))
    # 자체 점검: 2행 분할, 확인한 행·추정 행, Start
    t = convert({"fps": 24, "lists": {"cam": ["PAN"], "an": ["a"]}, "cuts": [{"id": "C001", "cut": "1", "split": 2, "frame": 9, "rf": [5, 4], "rfc": [True, False], "rs": [None, 120]}, {"id": "C002", "cut": "2", "split": 1, "frame": 30, "fc": True}]})
    assert t["frames"] == {"C001a": 5, "C001b": 4, "C002": 30} and t["estimated"] == ["C001b"] and t["starts"] == {"C001b": 120} and "cam" not in t["lists"] and t["splits"] == {"1": 2, "2": 1}, t
