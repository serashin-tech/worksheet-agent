"""게이트 화면 프로토타입에 넣을 데이터를 만든다(컷 JSON, 썸네일 JPEG, 에셋 목록, 요약).
사용: python -I console_data.py <docx> <thumbs_dir> <asset_xlsx> <out_dir>
"""
import sys, re, json, pathlib
from PIL import Image
HERE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import build_worksheet as B
from extract_assets import load_assets


def main(docx, thumbs, assets_xlsx, out):
    out = pathlib.Path(out); (out / "thumbs").mkdir(parents=True, exist_ok=True)
    parsed, rows, unknown = B.derive(docx, assets_xlsx, {})
    cuts = []
    for r in rows:
        if not r["first"]:
            continue
        n = r["cut"]
        split = 1 + sum(1 for x in rows if x["cut"] == n and not x["first"])
        text = r["description"] + " " + r["dialogue"]
        mains = [c for c in r["ch"] if not c.startswith("BabyPororo")]
        names = len(mains) + (1 if any(c.startswith("BabyPororo") for c in r["ch"]) else 0)
        ch_flag = bool(re.search(r"친구들|아기들|모두|다같이", text)) and len(mains) < 5
        tn = "C%02d" % n if n < 100 else "C%d" % n
        ok = (pathlib.Path(thumbs) / f"{tn}.png").exists()
        if ok:
            Image.open(pathlib.Path(thumbs) / f"{tn}.png").convert("L").resize((176, 99)).save(out / "thumbs" / ("C%03d.jpg" % n), quality=72)
        cuts.append({"id": "C%03d" % n, "cut": n, "place": parsed["cuts"][n - 1]["place"], "description": r["description"], "dialogue": r["dialogue"],
                     "cam": r["cam"], "same": r["same"], "bg": r["bg"], "ch": r["ch"], "pr": r["pr"], "count": r["count"],
                     "frame": B.estimate(r["feat"]), "split": split, "splitAuto": split > 1, "dup": r["dup"], "chFlag": ch_flag, "thumb": ok})
    a = load_assets(assets_xlsx)
    json.dump({"ep": "12", "baseFrame": B.estimate(B.features({"dialogue": "", "description": ""}, 0)), "cuts": cuts,
               "assets": {k: [{"name": n, "code": c} for n, c in v.items()] for k, v in a.items()}, "unknown": unknown},
              open(out / "cuts.json", "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    print(len(cuts), "cuts", sum(c["thumb"] for c in cuts), "thumbs", sum(c["splitAuto"] for c in cuts), "auto-split", sum(c["dup"] for c in cuts), "dup", sum(c["chFlag"] for c in cuts), "chFlag")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(*sys.argv[1:5])
