"""console.html의 시나리오 → 컷 목록 JS(SCENARIO-PARSE 블록)가 Python 원본(console_data.cuts_data)과 같은 결과를 내는지 견준다.
사용: python -I test_scenario_js.py <시나리오 docx|pptx|txt|md>   (node 18 이상 필요)
"""
import sys, json, pathlib, subprocess, tempfile

HERE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import console_data

KEYS = ["id", "cut", "place", "description", "dialogue", "cam", "same", "bg", "ch", "pr", "count", "frame", "split", "splitAuto", "dup", "chFlag"]


def js_cuts(path):
    html = (HERE.parent / "console" / "console.html").read_text(encoding="utf-8")
    block = html.split("// SCENARIO-PARSE-START")[1].split("// SCENARIO-PARSE-END")[0]
    run = block + "\nimport fs from 'node:fs';\nconst f = process.argv[2];\nconsole.log(JSON.stringify(await scenarioCuts(f, fs.readFileSync(f))));\n"
    with tempfile.TemporaryDirectory() as d:
        js = pathlib.Path(d) / "t.mjs"; js.write_text(run, encoding="utf-8")
        r = subprocess.run(["node", str(js), str(path)], capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def main(path):
    py, js = console_data.cuts_data(path), js_cuts(path)
    assert len(py["cuts"]) == len(js["cuts"]), f"컷 수 다름: Python {len(py['cuts'])}, JS {len(js['cuts'])}"
    bad = [(a["id"], k, a[k], b[k]) for a, b in zip(py["cuts"], js["cuts"]) for k in KEYS if a[k] != b[k]]
    assert py["baseFrame"] == js["baseFrame"], (py["baseFrame"], js["baseFrame"])
    for b in bad[:10]:
        print("다름:", b)
    assert not bad, f"{len(bad)}건 다름"
    print(f"일치: 컷 {len(js['cuts'])}개 × {len(KEYS)}개 항목")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv[1])
