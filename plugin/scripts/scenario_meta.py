"""시나리오(docx, pptx, txt, md)에서 1단계의 자동 입력값을 읽는다. 시나리오에 없는 값은 None으로 두어 담당자가 입력하게 한다.
읽는 곳: 에피소드 제목=첫 단락, 목표 컷 수·중복 연출 컷 수=머리말 "총 166컷 = 중복 연출 32컷 제외, 134컷",
애니메이션명=파일 이름의 첫 낱말(추정), 에피소드 번호·방영 길이=파일 이름이나 첫 두 단락에 적혀 있을 때만.
사용: python -I scenario_meta.py <docx|pptx|txt|md>
"""
import re, sys, pathlib

HERE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from parse_docx import parse, paragraphs, DEFAULT_SPEAKERS


def read(path):
    path = pathlib.Path(path)
    # docx는 첫 단락이 제목이다. 그 밖의 형식은 번호가 붙은 줄(=컷)을 제목·머리말로 오인하지 않게 뺀다.
    paras = [t for t, nid in paragraphs(path) if path.suffix.lower() == ".docx" or nid is None]
    head = " ".join(paras[:2])
    pick = lambda pat, text: (lambda m: m.group(1) if m else None)(re.search(pat, text))
    stem = path.stem
    return {
        "episode_title": (paras[0] if paras else None, "시나리오 첫 줄"),
        "target_cuts": (int(c) if (c := pick(r"총\s*(\d+)\s*컷", head)) else None, "시나리오 머리말"),
        "dup_cuts": (int(c) if (c := pick(r"중복\s*연출\s*(\d+)\s*컷", head)) else None, "시나리오 머리말"),
        "series": (stem.split()[0] if " " in stem else None, "파일 이름에서 추정"),
        "episode_no": (pick(r"(?:EP|에피소드)\s*(\d+)", stem + " " + head) or pick(r"(\d+)\s*화", stem + " " + head), "파일 이름이나 머리말"),
        "runtime_min": (pick(r"(\d+(?:\.\d+)?)\s*분", head), "머리말"),
        "counted_cuts": (len(parse(str(path), DEFAULT_SPEAKERS + ["노래", "다같이"])["cuts"]), "번호 목록을 센 값"),
    }


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    for k, (v, src) in read(sys.argv[1]).items():
        print(f"{k}: {v!r}  ({src})")
