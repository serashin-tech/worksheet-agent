"""시리즈 폴더와 화수 폴더의 파일 이름으로 BG·CH·PR 에셋 사전을 만든다.
토스트 대표와 협의한 방식이다: 이미지 매칭(벡터 DB)은 쓰지 않고, 이름이 같은 한글·영문 파일을 폴더에서 찾는다.

입력: [{"folder": "series" | "episode", "path": "BG/에디 집 내부.png"}, ...]  (Drive 목록에서 만든 파일 경로)
가정(확인 필요): 하위 폴더 이름이 BG·CH·PR(또는 배경·캐릭터·소품)이고, 파일 이름이 에셋 이름이다.
규칙
- 파일 이름(확장자 제외)이 영문이면 그것이 코드다. 한글이면 단어 사전으로 영문 코드로 바꾼다.
- 한글과 영문이 같은 에셋이면(에디.png, Eddy.png) 한 항목으로 합친다.
- 한글 이름을 영문 코드로 못 바꾸면 한글 이름을 코드로 쓰고 unresolved에 올린다(PD 확인).
- 같은 코드가 두 폴더에 있으면 화수 폴더 쪽을 쓴다.
사용: python -I folder_assets.py   (자체 점검만 실행)
"""
import re, sys

KIND = {"BG": "BG", "배경": "BG", "CH": "CH", "캐릭터": "CH", "PR": "PR", "소품": "PR", "프롭": "PR"}
WORD = {"에디": "Eddy", "뽀로로": "Pororo", "루피": "Loopy", "크롱": "Crong", "패티": "Petty", "포비": "Poby", "해리": "Harry",
        "내부": "In", "외부": "Out", "집": "", "눈길": "SnowRoad", "낮": "SkyDay", "밤": "SkyNight", "아침": "SkyMorning"}
IMAGE = re.compile(r"\.(png|jpe?g|svg|ai|psd|webp)$", re.I)
KO_TOKENS = sorted(WORD, key=len, reverse=True)


def to_code(stem):
    """영문이면 그대로, 한글이면 사전으로 변환. 변환 못 하면 None."""
    stem = stem.strip()
    if re.fullmatch(r"[A-Za-z0-9_\-]+", stem):
        return stem
    rest, out = re.sub(r"[\s_\-]+", " ", stem), []
    for w in rest.split():
        while w:
            t = next((k for k in KO_TOKENS if w.startswith(k)), None)
            if t is None:
                return None
            out.append(WORD[t]); w = w[len(t):]
    return "".join(out) or None


def kind_of(path):
    return next((KIND[p] for p in path.split("/")[:-1] if p in KIND), None)


def from_folders(files):
    """-> (assets {"BG": {이름: 코드}, ...}, unresolved [경로], skipped [경로])"""
    assets, src, unresolved, skipped = {"BG": {}, "CH": {}, "PR": {}}, {}, [], []
    for f in sorted(files, key=lambda f: f["folder"] == "episode"):  # 화수 폴더를 나중에 처리해 덮어씀
        path = f["path"].replace("\\", "/")
        kind = kind_of(path)
        if not kind or not IMAGE.search(path):
            skipped.append(path); continue
        stem = IMAGE.sub("", path.split("/")[-1])
        code = to_code(stem)
        if code is None:
            code = stem.strip(); unresolved.append(path)
        name = stem.strip() if not re.fullmatch(r"[A-Za-z0-9_\-]+", stem) else next(
            (n for n, c in assets[kind].items() if c == code), stem.strip())  # 한글 이름이 이미 있으면 그 이름 유지
        assets[kind][name] = code
        src[(kind, code)] = f["folder"]
    return assets, unresolved, skipped


if __name__ == "__main__":
    files = [{"folder": "series", "path": p} for p in
             ["CH/에디.png", "CH/Eddy.png", "CH/뽀로로.png", "BG/에디 집 내부.png", "BG/눈길.png", "PR/썰매.png", "메모.txt"]] \
        + [{"folder": "episode", "path": p} for p in ["PR/크리스마스트리.png", "BG/EddyIn.png"]]
    a, bad, skip = from_folders(files)
    assert a["CH"] == {"에디": "Eddy", "뽀로로": "Pororo"}, a["CH"]       # 한글·영문 같은 이름은 한 항목
    assert a["BG"] == {"에디 집 내부": "EddyIn", "눈길": "SnowRoad"}, a["BG"]  # 영문 EddyIn.png도 같은 코드로 합쳐짐
    assert bad == ["PR/썰매.png", "PR/크리스마스트리.png"] and skip == ["메모.txt"], (bad, skip)
    print("ok", a, bad, skip)
