"""Phase 1: 원작 대본 docx를 컷 번호 단위로 나누고 지문·대사를 분리한다.

docx의 컷 번호는 글자가 아니라 번호 매기기 목록(numPr)에 들어 있다. 본문 목록(가장 많이 쓰인 numId)의
n번째 단락이 n번 컷이다. 목록이 아닌 단락은 다음처럼 처리한다.
- '#'로 시작: 장소 제목. 이후 컷의 place가 됨
- '(WIPE)', '(CLOCK WIPE)' 같은 괄호 한 줄: 직전 컷의 transition
- 그 밖의 비목록 단락: 직전 컷의 지문·대사에 이어 붙임(extra에 기록)
대사 시작은 '화자:' 또는 'SE:' 패턴으로 찾는다. 화자 목록은 Asset의 CH와 PD 입력에서 받는다.
docx 말고 pptx·txt·md도 읽는다(`paragraphs`). 이 형식에는 번호 매기기 목록이 없으므로 줄 앞의
'12.', '12)', 'C012' 같은 번호를 컷 번호로 본다. pptx는 글머리 자동 번호도 컷 번호로 본다.
사용: python -I parse_docx.py <docx|pptx|txt|md> [out.json]
"""
import re, sys, json, collections, pathlib
import docx

DEFAULT_SPEAKERS = ["뽀로로", "에디", "루피", "패티", "크롱", "포비", "해리", "아기 뽀로로", "아기뽀로로", "산타", "SE", "소리"]
TRANSITION = re.compile(r"^\(\s*[A-Za-z ]*WIPE[^)]*\)$", re.I)


def build_pattern(speakers):
    spk = "(?:%s)" % "|".join(sorted(map(re.escape, speakers), key=len, reverse=True))
    one = spk + r"\s?[0-9~/]*(?:\([^)]*\))?"  # '아기 뽀로로1~10', '뽀로로(겹)' 같은 번호·괄호 꼬리
    return re.compile(r"(?<![가-힣A-Za-z/])(%s(?:/%s)*)\s*:" % (one, one))


def split_para(text, pat):
    m = pat.search(text)
    if not m:
        return text.strip(), ""
    return text[:m.start()].strip(), text[m.start():].strip()


# 목록이 없는 형식에서 컷 번호로 보는 줄머리: '12. ', '12) ', 'C012 ', 'C012: '. '12.5분', '10:30'은 걸리지 않는다.
NUM_PREFIX = re.compile(r"^\s*(?:[Cc]\s*0*\d{1,3}\s*[:.)\]]?|0*\d{1,3}\s*[.)\]])\s+(?=\S)")
SUFFIXES = (".docx", ".pptx", ".txt", ".md")


def _read_text(path):
    raw = pathlib.Path(path).read_bytes()
    for enc in ("utf-8-sig", "cp949"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            pass
    raise ValueError("텍스트 인코딩을 읽지 못했습니다. UTF-8이나 CP949로 저장해 주세요.")


def paragraphs(path):
    """(글, 번호 목록 표시) 목록. 번호 목록 표시는 docx는 numId, 나머지는 번호가 붙은 줄이면 'num', 아니면 None."""
    path = pathlib.Path(path); ext = path.suffix.lower()
    if ext == ".docx":
        return [(p.text.strip(), num_id(p)) for p in docx.Document(str(path)).paragraphs if p.text.strip()]
    if ext in (".txt", ".md"):
        lines = [x.strip() for x in _read_text(path).splitlines() if x.strip()]
    elif ext == ".pptx":
        from pptx import Presentation
        from pptx.oxml.ns import qn
        lines = []
        for slide in Presentation(str(path)).slides:
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                for para in shape.text_frame.paragraphs:
                    t = "".join(r.text for r in para.runs).strip()
                    if not t:
                        continue
                    pPr = para._p.pPr
                    auto = pPr is not None and pPr.find(qn("a:buAutoNum")) is not None
                    lines.append(("\0" + t) if auto else t)  # \0 = 글머리 자동 번호
    else:
        raise ValueError("지원하지 않는 형식입니다(%s). docx, pptx, txt, md만 읽습니다." % ext)
    out = []
    for t in lines:
        if t.startswith("\0"):
            out.append((t[1:], "num")); continue
        m = NUM_PREFIX.match(t)
        out.append((t[m.end():].strip(), "num") if m else (t, None))
    return out


def num_id(p):
    pPr = p._p.pPr
    if pPr is None or pPr.numPr is None or pPr.numPr.numId is None:
        return None
    return pPr.numPr.numId.val


def parse(path, speakers=DEFAULT_SPEAKERS):
    pat = build_pattern(speakers)
    paras = [(nid, t) for t, nid in paragraphs(path)]
    ids = collections.Counter(nid for nid, _ in paras if nid is not None)
    if not ids:
        raise ValueError("컷 번호 목록을 찾지 못했습니다. docx는 번호 매기기 목록, 그 밖의 형식은 줄 앞의 '12.' 같은 번호가 필요합니다.")
    body_id = ids.most_common(1)[0][0]
    cuts, place, summary = [], "", ""
    for nid, t in paras:
        if t.startswith("#"):
            place = t
        elif nid == body_id:
            desc, dia = split_para(t, pat)
            desc = " ".join(desc.split())  # 문단 안의 강제 줄바꿈(w:br)은 한 줄로 합친다
            cuts.append({"cut": len(cuts) + 1, "place": place, "description": desc, "dialogue": dia,
                         "transition": "", "extra": ""})
        elif nid is not None or not cuts:
            summary = summary or t  # 머리말(총 컷 수 안내 등)
        elif TRANSITION.match(t):
            cuts[-1]["transition"] = t
        else:
            cuts[-1]["extra"] = t
            cuts[-1]["dialogue"] = (cuts[-1]["dialogue"] + "\n" + t).strip()
    return {"summary": summary, "cuts": cuts}


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    out = parse(sys.argv[1])
    if len(sys.argv) > 2:
        open(sys.argv[2], "w", encoding="utf-8").write(json.dumps(out, ensure_ascii=False, indent=1))
    print(out["summary"], "cuts:", len(out["cuts"]),
          "transitions:", sum(bool(c["transition"]) for c in out["cuts"]),
          "extra:", [c["cut"] for c in out["cuts"] if c["extra"]])
