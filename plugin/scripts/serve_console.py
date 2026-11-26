"""Worksheet 콘솔을 로컬 사이트로 띄운다(Claude Code·Codex에서 쓴다).

콘솔은 Google Drive·Google Sheets 커넥터를 부르는데, 로컬에서는 이 서버가 같은 호출(search_files, create_file, get_values)을
Google Drive 데스크톱 앱의 동기화 폴더(G:\\내 드라이브\\…)에서 대신 처리한다. 커넥터의 10MB 한도 없이 큰 스토리보드도 다룬다.
- PD는 콘솔에 Drive 폴더 주소(…/folders/<id>)를 붙여 넣는다. 서버는 데스크톱 앱의 메타데이터 DB에서 그 id의 PC 경로를 찾는다.
  데스크톱 앱 드라이브(G:) 안의 PC 경로를 붙여 넣어도 그 폴더를 연결한 폴더로 받는다.
- 폴더·파일 id는 절대 경로를 base64url로 감싼 값("L"로 시작)이다.
- 콘솔에서 연결한 Drive 폴더와 --root로 준 폴더 안에서만 읽고 쓴다. 127.0.0.1에만 열린다.
- /status는 Claude Code가 확인한 커넥터(--connectors)와 데스크톱 앱 상태를 콘솔의 연결 설정 1단계에 준다.
- Google Sheet 대신 xlsx를 시트로 본다(작업 현황은 xlsx의 첫 탭 값을 읽는다).
- 콘솔이 쓰는 데이터(컷·썸네일·스토리보드 칸)는 PD가 파일을 올릴 때 /gen이 기존 스크립트로 만든다. 시연 데이터(--data)는 없어도 된다.
사용: python -I serve_console.py [--connectors "Google Drive,Google Sheets"] [--root <폴더>] [--data <시연 json 폴더>] [--port 8765] [--no-open]
렌더 이미지 모으기(worksheet-build 로컬 모드): python -I serve_console.py --export-images <cuts.json> <aimg 폴더>
완성한 xlsx 넣기(worksheet-build 로컬 모드): python -I serve_console.py --publish <cuts.json> <out.xlsx> <01_산출 폴더>
"""
import argparse, base64, binascii, json, os, pathlib, re, sqlite3, string, sys, threading, webbrowser
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

HERE = pathlib.Path(__file__).resolve().parent
CONSOLE = HERE.parent / "console"
FOLDER = "application/vnd.google-apps.folder"
SHEET = "application/vnd.google-apps.spreadsheet"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
HEAD = ('<!doctype html><html lang="ko"><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1"></head><body>')


class Fail(Exception):
    def __init__(self, code, message, status=400):
        super().__init__(message); self.code, self.status = code, status


def to_id(p):
    return "L" + base64.urlsafe_b64encode(str(pathlib.Path(p).resolve().as_posix()).encode()).decode().rstrip("=")


DRIVEFS = pathlib.Path(os.environ.get("LOCALAPPDATA", "")) / "Google" / "DriveFS"
TOPS = ("내 드라이브", "My Drive", "공유 드라이브", "Shared drives")


def drive_mount():
    """데스크톱 앱이 붙인 드라이브(보통 G:). 맨 위에 '내 드라이브' 같은 폴더가 있는 드라이브를 찾는다."""
    for c in string.ascii_uppercase[3:]:
        if any((pathlib.Path(f"{c}:/") / t).is_dir() for t in TOPS):
            return pathlib.Path(f"{c}:/")
    return None


def drive_status():
    return {"installed": DRIVEFS.is_dir(), "accounts": len(list(DRIVEFS.glob("*/metadata_sqlite_db"))) if DRIVEFS.is_dir() else 0,
            "mount": str(drive_mount() or "")}


def drive_path(fid):
    """Drive 폴더 id → PC 경로. 데스크톱 앱의 메타데이터 DB(items·stable_parents)로 상위 폴더 이름을 따라 올라간다.
    ponytail: Google이 공개하지 않은 DB 형식이다. 앱이 바꾸면 None이 되고, 콘솔은 폴더를 찾지 못했다고 알린다."""
    mount = drive_mount()
    if not mount or not DRIVEFS.is_dir():
        return None
    for db in DRIVEFS.glob("*/metadata_sqlite_db"):
        try:
            con = sqlite3.connect(f"file:{db.as_posix()}?mode=ro", uri=True)
            con.text_factory = lambda b: b.decode("utf-8", "replace")
            row = con.execute("select stable_id from items where id = ? and trashed = 0", (fid,)).fetchone()
            parts, sid = [], row and row[0]
            while sid and len(parts) < 64:
                t = con.execute("select local_title from items where stable_id = ?", (sid,)).fetchone()
                if not t:
                    break
                parts.append(t[0])
                q = con.execute("select parent_stable_id from stable_parents where item_stable_id = ?", (sid,)).fetchone()
                sid = q and q[0]
            con.close()
        except sqlite3.Error:
            continue
        parts.reverse()  # 맨 앞은 '내 드라이브'이거나 공유 드라이브 이름
        for p in (mount.joinpath(*parts), mount.joinpath("공유 드라이브", *parts), mount.joinpath("Shared drives", *parts)):
            if parts and p.is_dir():
                return p.resolve()
    return None


def drive_file_id(p):
    """PC 경로 → Drive 파일 id(drive_path의 반대). 데스크톱 앱이 아직 올리지 않은 파일은 None.
    같은 이름의 항목마다 상위 폴더 이름을 따라 올라가 경로가 같은지 본다."""
    mount, p = drive_mount(), pathlib.Path(p).resolve()
    if not mount or not DRIVEFS.is_dir() or mount.resolve() not in p.parents:
        return None
    for db in DRIVEFS.glob("*/metadata_sqlite_db"):
        try:
            con = sqlite3.connect(f"file:{db.as_posix()}?mode=ro", uri=True)
            con.text_factory = lambda b: b.decode("utf-8", "replace")
            for fid, sid in con.execute("select id, stable_id from items where local_title = ? and trashed = 0 and is_folder = 0", (p.name,)).fetchall():
                parts, cur = [], sid
                while cur and len(parts) < 64:
                    t = con.execute("select local_title from items where stable_id = ?", (cur,)).fetchone()
                    if not t:
                        break
                    parts.append(t[0])
                    q = con.execute("select parent_stable_id from stable_parents where item_stable_id = ?", (cur,)).fetchone()
                    cur = q and q[0]
                parts.reverse()
                if re.fullmatch(r"[\w-]{20,}", fid or "") and any(c == p for c in (mount.joinpath(*parts), mount.joinpath("공유 드라이브", *parts), mount.joinpath("Shared drives", *parts))):
                    con.close(); return fid
            con.close()
        except sqlite3.Error:
            continue
    return None


def sheet_link(inp, roots):
    """콘솔 결과 화면·작업 현황의 'Google Sheets에서 열기': 01_산출의 xlsx를 Drive에서 Google Sheets로 바로 편집하는 링크.
    편집한 내용은 같은 xlsx에 저장되고, 데스크톱 앱이 PC로 내려받으면 작업 현황이 읽는다. 아직 올라가지 않았으면 url이 비어 있다."""
    p = from_id(inp.get("fileId"), roots)
    fid = drive_file_id(p)
    return {"url": f"https://docs.google.com/spreadsheets/d/{fid}/edit?rtpof=true&sd=true" if fid else "", "inDrive": bool(drive_mount() and drive_mount().resolve() in p.parents)}


def from_id(i, roots):
    if i and not i.startswith("L") and re.fullmatch(r"[\w-]{10,}", i):  # Drive 폴더 id: 데스크톱 앱 폴더로 바꾸고, 연결한 폴더로 허용한다
        p = drive_path(i)
        if not p:
            raise Fail("not_found", "Google Drive 데스크톱 앱에서 이 폴더를 찾지 못했습니다. 앱에 로그인했는지, 폴더가 내 드라이브나 공유 드라이브에 있는지 확인해 주세요.", 404)
        if p not in roots:
            roots.append(p)
        return p
    if not i or not i.startswith("L"):
        raise Fail("not_found", f"로컬 폴더 id가 아닙니다: {i}", 404)
    try:
        p = pathlib.Path(base64.urlsafe_b64decode(i[1:] + "=" * (-len(i[1:]) % 4)).decode()).resolve()
    except (binascii.Error, UnicodeDecodeError, ValueError):
        raise Fail("not_found", "잘못된 id입니다.", 404)
    mount = drive_mount()
    if p.is_dir() and mount and mount.resolve() in p.parents and not any(p == r or r in p.parents for r in roots):
        roots.append(p)  # 데스크톱 앱 드라이브(G:) 안의 PC 경로: Drive 주소처럼 연결한 폴더로 받는다
    if not any(p == r or r in p.parents for r in roots):  # 연결한 폴더와 --root 밖은 읽지도 쓰지도 않는다
        raise Fail("permission_denied", f"연결할 수 없는 폴더입니다: {p}. Google Drive 데스크톱 앱 드라이브({mount or 'G:'}) 안의 폴더이거나 Drive 폴더 주소여야 합니다.", 403)
    return p


def mime(p):
    if p.is_dir():
        return FOLDER
    return SHEET if p.suffix.lower() == ".xlsx" else {".json": "application/json", ".pdf": "application/pdf", ".png": "image/png", ".jpg": "image/jpeg",
                                                     ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}.get(p.suffix.lower(), "application/octet-stream")


def meta(p):
    return {"id": to_id(p), "title": p.name, "mimeType": mime(p), "parentId": to_id(p.parent), "viewUrl": "/open?id=" + to_id(p),
            "modifiedTime": p.stat().st_mtime}


def search_files(inp, roots):
    """콘솔이 쓰는 질의만 푼다: parentId = 'x' | 'x' in parents | 'x' in ancestors, title = 't', mimeType =|!= 'm'"""
    q = inp.get("query", "")
    v = lambda pat: (re.search(pat, q) or [None, None])[1]
    unq = lambda s: s and s.replace("\\'", "'").replace("\\\\", "\\")
    parent, anc, title = unq(v(r"parentId = '((?:\\.|[^'])*)'") or v(r"'((?:\\.|[^'])*)' in parents")), unq(v(r"'((?:\\.|[^'])*)' in ancestors")), unq(v(r"title = '((?:\\.|[^'])*)'"))
    m_eq, m_ne = v(r"mimeType = '([^']*)'"), v(r"mimeType != '([^']*)'")
    base = parent or anc
    if not base:  # 연결 확인용 질의("trashed = false")
        return {"files": [meta(r) for r in roots][: int(inp.get("pageSize", 100))]}
    d = from_id(base, roots)
    if not d.is_dir():
        raise Fail("not_found", f"폴더가 없습니다: {d}", 404)
    it = d.rglob("*") if anc else d.iterdir()
    out = [p for p in it if not p.name.startswith((".", "~$")) and (not title or p.name == title) and (not m_eq or mime(p) == m_eq) and (not m_ne or mime(p) != m_ne)]
    out.sort(key=lambda p: -p.stat().st_mtime if "modifiedTime desc" in inp.get("orderBy", "") else 0)
    return {"files": [meta(p) for p in out[: int(inp.get("pageSize", 100))]]}


def create_file(inp, roots):
    d = from_id(inp.get("parentId"), roots)
    name = inp.get("title", "")
    if not name or "/" in name or "\\" in name or name in (".", ".."):
        raise Fail("invalid_argument", f"파일 이름이 올바르지 않습니다: {name}")
    p = d / name
    if inp.get("contentMimeType") == FOLDER:
        p.mkdir(parents=True, exist_ok=True)
    elif "textContent" in inp:
        p.write_text(inp["textContent"], encoding="utf-8")
    else:
        p.write_bytes(base64.b64decode(inp.get("base64Content") or ""))
    return meta(p)


def replace_file(inp, roots):
    """콘솔 3단계 '이미지 바꾸기': 연결한 폴더 안의 기존 그림 파일을 새 내용으로 덮어쓴다. 새 파일은 만들지 않는다.
    Google Drive 데스크톱 앱이 올리면 Drive의 버전 기록에 이전 파일이 남는다."""
    p = from_id(inp.get("fileId"), roots)
    if not p.is_file() or p.suffix.lower() not in (".png", ".jpg", ".jpeg", ".webp"):
        raise Fail("invalid_argument", f"바꿀 수 있는 그림 파일이 아닙니다: {p.name}")
    p.write_bytes(base64.b64decode(inp.get("base64Content") or ""))
    return meta(p)


def build_sheet(inp, roots):
    """콘솔 4단계 '확정하고 시트 만들기'(로컬): 02_작업/cuts.json으로 worksheet-build 로컬 모드(3절 조립·검증)를 돌려 01_산출에 xlsx를 넣는다.
    시나리오는 같은 에피소드 00_입력의 가장 최근 docx다. 검증 FAIL이 있어도 파일은 넣고, 결과를 그대로 돌려줘 콘솔이 보여 준다."""
    import subprocess, tempfile
    cuts = from_id(inp.get("fileId"), roots)
    if cuts.name != "cuts.json" or cuts.parent.name != "02_작업":
        raise Fail("invalid_argument", f"02_작업의 cuts.json이 아닙니다: {cuts}")
    ep = cuts.parent.parent
    docx = sorted((ep / "00_입력").glob("*.docx"), key=lambda p: p.stat().st_mtime)
    if not docx:
        raise Fail("not_found", f"시나리오(docx)가 없습니다: {ep / '00_입력'}", 404)
    out_dir = ep / "01_산출"; out_dir.mkdir(exist_ok=True)
    tmp = pathlib.Path(tempfile.mkdtemp(prefix="wsbuild_")); (tmp / "thumbs").mkdir()
    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    run = lambda *a: subprocess.run([sys.executable, "-I", *map(str, a)], cwd=HERE, env=env, capture_output=True, text=True, encoding="utf-8")
    r = run(HERE / "console_overrides.py", cuts, tmp / "o.json")
    if r.returncode:
        raise Fail("build_error", "cuts.json을 읽지 못했습니다: " + (r.stderr.strip().splitlines() or [""])[-1], 500)
    export_images(cuts, tmp / "aimg")
    b = run(HERE / "build_worksheet.py", docx[-1], tmp / "thumbs", "-", tmp / "out.xlsx", "--overrides", tmp / "o.json", "--asset-images", tmp / "aimg", *r.stdout.split())
    if b.returncode:
        raise Fail("build_error", "시트를 조립하지 못했습니다: " + (b.stderr.strip().splitlines() or [""])[-1], 500)
    v = run(HERE / "verify_worksheet.py", tmp / "out.xlsx")
    checks = [{"ok": ln.startswith("PASS"), "text": ln[5:]} for ln in v.stdout.splitlines() if ln.startswith(("PASS ", "FAIL "))]
    return {**meta(publish(cuts, tmp / "out.xlsx", out_dir)), "checks": checks}


def get_values(inp, roots):
    sid = inp.get("spreadsheetId", "")
    if not sid.startswith("L"):  # 시연용 Google Sheet id: 로컬에는 없다
        return {"values": []}
    import openpyxl
    ws = openpyxl.load_workbook(from_id(sid, roots), read_only=True, data_only=True).worksheets[0]
    rows = [["" if c is None else c for c in r] for r in ws.iter_rows(values_only=True)]
    while rows and not any(str(c).strip() for c in rows[-1]):
        rows.pop()
    return {"values": [[str(c) if not isinstance(c, (int, float)) else c for c in r] for r in rows[:400]]}


def gen_scenario(path):
    """시나리오 → 콘솔 cuts.json 내용(컷·장면·대사·분할·추정 Frame). 에셋 코드는 3단계 폴더에서 정하므로 비워 둔다."""
    sys.path.insert(0, str(HERE))
    import console_data
    return console_data.cuts_data(str(path))


def gen_storyboard(path):
    """스토리보드 PDF → 콘솔 thumbs.json·panels.json 내용. 컷 번호는 모르므로 컷의 첫 칸 순번(S001…)을 키로 주고, 컷 번호는 PD가 콘솔에서 짝짓는다.
    OCR(Windows 한국어)이 실패하면 글 없이 그림만 준다."""
    import io as _io, tempfile
    sys.path.insert(0, str(HERE))
    import pymupdf
    from label_sheets import label_slots
    from thumbnails import crop_first_panels
    import storyboard_panels as SP
    doc = pymupdf.open(str(path))
    seq = lambda k: "S%03d" % k
    thumbs = {}
    for cid, im in crop_first_panels(doc, label_slots(doc), seq).items():  # 컷의 첫 칸(썸네일): 콘솔 표의 THUMBNAIL
        buf = _io.BytesIO(); im.convert("L").resize((176, 99)).save(buf, "JPEG", quality=72)
        thumbs[SP.key(cid)] = "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()
    tmp = pathlib.Path(tempfile.mkdtemp(prefix="wsgen_"))
    SP.main(str(path), str(tmp / "panels.json"), seq)
    pj = json.loads((tmp / "panels.json").read_text(encoding="utf-8"))
    try:
        import storyboard_text
        storyboard_text.main(str(path), str(tmp / "text.json"), seq)
        pj["text"] = json.loads((tmp / "text.json").read_text(encoding="utf-8"))
    except Exception as e:  # ponytail: OCR은 Windows 전용. 다른 OS는 DIALOGUE·NOTE 없이 진행
        print("OCR 실패:", type(e).__name__, e)
        pj["text"] = {}
    return {"thumbs": thumbs, "panels": pj}


GEN = {"scenario": gen_scenario, "storyboard": gen_storyboard}


TOOLS = {("Google Drive", "search_files"): search_files, ("Google Drive", "create_file"): create_file, ("Google Drive", "replace_file"): replace_file, ("Worksheet", "build_sheet"): build_sheet, ("Worksheet", "sheet_link"): sheet_link, ("Google Sheets", "get_values"): get_values}


def handler(roots, data, connectors):
    class H(BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def send(self, code, body, ctype="application/json; charset=utf-8"):
            b = body if isinstance(body, bytes) else json.dumps(body, ensure_ascii=False).encode()
            self.send_response(code); self.send_header("Content-Type", ctype); self.send_header("Content-Length", str(len(b))); self.end_headers(); self.wfile.write(b)

        def do_GET(self):
            u = urlparse(self.path)
            if u.path in ("/", "/index.html"):
                return self.send(200, (HEAD + (CONSOLE / "console.html").read_text(encoding="utf-8") + "</body></html>").encode(), "text/html; charset=utf-8")
            if u.path == "/status":  # 연결 설정 1단계: Claude Code가 확인한 커넥터와 데스크톱 앱 상태
                return self.send(200, {"drive": drive_status(), "connectors": connectors})
            if u.path == "/img":  # 콘솔 3단계 에셋 목록의 이미지: 연결한 폴더 안의 그림 파일만 준다
                try:
                    p = from_id(parse_qs(u.query).get("id", [""])[0], roots)
                    ct = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".webp": "image/webp"}.get(p.suffix.lower())
                    if not ct or not p.is_file():
                        raise Fail("not_found", "그림 파일이 아닙니다.", 404)
                    return self.send(200, p.read_bytes(), ct)
                except Fail as e:
                    return self.send(e.status, {"code": e.code, "message": str(e)})
            if u.path == "/open":  # 콘솔의 폴더·파일 링크: 탐색기나 연결된 앱으로 연다
                try:
                    os.startfile(from_id(parse_qs(u.query).get("id", [""])[0], roots))  # noqa: Windows 전용
                    return self.send(200, (HEAD + "<p>탐색기에서 열었습니다. 이 탭은 닫아도 됩니다.</p></body></html>").encode(), "text/html; charset=utf-8")
                except Fail as e:
                    return self.send(e.status, {"code": e.code, "message": str(e)})
            name = u.path.lstrip("/")
            for base in [CONSOLE] + ([data] if data else []):  # 콘솔 파일(글꼴·로고)과 시연 데이터(cuts·thumbs·panels·work.json)
                p = (base / name).resolve()
                if base.resolve() in p.parents and p.is_file():
                    ctype = {".json": "application/json", ".png": "image/png", ".woff2": "font/woff2"}.get(p.suffix, "application/octet-stream")
                    return self.send(200, p.read_bytes(), ctype)
            self.send(404, {"code": "not_found", "message": name})

        def do_POST(self):
            if self.path == "/gen" and self.headers.get("Origin") in (None, f"http://127.0.0.1:{self.server.server_port}", f"http://localhost:{self.server.server_port}"):
                return self.gen()
            if self.path != "/mcp" or self.headers.get("Origin") not in (None, f"http://127.0.0.1:{self.server.server_port}", f"http://localhost:{self.server.server_port}"):
                return self.send(403, {"code": "forbidden", "message": "이 콘솔에서 온 요청만 받습니다."})
            try:
                req = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0))) or b"{}")
                fn = TOOLS.get((req.get("server"), req.get("tool")))
                if not fn:
                    raise Fail("not_in_manifest", f"로컬 콘솔이 지원하지 않는 호출입니다: {req.get('server')} {req.get('tool')}")
                self.send(200, fn(req.get("input") or {}, roots))
            except Fail as e:
                self.send(e.status, {"code": e.code, "message": str(e)})
            except Exception as e:  # 파일 잠김 등: 콘솔에 그대로 보인다
                self.send(500, {"code": "tool_error", "message": f"{type(e).__name__}: {e}"})
        def gen(self):
            """올린 파일 하나로 그 단계의 콘솔 데이터를 만든다: {kind: scenario|storyboard, name, base64}"""
            import tempfile
            try:
                req = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0))) or b"{}")
                fn = GEN.get(req.get("kind"))
                if not fn:
                    raise Fail("invalid_argument", f"모르는 종류입니다: {req.get('kind')}")
                name = pathlib.Path(req.get("name") or "input").name  # 경로는 버리고 이름만 쓴다
                p = pathlib.Path(tempfile.mkdtemp(prefix="wsin_")) / name
                p.write_bytes(base64.b64decode(req.get("base64") or ""))
                self.send(200, fn(p))
            except Fail as e:
                self.send(e.status, {"code": e.code, "message": str(e)})
            except Exception as e:
                self.send(500, {"code": "gen_error", "message": f"{type(e).__name__}: {e}"})
    return H


def export_images(cuts, out):
    """cuts.json의 assets에서 에셋 이미지를 out/<code>.<ext>로 모은다. PD가 콘솔에서 바꾼 이미지(img, PNG base64)가 있으면 그것을 쓰고,
    없으면 로컬 id(L…)의 파일을 복사한다. 둘 다 없는 에셋(손으로 더하고 이미지를 넣지 않은 것)은 건너뛴다."""
    import shutil
    out = pathlib.Path(out); out.mkdir(parents=True, exist_ok=True)
    n = 0
    for v in (json.loads(pathlib.Path(cuts).read_text(encoding="utf-8")).get("assets") or {}).values():
        for a in v:
            if a.get("img"):
                (out / f"{a['code']}.png").write_bytes(base64.b64decode(a["img"])); n += 1
                continue
            i = a.get("fileId") or ""
            if not i.startswith("L"):
                continue
            p = pathlib.Path(base64.urlsafe_b64decode(i[1:] + "=" * (-len(i[1:]) % 4)).decode())
            if p.suffix.lower() in (".png", ".jpg", ".jpeg", ".webp") and p.is_file():
                shutil.copyfile(p, out / f"{a['code']}{p.suffix.lower()}"); n += 1
    print(f"이미지 {n}장 → {out}")


def publish(cuts, xlsx, out_dir):
    """out.xlsx를 01_산출에 '<title>_YYYYMMDD.xlsx'로 복사한다. 같은 이름이 있으면 _HHMM을 붙이고, 그래도 있으면 숫자를 붙인다. 덮어쓰지 않는다.
    제목은 여기서 UTF-8로 읽는다(에이전트가 Windows 기본 인코딩으로 읽으면 한글 제목이 깨진다)."""
    import shutil, datetime
    now = datetime.datetime.now()
    title = re.sub(r'[\\/:*?"<>|]', "_", json.loads(pathlib.Path(cuts).read_text(encoding="utf-8")).get("title") or "Worksheet").strip()
    d = pathlib.Path(out_dir)
    cands = [f"{title}_{now:%Y%m%d}", f"{title}_{now:%Y%m%d_%H%M}"] + [f"{title}_{now:%Y%m%d_%H%M}_{i}" for i in range(2, 100)]
    dst = next(d / f"{c}.xlsx" for c in cands if not (d / f"{c}.xlsx").exists())
    shutil.copyfile(xlsx, dst)
    print(dst)
    return dst


def self_check(tmp):
    root = pathlib.Path(tmp).resolve(); roots = [root]
    rid = to_id(root)
    ep = create_file({"title": "EP12", "parentId": rid, "contentMimeType": FOLDER}, roots)["id"]
    create_file({"title": "cuts.json", "parentId": ep, "textContent": "{}"}, roots)
    assert search_files({"query": f"parentId = '{rid}' and title = 'EP12' and mimeType = '{FOLDER}'"}, roots)["files"][0]["id"] == ep
    assert [f["title"] for f in search_files({"query": f"'{rid}' in ancestors and mimeType != '{FOLDER}'"}, roots)["files"]] == ["cuts.json"]
    for bad in (to_id(root.parent), "L!!", "1AYqLq2Fn7e3cWHoiK9"):
        try:
            from_id(bad, roots); assert False, bad
        except Fail:
            pass
    try:
        create_file({"title": "../x", "parentId": rid}, roots); assert False
    except Fail:
        pass
    assert get_values({"spreadsheetId": "1AYqLq2Fn7e3cWHoiK9"}, roots) == {"values": []}
    img = root / "Pororo.png"; img.write_bytes(b"x")
    (root / "c.json").write_text(json.dumps({"assets": {"CH": [{"code": "Pororo", "fileId": to_id(img)}, {"code": "Hand"}, {"code": "Hat", "img": base64.b64encode(b"PNG").decode()}]}}), encoding="utf-8")
    export_images(root / "c.json", root / "aimg")
    assert [p.name for p in (root / "aimg").iterdir()] == ["Hat.png", "Pororo.png"]
    (root / "t.json").write_text(json.dumps({"title": "뽀로로 EP12 Worksheet"}, ensure_ascii=False), encoding="utf-8")
    a, b = publish(root / "t.json", img, root), publish(root / "t.json", img, root)
    assert a.name.startswith("뽀로로 EP12 Worksheet_") and a != b and a.exists() and b.exists(), (a, b)
    (root / "u.json").write_text(json.dumps({"title": "a\\b/c:d"}), encoding="utf-8")
    assert publish(root / "u.json", img, root).name.startswith("a_b_c_d_")
    print("self-check ok")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", action="append", default=[]); ap.add_argument("--connectors", help="Claude Code에서 연결을 확인한 커넥터(쉼표로 구분)"); ap.add_argument("--data"); ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--no-open", action="store_true"); ap.add_argument("--self-check"); ap.add_argument("--export-images", nargs=2, metavar=("CUTS", "OUT"))
    ap.add_argument("--publish", nargs=3, metavar=("CUTS", "XLSX", "OUT_DIR"))
    a = ap.parse_args()
    if a.publish:
        publish(*a.publish); sys.exit()
    if a.export_images:
        export_images(*a.export_images); sys.exit()
    if a.self_check:
        self_check(a.self_check); sys.exit()
    roots = [pathlib.Path(r).resolve() for r in a.root]
    if not all(r.is_dir() for r in roots):
        sys.exit("--root에 없는 폴더가 있습니다.")
    connectors = None if a.connectors is None else {n: n in [x.strip() for x in a.connectors.split(",")] for n in ("Google Drive", "Google Sheets")}
    data = pathlib.Path(a.data).resolve() if a.data else (CONSOLE if (CONSOLE / "cuts.json").exists() else None)
    if not data or not (data / "cuts.json").exists():
        data = None
        print("시연 데이터 없이 시작합니다. 시나리오와 스토리보드를 올리면 콘솔이 그 파일로 데이터를 만듭니다.")
    srv = ThreadingHTTPServer(("127.0.0.1", a.port), handler(roots, data, connectors))
    url = f"http://127.0.0.1:{a.port}/"
    print("콘솔:", url)
    for r in roots:
        print("허용 폴더:", r)
    print("데스크톱 앱:", drive_status())
    if not a.no_open:
        threading.Timer(0.5, webbrowser.open, [url]).start()
    srv.serve_forever()
