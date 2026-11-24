---
name: worksheet-build
description: 콘솔에서 확정한 에피소드의 Worksheet xlsx를 만들어 작업 폴더의 01_산출에 Google Sheet로 올린다. PD가 "시트 만들어 줘", "워크시트 올려 줘", "01_산출에 넣어 줘"라고 하거나, 콘솔이 02_작업에 cuts.json을 남긴 뒤 시트를 만들 때 쓴다.
---

# Worksheet 만들기와 01_산출 업로드

콘솔은 시트를 직접 만들지 못한다. 콘솔은 "시트 만들기"를 누르면 확정값을 `<작업 폴더>/<애니메이션명>/EP##/02_작업/cuts.json`에 남기고, 이 스킬이 그 파일로 xlsx를 만들어 같은 에피소드의 `01_산출`에 올린다. 숫자 계산은 모두 스크립트가 한다(ADR 0003).

스크립트 경로는 `${CLAUDE_PLUGIN_ROOT}/scripts/`다. Python은 늘 `python -I`로 돌린다. 내려받은 파일은 새 빈 폴더(예: 세션 임시 폴더의 `ws_EP12/`)에 둔다.

## 1. 에피소드 폴더 찾기

1. PD에게 작업 폴더 링크(`…/folders/<id>`)와 애니메이션명·에피소드 번호를 받는다. 콘솔의 결과 화면에 있는 폴더 링크를 받아도 된다.
2. Google Drive `search_files`로 폴더를 찾는다.
   - 애니메이션 폴더: `parentId = '<작업 폴더 id>' and title = '<애니메이션명>' and mimeType = 'application/vnd.google-apps.folder'`
   - 에피소드 폴더: 같은 방식으로 `EP12`(두 자리)
   - 그 아래 `00_입력`, `01_산출`, `02_작업`
3. `02_작업`에서 `title = 'cuts.json'`을 `orderBy: "modifiedTime desc"`로 찾아 가장 최근 파일을 쓴다. 없으면 콘솔에서 "시트 만들기"를 먼저 누르게 한다.

## 2. 입력 내려받기

1. `download_file_content`로 `cuts.json`과 `00_입력`의 시나리오(docx)를 받는다.
2. 스토리보드 PDF가 7MB를 넘으면 커넥터로 받지 못한다(ADR 0006). PD PC의 Drive 동기화 폴더(`G:\`)에서 같은 경로의 파일을 쓴다. 찾지 못하면 PD에게 로컬 경로를 묻는다.
3. 에셋 목록 xlsx(EP12는 정답 시트의 Asset 탭)는 PD에게 경로를 받는다.

## 3. 조립과 검증

```
python -I console_overrides.py <cuts.json> <o.json>
```
표준 출력으로 나온 인자(`--ep`, `--runtime-min`, `--target-cuts`)를 아래 명령에 붙인다.

```
python -I thumbnails.py <스토리보드.pdf> <에셋.xlsx> <thumbs폴더>
python -I build_worksheet.py <시나리오.docx> <thumbs폴더> <에셋.xlsx> <out.xlsx> --overrides <o.json> <위 인자>
python -I verify_worksheet.py <out.xlsx> <out.summary.json>
```

- FAIL이 있으면 올리기 전에 PD에게 항목을 그대로 보여 준다. 러닝타임 FAIL은 추정 Frame 때문일 수 있어 PD가 올릴지 정한다. 다른 FAIL은 고친 뒤 다시 조립한다.
- 썸네일 컷 번호 규칙은 EP12에서 육안으로 맞춘 것이다(`thumbnails.ep12_label`). 다른 에피소드는 label_sheets.py로 접촉 시트를 만들어 번호를 먼저 확인한다.

## 4. 01_산출에 올리기

1. `out.xlsx` 크기를 확인한다. 커넥터 요청 본문 한도가 10MB(base64 포함)라 7MB 이하만 올라간다. build_worksheet.py는 렌더 이미지를 480px, 썸네일을 240px JPEG로 줄여 EP12가 약 4.8MB다. 7MB를 넘으면 올리지 말고 PD에게 알린다. 동기화 폴더로 복사한 xlsx는 Google Sheet로 바뀌지 않아 작업현황에서 읽을 수 없다.
2. 파일 해시를 구한다: `python -I -c "import hashlib,sys;print(hashlib.sha256(open(sys.argv[1],'rb').read()).hexdigest())" <out.xlsx>`
3. Google Drive `create_file`로 올린다.
   - `title`: `cuts.json`의 `title`에 날짜를 붙인다(예: `뽀로로 EP12 Worksheet_20261124`). 같은 이름이 `01_산출`에 있으면 시각(`_1530`)을 더 붙인다. 기존 시트를 덮어쓰거나 지우지 않는다.
   - `parentId`: `01_산출` 폴더 id
   - `contentMimeType`: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
   - `base64Content`: `sha256:<해시>:<out.xlsx 절대 경로, 폴더 구분은 />`
   - `disableConversionToGoogleType`는 넣지 않는다. Google Sheet로 변환되어야 콘솔 작업현황이 읽는다.
4. 돌려받은 `viewUrl`을 PD에게 준다. 콘솔 작업현황의 "작업 폴더에서 찾기"를 누르면 이 시트가 목록에 나온다.

## 아직 넘기지 않는 값

콘솔에서 고친 Same Camera, Cam Move, BG·CH·PR은 cuts.json에 들어 있지만 build_worksheet.py가 시나리오에서 다시 도출한다. 이 값이 달라졌으면 올린 뒤 시트에서 고치게 PD에게 알린다.
