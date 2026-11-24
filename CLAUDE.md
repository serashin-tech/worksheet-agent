# worksheet-agent

@~/.claude/dev-rules-plugin.md

토스트 애니메이션 스튜디오 PD가 에피소드별 애니메이션 Worksheet(Google Sheet) 초안을 만들도록 돕는 Claude 플러그인 저장소다. 저장소 루트가 마켓플레이스(`worksheet-agent`)이고 `plugin/`이 플러그인이다. 별도 저장소인 `toast-agent`(정부사업·시장 정보 에이전트)와 코드, 설정, 문서를 공유하지 않는다.

## 먼저 읽을 것

- 개발 계획과 결정, PoC 결과: `docs/PLAN.md`
- 협의 내용(토스트 대표, 2026-11-23): 초안은 Description 열까지, BG·CH·PR은 스토리보드에 없으면 공란, 이미지·벡터 DB 매칭 제외, 요약은 별도 탭, 완료 체크는 작업자 수동

## 구조

- `plugin/scripts/`: 결정적 코드. 대본 파싱(`parse_docx.py`), 콘티 칸 분할(`panel_split.py`), 썸네일, 에셋 도출(`extract_assets.py`, `folder_assets.py`), 시트 조립(`build_worksheet.py`), 검증(`verify_worksheet.py`), 평가(`eval_rows.py`), Frame 추정 계수(`frame_model.json`)
- `plugin/console/`: 시연용 콘솔 artifact 원본(`console.html`)과 시연 데이터(JSON, 로고, Pretendard 글꼴). 게시본은 artifact `AGBNkhgLQkdopHBhSc3eSy`
- `plugin/assets/`: 토스트 로고
- `plugin/skills/`, `plugin/agents/`, `plugin/hooks/`: 아직 없음. PLAN.md §7 구성대로 만든다

## 고객 원본(이 저장소에 두지 않음)

EP12 원작 대본 docx, 콘티 PDF, 정답 시트는 `D:\12_BidProposal\20260629_AI 활용 가상융합 기술 지원 및 컨설팅\06_기술지원 및 컨설팅\토스트 애니메이션 스튜디오\01_레퍼런스 데이터\`에 있고 읽기만 한다. 스크립트는 경로를 인자로 받는다(`python -I <스크립트> <docx> <xlsx>`).

## 알려진 불일치

`docs/PLAN.md`에는 아직 24FPS 결정이 적혀 있다. 코드는 2026-11-23 협의대로 29.97(`ROUND(Frame/29.97,1)`)이며, 방영 길이(분)·목표 행 수 입력, 요약 탭, 합계 비교 검증(`verify_worksheet.py`, 허용 오차 0.5분은 토스트 확인 전 가정)을 반영했다. PLAN.md는 아직 고치지 않았다. EP12 추정 Frame의 총 러닝타임은 10.72분이라 방영 10분 검증이 FAIL로 나온다(정답 시트는 10.86분).

## 이전 위치

개발은 `12_BidProposal\…\07_Implementation\worksheet-pipeline\`(`poc\`)에서 시작했고, 2026-11-23에 이 저장소로 복사했다. 앞으로의 수정은 이 저장소에서만 한다. 원래 폴더는 지우지 않았다.
