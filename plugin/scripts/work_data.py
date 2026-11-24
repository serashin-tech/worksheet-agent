"""작업 현황 대시보드 시연 데이터: 정답 시트(EP12)의 공정 열(상태·작업자)을 읽어 work.json으로 만든다.
실제 운영에서는 새로 만든 시트의 같은 열을 Sheets 커넥터로 읽는다.
사용: python -I work_data.py <xlsx> <out.json>
"""
import sys, json, openpyxl
# (구분, 이름, 상태 열, 작업자 열) 열 번호는 1부터. Render TK02~04는 작업자 열이 따로 없어 TK01 작업자를 쓴다.
TAKES = [("LO", "LO", 15, 16)] + [("AN", f"AN TK0{i}", c, c + 1) for i, c in zip(range(1, 5), (18, 21, 24, 27))] \
    + [("Render", f"Render TK0{i}", c, 32) for i, c in zip(range(1, 5), (31, 34, 36, 38))] \
    + [("Comp", f"Comp TK0{i}", c, c + 1) for i, c in zip(range(1, 6), (40, 42, 44, 46, 48))]


def main(xlsx, out):
    ws = openpyxl.load_workbook(xlsx).worksheets[0]
    v = lambda r, c: (str(ws.cell(r, c).value).strip() if ws.cell(r, c).value not in (None, "") else "")
    rows = list(range(2, 181))
    data = {"ids": [v(r, 1) for r in rows], "takes": [{"g": g, "k": k} for g, k, _, _ in TAKES],
            "cells": {k: [[v(r, s), v(r, w)] for r in rows] for _, k, s, w in TAKES}}
    json.dump(data, open(out, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    print(len(rows), "rows", len(TAKES), "takes")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(*sys.argv[1:3])
