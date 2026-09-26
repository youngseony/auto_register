# -*- coding: utf-8 -*-
"""엑셀/CSV 파일 선택과 읽기 (두 작업 공통)"""
import csv
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import openpyxl

from common import ROOT
from common.console import input, log


@dataclass
class Loaded:
    """파일을 읽어 확인한 결과. 등록 전에 사용자에게 요약해서 보여준다."""
    rows: list = field(default_factory=list)      # 등록할 수 있는 줄
    invalid: list = field(default_factory=list)   # (위치, 항목1, 항목2, [문제점]) - 형식 문제로 제외된 줄
    per_file: list = field(default_factory=list)  # (파일명, 읽은 줄 수, 제외된 줄 수)
    notes: list = field(default_factory=list)     # 참고/주의 문구


def find_input_files():
    """프로젝트 루트 폴더의 엑셀(.xlsx/.xlsm)·CSV 파일 목록 (엑셀 임시파일 ~$ 제외)"""
    return [
        p for p in sorted(ROOT.iterdir())
        if p.is_file()
        and p.suffix.lower() in (".xlsx", ".xlsm", ".csv")
        and not p.name.startswith("~$")
    ]


def choose_input_files(label):
    """파일 탐색기 창에서 파일(엑셀/CSV)을 선택한다. 창을 열 수 없으면 루트 폴더 자동 탐색으로 대체."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        win = tk.Tk()
        win.withdraw()
        win.attributes("-topmost", True)
        paths = filedialog.askopenfilenames(
            parent=win,
            title=f"{label}(엑셀/CSV)을 선택하세요",
            initialdir=str(ROOT),
            filetypes=[(label, "*.xlsx *.xlsm *.csv"), ("모든 파일", "*.*")],
        )
        win.destroy()
    except Exception as e:  # tkinter 미설치/화면 없음 등
        log.warning(f"파일 선택 창을 열 수 없어 루트 폴더에서 찾습니다: {e}")
        return choose_input_files_from_root(label)
    if not paths:
        raise SystemExit("파일을 선택하지 않아 종료합니다.")
    return [Path(p) for p in paths]


def choose_input_files_from_root(label):
    files = find_input_files()
    if not files:
        raise SystemExit(f"루트 폴더에 엑셀(.xlsx) 또는 CSV(.csv) 파일이 없습니다. {label}을 넣고 다시 실행하세요.")
    if len(files) == 1:
        return files
    print("\n루트 폴더에서 여러 개의 파일을 찾았습니다:")
    for i, p in enumerate(files, 1):
        print(f"  {i}. {p.name}")
    answer = input("처리할 파일 번호를 입력하세요 (쉼표로 여러 개, a=전체): ").strip().lower()
    if answer == "a":
        return files
    picked = [files[int(n) - 1] for n in re.split(r"[,\s]+", answer) if n.isdigit() and 1 <= int(n) <= len(files)]
    if not picked:
        raise SystemExit("올바른 번호가 입력되지 않아 종료합니다.")
    return picked


def _read_csv_rows(path):
    for enc in ("utf-8-sig", "cp949"):
        try:
            with open(path, newline="", encoding=enc) as f:
                return list(csv.DictReader(f))
        except UnicodeDecodeError:
            continue
        except PermissionError:
            raise SystemExit(f"{path.name}: 파일을 읽을 수 없습니다. 엑셀 등에서 열려 있으면 닫고 다시 시도하세요.")
    raise SystemExit(f"{path.name}: CSV 인코딩을 읽지 못했습니다. UTF-8 또는 CP949(엑셀 기본)로 저장해주세요.")


def _read_excel_rows(path, key_headers):
    """엑셀에서 key_headers(필요한 열 이름들)가 모두 있는 시트를 찾아 행(dict 목록)을 읽는다.
    (마지막에 저장한 시트가 아닐 수 있으므로 전체 시트를 확인. 못 찾으면 첫 시트를 돌려준다)"""
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except PermissionError:
        raise SystemExit(f"{path.name}: 파일을 읽을 수 없습니다. 엑셀에서 열려 있으면 닫고 다시 시도하세요.")
    except Exception as e:
        raise SystemExit(
            f"{path.name}: 엑셀 파일을 열 수 없습니다 ({type(e).__name__}). "
            "파일이 손상되었거나 .xlsx 형식이 아닐 수 있습니다. 엑셀에서 '다른 이름으로 저장 > Excel 통합 문서(.xlsx)'로 다시 저장해 보세요."
        )
    first_rows = None
    for ws in [wb.active] + [s for s in wb.worksheets if s is not wb.active]:
        headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
        raw_rows = [dict(zip(headers, r)) for r in ws.iter_rows(min_row=2, values_only=True)]
        if first_rows is None:
            first_rows = raw_rows
        if all(h in headers for h in key_headers):
            return raw_rows
    return first_rows or []


def read_table(path, key_headers):
    """엑셀/CSV 파일의 1행(제목)을 열 이름으로 삼아 행(dict 목록)을 읽는다."""
    if path.suffix.lower() == ".csv":
        return _read_csv_rows(path)
    return _read_excel_rows(path, key_headers)


def txt(value):
    """셀 값을 글자로 바꾼다. 빈 칸(None)은 빈 글자로, 숫자로 저장된 ID(610729.0)는 정수 글자로."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def clean_row(raw):
    """열 이름 앞뒤 공백을 없애고, 열 이름이 없는 칸은 버린다."""
    return {str(k).strip(): v for k, v in raw.items() if k is not None}


_WEEKDAYS = "월화수목금토일"


def fmt_ref(value):
    """안내 문구용 참고 값(일자 등)을 글자로. 엑셀 날짜 셀은 '2026. 09. 28' 모양으로."""
    if isinstance(value, datetime):
        return value.strftime("%Y. %m. %d")
    return txt(value)


def fmt_date_weekday(date_value, weekday_value):
    """일자 글자에 요일을 붙인다: '2026. 09. 28 (월)'. 요일 칸이 비어 있으면 일자로 계산하고, 계산도 안 되면 일자만."""
    text = fmt_ref(date_value)
    weekday = txt(weekday_value).replace("요일", "")[:1]
    if not weekday or weekday not in _WEEKDAYS:
        weekday = ""
        m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", text)
        if m:
            try:
                weekday = _WEEKDAYS[datetime(int(m[1]), int(m[2]), int(m[3])).weekday()]
            except ValueError:
                pass
    return f"{text} ({weekday})" if text and weekday else text
