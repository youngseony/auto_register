# -*- coding: utf-8 -*-
"""등록이 끝난 뒤 결과 요약 출력 + 결과 파일(CSV) 저장, 등록 전 확인 요약의 공통 부분"""
import csv
from datetime import datetime

import config
from common import ROOT
from common.console import issues


def out_dir():
    path = ROOT / "결과"
    path.mkdir(exist_ok=True)
    return path


def mode_line():
    """등록 전 요약에 넣는 '실행 방식' 안내 문구."""
    if getattr(config, "DRY_RUN", False):
        return "시험 모드 (입력까지만 하고 저장하지 않습니다)"
    if not getattr(config, "AUTO_SUBMIT", True):
        return "실제 등록 (한 건마다 화면을 확인하고 Enter로 저장합니다)"
    return "실제 등록 (진행하면 사이트에 실제로 등록·저장됩니다)"


def print_load_summary(title, unit, loaded, detail_lines, note_lines=()):
    """파일을 읽은 결과(건수, 오류 항목)를 한눈에 보여준다.
    detail_lines: 등록 가능한 항목을 설명하는 줄들 / unit: '건' 또는 '명'"""
    bar = "=" * 60
    print("\n" + bar)
    print(f" 파일 확인 결과 - {title}")
    print(bar)
    for name, total, bad in loaded.per_file:
        state = f"오류 {bad}{unit} 제외" if bad else "이상 없음"
        print(f" 파일: {name}  (읽은 줄 {total}{unit} / {state})")
    print(f"\n 등록할 수 있는 항목: {len(loaded.rows)}{unit}")
    for line in detail_lines:
        print(f"   {line}")
    if loaded.invalid:
        print(f"\n [!] 오류가 있는 줄: {len(loaded.invalid)}{unit}  (진행하면 이 줄은 자동 등록에서 제외됩니다)")
        for where, a, b, problems in loaded.invalid:
            print(f"   - {where}  {a} / {b}")
            for p in problems:
                print(f"       · {p}")
    for line in list(loaded.notes) + list(note_lines):
        print(f"\n [참고] {line}")
    print(f"\n 실행 방식: {mode_line()}")
    print(bar)


def print_report(*, total, done, unit, group_title, key_names, results=None):
    """모든 등록이 끝난 뒤, 결과 요약과 문제가 있었던 항목만 모아 출력하고 CSV로 저장한다.
    unit: '건'/'명' / group_title: '수동으로 수정이 필요한 교육과정' 등 / key_names: 항목 키 이름 3개(번호, 항목1, 항목2)
    results: (교육생 등록처럼) 전체 처리 결과를 파일로도 남길 때의 dict 목록"""
    issues.current = None
    stamp = f"{datetime.now():%Y%m%d_%H%M%S}"
    print("\n" + "=" * 60)
    print(f"처리 결과: 총 {total}{unit} 중 {done}{unit} 등록 완료")

    if results:
        counts = {}
        for r in results:
            counts[r["결과"]] = counts.get(r["결과"], 0) + 1
        print("  " + ", ".join(f"{k} {v}{unit}" for k, v in counts.items()))
        path = out_dir() / f"등록결과_{stamp}.csv"
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=list(results[0].keys()))
            writer.writeheader()
            writer.writerows(results)
        print(f"  전체 결과 파일: {path}")

    if not issues.issues:
        print("오류/수동 확인이 필요한 항목이 없습니다.")
        print("=" * 60)
        return

    grouped = {}
    seen = set()
    for key, level, message in issues.issues:
        if (key, message) in seen:
            continue
        seen.add((key, message))
        grouped.setdefault(key, []).append((level, message))

    print(f"{group_title}: {len(grouped)}{unit}")
    print("-" * 60)
    for (no, a, b), items in grouped.items():
        print(f"[{no}] {a} / {b}" if isinstance(no, str) else f"[{no}번째] {a} / {b}")
        for level, message in items:
            print(f"   - ({level}) {message}")

    out_path = out_dir() / f"수동작업필요_{stamp}.csv"
    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow([*key_names, "구분", "내용"])
        for (no, a, b), items in grouped.items():
            for level, message in items:
                writer.writerow([no, a, b, level, message])
    print("-" * 60)
    print(f"위 내용은 파일로도 저장했습니다: {out_path}")
    print("=" * 60)
