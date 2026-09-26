# -*- coding: utf-8 -*-
"""
AI디지털배움터(www.xn--2z1bw8k1pjz5ccumkb.kr) 등록 업무 자동화 (auto_register)

사용법: README.md 참고
    1) setup.bat 실행 (Python/가상환경/라이브러리 설치, 처음 한 번)
    2) run.bat 실행 (= venv의 python 으로 이 파일 실행)
    3) 작업 선택: 1. 교육계획 등록  2. 교육생 등록
    4) 파일 선택 창에서 엑셀/CSV를 고른다
    5) 파일 확인 결과(건수, 오류가 있는 줄)를 보고 '진행할까요?' 에 답한다
       (Y=진행 / F=다른 파일 선택 / N=취소)
    6) 크롬이 열리면 직접 로그인 후, 터미널에서 Enter -> 자동 등록 -> 결과 요약

주의: 이 프로그램은 아이디/비밀번호를 다루지 않습니다. 로그인은 항상 사용자가 직접 합니다.
"""

import importlib
import sys

from selenium.common.exceptions import InvalidSessionIdException, NoSuchWindowException

import config
from common.console import input, log, setup_logging
from common.files import choose_input_files

# 메뉴 번호 -> (작업 이름 표시, 작업 모듈, 명령줄에서 쓸 이름)
TASKS = {
    "1": ("교육계획 등록", "tasks.edu_plan", "plan"),
    "2": ("교육생 등록", "tasks.student", "student"),
}

_YES = ("y", "yes", "ㅛ")
_NO = ("n", "no", "q", "x", "ㅜ", "ㅂ", "ㅌ")
_OTHER_FILE = ("f", "ㄹ")


def choose_task():
    """메뉴에서 작업을 고른다. (명령줄에 1/2/plan/student 를 주면 메뉴를 건너뛴다)"""
    if len(sys.argv) > 1:
        wanted = sys.argv[1].strip().lower()
        for number, (_, module, short) in TASKS.items():
            if wanted in (number, short):
                return importlib.import_module(module)
        raise SystemExit(f"알 수 없는 작업입니다: {sys.argv[1]}  (1 또는 plan = 교육계획 등록, 2 또는 student = 교육생 등록)")

    print("\n" + "=" * 60)
    print(" AI디지털배움터 등록 업무 자동화")
    print("=" * 60)
    print(" 어떤 작업을 진행할까요?")
    for number, (label, _, _) in TASKS.items():
        print(f"   {number}. {label}")
    print("   q. 종료")
    while True:
        answer = input("\n번호를 입력하고 Enter: ").strip().lower()
        if answer in TASKS:
            return importlib.import_module(TASKS[answer][1])
        if answer in _NO:
            raise SystemExit("종료합니다.")
        print("  1, 2, q 중에서 입력해 주세요.")


def confirm_and_load(task):
    """파일을 고르고 내용을 확인해서 보여준 뒤, '진행할까요?'에 답을 받는다.
    반환: 등록할 내용(Loaded)"""
    paths = choose_input_files(task.FILE_LABEL)
    while True:
        try:
            loaded = task.load(paths)
        except SystemExit as e:  # 파일이 열려 있거나 열 이름이 다른 경우 등: 안내 후 다시 시도할 수 있게 한다
            print(f"\n[!] {e}")
            loaded = None
        else:
            task.show_summary(loaded)

        can_go = loaded is not None and bool(loaded.rows)
        if loaded is not None and loaded.invalid:
            print("\n 진행할까요? (오류사항은 수동으로 입력하셔야 합니다)")
        else:
            print("\n 진행할까요?")
        if can_go:
            if loaded.invalid:
                print(f"   Y : 진행 (오류가 있는 {len(loaded.invalid)}줄은 제외하고 등록합니다)")
            else:
                print("   Y : 진행")
        print("   F : 다른 파일 선택")
        print("   N : 취소")
        while True:
            answer = input("선택하고 Enter (Y/F/N): ").strip().lower()
            if answer in _YES and can_go:
                return loaded
            if answer in _YES:
                print("  등록할 수 있는 항목이 없어 진행할 수 없습니다. F, N 중에서 선택해 주세요.")
            elif answer in _OTHER_FILE:
                paths = choose_input_files(task.FILE_LABEL)
                break
            elif answer in _NO:
                raise SystemExit("사용자가 취소했습니다. 등록을 진행하지 않았습니다.")
            else:
                print("  Y, F, N 중에서 입력해 주세요.")


def main():
    setup_logging()
    task = choose_task()
    print(f"\n[{task.TITLE}]")
    loaded = confirm_and_load(task)
    task.run(loaded)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except KeyboardInterrupt:
        log.info("Ctrl+C 로 중단되었습니다.")
        print(
            "\n[중단] Ctrl+C 가 눌려 프로그램을 멈췄습니다.\n"
            "  run.bat 을 다시 실행하면 처음부터 다시 할 수 있습니다."
        )
        raise SystemExit(1)
    except (InvalidSessionIdException, NoSuchWindowException) as e:
        log.error(f"크롬 창이 닫혀 중단되었습니다: {type(e).__name__}")
        print(
            "\n[중단] 자동 등록용 크롬 창이 닫혀서 더 진행할 수 없습니다.\n"
            "  프로그램이 연 크롬 창은 끝날 때까지 닫지 말고 그대로 두세요.\n"
            "  run.bat 을 다시 실행하면 처음부터 다시 할 수 있습니다."
        )
        raise SystemExit(1)
    except Exception as e:
        log.exception("예상하지 못한 오류가 발생했습니다.")
        print(
            "\n[오류] 예상하지 못한 문제가 생겨 중단되었습니다.\n"
            f"  내용: {type(e).__name__}: {e}\n"
            f"  같은 폴더의 {config.LOG_FILE} 파일을 확인하세요. 해결되지 않으면 화면 내용과 로그 파일을 GitHub 이슈로 알려주세요.\n"
            "  (그때까지 등록된 내용은 크롬의 '교육계획 관리' 또는 '교육생관리(실시)' 목록에서 확인할 수 있습니다.)"
        )
        raise SystemExit(1)
