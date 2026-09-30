# -*- coding: utf-8 -*-
"""로그 기록, 오류/수동작업 항목 수집, 수동 개입(">>" 안내) 입력 대기"""
import logging
import re
import sys
import time

import config
from common import ROOT

try:
    import msvcrt  # Windows 콘솔에서 제한 시간 입력 대기에 사용
except ImportError:
    msvcrt = None

log = logging.getLogger("auto_register")


def setup_logging():
    """로그 파일(config.LOG_FILE)과 화면 출력을 설정한다. 여러 번 불러도 한 번만 설정된다."""
    root = logging.getLogger()
    if getattr(root, "_auto_register_ready", False):
        return
    root._auto_register_ready = True
    logging.basicConfig(
        filename=ROOT / config.LOG_FILE,
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        encoding="utf-8",
    )
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    console.setFormatter(logging.Formatter("%(message)s"))
    root.addHandler(console)
    root.addHandler(issues)
    # 종료 시 selenium이 chromedriver를 정리하다 남기는 무해한 오류 로그가 화면에 나오지 않게 한다.
    logging.getLogger("selenium").setLevel(logging.CRITICAL)


def mask_id(member_id):
    """로그에 남기는 회원ID는 앞 3글자만 보이게 가린다."""
    member_id = member_id or ""
    return (member_id[:3] + "***") if member_id else "(ID없음)"


def mask_phone(text):
    """휴대폰 번호(숫자만)는 끝 4자리만 남긴다. 그 밖의 글자는 그대로 둔다."""
    return f"휴대폰***{text[-4:]}" if re.fullmatch(r"01\d{8,9}", text or "") else text


# ------------------------------------------- 오류/수동작업 항목 수집·정리 ----
# 진행 중 발생하는 경고/오류 로그와 수동 개입 안내(">>" 프롬프트)를 항목(교육과정/교육생)별로 모아 두었다가
# 모든 등록이 끝난 뒤 한 번에 정리해서 출력한다.
class IssueCollector(logging.Handler):
    def __init__(self):
        super().__init__(level=logging.WARNING)
        self.current = None  # (번호, 항목1, 항목2) - 지금 처리 중인 항목
        self.issues = []     # ((번호, 항목1, 항목2), 구분, 내용)
        self.mark = 0        # 직전 수동 안내 이후에 쌓인 항목만 '해결됨' 판정 대상으로 삼기 위한 위치
        self.verify = None   # 함수(항목이름) -> 채워졌으면 True, 비었으면 False, 알 수 없으면 None

    def emit(self, record):
        if self.current:
            self.issues.append((self.current, record.levelname, record.getMessage().strip()))


issues = IssueCollector()

_builtin_input = input


def _input_with_timeout(prompt, timeout):
    """Enter를 기다리되 timeout(초)이 지나면 None을 반환 (Windows 콘솔 전용)."""
    sys.stdout.write(prompt)
    sys.stdout.flush()
    buf = []
    end = time.time() + timeout
    while time.time() < end:
        if not msvcrt.kbhit():
            time.sleep(0.05)
            continue
        ch = msvcrt.getwche()
        if ch in ("\r", "\n"):
            print()
            return "".join(buf)
        if ch == "\x03":
            raise KeyboardInterrupt
        if ch in ("\x00", "\xe0"):  # 방향키 등 특수키는 두 글자로 들어오므로 나머지를 버린다
            msvcrt.getwch()
        elif ch == "\b":
            if buf:
                buf.pop()
            sys.stdout.write(" \b")
            sys.stdout.flush()
        else:
            buf.append(ch)
    return None


# 대괄호 이름이 없는 안내(교육시간표·지역)는 이 낱말이 들어 있는 항목끼리 짝을 짓는다. (교육계획 등록용)
_KEYWORDS = {
    "교육시간표": ("교육시간표",),
    "교육시간표 일차": ("교육시간표", "일차"),
    "지역": ("시/도", "지역 선택", "읍면동"),
}


def _entry_label(message):
    """'[담당 배움터] ...' 처럼 맨 앞 대괄호 안의 항목 이름을 꺼낸다. 없으면 None."""
    m = re.match(r"\s*\[(.+?)\]", message)
    return m.group(1) if m else None


def _prompt_key(prompt_text):
    """수동 안내(>> ...)가 어느 항목에 대한 것인지 알아낸다. 알 수 없으면 None."""
    head = prompt_text.split(">>", 1)[1]
    m = re.match(r"\s*\[(.+?)\]", head)
    if m:
        return m.group(1)
    if "교육시간표를 직접" in head:
        return "교육시간표"
    if "2일차 이후" in head:
        return "교육시간표 일차"
    if "지역 선택" in head:
        return "지역"
    return None


def _resolve_issues(key):
    """직접 처리해서 해결된 것이 확인된 항목을 현재 항목의 수집 목록에서 뺀다."""
    kept = []
    for idx, (row_key, level, message) in enumerate(issues.issues):
        if row_key == issues.current and idx >= issues.mark:
            entry = _entry_label(message)
            if entry is not None:
                match = entry == key
            else:
                match = any(word in message for word in _KEYWORDS.get(key, ()))
            if match:
                continue
        kept.append((row_key, level, message))
    issues.issues[:] = kept


def input_timeout(prompt, seconds):
    """Enter를 seconds초 기다려 입력값을 돌려준다. 시간이 지나면 None (seconds가 0이거나 콘솔이 아니면 끝없이 기다림)."""
    if not seconds or msvcrt is None:
        return _builtin_input(prompt)
    return _input_with_timeout(prompt, seconds)


def input(prompt=""):  # noqa: A001 - 수동 개입 요청(">> ...")을 기록·판정하고 제한 시간을 두는 래퍼
    text = str(prompt)
    intervention = ">>" in text  # 자동 처리에 실패해 사용자에게 직접 처리를 요청하는 안내
    if issues.current and intervention:
        issues.issues.append((issues.current, "수동개입", text.replace(">>", "").strip()))

    wait = getattr(config, "MANUAL_WAIT_SECONDS", 30)
    timed_out = False
    if (intervention or "계속할까요" in text) and wait and msvcrt is not None:
        print(f"  (안내에 따라 {wait}초 안에 처리하고 Enter를 누르세요. 그냥 Enter를 누르거나, 시간이 지나면 다음 작업으로 넘어갑니다)")
        answer = _input_with_timeout(text, wait)
        if answer is None:
            timed_out = True
            answer = ""
            print(f"\n  ({wait}초가 지나 다음 작업으로 넘어갑니다)")
    else:
        answer = _builtin_input(prompt)

    if intervention and issues.current:
        if timed_out:
            issues.issues.append((issues.current, "시간초과", "직접 처리 없이 넘어갔습니다. 해당 항목을 확인하고 수동으로 처리하세요."))
        else:
            # Enter를 눌렀다고 해결된 것으로 보지 않는다. 화면에서 실제로 처리된 것이 확인될 때만
            # 목록에서 빼고, 처리되지 않았거나 확인할 수 없으면(애매하면) 수동 처리 대상으로 남긴다.
            key = _prompt_key(text)
            done = issues.verify(key) if (issues.verify and key) else None
            if done is True:
                _resolve_issues(key)
            elif done is False:
                issues.issues.append((issues.current, "미해결", "직접 처리했다고 하셨지만 화면에서 처리된 것이 확인되지 않습니다. 수동으로 확인하세요."))
            else:
                issues.issues.append((issues.current, "확인불가", "직접 처리했는지 화면에서 확인할 수 없어 목록에 남깁니다. 결과를 확인하고 수동으로 정리하세요."))
        issues.mark = len(issues.issues)
    return answer
