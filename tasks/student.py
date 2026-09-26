# -*- coding: utf-8 -*-
"""
작업 2. 교육생 등록 (강사페이지 > 교육실시 관리 > 교육생실시관리)

엑셀/CSV의 교육생 목록(교육실시ID, 일자, 요일, 과정명, 기관명, 이름, 회원ID, 휴대폰)을 읽어 교육별로
'교육생실시관리' 클릭 -> 팝업 확인 후, 교육생마다 '교육생등록' -> '회원검색'(회원ID -> 휴대폰 -> 이름 순) -> 선택
-> 기자재/영상교육 입력 -> '저장' 까지 진행한다.
교육실시ID로 '교육실시 관리' 목록에서 교육을 자동으로 찾아 체크하고(COURSE_SELECT_MODE="auto"), 못 찾으면
사용자가 교육 1개를 체크하고 Enter를 누른다. 자동 처리에 실패한 단계는 터미널 안내에 따라 브라우저에서 직접 마무리하고
Enter를 누르면 이어서 진행되며(30초 안에 처리하지 않으면 다음 작업으로 자동 진행),
모든 등록이 끝나면 문제가 있던 항목만 모아 출력/저장한다(결과 폴더).

주의: 아이디/비밀번호를 다루지 않습니다. 로그인은 항상 사용자가 직접 합니다.
"""

import re
import time

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
    WebDriverException,
)

import config
from common.browser import BaseRegistrar, start_registrar
from common.console import input, issues, log, mask_id, mask_phone
from common.files import Loaded, clean_row, fmt_date_weekday, read_table, txt as _txt
from common.report import print_load_summary, print_report

TITLE = "교육생 등록"
FILE_LABEL = "교육생 파일"

# 저장 뒤 사이트가 띄우는 안내창 중 '필수 항목 누락/실패'를 뜻하는 낱말 (예: '거주지역 시군구를 선택해주십시요.')
_SAVE_FAIL_WORDS = ("선택해", "입력해", "필수", "오류", "실패", "올바르", "확인해")

# 오류 유형: 마지막 '오류 유형별 정리'에서 이 순서로 묶어 보여 준다.
KIND_FILE = "파일 형식 오류"
KIND_COURSE = "교육 선택 실패"
KIND_MEMBER = "회원 검색·선택 실패"
KIND_RESIDENCE = "거주지역 미입력"
KIND_INPUT = "항목 자동 입력 실패"
KIND_SAVE = "저장 실패"
KIND_ERROR = "처리 중 오류"
_KIND_HELP = {
    KIND_FILE: "파일의 해당 줄을 고친 뒤 다시 실행하거나 직접 등록하세요.",
    KIND_COURSE: "교육실시ID가 맞는지 확인하고 직접 등록하세요.",
    KIND_MEMBER: "회원ID·휴대폰이 맞는지 확인하고 직접 등록하세요. (동명이인, 미가입 등)",
    KIND_RESIDENCE: "교육생 등록 화면에서 거주지역(시/도, 시군구)을 직접 등록해주세요.",
    KIND_INPUT: "크롬의 교육생등록 화면에서 해당 항목(교육보유기자재 등)을 확인하고 직접 등록하세요.",
    KIND_SAVE: "사이트 안내(정원 초과 등)를 확인하고 직접 저장하세요.",
    KIND_ERROR: "auto_register.log 를 확인하고 직접 등록하세요.",
}


def _wait_for_manual():
    """예상하지 못한 오류(회원 정보 부족, 사이트 안내창, 저장 확인 실패 등)를 만났을 때 직접 처리하도록 기다릴지
    (config.ERROR_ACTION == "wait"). 기본은 기다리지 않고 건너뛰어, 마지막 '오류 유형별 정리'에 모아 알려 준다."""
    return getattr(config, "ERROR_ACTION", "skip") == "wait"


# ------------------------------------------------------- 엑셀/CSV 읽기 ----
def normalize_phone(value):
    """휴대폰 번호에서 숫자만 남긴다. 엑셀이 앞의 0을 지운 경우(1082237520)는 0을 되살린다."""
    digits = re.sub(r"\D", "", _txt(value))
    if len(digits) == 10 and digits.startswith("10"):
        digits = "0" + digits
    return digits


def _read_file(path, loaded, seen):
    """엑셀/CSV 파일 하나에서 교육생 행을 읽어 loaded 에 더한다. 헤더는 1행, 열 이름은 config.COL_* 기준.
    rows 항목: {'edu_id', 'mem_id', 'name', 'phone', 'ref', 'where'}  (phone은 숫자만, ref는 안내 문구용 일자/과정명/기관명)"""
    raw_rows = read_table(path, [config.COL_EDU_ID])
    if not raw_rows:
        raise SystemExit(f"{path.name}: 데이터가 없습니다. 1행은 제목, 2행부터 교육생을 적어주세요.")

    headers = {str(k).strip() for k in raw_rows[0].keys() if k is not None}
    phone_cols = [c for c in config.COL_PHONES if c in headers]
    missing = [c for c in (config.COL_EDU_ID,) if c not in headers]
    if not ({config.COL_MEM_ID, config.COL_MEM_NAME} & headers) and not phone_cols:
        missing.append(f"{config.COL_MEM_ID} / {config.COL_MEM_NAME} / {'·'.join(config.COL_PHONES)} 중 하나")
    if missing:
        raise SystemExit(
            f"{path.name}: 필요한 열이 없습니다 -> {missing}\n"
            f"  파일의 1행(제목) 이름이 README의 표와 똑같은지 확인하세요. 현재 제목: {sorted(headers)}"
        )

    total = bad = 0
    for n, raw in enumerate(raw_rows, start=2):  # 엑셀 행 번호 (1행은 제목)
        row = clean_row(raw)
        edu_id = _txt(row.get(config.COL_EDU_ID)).upper()
        mem_id = _txt(row.get(config.COL_MEM_ID))
        name = _txt(row.get(config.COL_MEM_NAME))
        phone = next((normalize_phone(row.get(c)) for c in phone_cols if _txt(row.get(c))), "")
        if not edu_id and not mem_id and not name and not phone:
            continue  # 빈 줄
        total += 1

        where = f"{path.name} {n}행"
        who = mem_id or name or phone or "(회원 없음)"
        problems = []
        if not edu_id:
            problems.append(f"'{config.COL_EDU_ID}'이(가) 비어 있습니다")
        elif not re.fullmatch(r"[A-Z0-9]+", edu_id):
            problems.append(f"'{config.COL_EDU_ID}' 형식이 올바르지 않습니다 (값: {edu_id!r}, 예: E2600997582)")
        if not mem_id and not name and not phone:
            problems.append(f"'{config.COL_MEM_ID}', '{config.COL_MEM_NAME}', 휴대폰이 모두 비어 있습니다")
        if phone and not re.fullmatch(r"01\d{8,9}", phone):
            problems.append("휴대폰 번호 형식이 올바르지 않습니다 (예: 010-1234-5678)")
        key = (edu_id, mem_id.lower() or phone or name)
        if not problems and key in seen:
            problems.append("같은 교육실시ID에 같은 교육생이 중복되어 있습니다")
        if problems:
            bad += 1
            loaded.invalid.append((where, edu_id or "(교육실시ID 없음)", mask_phone(who), problems))
            continue
        seen.add(key)
        loaded.rows.append({
            "edu_id": edu_id, "mem_id": mem_id, "name": name, "phone": phone, "where": where,
            "ref": {
                "date": fmt_date_weekday(
                    row.get(config.COL_EDU_DATE),
                    next((row.get(c) for c in config.COL_WEEKDAY if _txt(row.get(c))), ""),
                ),
                "course": _txt(row.get(config.COL_EDU_NAME)),
                "org": _txt(row.get(config.COL_ORG)),
            },
        })
    loaded.per_file.append((path.name, total, bad))


def load(paths):
    """파일들을 읽어 등록할 수 있는 교육생과 오류가 있는 줄을 나눈다. (파일 자체를 읽을 수 없으면 SystemExit)"""
    loaded = Loaded()
    seen = set()
    for path in paths:
        _read_file(path, loaded, seen)

    limit = getattr(config, "TEST_LIMIT", None)
    if limit and len(loaded.rows) > limit:
        loaded.notes.append(f"TEST_LIMIT={limit} 설정으로 앞의 {limit}명만 등록합니다. (전체를 등록하려면 config.py 에서 TEST_LIMIT = None)")
        loaded.rows = loaded.rows[:limit]

    name_only = sum(1 for s in loaded.rows if not s["mem_id"] and not s["phone"])
    if name_only:
        loaded.notes.append(
            f"회원ID·휴대폰 없이 이름만 있는 교육생이 {name_only}명 있습니다. 동명이인이면 자동 선택하지 않고 수동 안내가 나옵니다. "
            "가능하면 회원ID나 휴대폰을 적어 주세요."
        )
    return loaded


def group_by_education(students):
    """교육실시ID별로 묶는다 (파일에 처음 나온 순서 유지). 같은 교육은 팝업을 한 번만 연다."""
    groups = {}
    for s in students:
        groups.setdefault(s["edu_id"], []).append(s)
    return groups


def show_summary(loaded):
    """등록 전 확인 화면: 교육별 인원과 오류가 있는 줄을 보여준다."""
    groups = group_by_education(loaded.rows)
    lines = [f"교육 {len(groups)}개  (교육실시ID / 인원 / 일자 / 과정명 / 기관명)"]
    for edu_id, students in groups.items():
        ref = next((s["ref"] for s in students if s["ref"]["date"] or s["ref"]["course"]), {})
        info = " / ".join(x for x in (ref.get("date", ""), ref.get("course", ""), ref.get("org", "")) if x)
        lines.append(f"- {edu_id}  {len(students)}명" + (f"  {info}" if info else ""))
    print_load_summary(TITLE, "명", loaded, lines)


results = []  # {번호, 교육실시ID, 회원ID, 이름, 휴대폰, 결과, 유형, 비고}


def record_result(no, edu_id, student, status, note="", kind=""):
    results.append({
        "번호": no,
        "교육실시ID": edu_id,
        "회원ID": student["mem_id"],
        "이름": student["name"],
        "휴대폰": student["phone"],
        "결과": status,
        "유형": kind,
        "비고": note,
    })


def print_kind_summary(invalid):
    """맨 마지막에, 오류가 있었던 교육생을 유형별로 묶어 '교육실시ID / 이름 / 오류내용'으로 보여 준다."""
    groups = {}
    for where, edu_id, who, problems in invalid:
        groups.setdefault(KIND_FILE, []).append((edu_id, who, f"{where}: " + "; ".join(problems)))
    for r in results:
        if r["유형"]:
            shown = r["이름"] or r["회원ID"] or "(이름 없음)"
            if r["이름"] and r["회원ID"]:
                shown = f"{r['이름']}({r['회원ID']})"
            groups.setdefault(r["유형"], []).append((r["교육실시ID"], shown, r["비고"]))
    if not groups:
        return
    total = sum(len(v) for v in groups.values())
    print("\n" + "=" * 60)
    print(f" 오류 유형별 정리 - 직접 확인/등록이 필요한 교육생 {total}명")
    print("=" * 60)
    first = True
    for kind in _KIND_HELP:
        rows = groups.get(kind)
        if not rows:
            continue
        print(("" if first else "\n") + f"  ※ {_KIND_HELP[kind]}")
        first = False
        print(f"\n[{kind}] {len(rows)}명")
        print("     교육실시ID / 이름 / 오류내용")
        for edu_id, shown, note in rows:
            print(f"   - {edu_id} / {shown} / {note}")
    print("\n" + "=" * 60)


# --------------------------------------------------------------- 본체 ----
class StudentRegistrar(BaseRegistrar):
    # 교육생관리(실시) 팝업 목록(#stdntTB)에서 등록된 교육생의 회원ID/회원 고유번호를 읽는 스크립트
    _LIST_JS = """
        var out = [];
        document.querySelectorAll('#stdntTB > tr').forEach(function (tr) {
            var a = tr.querySelector('td[name="%s"]');
            var b = tr.querySelector('td[name="%s"]');
            out.push([a ? a.textContent.trim() : '', b ? b.textContent.trim() : '']);
        });
        return out;
    """ % (config.LIST_TD_MEM_ID, config.LIST_TD_MEM_UID)

    # 화면에 보이는 '...등록하시겠습니까' 안내 레이어에서 '확인' 버튼을 찾는다 (alert가 아닌 레이어 팝업 대비)
    _CONFIRM_BTN_JS = """
        function vis(e) { return e.getClientRects().length > 0 && getComputedStyle(e).visibility !== 'hidden'; }
        var snap = document.evaluate("//*[contains(text(),'등록하시겠습니까')]", document, null,
                                     XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
        for (var i = 0; i < snap.snapshotLength; i++) {
            var n = snap.snapshotItem(i);
            if (!vis(n)) continue;
            var el = n;
            for (var depth = 0; depth < 8 && el; depth++, el = el.parentElement) {
                var btns = el.querySelectorAll('button, a, input[type=button]');
                for (var j = 0; j < btns.length; j++) {
                    var t = (btns[j].innerText || btns[j].value || '').trim();
                    if (t === '확인' && vis(btns[j])) return btns[j];
                }
            }
        }
        return null;
    """

    def __init__(self):
        super().__init__(unhandled_prompt="ignore")
        self.edu_id = None           # 지금 팝업을 열어 둔 교육실시ID
        self.registered_ids = set()  # 이 교육에 이미 등록된 회원ID
        self.registered_uids = set() # 이 교육에 이미 등록된 회원 고유번호
        self.cur_mem_id = ""         # 지금 처리 중인 교육생 (수동 처리 후 확인용)
        self.cur_uid = None
        self._pending_edu = ""       # 수동 안내 후 확인할 교육실시ID
        self._info = {}              # 안내 문구용 교육 정보 (일자/과정명/기관명)

    # ---- 공통 유틸 -------------------------------------------------
    def _visible(self, by, sel):
        """조건에 맞는 요소 중 화면에 보이는 첫 번째 요소 (없으면 None). id가 중복된 템플릿 DOM 대비."""
        try:
            for el in self.driver.find_elements(by, sel):
                if el.is_displayed():
                    return el
        except StaleElementReferenceException:
            pass
        return None

    def _wait_visible(self, by, sel, timeout=None):
        WebDriverWait(self.driver, timeout or config.WAIT_TIME).until(lambda d: self._visible(by, sel) is not None)
        return self._visible(by, sel)

    def _value(self, css):
        try:
            return (self.driver.find_element(By.CSS_SELECTOR, css).get_attribute("value") or "").strip()
        except NoSuchElementException:
            return ""

    def _alert_text(self, timeout=1.0):
        """alert/confirm 창이 떠 있으면 내용을 읽고 확인을 눌러 닫는다. 없으면 None."""
        try:
            WebDriverWait(self.driver, timeout).until(EC.alert_is_present())
            alert = self.driver.switch_to.alert
            text = alert.text.strip()
            alert.accept()
            return text
        except TimeoutException:
            return None

    # ---- 교육실시 관리: 교육 선택 -> '교육생실시관리' 팝업 열기 -------------
    def _layer_open(self):
        """'교육생관리(실시)' 팝업(레이어)이 화면에 열려 있는지."""
        try:
            return any(e.is_displayed() and "교육생관리" in e.text for e in self._find_all(By.CSS_SELECTOR, "strong.tit"))
        except (StaleElementReferenceException, WebDriverException):
            return False

    def _layer_ids(self):
        """열려 있는 팝업이 가리키는 교육의 ID들 (교육실시ID #edc_oprtn_id + 교육계획ID #edc_plan_id)."""
        ids = set()
        try:
            for input_id in ("edc_oprtn_id", "edc_plan_id"):
                for e in self._find_all(By.ID, input_id):
                    value = (e.get_attribute("value") or "").strip().upper()
                    if value:
                        ids.add(value)
        except (StaleElementReferenceException, WebDriverException):
            pass
        return ids

    def _layer_matches(self, edu_id):
        """팝업이 열려 있고, 그 교육이 파일의 교육실시ID와 같은지 (두 종류의 ID 중 하나라도 같으면 같은 교육)."""
        return self._layer_open() and edu_id.upper() in self._layer_ids()

    def _checked_count(self):
        return len(self._find_all(By.CSS_SELECTOR, "input.mngCheck:checked"))

    def _uncheck_all_courses(self):
        try:
            self.driver.execute_script(
                "document.querySelectorAll('input.mngCheck:checked').forEach(function (c) { c.click(); });"
            )
        except WebDriverException:
            pass

    def close_stdnt_layer(self):
        """'교육생관리(실시)' 팝업을 닫고, 목록의 교육 체크를 모두 푼다."""
        try:
            for e in self._find_all(By.CSS_SELECTOR, "strong.tit"):
                if e.is_displayed() and "교육생관리" in e.text:
                    btn = self.driver.execute_script(
                        "var b = arguments[0].closest('.inBox');"
                        "return b ? b.querySelector('a[onclick*=\"fn_layer_close\"]') : null;", e)
                    if btn is not None:
                        self._js_click(btn)
                        time.sleep(0.5)
                    break
        except WebDriverException:
            pass
        self._uncheck_all_courses()
        self.edu_id = None

    def _set_list_page_size(self, size):
        """교육실시 관리 목록의 표시 개수(#sort_psize)를 바꾼다. 실패해도 무시(페이지 이동으로 대체)."""
        try:
            el = self._find(By.ID, "sort_psize", timeout=3)
            if Select(el).first_selected_option.get_attribute("value") == str(size):
                return
            before = self._checked_total()
            Select(el).select_by_value(str(size))
            self.driver.execute_script("arguments[0].dispatchEvent(new Event('change', {bubbles:true}));", el)
            # 목록이 다시 그려질 때까지 기다린다 (교육이 적어 줄 수가 그대로면 제한 시간까지만 기다림)
            end = time.time() + 4
            while time.time() < end:
                time.sleep(0.4)
                try:
                    if self._checked_total() != before:
                        break
                except StaleElementReferenceException:
                    pass
            time.sleep(0.5)
        except (TimeoutException, NoSuchElementException, StaleElementReferenceException):
            pass

    def _checked_total(self):
        """교육실시 관리 목록에 지금 보이는 교육(체크박스) 수."""
        return len(self._find_all(By.CSS_SELECTOR, "input.mngCheck"))

    def _current_page(self, pager_css):
        for el in self._find_all(By.CSS_SELECTOR, f"{pager_css} .page strong"):
            digits = re.sub(r"\D", "", el.text)
            if digits:
                return int(digits)
        return None

    def _goto_next_page(self, pager_css):
        """페이지 번호 목록에서 다음 페이지로 이동. 더 이동할 수 없으면 False."""
        cur = self._current_page(pager_css)
        if cur is None:
            return False
        target = None
        for a in self._find_all(By.CSS_SELECTOR, f"{pager_css} .page a"):
            if a.text.strip() == str(cur + 1):
                target = a
                break
        if target is None:  # 다음 묶음(예: 5 -> 6)은 '다음페이지' 버튼으로
            nxt = self._find_all(By.CSS_SELECTOR, f"{pager_css} .btnNext")
            if not nxt:
                return False
            # '다음페이지'는 마지막 페이지에서도 남아 있으므로, 이동할 페이지 번호가 현재보다 큰 경우에만 누른다.
            onclick = nxt[0].get_attribute("onclick") or ""
            m = re.search(r"pno=(\d+)", onclick) or re.search(r"\(\s*'(\d+)'", onclick)
            if not m or int(m.group(1)) <= cur:
                return False
            target = nxt[0]
        self._js_click(target)
        end = time.time() + 8
        while time.time() < end:
            time.sleep(0.4)
            try:
                now = self._current_page(pager_css)
            except StaleElementReferenceException:
                continue
            if now is not None and now != cur:
                return True
        return False

    def _find_course_checkbox(self, edu_id):
        """(auto 모드) 교육실시 관리 목록에서 edu_id 교육의 체크박스를 찾는다.
        체크박스 value(교육실시ID)와, 같은 칸의 숨은 값(교육계획ID) 어느 쪽이든 edu_id와 같으면 그 교육으로 본다."""
        xpath = f"//input[contains(@class,'mngCheck')][@value='{edu_id}' or ../input[@type='hidden']/@value='{edu_id}']"

        def look():
            els = self._find_all(By.XPATH, xpath)
            return els[0] if els else None

        self._find(By.ID, "sort_psize")  # 목록 화면이 뜰 때까지 대기
        time.sleep(0.5)
        self._set_list_page_size(config.LIST_PAGE_SIZE)  # 한 화면에 보이는 교육 수를 늘려 페이지 이동을 줄인다
        cb = look()
        if cb:
            return cb
        for _ in range(60):
            if not self._goto_next_page("#paging"):
                break
            cb = look()
            if cb:
                return cb
        return None

    def _select_only_course(self, cb):
        """목록의 다른 체크를 모두 풀고 이 교육만 체크한다."""
        self._uncheck_all_courses()
        if not cb.is_selected():
            labels = self._find_all(By.CSS_SELECTOR, f"label[for='{cb.get_attribute('id')}']")
            if labels:
                self._js_click(labels[0])
        if not cb.is_selected():
            self.driver.execute_script("arguments[0].click();", cb)

    def _open_layer_from_checked(self):
        """목록에서 체크된 교육의 '교육생실시관리'를 눌러 팝업을 연다. 열렸으면 True."""
        try:
            self._click(By.CSS_SELECTOR, "button[onclick*='fn_stdnt_mng']", timeout=3)
        except TimeoutException:
            return False
        alert = self._alert_text(1.0)
        if alert:
            log.warning(f"  [교육생실시관리] 안내창: {alert}")
            return False
        try:
            WebDriverWait(self.driver, config.WAIT_TIME).until(lambda d: self._layer_open())
            return True
        except TimeoutException:
            return False

    def _open_course_automatically(self, edu_id, first):
        """사용자 선택 없이 팝업을 열어 본다. auto 모드(첫 시도)는 목록에서 교육을 찾아 체크하고,
        그 밖에는 이미 체크된 교육이 1개일 때 그 교육의 팝업을 다시 연다. 열린 교육이 맞으면 True."""
        try:
            if config.COURSE_SELECT_MODE == "auto" and first:
                self.driver.get(config.EDU_OPRTN_URL)
                cb = self._find_course_checkbox(edu_id)
                if cb is None:
                    log.warning(f"  [교육생실시관리] 교육실시ID {edu_id} 를 교육실시 관리 목록에서 찾지 못했습니다.")
                    return False
                self._select_only_course(cb)
            elif self._layer_open() or self._checked_count() != 1:
                return False
            return self._open_layer_from_checked() and self._layer_matches(edu_id)
        except (TimeoutException, NoSuchElementException, StaleElementReferenceException) as e:
            log.warning(f"  [교육생실시관리] 자동으로 열지 못했습니다: {type(e).__name__}")
            return False

    def select_course_manually(self, edu_id, info):
        """사용자가 교육실시 관리 목록에서 교육을 체크하면 '교육생실시관리'를 눌러 팝업을 연다.
        Enter를 기다리는 시간은 제한하지 않는다. 반환: 파일의 교육실시ID와 같은 교육의 팝업이 열렸으면 True, 건너뛰면 False."""
        print("\n" + "-" * 60)
        print("  [교육 선택] 크롬의 '교육실시 관리' 목록에서 아래 교육을 1개만 체크해 주세요.")
        print(f"    교육실시ID : {edu_id}")
        for title, key in (("일자", "date"), ("과정명", "course"), ("기관명", "org")):
            if info.get(key):
                print(f"    {title:<9}: {info[key]}")
        print("  (체크하고 이 창으로 돌아와 Enter를 누르면, '교육생실시관리' 클릭부터 자동으로 진행합니다.)")
        print("-" * 60)
        while True:
            answer = input("  교육을 체크했으면 Enter (이 교육을 건너뛰려면 s + Enter): ").strip().lower()
            if answer == "s":
                return False
            if not self._layer_open():
                n = self._checked_count()
                if n != 1:
                    print(f"  ! 체크된 교육이 {n}개입니다. 교육실시 관리 목록에서 1개만 체크해 주세요.")
                    continue
                if not self._open_layer_from_checked():
                    print("  ! '교육생실시관리' 팝업이 열리지 않았습니다. 화면을 확인하고 다시 Enter를 눌러주세요.")
                    continue
            if self._layer_matches(edu_id):
                return True
            print("  ! 열린 팝업의 교육이 파일의 교육실시ID와 다릅니다. 팝업을 닫고 올바른 교육을 체크한 뒤 다시 Enter를 눌러주세요.")
            self.close_stdnt_layer()

    def open_stdnt_layer(self, edu_id, info=None, first=True):
        """edu_id 교육의 '교육생관리(실시)' 팝업을 연다 (기본은 사용자가 교육을 체크). 열리면 이미 등록된 교육생도 읽어 둔다."""
        label = "교육생실시관리"
        self.edu_id = None
        self._pending_edu = edu_id
        if info is not None:
            self._info = info
        info = self._info
        opened = self._layer_matches(edu_id) or self._open_course_automatically(edu_id, first)
        if not opened:
            if self._layer_open() and not self._layer_matches(edu_id):
                self.close_stdnt_layer()  # 다른 교육의 팝업이 열려 있으면 닫고 다시 고르게 한다
            opened = self.select_course_manually(edu_id, info)
        if opened and self._layer_matches(edu_id):
            self.edu_id = edu_id
            log.info(f"  [{label}] 팝업 열림: {edu_id}")
            self.registered_ids, self.registered_uids = self._read_registered()
            return True
        return False

    def ensure_layer(self, edu_id):
        """팝업이 닫혔거나 다른 교육이면 다시 연다 (체크된 교육이 1개면 자동, 아니면 사용자에게 안내)."""
        if self.edu_id == edu_id and self._layer_matches(edu_id):
            return True
        self._close_member_popup_if_open()
        return self.open_stdnt_layer(edu_id, first=False)

    # ---- 이미 등록된 교육생 목록 읽기 ---------------------------------
    def _ensure_layer_page_size(self, size=100):
        el = self._visible(By.ID, "pageSize")
        if el is None:
            return
        try:
            if Select(el).first_selected_option.get_attribute("value") == str(size):
                return
            Select(el).select_by_value(str(size))
            self.driver.execute_script("arguments[0].dispatchEvent(new Event('change', {bubbles:true}));", el)
            time.sleep(1.5)
        except (NoSuchElementException, StaleElementReferenceException):
            pass

    def _read_list_rows(self):
        return self.driver.execute_script(self._LIST_JS) or []

    def _read_registered(self, quiet=False):
        """팝업 목록의 모든 페이지에서 이미 등록된 교육생의 (회원ID 집합, 회원 고유번호 집합)을 읽는다.
        quiet=True 는 저장 뒤 반복해서 확인할 때 쓰며, '이미 등록된 교육생 N명 확인' 안내를 화면에 내지 않는다."""
        ids, uids = set(), set()
        try:
            self._ensure_layer_page_size()
            # 1페이지에서 시작
            cur = self._current_page("#pop_paging")
            if cur is not None and cur != 1:
                first = self._find_all(By.CSS_SELECTOR, "#pop_paging .btnFirst")
                if first:
                    self._js_click(first[0])
                    time.sleep(1.2)
            for _ in range(50):
                for mid, uid in self._read_list_rows():
                    if mid:
                        ids.add(mid.lower())
                    if uid:
                        uids.add(uid)
                if not self._goto_next_page("#pop_paging"):
                    break
                time.sleep(0.3)
        except (WebDriverException, TimeoutException) as e:
            log.warning(f"  [교육생목록] 등록된 교육생 목록을 읽지 못했습니다(중복 확인이 불완전할 수 있음): {type(e).__name__}")
        (log.debug if quiet else log.info)(f"  이미 등록된 교육생 {len(uids or ids)}명 확인")
        return ids, uids

    def _is_registered(self, mem_id, uid):
        return bool((uid and uid in self.registered_uids) or (mem_id and mem_id.lower() in self.registered_ids))

    # ---- 회원검색 ---------------------------------------------------
    def _close_member_popup_if_open(self):
        """회원검색 팝업이 화면에 남아 있으면 닫는다."""
        inp = self._visible(By.ID, "sch_mem_id")
        if inp is None:
            return
        try:
            layer = self.driver.execute_script("return arguments[0].closest('.layerPop');", inp)
            if layer is None:
                return
            for btn in layer.find_elements(By.XPATH, ".//*[contains(text(),'닫기')]"):
                if btn.is_displayed():
                    self._js_click(btn)
                    time.sleep(0.4)
                    return
        except WebDriverException:
            pass

    def _search_member(self, box, kind, key, name):
        """회원검색 팝업에서 한 번 검색한다. kind: '아이디' / '휴대폰' / '이름'.
        Total 1건이고 정보가 파일과 일치하면 그 행의 '선택' 버튼을 돌려준다. 반환: (선택 버튼 또는 None, 문제 설명)"""
        id_input = box.find_element(By.ID, "sch_mem_id")
        name_input = box.find_element(By.ID, "sch_mem_nm")
        id_input.clear()
        name_input.clear()
        # 휴대폰은 비회원 아이디 형식(GUEST_휴대폰번호, 예: GUEST_01022279306)으로 아이디 칸에 입력한다. 숫자만 넣으면 검색되지 않는다.
        (name_input if kind == "이름" else id_input).send_keys(f"GUEST_{key}" if kind == "휴대폰" else key)

        # 이전 검색 결과가 남아 있어 새 결과로 착각하지 않도록 비운다.
        self.driver.execute_script(
            "arguments[0].querySelector('#totalCnt_sch').textContent = '';"
            "arguments[0].querySelector('#mainTbody').innerHTML = '';", box)
        search_btn = None
        for b in box.find_elements(By.CSS_SELECTOR, "button[onclick*='fn_sch_mem(']"):
            if "검색하기" in b.text:
                search_btn = b
                break
        if search_btn is None:
            raise NoSuchElementException("회원검색 팝업의 '검색하기' 버튼")
        self._js_click(search_btn)

        total_el = box.find_element(By.ID, "totalCnt_sch")
        WebDriverWait(self.driver, config.WAIT_TIME).until(lambda d: total_el.text.strip() != "")
        total = int(re.sub(r"\D", "", total_el.text) or 0)

        cands = []
        for tr in box.find_element(By.ID, "mainTbody").find_elements(By.TAG_NAME, "tr"):
            btns = tr.find_elements(By.CSS_SELECTOR, "button[onclick*='chooseUid']")
            cells = [td.text.strip() for td in tr.find_elements(By.TAG_NAME, "td")]
            if btns and len(cells) >= 3:
                cands.append((cells, btns[0]))

        def row_ok(cells):  # cells: [번호, 아이디, 이름, 출생년도, 주소, 선택]
            if kind == "아이디" and cells[1].lower() != key.lower():
                return False
            if kind == "휴대폰" and re.sub(r"\D", "", cells[1]) != key:  # 예: GUEST_01082237520
                return False
            if name and cells[2] != name:
                return False
            return True

        if total == 1 and len(cands) == 1:
            if row_ok(cands[0][0]):
                return cands[0][1], ""
            return None, "검색된 회원의 아이디/이름이 파일의 내용과 다릅니다"
        if total == 0:
            return None, "검색결과가 0건입니다"
        if kind != "이름":
            # 아이디/휴대폰 번호는 회원마다 유일하므로, 정확히 같은 행이 딱 하나면 임의 선택이 아니다.
            exact = [btn for cells, btn in cands if row_ok(cells)]
            if len(exact) == 1:
                log.info(f"  [회원검색] {kind} 검색결과 {total}건 중 정확히 같은 1건을 선택합니다.")
                return exact[0], ""
        return None, f"검색결과가 {total}건이라 자동 선택하지 않았습니다"

    def select_member(self, student):
        """
        '회원검색' 팝업: 회원ID로 검색 -> (1명으로 못 찾으면) 휴대폰 번호(숫자만)로 검색 -> (둘 다 없을 때만) 이름으로 검색.
        Total 1건이고 아이디/이름이 파일과 일치하면 '선택'을 자동으로 누른다.
        모든 검색으로도 못 찾으면 임의로 고르지 않고 수동 안내(30초 후 다음 작업)로 넘긴다.
        반환: (회원ID 칸이 채워졌는지, 선택한 회원 고유번호 또는 None)
        """
        label = "회원검색"
        mem_id, phone, name = student["mem_id"], student["phone"], student["name"]
        attempts = []
        if mem_id:
            attempts.append(("아이디", mem_id))
        if phone:
            attempts.append(("휴대폰", phone))
        if not attempts:
            attempts.append(("이름", name))
        who = mask_id(mem_id)
        problems = []
        try:
            self._click(By.ID, "mbrSchBtn")
            id_input = self._wait_visible(By.ID, "sch_mem_id")
            box = self.driver.execute_script("return arguments[0].closest('.cont');", id_input)
            for kind, key in attempts:
                target, problem = self._search_member(box, kind, key, name)
                if target is not None:
                    m = re.search(r"chooseUid\('(\d+)'\)", target.get_attribute("onclick") or "")
                    uid = m.group(1) if m else None
                    self._js_click(target)
                    try:
                        WebDriverWait(self.driver, config.WAIT_TIME).until(lambda d: self._value("#mem_id") != "")
                    except TimeoutException:
                        pass
                    if self._value("#mem_id"):
                        self._close_member_popup_if_open()
                        log.info(f"  [{label}] 자동 선택 완료 ({kind} 검색): {who}")
                        return True, uid
                    problems.append(f"{kind} 검색: '선택'을 눌렀지만 회원ID 칸이 채워지지 않았습니다")
                    break
                problems.append(f"{kind} 검색: {problem}")
                if kind == "아이디" and phone:
                    log.info(f"  [{label}] 아이디로 1명을 찾지 못해 휴대폰 번호로 다시 검색합니다.")
        except (NoSuchElementException, TimeoutException, StaleElementReferenceException) as e:
            problems.append(f"자동 검색 실패 ({type(e).__name__}: {(str(e).strip().splitlines() or [''])[0]})")

        detail = " / ".join(problems)
        log.warning(f"  [{label}] {detail}")
        keyword = mem_id or name or f"휴대폰 끝 4자리 {phone[-4:]}"
        input(
            f"\n  >> [{label}] {detail}. 브라우저에서 '{keyword}' 회원을 직접 검색/선택한 뒤 "
            f"Enter를 눌러주세요..."
        )
        self._close_member_popup_if_open()
        return bool(self._value("#mem_id")), None

    # ---- 저장 -------------------------------------------------------
    def _find_confirm_button(self):
        try:
            return self.driver.execute_script(self._CONFIRM_BTN_JS)
        except WebDriverException:
            return None

    def _handle_save_dialogs(self):
        """
        '저장' 클릭 후의 '입력한 자료로 교육생을 등록하시겠습니까?' 확인창(alert 또는 레이어 팝업)에서
        '확인'을 누르고, 이어지는 안내창을 읽는다. 반환: (확인창을 처리했는지, 안내 메시지 목록)
        """
        messages = []
        confirmed = False
        for _ in range(6):
            text = self._alert_text(2.0 if not confirmed else 1.0)
            if text is not None:
                messages.append(text)
                if "등록하시겠습니까" in text:
                    confirmed = True
                continue
            btn = self._find_confirm_button()
            if btn is not None:
                messages.append("(확인창)")
                self._js_click(btn)
                confirmed = True
                time.sleep(0.6)
                continue
            break
        return confirmed, messages

    def _wait_until_registered(self, mem_id, uid, seconds=8):
        """저장 후 목록이 갱신되어 이 교육생이 보일 때까지 기다린다(목록을 다시 읽어 확인)."""
        end = time.time() + seconds
        while True:
            time.sleep(1.2)
            self.registered_ids, self.registered_uids = self._read_registered(quiet=True)
            if self._is_registered(mem_id, uid):
                return True
            if time.time() > end:
                return False

    # ---- 거주지역 (회원이 가입 때 입력하지 않았으면 사이트가 저장을 막는다) ----------
    _RESIDENCE = (("시/도", "#resdnc_area_cd"), ("시군구", "#resdnc_signgu_cd"))

    def _missing_residence(self, wait=2.5):
        """회원을 선택한 뒤 거주지역(시/도, 시군구)이 채워졌는지 확인해 비어 있는 항목 이름 목록을 돌려준다.
        회원 정보가 화면에 채워지는 데 시간이 걸릴 수 있어 wait(초) 동안은 채워지길 기다린다.
        화면에 이 칸이 없으면(사이트 개편 등) 확인할 수 없으므로 빈 목록을 돌려준다."""
        end = time.time() + wait
        while True:
            missing = []
            for label, css in self._RESIDENCE:
                if not self._find_all(By.CSS_SELECTOR, css):
                    return []
                if not self._value(css):
                    missing.append(label)
            if not missing or time.time() >= end:
                return missing
            time.sleep(0.3)

    # ---- 직접 처리 후 화면 확인 -----------------------------------------
    def field_filled(self, key):
        """
        수동 안내 항목(key)이 화면에서 실제로 처리되었는지 확인한다.
        처리되었으면 True, 아니면 False, 확인할 수 없으면 None (확인할 수 없으면 해결로 보지 않는다).
        """
        try:
            if key == "회원검색":
                return bool(self._value("#mem_id"))
            if key == "교육생실시관리":
                return self._layer_matches(getattr(self, "_pending_edu", ""))
            if key == "교육보유기자재":
                return all(self._find(By.ID, config.EQUIPMENT_IDS[n], 2).is_selected()
                           for n in config.STUDENT_DEFAULT_VALUES["교육보유기자재"])
            if key == "양방향영상교육가능":
                return self._find(By.ID, config.BIDIRECTIONAL_IDS[config.STUDENT_DEFAULT_VALUES["양방향영상교육가능"]], 2).is_selected()
            if key == "거주지역":
                return not self._missing_residence(wait=0)
            if key == "저장":
                self.registered_ids, self.registered_uids = self._read_registered()
                return self._is_registered(self.cur_mem_id, self.cur_uid)
        except Exception:
            return None
        return None

    def recover_after_error(self):
        """오류가 난 뒤 다음 교육생을 처리할 수 있게 화면을 정리한다: 떠 있는 안내창을 닫고(그 내용을 돌려줌) 회원검색 팝업을 닫는다."""
        text = None
        try:
            text = self._alert_text(0.5)
        except Exception:
            pass
        try:
            self._close_member_popup_if_open()
        except Exception:
            pass
        return text

    # ---- 교육생 한 명 등록 ---------------------------------------------
    def register_student(self, student):
        """반환: (결과, 비고, 오류유형)  결과: 등록완료 / 이미등록 / 실패 / 건너뜀 / 시험(저장안함). 오류가 아니면 오류유형은 빈 글자"""
        edu_id, mem_id, name = student["edu_id"], student["mem_id"], student["name"]
        who = mask_id(mem_id)
        self.cur_mem_id, self.cur_uid = mem_id, None

        if not self.ensure_layer(edu_id):
            log.error(f"  [교육생실시관리] {edu_id} 팝업을 열지 못해 이 교육생을 등록하지 못했습니다.")
            return "실패", "교육생실시관리 팝업을 열지 못함", KIND_COURSE

        # 이미 등록된 회원ID로 시작한 경우에는 회원검색 전에 바로 알 수 있다.
        if mem_id and self._is_registered(mem_id, None):
            log.info(f"  이미 등록된 교육생입니다: {who} -> 건너뜀")
            return "이미등록", "교육생 목록에 이미 있음", ""

        # 1) '교육생등록' 버튼
        try:
            self._click(By.ID, "btn_stdnt_save")
        except TimeoutException:
            log.error("  [교육생등록] '교육생등록' 버튼을 누를 수 없습니다 (수정 불가 상태일 수 있음).")
            return "실패", "교육생등록 버튼을 누를 수 없음", KIND_SAVE
        alert = self._alert_text(0.8)
        if alert:
            log.error(f"  [교육생등록] 안내창: {alert}")
            return "실패", f"교육생등록 안내창: {alert}", KIND_SAVE
        self._wait_visible(By.ID, "mbrSchBtn")

        # 2) 회원검색 -> 선택
        ok, uid = self.select_member(student)
        if not ok:
            log.error(f"  [회원검색] 회원이 선택되지 않아 이 교육생은 등록하지 않고 넘어갑니다: {who}")
            return "건너뜀", "회원 미선택(0건/2건 이상/시간초과)", KIND_MEMBER
        selected_id = self._value("#mem_id")
        self.cur_mem_id, self.cur_uid = selected_id or mem_id, uid

        if self._is_registered(selected_id, uid):
            log.info(f"  이미 등록된 교육생입니다: {mask_id(selected_id)} -> 건너뜀")
            return "이미등록", "교육생 목록에 이미 있음", ""

        # 3) 교육보유기자재 / 양방향영상교육가능
        for item in config.STUDENT_DEFAULT_VALUES["교육보유기자재"]:
            try:
                self._check_input(config.EQUIPMENT_IDS[item])
            except (TimeoutException, NoSuchElementException):
                pass
        try:
            self._check_input(config.BIDIRECTIONAL_IDS[config.STUDENT_DEFAULT_VALUES["양방향영상교육가능"]])
        except (TimeoutException, NoSuchElementException):
            pass
        not_filled = []
        for key in ("교육보유기자재", "양방향영상교육가능"):
            if self.field_filled(key) is not True:
                log.warning(f"  [{key}] 자동 입력이 확인되지 않습니다.")
                if _wait_for_manual():
                    input(f"\n  >> [{key}] 자동 입력 실패. 브라우저에서 직접 체크한 뒤 Enter를 눌러주세요...")
                else:
                    not_filled.append(key)
        if not_filled:
            log.error(f"  [{', '.join(not_filled)}] 자동 입력에 실패해 이 교육생은 저장하지 않고 넘어갑니다. (마지막 '오류 유형별 정리'에서 확인하세요): {mask_id(selected_id)}")
            return "실패", f"자동 입력 실패: {', '.join(not_filled)}", KIND_INPUT

        # 4) 거주지역: 회원이 가입 때 입력하지 않았으면 사이트가 저장을 막는다. 저장을 누르기 전에 미리 알아내
        #    (기본) 기다리지 않고 건너뛰어 마지막 '오류 유형별 정리'에 모아 알려 주고, ERROR_ACTION="wait" 이면 직접 선택하도록 기다린다.
        missing = self._missing_residence()
        if missing:
            names = ", ".join(missing)
            problem = f"  [거주지역] 거주지역 선택이 비어 있습니다 ({names}): {mask_id(selected_id)}"
            skipping = f"     이 교육생은 저장하지 않고 넘어갑니다. (마지막 '오류 유형별 정리'에서 확인하세요): {mask_id(selected_id)}"
            if _wait_for_manual():
                input(f"\n  >> [거주지역] 거주지역 선택이 비어 있습니다 ({names}). 브라우저에서 직접 선택한 뒤 Enter를 눌러주세요...")
                missing = self._missing_residence(wait=0.5)
            if missing:
                # 오류가 두 건처럼 보이지 않도록 한 번에 기록한다 (둘째 줄은 이어지는 설명)
                log.error(problem + "\n" + skipping)
                return "건너뜀", f"거주지역({names}) 미선택", KIND_RESIDENCE

        # 5) 저장 (설정에 따라 시험 모드 / 교육생마다 사용자 확인)
        if getattr(config, "DRY_RUN", False):
            log.info(f"  [시험 모드] 저장 직전까지 입력했습니다. 저장하지 않고 넘어갑니다: {mask_id(selected_id)}")
            return "시험(저장안함)", "DRY_RUN", ""
        if not getattr(config, "AUTO_SUBMIT", True):
            print(f"\n  >>> [{edu_id} / {mask_id(selected_id)}] 입력 완료. 화면을 확인해주세요.")
            answer = input("      저장하시려면 Enter, 이 교육생을 건너뛰려면 s + Enter: ").strip().lower()
            if answer == "s":
                log.info(f"[건너뜀] {edu_id} / {mask_id(selected_id)}")
                return "건너뜀", "사용자가 건너뜀", ""

        try:
            self._click(By.CSS_SELECTOR, "#btnBox button[onclick*='fn_stdnt_save_btn']")
        except TimeoutException:
            log.error("  [저장] '저장' 버튼을 찾지 못했습니다.")
            if not _wait_for_manual():
                return "실패", "저장 버튼을 찾지 못함", KIND_SAVE
            input("  >> [저장] 저장 버튼을 직접 눌러 등록을 마친 뒤 Enter를 눌러주세요...")
        else:
            confirmed, messages = self._handle_save_dialogs()
            site_errors = [m for m in messages if any(w in m for w in _SAVE_FAIL_WORDS)]
            if site_errors:
                # 사이트가 필수 항목 누락 등을 알려 준 경우: 목록을 반복해서 읽으며 기다리지 않고 바로 알린다.
                reason = " / ".join(site_errors)
                log.error(f"  [저장] 사이트 안내: {reason}")
                if _wait_for_manual():
                    input(f"\n  >> [저장] 사이트 안내 '{reason}' 를 확인하고, 브라우저에서 직접 처리해 저장한 뒤 Enter를 눌러주세요...")
                    self.registered_ids, self.registered_uids = self._read_registered()
                    if self._is_registered(selected_id, uid):
                        return "등록완료", "직접 처리 후 목록에서 확인됨", ""
                else:
                    log.error(f"  [저장] 이 교육생은 저장하지 못하고 넘어갑니다. (마지막 '오류 유형별 정리'에서 확인하세요): {mask_id(selected_id)}")
                return "실패", f"저장 안내창: {reason}", KIND_SAVE
            if not confirmed:
                log.warning(f"  [저장] 등록 확인창을 처리하지 못했습니다. 안내창: {messages}")

        if self._wait_until_registered(selected_id, uid):
            log.info(f"[등록완료] {edu_id} / {mask_id(selected_id)}")
            return "등록완료", "", ""

        log.error(f"[등록실패] {edu_id} / {mask_id(selected_id)} - 저장 후 교육생 목록에서 확인되지 않습니다.")
        if _wait_for_manual():
            input("  >> [저장] 화면의 안내 메시지를 확인하고 필요하면 직접 저장한 뒤 Enter를 눌러주세요...")
            self.registered_ids, self.registered_uids = self._read_registered()
            if self._is_registered(selected_id, uid):
                return "등록완료", "직접 처리 후 목록에서 확인됨", ""
        else:
            log.error(f"  [저장] 이 교육생은 저장이 확인되지 않아 넘어갑니다. (마지막 '오류 유형별 정리'에서 확인하세요): {mask_id(selected_id)}")
        return "실패", "저장 후 교육생 목록에서 확인되지 않음", KIND_SAVE

def run(loaded):
    """확인이 끝난 교육생들을 크롬으로 등록한다."""
    targets = loaded.rows
    # 형식 문제로 제외된 줄은 마지막 '수동 확인' 목록에 함께 정리한다.
    for where, edu_id, who, problems in loaded.invalid:
        issues.issues.append(((where, edu_id, who), "형식오류", "; ".join(problems) + " -> 이 줄은 자동 등록에서 제외되었으니 직접 등록하세요."))
    log.info(f"총 {len(targets)}명 등록 예정")
    if getattr(config, "DRY_RUN", False):
        log.info("DRY_RUN 모드: 저장 직전까지만 진행하고 실제로 저장하지 않습니다.")

    groups = group_by_education(targets)
    for i, s in enumerate(targets, 1):
        s["no"] = i

    registrar = start_registrar(StudentRegistrar)
    issues.verify = registrar.field_filled
    registrar.login_and_wait()

    def key_of(s):
        return (s["no"], s["edu_id"], mask_phone(s["mem_id"] or s["name"] or s["phone"]))

    done = 0
    stop = False
    consecutive_errors = 0  # 연속으로 예외가 난 횟수 (브라우저가 닫히는 등 계속 실패할 때 멈추기 위함)
    try:
        for edu_id, students in groups.items():
            first = students[0]
            issues.current = key_of(first)
            issues.mark = len(issues.issues)
            log.info(f"\n===== 교육실시ID {edu_id} ({len(students)}명) =====")
            if not registrar.open_stdnt_layer(edu_id, first["ref"]):
                for s in students:  # 교육 팝업을 열지 못했거나 건너뛰었으면 이 교육의 교육생은 모두 미처리로 정리
                    issues.current = key_of(s)
                    log.warning(f"  [교육생실시관리] {edu_id} 교육을 선택하지 않았거나 팝업을 열지 못해 등록하지 못했습니다.")
                    record_result(s["no"], edu_id, s, "건너뜀", "교육 선택 건너뜀 또는 팝업을 열지 못함", KIND_COURSE)
                continue

            for s in students:
                issues.current = key_of(s)
                issues.mark = len(issues.issues)
                log.info(f"\n[{s['no']}/{len(targets)}] {edu_id} / {mask_id(s['mem_id'])} 처리 중...")
                try:
                    status, note, kind = registrar.register_student(s)
                except Exception as e:
                    reason = f"{type(e).__name__}: {(str(e).strip().splitlines() or [''])[0][:120]}"
                    log.error(f"  처리 중 오류 발생: {reason}")
                    alert = registrar.recover_after_error()
                    record_result(s["no"], edu_id, s, "실패", f"처리 중 오류: {reason}" + (f" (사이트 안내: {alert})" if alert else ""), KIND_ERROR)
                    consecutive_errors += 1
                    if _wait_for_manual():
                        cont = input("  오류가 발생했습니다. 다음 교육생으로 계속할까요? (Enter=계속, x=중단): ").strip().lower()
                        if cont == "x":
                            stop = True
                            break
                    elif consecutive_errors >= 3:
                        log.error("  같은 종류의 처리 중 오류가 3번 연속 발생해 등록을 중단합니다. 크롬 창과 로그인 상태를 확인하세요.")
                        stop = True
                        break
                else:
                    consecutive_errors = 0
                    record_result(s["no"], edu_id, s, status, note, kind)
                    if status == "등록완료":
                        done += 1
                time.sleep(config.DELAY_BETWEEN_ACTIONS)
            registrar.close_stdnt_layer()  # 다음 교육을 고를 수 있게 팝업을 닫고 체크를 푼다
            if stop:
                break
    except KeyboardInterrupt:
        print("\n사용자가 중단했습니다. 지금까지의 결과를 정리합니다.")

    log.info(f"\n===== 완료: {done}명 등록 / 총 {len(targets)}명 대상 =====")
    print_report(
        total=len(targets), done=done, unit="명",
        group_title="수동으로 확인/처리가 필요한 교육생",
        key_names=("번호", "교육실시ID", "회원"),
        results=results,
    )
    print_kind_summary(loaded.invalid)
