# -*- coding: utf-8 -*-
"""
작업 1. 교육계획 등록 (강사페이지 > 교육계획 관리 > 교육계획 등록)

엑셀/CSV의 교육과정을 한 줄씩 읽어 교육계획 등록 화면을 채우고 '등록'까지 진행한다.
자동 처리에 실패한 단계는 터미널 안내에 따라 브라우저에서 직접 마무리하고 Enter를 누르면 이어서 진행되며,
모든 등록이 끝나면 문제가 있던 행만 모아 출력/저장한다(결과 폴더).

주의: 아이디/비밀번호를 다루지 않습니다. 로그인은 항상 사용자가 직접 합니다.
"""

import re
import time
from datetime import datetime

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementClickInterceptedException,
)

import config
from common.browser import BaseRegistrar, start_registrar
from common.console import input, issues, log
from common.files import Loaded, clean_row, read_table, txt as _txt
from common.report import print_load_summary, print_report

TITLE = "교육계획 등록"
FILE_LABEL = "교육과정 파일"


# ---------------------------------------------------------- 엑셀 읽기 ----
def normalize_date(value) -> str:
    """엑셀 '일자' 값을 'YYYY-MM-DD' 문자열로 정규화"""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    s = str(value).strip().replace(" ", "")
    parts = [p for p in re.split(r"[.\-/]", s) if p]
    y, m, d = parts[0], parts[1], parts[2]
    return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"


def normalize_time(value) -> str:
    """엑셀 '시작'/'종료' 값을 (시, 분) 튜플로 정규화"""
    if hasattr(value, "hour") and hasattr(value, "minute"):
        # datetime.datetime 또는 datetime.time 모두 처리
        return value.hour, value.minute
    s = str(value).strip()
    h, m = s.split(":")[:2]
    return int(h), int(m)


def _read_file(path, loaded):
    """엑셀/CSV 파일 하나에서 교육과정 행을 읽어 loaded 에 더한다. 헤더는 1행, 컬럼명은 config.COL_* 기준."""
    raw_rows = read_table(path, [config.COL_COURSE, config.COL_DATE])

    if not raw_rows:
        raise SystemExit(f"{path.name}: 데이터가 없습니다. 1행은 제목, 2행부터 교육과정을 적어주세요.")

    # '보조강사' 열이 없으면 열 이름에 '보조'가 들어 있는 열(예: 보조1)을 보조강사 열로 읽는다.
    keyword = getattr(config, "COL_ASSISTANT_KEYWORD", "")
    existing = [str(k).strip() for k in raw_rows[0].keys() if k is not None]
    if keyword and config.COL_ASSISTANT not in existing:
        found = [h for h in existing if keyword in h]
        if found:
            if len(found) > 1:
                loaded.notes.append(f"{path.name}: '{keyword}'가 들어간 열이 여러 개입니다 {found} -> 맨 앞의 '{found[0]}' 열을 보조강사로 읽습니다.")
            raw_rows = [{(config.COL_ASSISTANT if str(k).strip() == found[0] else k): v for k, v in raw.items()} for raw in raw_rows]

    # 필요한 열(제목)이 있는지 먼저 확인 (제목이 다르면 어떤 열이 없는지 알려준다)
    headers = {str(k).strip() for k in raw_rows[0].keys() if k is not None}
    required = [config.COL_COURSE, config.COL_DATE, config.COL_START, config.COL_END,
                config.COL_ADDRESS, config.COL_ADDRESS_DETAIL, config.COL_SIDO, config.COL_SIGUNGU,
                config.COL_DONG, config.COL_ASSISTANT]
    missing = [c for c in required if c not in headers]
    if missing:
        raise SystemExit(
            f"{path.name}: 필요한 열이 없습니다 -> {missing}\n"
            f"  파일의 1행(제목) 이름이 README의 표와 똑같은지 확인하세요. 현재 제목: {sorted(headers)}"
        )

    total = bad = 0
    for n, raw in enumerate(raw_rows, start=2):  # 엑셀 행 번호 (1행은 제목)
        row = {k: (v.strip() if isinstance(v, str) else v) for k, v in clean_row(raw).items()}
        if not _txt(row.get(config.COL_COURSE)) and not _txt(row.get(config.COL_DATE)):
            continue  # 빈 줄
        total += 1

        problems = []
        try:
            normalize_date(row.get(config.COL_DATE))
        except Exception:
            problems.append(f"'{config.COL_DATE}'을(를) 읽을 수 없습니다 (값: {_txt(row.get(config.COL_DATE))!r}, 예: 2026. 11. 03)")
        for col in (config.COL_START, config.COL_END):
            try:
                normalize_time(row.get(col))
            except Exception:
                problems.append(f"'{col}'을(를) 읽을 수 없습니다 (값: {_txt(row.get(col))!r}, 예: 14:00)")
        for col in (config.COL_COURSE, config.COL_ADDRESS, config.COL_ADDRESS_DETAIL, config.COL_SIDO, config.COL_SIGUNGU, config.COL_DONG):
            if not _txt(row.get(col)):
                problems.append(f"'{col}'이(가) 비어 있습니다")

        if problems:
            bad += 1
            where = f"{path.name} {n}행"
            loaded.invalid.append((where, _txt(row.get(config.COL_COURSE)) or "(과정명 없음)", _txt(row.get(config.COL_DATE)), problems))
            continue
        loaded.rows.append(row)
    loaded.per_file.append((path.name, total, bad))


def load(paths):
    """파일들을 읽어 등록할 수 있는 줄과 오류가 있는 줄을 나눈다. (파일 자체를 읽을 수 없으면 SystemExit)"""
    loaded = Loaded()
    for path in paths:
        _read_file(path, loaded)

    # 같은 과정·일자·시간이 두 번 나오면 실수로 중복 입력했을 수 있으므로 알려준다. (등록은 막지 않음)
    seen = {}
    for row in loaded.rows:
        key = (_txt(row.get(config.COL_COURSE)), normalize_date(row[config.COL_DATE]),
               normalize_time(row[config.COL_START]), normalize_time(row[config.COL_END]))
        seen[key] = seen.get(key, 0) + 1
    for (course, date, _, _), count in seen.items():
        if count > 1:
            loaded.notes.append(f"같은 과정·일자·시간이 {count}번 들어 있습니다: {date} / {course} (중복 입력이 아닌지 확인하세요)")

    limit = getattr(config, "TEST_LIMIT", None)
    if limit and len(loaded.rows) > limit:
        loaded.notes.append(f"TEST_LIMIT={limit} 설정으로 앞의 {limit}건만 등록합니다. (전체를 등록하려면 config.py 에서 TEST_LIMIT = None)")
        loaded.rows = loaded.rows[:limit]
    return loaded


def show_summary(loaded):
    """등록 전 확인 화면: 등록 대상 목록과 오류가 있는 줄을 보여준다."""
    lines = []
    for n, row in enumerate(loaded.rows, 1):
        sh, sm = normalize_time(row[config.COL_START])
        eh, em = normalize_time(row[config.COL_END])
        assistant = _txt(row.get(config.COL_ASSISTANT))
        lines.append(
            f"{n:>2}. {normalize_date(row[config.COL_DATE])} {sh:02d}:{sm:02d}~{eh:02d}:{em:02d}"
            f" / {_txt(row[config.COL_COURSE])}" + (f" / 보조 {assistant}" if assistant else "")
        )
    print_load_summary(TITLE, "건", loaded, lines)


# --------------------------------------------------------------- 본체 ----
class EduPlanRegistrar(BaseRegistrar):
    def __init__(self):
        super().__init__()
        self.place_name = config.PLAN_DEFAULT_VALUES["담당배움터"]

    # ---- 공통 유틸 -------------------------------------------------
    def _set_text(self, by, sel, value):
        el = self._find(by, sel)
        el.clear()
        el.send_keys(value)
        return el

    def _set_value_js(self, by, sel, date_str):
        """readonly input에 JS로 직접 값 주입 + input/change/blur 이벤트 발생 (날짜, 상세주소 등 공용)"""
        el = self._find(by, sel)
        self.driver.execute_script(
            """
            var el = arguments[0];
            el.removeAttribute('readonly');
            el.value = arguments[1];
            el.dispatchEvent(new Event('input', {bubbles:true}));
            el.dispatchEvent(new Event('change', {bubbles:true}));
            el.dispatchEvent(new Event('blur', {bubbles:true}));
            """,
            el,
            date_str,
        )
        return el

    def _select_by_text(self, by, sel, text, fire_change=True):
        el = self._find(by, sel)
        Select(el).select_by_visible_text(text)
        if fire_change:
            self.driver.execute_script(
                "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));", el
            )
        return el

    def _select_number(self, by, sel, number):
        """
        시/분 드롭다운에서 숫자가 같은 옵션을 선택한다. 옵션 표기가 '00', '0', '00분', '14시' 등
        무엇이든 value/텍스트에서 숫자만 뽑아 비교한다.
        """
        el = self._find(by, sel)
        select = Select(el)
        for opt in select.options:
            for cand in (opt.get_attribute("value"), opt.text):
                digits = re.sub(r"\D", "", cand or "")
                if digits != "" and int(digits) == number:
                    select.select_by_index(select.options.index(opt))
                    self.driver.execute_script(
                        "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));", el
                    )
                    return el
        options = [f"{o.get_attribute('value')}|{o.text}" for o in select.options]
        raise NoSuchElementException(f"{sel} 에서 숫자 {number} 옵션을 찾지 못함. 옵션: {options}")

    # ---- 팝업(과정/배움터/보조강사 검색) 공통 처리 ------------------
    # showPopup()이 여는 팝업은 새 창/iframe이 아니라 같은 페이지 안의
    # <div class="layerPop {popup_class} on"> 레이어다. 팝업이 열려도 이전에
    # 열렸던 다른 팝업의 템플릿 DOM(id 중복)이 남아있을 수 있으므로,
    # 항상 현재 열린 팝업 요소(popup) 안에서만 find_element를 호출해 범위를 좁힌다.

    def _wait_popup_open(self, popup_class, timeout=None):
        w = WebDriverWait(self.driver, timeout or config.WAIT_TIME)
        return w.until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, f".layerPop.{popup_class}.on"))
        )

    def select_course(self, course_name):
        """
        교육과정명 검색 팝업(crsePop): 과정명 입력 -> '검색하기' 클릭 -> 결과 목록에서 '선택'.
        (실제 팝업 HTML 확인 완료: #sch_edc_crse_nm 입력, id='mainTbody' 결과, onclick="chooseCrse(N)")
        """
        label = "교육과정명"
        self._click(By.CSS_SELECTOR, "button[onclick*=\"showPopup('crsePop'\"]")
        try:
            popup = self._wait_popup_open("crsePop")
            time.sleep(0.3)
            inp = popup.find_element(By.ID, "sch_edc_crse_nm")
            inp.clear()
            inp.send_keys(course_name)

            search_btn = None
            for b in popup.find_elements(By.TAG_NAME, "button"):
                if "검색하기" in b.text:
                    search_btn = b
                    break
            if search_btn:
                try:
                    search_btn.click()
                except ElementClickInterceptedException:
                    self.driver.execute_script("arguments[0].click();", search_btn)
            time.sleep(1.2)

            tbody = popup.find_element(By.ID, "mainTbody")
            rows = tbody.find_elements(By.TAG_NAME, "tr")
            target_btn = self._pick_row_button(rows, course_name, "chooseCrse")
            if target_btn:
                self._js_click(target_btn)
                time.sleep(0.6)
                self._close_popup_if_open("crsePop")
                if self._find(By.ID, "edc_crse_nm").get_attribute("value"):
                    log.info(f"  [{label}] 자동 선택 완료: {course_name}")
                    return True
        except (NoSuchElementException, TimeoutException) as e:
            log.warning(f"  [{label}] 자동 검색 실패: {e}")

        input(
            f"\n  >> [{label}] 자동 선택 실패. 브라우저에서 '{course_name}' 를 직접 검색/선택한 뒤 "
            f"Enter를 눌러주세요..."
        )
        return False

    def select_place(self, place_name):
        """
        담당 배움터 검색 팝업(plcPop): 배움터명 입력 -> '검색하기' -> 결과에서 배움터명이 정확히 같은 행의 '선택'.
        (실제 팝업 HTML 확인: #sch_edc_place_nm 입력, id='mainTbody' 결과, onclick="choosePlc(...)")
        정확히 일치하는 행이 없으면 결과가 1건일 때만 그 행을 선택한다.
        """
        label = "담당 배움터"
        self._click(By.CSS_SELECTOR, "button[onclick*=\"showPopup('plcPop'\"]")
        try:
            popup = self._wait_popup_open("plcPop")
            time.sleep(0.3)
            inp = popup.find_element(By.ID, "sch_edc_place_nm")
            inp.clear()
            inp.send_keys(place_name)
            for b in popup.find_elements(By.TAG_NAME, "button"):
                if "검색하기" in b.text:
                    self._js_click(b)
                    break
            time.sleep(1.2)

            rows = popup.find_element(By.ID, "mainTbody").find_elements(By.TAG_NAME, "tr")
            candidates = [
                (r, r.find_elements(By.CSS_SELECTOR, "button[onclick*='choosePlc']")) for r in rows
            ]
            candidates = [(r, b[0]) for r, b in candidates if b]
            target = None
            for r, b in candidates:
                names = [td.text.strip() for td in r.find_elements(By.CSS_SELECTOR, "td.title")]
                if place_name in names:
                    target = b
                    break
            if target is None and len(candidates) == 1:
                target = candidates[0][1]
            if target is not None:
                self._js_click(target)
                time.sleep(0.6)
                self._close_popup_if_open("plcPop")
                if self._find(By.ID, "edc_place_nm").get_attribute("value"):
                    log.info(f"  [{label}] 자동 선택 완료: {place_name}")
                    return True
            else:
                log.warning(f"  [{label}] '{place_name}' 과(와) 일치하는 배움터를 찾지 못했습니다 (검색결과 {len(candidates)}건).")
        except (NoSuchElementException, TimeoutException) as e:
            log.warning(f"  [{label}] 자동 검색 실패: {e}")

        input(
            f"\n  >> [{label}] 자동 선택 실패. 브라우저에서 '{place_name}' 를 직접 검색/선택한 뒤 "
            f"Enter를 눌러주세요..."
        )
        return False

    def select_assistant(self, name):
        """
        보조강사/가이드 검색 팝업(sptsPop): 이름 입력(#sch_tcr_nm) -> '검색하기' -> 이름이 정확히 같은 행의 '선택'.
        팝업의 배움터 칸(#sch_edc_place_nm)은 앞서 선택한 담당 배움터가 자동 입력되어
        해당 배움터에 등록된 보조강사만 조회된다. 동명이인 등 2건 이상이면 임의로 고르지 않는다.
        """
        label = "보조강사/가이드명"
        self._click(By.ID, "sptsPopBtn")
        try:
            popup = self._wait_popup_open("sptsPop")
            time.sleep(0.3)
            inp = popup.find_element(By.ID, "sch_tcr_nm")
            inp.clear()
            inp.send_keys(name)
            for b in popup.find_elements(By.TAG_NAME, "button"):
                if "검색하기" in b.text:
                    self._js_click(b)
                    break
            time.sleep(1.2)

            rows = popup.find_element(By.ID, "mainTbody").find_elements(By.TAG_NAME, "tr")
            matches = []
            for r in rows:
                btns = r.find_elements(By.CSS_SELECTOR, "button[onclick*='chooseTcr']")
                cells = [td.text.strip() for td in r.find_elements(By.TAG_NAME, "td")]
                if btns and name in cells:
                    matches.append(btns[0])
            if len(matches) == 1:
                self._js_click(matches[0])
                time.sleep(0.6)
                self._close_popup_if_open("sptsPop")
                if self._find(By.ID, "spts_nm").get_attribute("value"):
                    log.info(f"  [{label}] 자동 선택 완료: {name}")
                    return True
            elif len(matches) == 0:
                log.warning(f"  [{label}] '{name}' 이(가) 담당 배움터의 보조강사 목록에 없습니다.")
            else:
                log.warning(f"  [{label}] '{name}' 과(와) 이름이 같은 보조강사가 {len(matches)}명이라 자동 선택하지 않았습니다.")
        except (NoSuchElementException, TimeoutException) as e:
            log.warning(f"  [{label}] 자동 검색 실패: {e}")

        input(
            f"\n  >> [{label}] 자동 선택 실패. 브라우저에서 '{name}' 를 직접 찾아 선택한 뒤 "
            f"Enter를 눌러주세요..."
        )
        return False

    def search_list_and_pick(self, open_by, open_sel, popup_class, keyword, label, verify_field_id=None):
        """
        담당 배움터(plcPop) / 보조강사(sptsPop) 검색 팝업 공통 처리.
        등록방법.txt 기준: 입력 없이 '검색' 버튼으로 팝업을 연 뒤, 결과 목록에서
        keyword가 포함된 행을 찾아 '선택'을 누르는 방식. 목록에 keyword가 안 보이면
        팝업 내 검색창(있는 경우)에 keyword를 입력하고 검색 버튼을 눌러본 뒤 다시 탐색한다.
        팝업의 정확한 HTML 구조를 모르므로 자동 실패 시 수동 처리로 전환한다.
        """
        self._click(open_by, open_sel)
        try:
            popup = self._wait_popup_open(popup_class)
            time.sleep(0.8)  # 목록이 ajax로 채워지는 경우 대비

            def find_and_click_row():
                rows = popup.find_elements(By.TAG_NAME, "tr")
                for row in rows:
                    if keyword in row.text:
                        btns = [
                            b for b in row.find_elements(By.TAG_NAME, "button")
                            if "선택" in b.text
                        ]
                        if not btns:
                            btns = [
                                a for a in row.find_elements(By.TAG_NAME, "a")
                                if "선택" in a.text
                            ]
                        if btns:
                            self._js_click(btns[0])
                            return True
                return False

            if find_and_click_row():
                time.sleep(0.5)
                self._close_popup_if_open(popup_class)
                if not verify_field_id or self._find(By.ID, verify_field_id).get_attribute("value"):
                    log.info(f"  [{label}] 자동 선택 완료: {keyword}")
                    return True

            # 목록에 바로 없으면 검색창을 찾아 입력 후 재시도
            search_input = None
            for inp in popup.find_elements(By.CSS_SELECTOR, "input[type='text']"):
                if inp.is_displayed():
                    search_input = inp
                    break
            if search_input:
                search_input.clear()
                search_input.send_keys(keyword)
                for b in popup.find_elements(By.TAG_NAME, "button"):
                    if "검색" in b.text:
                        self._js_click(b)
                        break
                time.sleep(1.2)
                if find_and_click_row():
                    time.sleep(0.5)
                    self._close_popup_if_open(popup_class)
                    if not verify_field_id or self._find(By.ID, verify_field_id).get_attribute("value"):
                        log.info(f"  [{label}] 자동 선택 완료: {keyword}")
                        return True
        except (NoSuchElementException, TimeoutException) as e:
            log.warning(f"  [{label}] 자동 검색 실패: {e}")

        input(
            f"\n  >> [{label}] 자동 선택 실패. 브라우저에서 '{keyword}' 를 직접 찾아 선택한 뒤 "
            f"Enter를 눌러주세요..."
        )
        return False

    def _pick_row_button(self, rows, keyword, choose_js_prefix):
        for row in rows:
            if keyword in row.text:
                btns = row.find_elements(By.CSS_SELECTOR, f"button[onclick*='{choose_js_prefix}']")
                if btns:
                    return btns[0]
        if rows:
            btns = rows[0].find_elements(By.CSS_SELECTOR, f"button[onclick*='{choose_js_prefix}']")
            if btns:
                return btns[0]
        return None

    def _close_popup_if_open(self, popup_class):
        """선택 후 팝업이 자동으로 안 닫히면 닫기 버튼을 눌러 화면을 가리지 않게 한다."""
        popups = self._find_all(By.CSS_SELECTOR, f".layerPop.{popup_class}.on")
        for p in popups:
            if p.is_displayed():
                for btn in p.find_elements(By.XPATH, ".//*[contains(text(),'닫기')]"):
                    if btn.is_displayed():
                        self._js_click(btn)
                        time.sleep(0.3)
                        return

    # ---- 실제교육장 주소검색 (juso.go.kr 별도 창) ---------------------
    def fill_address_via_popup(self, address, detail):
        """
        '주소검색' 버튼 -> 새 창에서 주소 입력 -> 돋보기 -> 결과 클릭 -> 상세주소 입력 -> '주소입력'.
        완료되면 새 창이 닫히고 부모 창의 우편번호/주소/상세주소 칸이 채워진다.
        """
        label = "실제교육장 주소"
        main = self.driver.current_window_handle
        before = set(self.driver.window_handles)
        w = lambda: WebDriverWait(self.driver, config.WAIT_TIME)
        try:
            self._click(By.ID, "find_addr")  # '직접입력' 선택 후에 보이는 버튼
            w().until(lambda d: len(d.window_handles) > len(before))
            popup = [h for h in self.driver.window_handles if h not in before][0]
            self.driver.switch_to.window(popup)

            keyword = w().until(EC.element_to_be_clickable((By.ID, "keyword")))
            keyword.clear()
            keyword.send_keys(address)
            self._click(By.CSS_SELECTOR, "input[type='button'][title='검색']")

            first = w().until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "#resultList td.subj a[href^='javascript:setMaping']")
            ))
            self._js_click(first)

            detail_input = w().until(EC.element_to_be_clickable((By.ID, "rtAddrDetail")))
            detail_input.clear()
            detail_input.send_keys(detail)
            self._click(By.CSS_SELECTOR, "a.btn-bl")  # '주소입력'

            w().until(lambda d: len(d.window_handles) == len(before))
            self.driver.switch_to.window(main)
            if self._find(By.ID, "real_edc_plc").get_attribute("value"):
                log.info(f"  [{label}] 자동 입력 완료: {address} {detail}")
                return True
            log.warning(f"  [{label}] 팝업은 닫혔지만 주소칸이 비어 있습니다.")
        except (TimeoutException, NoSuchElementException, IndexError) as e:
            log.warning(f"  [{label}] 자동 입력 실패: {e}")

        if main in self.driver.window_handles and self.driver.current_window_handle != main:
            input(
                f"\n  >> [{label}] 주소 검색 창에서 '{address}' 를 직접 검색/선택하고 상세주소를 입력한 뒤 "
                f"'주소입력'까지 누르고 Enter를 눌러주세요..."
            )
        else:
            input(f"\n  >> [{label}] 브라우저에서 주소를 직접 입력한 뒤 Enter를 눌러주세요...")
        if main in self.driver.window_handles:
            self.driver.switch_to.window(main)
        return False

    # ---- 직접 처리 후 칸이 실제로 채워졌는지 확인 ---------------------
    _FIELD_IDS = {
        "교육과정명": "edc_crse_nm",
        "담당 배움터": "edc_place_nm",
        "보조강사/가이드명": "spts_nm",
        "실제교육장 주소": "real_edc_plc",
    }

    _TIMETABLE_FIELDS = (
        ".scheduleDate",
        "select[name='edc_bgn_hour[]']", "select[name='edc_bgn_minute[]']",
        "select[name='edc_end_hour[]']", "select[name='edc_end_minute[]']",
        "input.toolsUsed[name='eqpmn_nm[]']",
    )

    def field_filled(self, key):
        """
        수동 안내 항목(key)의 칸이 화면에서 실제로 채워졌는지 확인한다.
        채워졌으면 True, 비어 있으면 False, 확인할 수 없으면 None (확인할 수 없으면 해결로 보지 않는다).
        """
        try:
            try:
                current = self.driver.current_window_handle
            except Exception:  # 주소 검색 창이 이미 닫힌 경우
                current = None
            self.driver.switch_to.window(self.driver.window_handles[0])  # 등록 화면(메인 창)

            def value(css):
                return (self.driver.find_element(By.CSS_SELECTOR, css).get_attribute("value") or "").strip()

            if key in self._FIELD_IDS:
                result = bool(value(f"#{self._FIELD_IDS[key]}"))
            elif key == "교육시간표":
                result = all(value(css) for css in self._TIMETABLE_FIELDS)
            elif key == "교육시간표 일차":
                result = len(self._find_all(By.CSS_SELECTOR, "ul.eduTime > li")) == 1
            elif key == "지역":
                result = all(value(f"#{i}") for i in ("real_area_cd", "real_signgu_cd", "real_dong_cd"))
            else:
                result = None

            if current and current in self.driver.window_handles:
                self.driver.switch_to.window(current)
            return result
        except Exception:
            return None

    # ---- 교육시간표: 2일차/3일차 삭제 ---------------------------------
    def remove_extra_day_rows(self):
        """교육시간표에서 1일차만 남기고 2일차~N일차를 삭제 (삭제 버튼: button.iconClose)"""
        # 과정 선택 후 교육시간표가 ajax로 채워질 때까지 대기
        try:
            WebDriverWait(self.driver, config.WAIT_TIME).until(
                lambda d: len(d.find_elements(By.CSS_SELECTOR, "ul.eduTime > li")) >= 1
            )
        except TimeoutException:
            log.warning("  교육시간표(ul.eduTime > li)가 나타나지 않았습니다.")
            return

        for _ in range(10):
            items = self._find_all(By.CSS_SELECTOR, "ul.eduTime > li")
            if len(items) <= 1:
                break
            btns = items[-1].find_elements(By.CSS_SELECTOR, "button.iconClose")
            if not btns:
                log.warning("  마지막 일차 항목에서 삭제 버튼(button.iconClose)을 찾지 못했습니다.")
                break
            self._js_click(btns[0])
            self._accept_alert_if_present()
            time.sleep(0.5)

        if len(self._find_all(By.CSS_SELECTOR, "ul.eduTime > li")) > 1:
            log.warning("  교육시간표에 1일차 외 항목이 남아있습니다. 화면을 확인해주세요.")
            input("  >> 필요하면 브라우저에서 2일차 이후를 직접 삭제한 뒤 Enter를 눌러주세요...")

    def _accept_alert_if_present(self):
        """삭제 시 confirm/alert 창이 뜨면 확인을 누른다."""
        try:
            WebDriverWait(self.driver, 1).until(EC.alert_is_present())
            self.driver.switch_to.alert.accept()
        except TimeoutException:
            pass

    # ---- 한 행(row) 등록 ----------------------------------------------
    def fill_and_submit(self, row):
        course = str(row[config.COL_COURSE]).strip()
        date_str = normalize_date(row[config.COL_DATE])
        sh, sm = normalize_time(row[config.COL_START])
        eh, em = normalize_time(row[config.COL_END])
        address = str(row[config.COL_ADDRESS]).strip()
        address_detail = _txt(row.get(config.COL_ADDRESS_DETAIL))
        sido = str(row[config.COL_SIDO]).strip()
        sigungu = str(row[config.COL_SIGUNGU]).strip()
        dong = str(row[config.COL_DONG]).strip()
        assistant = _txt(row.get(config.COL_ASSISTANT))
        # '담당 배움터' 열(공백 유무 무관)이 있으면 그 값을, 없거나 비어 있으면 config 기본값을 사용
        place = next(
            (str(v).strip() for k, v in row.items()
             if k.replace(" ", "") == config.COL_PLACE.replace(" ", "") and v not in (None, "")),
            self.place_name,
        )

        self.driver.get(config.EDU_PLAN_URL)
        time.sleep(config.DELAY_BETWEEN_ACTIONS)

        # 1) 교육과정명 검색/선택
        self.select_course(course)

        # 2) PC/모바일 -> 모바일
        try:
            self._check_input("edc_env_se_cd_02")
        except TimeoutException:
            log.warning("  모바일 라디오 버튼(#edc_env_se_cd_02)을 찾지 못했습니다.")

        # 3) 교육시간표: 2/3일차 삭제 후 1일차 채우기
        self.remove_extra_day_rows()
        try:
            self._set_value_js(By.CSS_SELECTOR, ".scheduleDate", date_str)
            self._select_number(By.CSS_SELECTOR, "select[name='edc_bgn_hour[]']", sh)
            self._select_number(By.CSS_SELECTOR, "select[name='edc_bgn_minute[]']", sm)
            self._select_number(By.CSS_SELECTOR, "select[name='edc_end_hour[]']", eh)
            self._select_number(By.CSS_SELECTOR, "select[name='edc_end_minute[]']", em)
            self._set_text(By.CSS_SELECTOR, "input.toolsUsed[name='eqpmn_nm[]']", config.PLAN_DEFAULT_VALUES["사용기자재"])
        except (TimeoutException, NoSuchElementException) as e:
            log.warning(f"  교육시간표 입력 중 일부 필드를 찾지 못했습니다: {e}")
            input("  >> 교육시간표를 직접 확인/수정한 뒤 Enter를 눌러주세요...")

        # 4) 접수기간 / 교육일자
        rcept_bgn = config.RCEPT_BGN_DATE or date_str[:8] + "01"  # 미지정 시 교육일이 속한 달의 1일
        self._set_value_js(By.ID, "rcept_bgn_dt", rcept_bgn)
        self._set_value_js(By.ID, "rcept_end_dt", date_str)
        self._set_value_js(By.ID, "edc_bgn_dt", date_str)
        self._set_value_js(By.ID, "edc_end_dt", date_str)

        # 5) 담당 배움터
        self.select_place(place)

        # 6) 실제 교육장 - 직접입력 + 상세주소 + 시/도/시군구/읍면동
        try:
            self._check_input("real_edc_plc_cd_WR")
        except TimeoutException:
            log.warning("  실제교육장 '직접입력' 라디오(#real_edc_plc_cd_WR)를 찾지 못했습니다.")

        self.fill_address_via_popup(address, address_detail)

        try:
            self._select_by_text(By.ID, "real_area_cd", sido)
            time.sleep(0.8)  # 하위 select(시군구) 옵션 로드 대기
            self._select_by_text(By.ID, "real_signgu_cd", sigungu)
            time.sleep(0.8)  # 하위 select(읍면동) 옵션 로드 대기
            self._select_by_text(By.ID, "real_dong_cd", dong)
        except (TimeoutException, NoSuchElementException) as e:
            log.warning(f"  시/도-시군구-읍면동 선택 중 문제 발생: {e}")
            input("  >> 지역 선택(시/도/시군구/읍면동)을 직접 확인한 뒤 Enter를 눌러주세요...")

        # 7) 보조강사/가이드명
        if assistant:  # 보조강사 칸이 비어 있으면 (사이트에서도 선택 항목이므로) 건너뜀
            self.select_assistant(assistant)

        # 8) 교육정원 / 교육대상자유형 / 취약계층
        try:
            self._set_text(By.ID, "edc_nmpr_co", config.PLAN_DEFAULT_VALUES["교육정원"])
        except TimeoutException:
            log.warning("  교육정원(#edc_nmpr_co)을 찾지 못했습니다.")

        try:
            self._select_by_text(By.ID, "edc_trgter_ty_cd", config.PLAN_DEFAULT_VALUES["교육대상자유형"])
        except (TimeoutException, NoSuchElementException):
            log.warning("  교육대상자유형(#edc_trgter_ty_cd)을 찾지 못했습니다.")

        if config.PLAN_DEFAULT_VALUES["취약계층"]:
            try:
                self._check_input("vulnerable_yn")
            except TimeoutException:
                log.warning("  취약계층 체크박스(#vulnerable_yn)를 찾지 못했습니다.")

        # 9) 등록 (설정에 따라 행마다 사용자 확인)
        if getattr(config, "DRY_RUN", False):
            print(f"\n  [{course} / {date_str}] 입력 완료 (테스트 모드: 저장하지 않습니다). 화면을 확인해주세요.")
            input("      확인했으면 Enter를 눌러 다음 행으로 넘어가세요: ")
            log.info(f"[저장 안 함] {course} / {date_str}")
            return False

        if not getattr(config, "AUTO_SUBMIT", True):
            print(f"\n  >>> [{course} / {date_str}] 입력 완료. 화면을 확인해주세요.")
            answer = input("      등록(저장)하시려면 Enter, 이 행을 건너뛰려면 s + Enter: ").strip().lower()
            if answer == "s":
                log.info(f"[건너뜀] {course} / {date_str}")
                return False

        try:
            self._click(By.ID, "saveBtn")
        except TimeoutException:
            log.error(f"[등록버튼 못찾음] {course} / {date_str} - #saveBtn 확인 필요")
            input("  >> 등록 버튼을 직접 눌러주세요. 완료 후 Enter...")
            return True

        ok, messages = self._handle_save_alerts()
        if ok:
            log.info(f"[등록완료] {course} / {date_str}")
            return True
        log.error(f"[등록실패] {course} / {date_str} - 안내창: {messages}")
        input("  >> 화면의 안내 메시지를 확인하고 필요하면 직접 수정/등록한 뒤 Enter를 눌러주세요...")
        return False

    def _handle_save_alerts(self):
        """
        등록 버튼 클릭 후 나타나는 확인창('등록하시겠습니까?')은 확인을 누르고,
        이어지는 안내창을 읽는다. 필수값 누락 등 실패성 메시지가 있으면 ok=False.
        """
        failure_words = ("입력", "선택", "확인해", "필수", "오류", "실패", "없습니다", "올바르")
        messages = []
        confirmed = False
        failed = False
        for _ in range(4):
            try:
                WebDriverWait(self.driver, 4).until(EC.alert_is_present())
            except TimeoutException:
                break
            alert = self.driver.switch_to.alert
            text = alert.text.strip()
            messages.append(text)
            alert.accept()
            if "등록하시겠습니까" in text:
                confirmed = True
            elif any(w in text for w in failure_words):
                failed = True
            time.sleep(0.5)
        return (confirmed and not failed), messages


def run(loaded):
    """확인이 끝난 교육과정들을 크롬으로 등록한다."""
    targets = loaded.rows
    # 형식 문제로 제외된 줄은 마지막 '수동으로 수정이 필요한 교육과정' 목록에 함께 정리한다.
    for where, course, date, problems in loaded.invalid:
        issues.issues.append(((where, course, date), "형식오류", "; ".join(problems) + " -> 이 줄은 자동 등록에서 제외되었으니 직접 등록하세요."))
    log.info(f"총 {len(targets)}건 등록 예정")

    registrar = start_registrar(EduPlanRegistrar)
    issues.verify = registrar.field_filled
    registrar.login_and_wait()

    done, failed = 0, []
    try:
        for i, row in enumerate(targets, 1):
            course = str(row[config.COL_COURSE]).strip()
            date_str = normalize_date(row[config.COL_DATE])
            issues.current = (i, course, date_str)
            issues.mark = len(issues.issues)
            log.info(f"\n[{i}/{len(targets)}] {course} / {date_str} 처리 중...")
            try:
                if registrar.fill_and_submit(row):
                    done += 1
            except Exception as e:
                log.error(f"  처리 중 오류 발생: {e}")
                failed.append((course, date_str, str(e)))
                cont = input("  오류가 발생했습니다. 다음 행으로 계속할까요? (Enter=계속, x=중단): ").strip().lower()
                if cont == "x":
                    break
            time.sleep(config.DELAY_BETWEEN_ACTIONS)
    except KeyboardInterrupt:
        print("\n사용자가 중단했습니다. 지금까지의 결과를 정리합니다.")

    log.info(f"\n===== 완료: {done}건 등록 / 총 {len(targets)}건 대상 =====")
    if failed:
        log.info("실패한 항목:")
        for c, d, e in failed:
            log.info(f"  - {c} / {d} : {e}")

    print_report(
        total=len(targets), done=done, unit="건",
        group_title="수동으로 수정이 필요한 교육과정",
        key_names=("번호", "과정명", "일자"),
    )
