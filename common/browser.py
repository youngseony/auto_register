# -*- coding: utf-8 -*-
"""크롬(Selenium) 실행과 화면 조작 공통 기능 (두 작업 공통)"""

from selenium import webdriver
from selenium.common.exceptions import ElementClickInterceptedException, InvalidSessionIdException, NoSuchWindowException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

import config
from common.console import input


class BaseRegistrar:
    def __init__(self, unhandled_prompt=None):
        self._unhandled_prompt = unhandled_prompt
        self._start_chrome()

    def _start_chrome(self):
        """크롬을 새로 띄운다. (처음 시작할 때, 로그인 대기 중 크롬 연결이 끊겼을 때)"""
        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)  # 스크립트 종료 후에도 브라우저 유지
        if self._unhandled_prompt:
            # 확인창(confirm)이 우리가 읽기 전에 자동으로 '취소' 처리되지 않도록 한다.
            options.set_capability("unhandledPromptBehavior", self._unhandled_prompt)
        if config.CHROMEDRIVER_PATH:
            from selenium.webdriver.chrome.service import Service
            self.driver = webdriver.Chrome(service=Service(config.CHROMEDRIVER_PATH), options=options)
        else:
            self.driver = webdriver.Chrome(options=options)  # Selenium 4.6+ 자동 드라이버 관리
        self.driver.maximize_window()
        self.wait = WebDriverWait(self.driver, config.WAIT_TIME)

    # ---- 공통 유틸 -------------------------------------------------
    def _js_click(self, el):
        try:
            el.click()
        except ElementClickInterceptedException:
            self.driver.execute_script("arguments[0].click();", el)

    def _click(self, by, sel, timeout=None):
        w = WebDriverWait(self.driver, timeout or config.WAIT_TIME)
        el = w.until(EC.element_to_be_clickable((by, sel)))
        self._js_click(el)
        return el

    def _find(self, by, sel, timeout=None):
        w = WebDriverWait(self.driver, timeout or config.WAIT_TIME)
        return w.until(EC.presence_of_element_located((by, sel)))

    def _find_all(self, by, sel):
        return self.driver.find_elements(by, sel)

    def _check_input(self, input_id):
        """
        커스텀 스타일 radio/checkbox(input이 숨겨져 있고 label만 보이는 형태)를 체크한다.
        input이 동적으로 생성될 수 있으므로 존재할 때까지 기다린 뒤, label 클릭 -> JS 클릭 순으로 시도.
        """
        el = self._find(By.ID, input_id)
        if el.is_selected():
            return el
        labels = self._find_all(By.CSS_SELECTOR, f"label[for='{input_id}']")
        if labels:
            self._js_click(labels[0])
        if not el.is_selected():
            self.driver.execute_script("arguments[0].click();", el)
        return el

    # ---- 로그인 대기 -------------------------------------------------
    def login_and_wait(self):
        """로그인 대기는 시간 제한이 없다. Enter 후 크롬 연결이 끊겨 있으면 크롬을 새로 열고 다시 로그인하게 한다."""
        while True:
            self.driver.get(config.LOGIN_URL)
            answer = input("\n열린 크롬 창에서 로그인을 완료한 뒤, 이 창을 클릭하고 Enter를 눌러주세요...")
            # 로그인 없이 Enter를 누르면 모든 줄이 실패하므로, 아직 로그인 화면이면 다시 안내한다.
            while answer.strip().lower() != "s" and self._on_login_page():
                answer = input(
                    "\n아직 로그인 화면입니다. 크롬에서 로그인을 완료한 뒤 Enter를 눌러주세요.\n"
                    "(이미 로그인했는데 이 안내가 계속 나오면 s + Enter 로 그대로 진행)... "
                )
            if self._browser_alive():
                return
            print("\n  ! 자동 등록용 크롬 창과 연결이 끊겼습니다. 크롬을 새로 열 테니, 새 창에서 다시 로그인해 주세요.")
            try:
                self.driver.quit()
            except Exception:
                pass
            self._start_chrome()

    def _browser_alive(self):
        """프로그램이 연 크롬 창이 아직 열려 있는지. 확인할 수 없으면 True(그대로 진행)."""
        try:
            self.driver.current_url
            return True
        except (InvalidSessionIdException, NoSuchWindowException):
            return False
        except Exception:
            return True

    def _on_login_page(self):
        """현재 크롬 화면이 로그인 페이지인지. 확인할 수 없으면 False(그대로 진행)."""
        try:
            return "login" in self.driver.current_url.lower()
        except Exception:
            return False


def start_registrar(cls):
    """크롬을 띄워 등록 담당 객체(cls)를 만든다. 크롬을 시작하지 못하면 안내 후 종료."""
    try:
        return cls()
    except Exception as e:
        first_line = (str(e).strip().splitlines() or [""])[0]
        raise SystemExit(
            "Chrome을 시작하지 못했습니다. Chrome이 설치되어 있는지, 인터넷에 연결되어 있는지 확인하세요.\n"
            f"({type(e).__name__}: {first_line})"
        )
