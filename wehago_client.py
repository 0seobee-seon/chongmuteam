"""Wehago(wehago.co.kr) 로그인·증명서 이메일 발송 자동화.

PRD 문서: docs/PRD-wehago-certificate.md (§0-2에 화면 캡처로 확인한 실제 경로 기록)

이 모듈은 실제 페이지 HTML을 확인하지 않고 화면 캡처만 보고 작성했다.
아래 두 그룹으로 나눠 신뢰도를 구분한다.

- 화면에 보이는 한글 문구를 그대로 클릭하는 부분(로그인, 급여관리, 조회,
  연말정산 근로소득원천징수영수증, 급여자료입력, 이메일 보내기 등)은
  캡처에서 직접 확인한 문구라 신뢰도가 높다.
- `# TODO(확인필요)`로 표시한 부분(더보기 메뉴 아이콘, 날짜 선택 위젯,
  사원 목록 검색/스크롤 방식)은 캡처에 DOM 구조가 드러나지 않아 추정으로
  작성했다. 첫 실행은 반드시 화면을 보면서 진행하고, 여기서 실패하면
  오류 메시지에 찍히는 단계를 보고 셀렉터를 조정해야 한다.
"""

import os
import sys
import time
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


def app_dir():
    """실행 파일(exe) 또는 스크립트가 실제로 있는 폴더.

    PyInstaller --onefile로 묶으면 __file__은 실행할 때마다 새로 생기고
    종료 시 지워지는 임시 폴더(_MEIxxxxxx)를 가리킨다 — 그 안에 로그나
    스크린샷을 저장하면 프로그램을 닫는 순간 사라진다. sys.frozen일 때는
    sys.executable(exe의 실제 위치)을 기준으로 삼아야 한다.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


LOGIN_URL = "https://wehago.com/#/login"
MAIN_URL_FRAGMENT = "#/main"
PAYROLL_APP_HOST = "smarta.wehago.com"
DEFAULT_TIMEOUT = 15
DEBUG_DIR = os.path.join(app_dir(), "debug")


class WehagoError(Exception):
    """Wehago 자동화 중 발생한 오류의 공통 기반 클래스."""


class LoginFailedError(WehagoError):
    pass


class TwoFactorRequiredError(WehagoError):
    """QR 2차인증 화면이 나타나면 자동 로그인이 불가능하다 (PRD §8 R2)."""


class EmployeeNotFoundError(WehagoError):
    pass


class BrowserStartFailedError(WehagoError):
    """Chrome/chromedriver 자체를 켜는 단계에서 실패했다 — 로그인 시도 이전 단계."""


class WehagoClient:
    def __init__(self, on_status=None, headless=False):
        self.on_status = on_status or (lambda msg: None)
        try:
            options = webdriver.ChromeOptions()
            if headless:
                options.add_argument("--headless=new")
                options.add_argument("--disable-gpu")
                options.add_argument("--no-sandbox")
            options.add_argument("--window-size=1400,900")
            # Wehago 로그인 직후 "알림 표시 허용" 네이티브 권한 팝업이 뜨는데,
            # 이건 페이지 DOM 밖의 브라우저 UI라 클릭으로 닫을 수 없다 —
            # 아예 요청 자체가 뜨지 않도록 차단해둔다.
            options.add_experimental_option(
                "prefs", {"profile.default_content_setting_values.notifications": 2}
            )
            # 기본 전략("normal")은 모든 하위 리소스가 끝날 때까지 기다리다
            # 느린 배경 스크립트 하나 때문에 driver.get() 자체가 오래 멈출 수
            # 있다 — 화면(DOM)이 뜨면 바로 진행하는 "eager"로 완화한다.
            options.page_load_strategy = "eager"
            self.driver = webdriver.Chrome(options=options)
        except Exception as e:
            raise BrowserStartFailedError(f"브라우저(Chrome) 시작에 실패했습니다: {e}")
        self.wait = WebDriverWait(self.driver, DEFAULT_TIMEOUT)

    def _status(self, msg):
        self.on_status(msg)

    def _click_by_text(self, text, timeout=DEFAULT_TIMEOUT):
        """텍스트가 정확히 일치하는 요소를 찾아 클릭한다.

        같은 텍스트를 가진 요소가 여러 개일 수 있고(예: 화면에는 안 보이는
        템플릿/툴팁용 요소), 그중 문서상 첫 번째가 항상 실제로 보이는
        것이라는 보장이 없다 — Selenium의 기본 element_to_be_clickable은
        XPath의 첫 매치 하나만 확인하므로, 보이지 않는 첫 매치 때문에
        영원히 타임아웃 나는 문제가 있었다. 모든 후보를 확인해 실제로
        보이고 활성화된 것을 찾는다.
        """
        xpath = f"//*[normalize-space(text())='{text}']"

        def _find_visible(driver):
            for candidate in driver.find_elements(By.XPATH, xpath):
                try:
                    if candidate.is_displayed() and candidate.is_enabled():
                        return candidate
                except Exception:
                    continue
            return False

        el = WebDriverWait(self.driver, timeout).until(_find_visible)
        target = self._nearest_clickable_ancestor(el)
        target.click()
        return target

    def _nearest_clickable_ancestor(self, el):
        """el(보통 아이콘 아래 캡션 텍스트)이 <a>/<button>이 아니면, 실제
        내비게이션을 담당하는 가장 가까운 <a>/<button> 조상을 찾아 그것을
        반환한다. 캡션 <p>를 눌러도 클릭 자체는 성공하지만 새 탭이 열리지
        않는 문제가 있어(부모 <a class="wehagoApp__sectionItem">가 실제
        대상이었음) 추가했다. 그런 조상이 없으면 el 그대로 반환한다.
        """
        try:
            found = self.driver.execute_script(
                """
                var p = arguments[0];
                for (var i = 0; i < 6 && p; i++) {
                    var tag = p.tagName ? p.tagName.toLowerCase() : '';
                    if (tag === 'a' || tag === 'button') return p;
                    p = p.parentElement;
                }
                return null;
                """,
                el,
            )
            return found if found is not None else el
        except Exception:
            return el

    def _save_debug_screenshot(self, label):
        """실패 지점을 눈으로 다시 확인할 수 있도록 스크린샷을 남긴다."""
        try:
            os.makedirs(DEBUG_DIR, exist_ok=True)
            path = os.path.join(DEBUG_DIR, f"{label}_{datetime.now():%Y%m%d_%H%M%S}.png")
            self.driver.save_screenshot(path)
            return path
        except Exception:
            return None

    def _handle_duplicate_login(self):
        """다른 기기에서 이미 로그인되어 있으면 Wehago가 "중복 로그인" 확인을
        묻는다 — 사용자 확인에 따라 항상 "확인"을 눌러 그 세션을 종료하고
        계속 진행한다.
        """
        try:
            WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(),'중복 로그인')]"))
            )
        except TimeoutException:
            return False
        self._status("다른 기기의 Wehago 세션을 종료하고 계속 진행합니다...")
        self._click_by_text("확인")
        time.sleep(2)  # 확인 클릭 직후 페이지 전환이 끝날 때까지 대기
        return True

    # ---- 로그인 ----

    def login(self, user_id, password):
        try:
            self._login_impl(user_id, password)
        except (LoginFailedError, TwoFactorRequiredError):
            raise
        except Exception as e:
            # 어떤 예외든 실패 지점 화면을 남겨야 다음에 원인을 알 수 있다.
            shot = self._save_debug_screenshot("login_error")
            hint = f" (스크린샷: {shot})" if shot else ""
            raise LoginFailedError(f"로그인 중 예상치 못한 오류: {e}{hint}")

    def _login_impl(self, user_id, password):
        self._status("Wehago 로그인 페이지로 이동 중...")
        self.driver.get(LOGIN_URL)

        # TODO(확인필요): 아이디/비밀번호 입력창을 type 속성만으로 구분한다.
        # 화면에는 두 개의 입력창만 보였으므로 우선 이렇게 작성했다.
        id_input = self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']"))
        )
        pw_input = self.driver.find_element(By.CSS_SELECTOR, "input[type='password']")

        id_input.clear()
        id_input.send_keys(user_id)
        pw_input.clear()
        pw_input.send_keys(password)

        self._status("로그인 시도 중...")
        # 화면 상단 메뉴에도 "로그인"이라는 문구가 있어 텍스트로 버튼을 찾으면
        # 엉뚱한 요소(메뉴 링크)를 누를 수 있다 — 비밀번호 입력란에서 Enter를
        # 눌러 폼을 제출하는 방식이 더 안전하다.
        pw_input.send_keys(Keys.RETURN)
        time.sleep(1)  # 제출 직후 페이지 반응이 나타날 때까지 짧게 대기
        self._handle_duplicate_login()

        try:
            WebDriverWait(self.driver, DEFAULT_TIMEOUT).until(
                lambda d: MAIN_URL_FRAGMENT in d.current_url
            )
        except TimeoutException:
            page_text = self.driver.find_element(By.TAG_NAME, "body").text
            shot = self._save_debug_screenshot("login_fail")
            if "QR" in page_text or "2차" in page_text or "인증" in page_text:
                raise TwoFactorRequiredError(
                    "로그인 후 2차인증(QR) 화면이 나타났습니다. 자동 로그인을 진행할 수 없습니다."
                )
            hint = f" (스크린샷: {shot})" if shot else ""
            raise LoginFailedError(
                f"로그인 실패로 보입니다 — 현재 URL: {self.driver.current_url}{hint}"
            )

        self._status("로그인 성공")
        time.sleep(1.5)  # 안내 팝업이 URL 전환보다 살짝 늦게 뜨는 경우가 있어 대기
        self._dismiss_popups()

    def _is_actually_on_top(self, el):
        """el이 놓인 좌표에서 실제로 화면 맨 위에 있는 요소가 el(또는 그
        자손/조상)인지 document.elementFromPoint로 확인한다. 겹쳐진 다른
        요소를 잘못 클릭하는 것을 막기 위한 용도.
        """
        try:
            rect = el.rect
            x = rect["x"] + rect["width"] / 2
            y = rect["y"] + rect["height"] / 2
            return self.driver.execute_script(
                """
                var el = arguments[0], x = arguments[1], y = arguments[2];
                var top = document.elementFromPoint(x, y);
                return top === el || el.contains(top) || (top && top.contains(el));
                """,
                el, x, y,
            )
        except Exception:
            return False

    def _dismiss_popups(self, attempts=3):
        """로그인 직후나 화면 전환 뒤 뜨는 광고 배너·2차인증 유도 팝업 등을
        닫는다. 화면 캡처에서 확인한 문구들:
        - 광고 배너·2차인증 유도 팝업 → "닫기"
        - 메인 화면 진입 시 자동으로 뜨는 "버튼을 제거하시겠습니까?" 확인창
          → "취소"를 눌러 원래 상태를 그대로 둔다(이 자동화가 사용자의
          화면 구성을 바꾸면 안 되므로)

        페이지에 같은 텍스트를 가진 요소가 여러 개(숨겨진 메뉴 등) 있을 수
        있고, is_displayed()만으로는 그 자리에 실제로 다른 요소가 겹쳐
        있는지까지는 걸러내지 못해 elementFromPoint로 그 좌표의 실제
        최상단 요소가 맞는지 확인한 뒤에만 클릭한다.
        """
        for _ in range(attempts):
            candidates = self.driver.find_elements(
                By.XPATH,
                "//*[normalize-space(text())='닫기' or normalize-space(text())='취소']",
            )
            clicked = False
            for el in candidates:
                try:
                    if el.is_displayed() and self._is_actually_on_top(el):
                        el.click()
                        clicked = True
                        time.sleep(0.3)
                        break
                except Exception:
                    continue
            if not clicked:
                break

    def open_payroll_app(self):
        """"급여관리" 클릭 결과가 새 탭으로 열릴 때도, 같은 탭 안에서
        전환될 때도 있어(실제로 둘 다 관찰됨) 두 경우 모두 대응한다.
        """
        self._status("급여관리 메뉴로 진입 중...")
        self._dismiss_popups()  # 팝업이 첫 시도 이후에 뒤늦게 뜨는 경우를 한 번 더 대비
        original_handles = set(self.driver.window_handles)
        self._click_by_text("급여관리")

        def _reached_payroll(driver):
            if len(driver.window_handles) > len(original_handles):
                return "new_tab"
            if PAYROLL_APP_HOST in driver.current_url:
                return "same_tab"
            try:
                if driver.find_elements(By.XPATH, "//*[contains(text(),'근로소득관리')]"):
                    return "same_tab"
            except Exception:
                pass
            return False

        result = WebDriverWait(self.driver, DEFAULT_TIMEOUT).until(_reached_payroll)
        if result == "new_tab":
            new_handle = (set(self.driver.window_handles) - original_handles).pop()
            self.driver.switch_to.window(new_handle)
            self.wait.until(lambda d: PAYROLL_APP_HOST in d.current_url)

        self._status("급여관리 화면 로드 완료")
        self._dismiss_popups()

    def set_fiscal_year(self, year, max_scroll=10):
        """상단의 연도 배지("2026 ▾" 형태)를 눌러 대상 연도로 전환한다.

        배지 자체의 정확한 선택자는 알 수 없어(캡처로는 텍스트만 보임),
        "4자리 연도처럼 보이는 클릭 가능 요소"를 찾는 방식으로 구현했다.
        누른 뒤 나오는 목록은 화면 캡처에서 확인한 대로 연도 텍스트를
        그대로 클릭하면 되는 단순 목록이라 기존 _click_by_text를 재사용한다.
        """
        self._status(f"{year}년으로 연도 전환 중...")

        badge = None
        for el in self.driver.find_elements(By.XPATH, "//button | //a | //div | //span"):
            text = el.text.strip().replace("▾", "").replace("v", "").strip()
            if text.isdigit() and 2015 <= int(text) <= 2035:
                badge = el
                break
        if badge is None:
            raise WehagoError("연도 선택 배지를 찾지 못했습니다.")
        badge.click()

        for _ in range(max_scroll):
            try:
                self._click_by_text(str(year), timeout=2)
                time.sleep(1)
                return
            except TimeoutException:
                lists = self.driver.find_elements(By.XPATH, "//ul | //*[contains(@class,'scroll')]")
                if lists:
                    self.driver.execute_script("arguments[0].scrollTop += 100", lists[-1])
                time.sleep(0.3)
        raise WehagoError(f"연도 목록에서 {year}년을 찾지 못했습니다.")

    # ---- 원천징수영수증 ----

    def go_to_withholding_menu(self):
        self._status("연말정산 근로소득원천징수영수증 메뉴로 이동 중...")
        self._click_by_text("연말정산 근로소득원천징수영수증")
        time.sleep(1)  # 화면 전환 대기

    def set_period_withholding(self, year):
        # TODO(확인필요): 정산년월 범위 입력창(달력 아이콘)이 직접 타이핑을
        # 허용하는지 캡처만으로는 알 수 없다. 우선 텍스트 입력 후 조회를 시도한다.
        self._status(f"{year}년 원천징수영수증 조회 중...")
        self._click_by_text("연말")
        self._click_by_text("조회")
        time.sleep(1)

    def select_employee_withholding(self, employee_name):
        self._select_employee_in_list(employee_name)

    def send_withholding_email(self, expected_email=None):
        """⋮ 메뉴 → 영수증/소득공제신고서 이메일 전송 팝업을 열고 발송한다.

        Wehago가 표시하는 등록 이메일이 expected_email과 다르면 보내지 않고
        ("mismatch", 등록된 이메일)을 반환해 담당자가 나중에 직접 판단하게
        한다 — PRD FR-4의 "발송 전 교차 확인"에 해당한다.
        """
        self._open_kebab_menu()
        self._click_by_text("영수증/소득공제신고서 이메일 전송")
        return self._handle_email_popup(expected_email, send_button_text="이메일 보내기")

    # ---- 급여명세서 ----

    def go_to_payslip_menu(self):
        self._status("급여자료입력 메뉴로 이동 중...")
        self._click_by_text("급여자료입력")
        time.sleep(1)

    def set_period_payslip(self, year, month):
        # TODO(확인필요): 귀속연월 입력창도 정산년월과 마찬가지로 달력 위젯이라
        # 직접 타이핑이 통하는지 확인이 필요하다.
        self._status(f"귀속연월 {year}.{month:02d} 조회 중...")
        period_input = self.wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//label[contains(text(),'귀속연월')]/following::input[1]")
            )
        )
        period_input.click()
        period_input.send_keys(Keys.CONTROL, "a")
        period_input.send_keys(f"{year}.{month:02d}")
        period_input.send_keys(Keys.ESCAPE)
        self._click_by_text("조회")
        time.sleep(1)

    def select_employee_payslip(self, employee_name):
        self._select_employee_in_list(employee_name, use_checkbox=True)

    def send_payslip_email(self, expected_email=None):
        self._click_by_text("급여명세서 보내기")
        return self._handle_email_popup(expected_email, send_button_text="보내기")

    # ---- 공용 ----

    def _select_employee_in_list(self, employee_name, use_checkbox=False, max_scroll=15):
        # TODO(확인필요): 673명 규모 목록이 가상 스크롤(virtualized)일 가능성이
        # 높아, 화면에 없는 사원은 스크롤을 내리며 다시 찾는다. 실제 스크롤
        # 컨테이너 선택자는 확인 후 조정이 필요할 수 있다.
        for _ in range(max_scroll):
            rows = self.driver.find_elements(
                By.XPATH, f"//tr[.//text()[contains(., '{employee_name}')]]"
            )
            if rows:
                row = rows[0]
                if use_checkbox:
                    checkbox = row.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
                    if not checkbox.is_selected():
                        checkbox.click()
                else:
                    row.click()
                return
            self.driver.execute_script(
                "arguments[0].scrollTop += 300", self.driver.find_element(By.CSS_SELECTOR, "table")
            )
            time.sleep(0.3)
        raise EmployeeNotFoundError(f"목록에서 '{employee_name}'을(를) 찾지 못했습니다.")

    def _open_kebab_menu(self):
        # TODO(확인필요): "⋮" 더보기 아이콘은 텍스트가 없어 캡처만으로는
        # 정확한 선택자를 알 수 없다. 최초 실행 시 가장 먼저 확인해야 할 지점.
        kebab = self.wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(@class,'more') or contains(@aria-label,'더보기')]")
            )
        )
        kebab.click()

    def _handle_email_popup(self, expected_email, send_button_text):
        self.wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(),'이메일')]")))
        time.sleep(0.5)

        email_cells = self.driver.find_elements(By.XPATH, "//td[contains(text(),'@')]")
        actual_email = email_cells[0].text.strip() if email_cells else ""

        if expected_email and actual_email and expected_email.strip().lower() != actual_email.lower():
            self._click_by_text("닫기(Esc)")
            return "mismatch", actual_email

        self._click_by_text(send_button_text)
        time.sleep(1)
        return "sent", actual_email

    def close(self):
        try:
            self.driver.quit()
        except Exception:
            pass
