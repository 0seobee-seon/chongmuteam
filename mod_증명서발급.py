"""Wehago 증명서(원천징수영수증/급여명세서) 자동발급 — 허브용 모듈.

PRD: docs/PRD-wehago-certificate.md

담당자가 신청 정보를 여러 건 입력해 목록에 모은 뒤, 목록을 확인하고
"일괄 처리" 버튼을 눌러야 Wehago 로그인·조회·이메일 발송이 진행된다
(반자동 — 사람 확인 없이 자동 발급하지 않는다).
"""

import csv
import os
import threading
import tkinter as tk
from datetime import datetime
from tkinter import simpledialog, messagebox, ttk

import keyring

from wehago_client import (
    WehagoClient,
    LoginFailedError,
    TwoFactorRequiredError,
    EmployeeNotFoundError,
    WehagoError,
    app_dir,
)

SERVICE = "WehagoCertificate"
DOC_TYPES = ["원천징수영수증", "급여명세서"]
# exe로 빌드된 경우 __file__ 기준 경로는 실행할 때마다 사라지는 임시 폴더를
# 가리키므로(wehago_client.app_dir 참고), 로그도 같은 방식으로 실제 exe 위치를 쓴다.
LOG_DIR = os.path.join(app_dir(), "logs")


def _get_credentials():
    user_id = keyring.get_password(SERVICE, "user_id")
    password = keyring.get_password(SERVICE, "password") if user_id else None
    return user_id, password


def _save_credentials(user_id, password):
    keyring.set_password(SERVICE, "user_id", user_id)
    keyring.set_password(SERVICE, "password", password)


class App(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Wehago 증명서 자동발급")
        self.resizable(False, False)
        self.geometry("760x560")
        self.pending = []  # 각 항목: dict(name, doc_type, year, month, email, status, note)
        self._batch_running = False
        self._build_ui()
        self._refresh_account_label()

    # ---- UI 구성 ----

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        # 계정 영역
        frm_account = tk.LabelFrame(self, text="Wehago 계정", font=("맑은 고딕", 10, "bold"))
        frm_account.pack(fill="x", **pad)
        self.account_var = tk.StringVar()
        tk.Label(frm_account, textvariable=self.account_var, font=("맑은 고딕", 10)).pack(
            side="left", padx=8, pady=6
        )
        tk.Button(frm_account, text="계정 설정", command=self._prompt_credentials).pack(
            side="right", padx=8, pady=6
        )

        # 입력 폼
        frm_input = tk.LabelFrame(self, text="신청 정보 입력", font=("맑은 고딕", 10, "bold"))
        frm_input.pack(fill="x", **pad)

        tk.Label(frm_input, text="사원명", font=("맑은 고딕", 10)).grid(
            row=0, column=0, sticky="e", padx=8, pady=6
        )
        self.name_var = tk.StringVar()
        tk.Entry(frm_input, textvariable=self.name_var, width=14).grid(row=0, column=1, padx=4)

        tk.Label(frm_input, text="문서 종류", font=("맑은 고딕", 10)).grid(
            row=0, column=2, sticky="e", padx=8
        )
        self.doc_type_var = tk.StringVar(value=DOC_TYPES[0])
        doc_combo = ttk.Combobox(
            frm_input, textvariable=self.doc_type_var, values=DOC_TYPES, width=14, state="readonly"
        )
        doc_combo.grid(row=0, column=3, padx=4)
        doc_combo.bind("<<ComboboxSelected>>", self._on_doc_type_changed)

        tk.Label(frm_input, text="대상 연도", font=("맑은 고딕", 10)).grid(
            row=1, column=0, sticky="e", padx=8, pady=6
        )
        this_year = datetime.now().year
        MIN_YEAR = 2018
        self.year_var = tk.StringVar(value=str(this_year))
        ttk.Combobox(
            frm_input, textvariable=self.year_var,
            values=[str(y) for y in range(this_year, MIN_YEAR - 1, -1)],
            width=8, state="readonly",
        ).grid(row=1, column=1, padx=4, sticky="w")

        tk.Label(frm_input, text="대상 월", font=("맑은 고딕", 10)).grid(
            row=1, column=2, sticky="e", padx=8
        )
        self.month_var = tk.StringVar(value=f"{datetime.now().month:02d}")
        self.month_combo = ttk.Combobox(
            frm_input, textvariable=self.month_var,
            values=[f"{m:02d}" for m in range(1, 13)], width=6, state="readonly",
        )
        self.month_combo.grid(row=1, column=3, padx=4, sticky="w")

        tk.Label(frm_input, text="신청자 이메일", font=("맑은 고딕", 10)).grid(
            row=2, column=0, sticky="e", padx=8, pady=6
        )
        self.email_var = tk.StringVar()
        tk.Entry(frm_input, textvariable=self.email_var, width=30).grid(
            row=2, column=1, columnspan=3, padx=4, sticky="w"
        )
        tk.Label(
            frm_input, text="(Wehago 등록 이메일과 다르면 발송 전 경고합니다)",
            font=("맑은 고딕", 8), fg="gray",
        ).grid(row=3, column=1, columnspan=3, sticky="w", padx=4)

        tk.Button(
            frm_input, text="목록에 추가", command=self._add_item,
            font=("맑은 고딕", 10, "bold"), bg="#1F497D", fg="white", width=12,
        ).grid(row=0, column=4, rowspan=2, padx=(16, 8))

        self._on_doc_type_changed()

        # 실행 영역 — 처리 대기 목록(Treeview)이 expand=True로 남은 공간을 다
        # 차지하면 고정 크기 창에서 이 버튼이 밀려 안 보일 수 있어, side="bottom"으로
        # 목록보다 먼저 공간을 확보해둔다.
        frm_run = tk.Frame(self)
        frm_run.pack(fill="x", side="bottom", **pad)
        self.status_var = tk.StringVar(value="대기 중")
        tk.Label(frm_run, textvariable=self.status_var, font=("맑은 고딕", 9), fg="#444").pack(
            side="left", padx=8
        )
        self.run_btn = tk.Button(
            frm_run, text="일괄 처리 시작", command=self._start_batch,
            font=("맑은 고딕", 11, "bold"), bg="#1F497D", fg="white", width=16, height=2,
        )
        self.run_btn.pack(side="right", padx=8)

        self.headless_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            frm_run, text="브라우저 화면 숨기고 실행(백그라운드)",
            variable=self.headless_var, font=("맑은 고딕", 9),
        ).pack(side="right", padx=8)

        # 처리 대기 목록
        frm_list = tk.LabelFrame(self, text="처리 대기 목록", font=("맑은 고딕", 10, "bold"))
        frm_list.pack(fill="both", expand=True, **pad)

        columns = ("name", "doc_type", "period", "email", "status")
        self.tree = ttk.Treeview(frm_list, columns=columns, show="headings", height=12)
        headers = {"name": "사원명", "doc_type": "문서종류", "period": "대상기간", "email": "신청자 이메일", "status": "상태"}
        widths = {"name": 90, "doc_type": 110, "period": 90, "email": 200, "status": 160}
        for col in columns:
            self.tree.heading(col, text=headers[col])
            self.tree.column(col, width=widths[col], anchor="center")
        self.tree.pack(fill="both", expand=True, padx=6, pady=6)

        frm_list_btn = tk.Frame(frm_list)
        frm_list_btn.pack(fill="x", padx=6, pady=(0, 6))
        tk.Button(frm_list_btn, text="선택 삭제", command=self._remove_selected).pack(side="left", padx=4)
        tk.Button(frm_list_btn, text="전체 삭제", command=self._clear_all).pack(side="left", padx=4)

    def _on_doc_type_changed(self, event=None):
        # 원천징수영수증은 연간 문서라 월 선택이 필요 없다.
        if self.doc_type_var.get() == "원천징수영수증":
            self.month_combo.configure(state="disabled")
        else:
            self.month_combo.configure(state="readonly")

    def _refresh_account_label(self):
        user_id, _ = _get_credentials()
        if user_id:
            self.account_var.set(f"등록된 계정: {user_id}")
        else:
            self.account_var.set("계정이 설정되지 않았습니다 — 먼저 '계정 설정'을 눌러주세요.")

    # ---- 계정 설정 ----

    def _prompt_credentials(self):
        user_id, _ = _get_credentials()
        new_id = simpledialog.askstring(
            "Wehago 계정 설정", "Wehago 아이디:", initialvalue=user_id or "", parent=self
        )
        if not new_id:
            return
        new_pw = simpledialog.askstring(
            "Wehago 계정 설정", "Wehago 비밀번호:", show="*", parent=self
        )
        if not new_pw:
            return
        _save_credentials(new_id, new_pw)
        self._refresh_account_label()
        messagebox.showinfo("완료", "Wehago 계정 정보가 저장되었습니다.", parent=self)

    # ---- 목록 관리 ----

    def _add_item(self):
        name = self.name_var.get().strip()
        email = self.email_var.get().strip()
        year = self.year_var.get().strip()
        doc_type = self.doc_type_var.get()
        month = self.month_var.get().strip() if doc_type == "급여명세서" else ""

        if not name:
            messagebox.showwarning("입력 확인", "사원명을 입력하세요.", parent=self)
            return
        # 신청자 이메일은 필수가 아니다 — 입력하면 Wehago 등록 이메일과 다를 때
        # 경고해주는 교차 확인용으로만 쓰인다(FR-4). 비워두면 그 확인을 건너뛴다.
        if email and ("@" not in email or "." not in email.split("@")[-1]):
            messagebox.showwarning("입력 확인", "이메일 형식이 올바르지 않습니다.", parent=self)
            return

        period = f"{year}" if doc_type == "원천징수영수증" else f"{year}.{month}"
        item = {
            "name": name, "doc_type": doc_type, "year": year, "month": month,
            "email": email, "period": period, "status": "대기", "note": "",
        }
        self.pending.append(item)
        self.tree.insert("", "end", values=(name, doc_type, period, email, "대기"))

        self.name_var.set("")
        self.email_var.set("")

    def _remove_selected(self):
        if self._batch_running:
            return
        for iid in self.tree.selection():
            idx = self.tree.index(iid)
            self.tree.delete(iid)
            del self.pending[idx]

    def _clear_all(self):
        if self._batch_running:
            return
        self.tree.delete(*self.tree.get_children())
        self.pending.clear()

    # ---- 일괄 처리 ----

    def _start_batch(self):
        if self._batch_running:
            return
        if not self.pending:
            messagebox.showinfo("알림", "처리할 항목이 없습니다.", parent=self)
            return

        user_id, password = _get_credentials()
        if not user_id or not password:
            messagebox.showwarning("알림", "먼저 Wehago 계정을 설정하세요.", parent=self)
            return

        if not messagebox.askyesno(
            "일괄 처리 확인",
            f"{len(self.pending)}건을 처리합니다.\nWehago에 로그인해 순차적으로 조회·발송합니다.\n계속하시겠습니까?",
            parent=self,
        ):
            return

        self._batch_running = True
        self.run_btn.configure(state="disabled")
        for iid in self.tree.get_children():
            self.tree.set(iid, "status", "대기")

        thread = threading.Thread(
            target=self._run_batch, args=(user_id, password, self.headless_var.get()), daemon=True
        )
        thread.start()

    def _set_status(self, msg):
        self.after(0, lambda: self.status_var.set(msg))

    def _set_item_status(self, index, status, note=""):
        def _update():
            self.pending[index]["status"] = status
            self.pending[index]["note"] = note
            iid = self.tree.get_children()[index]
            self.tree.set(iid, "status", status)
        self.after(0, _update)

    def _run_batch(self, user_id, password, headless):
        client = None
        results = []
        try:
            client = WehagoClient(on_status=self._set_status, headless=headless)
            client.login(user_id, password)
            client.open_payroll_app()
        except (LoginFailedError, TwoFactorRequiredError, WehagoError) as e:
            # FR-3: 로그인 실패 시 전체 배치를 즉시 중단한다.
            for i in range(len(self.pending)):
                self._set_item_status(i, "실패", f"로그인 실패: {e}")
            self._finish_batch(client, results, aborted=True, reason=str(e))
            return
        except Exception as e:
            for i in range(len(self.pending)):
                self._set_item_status(i, "실패", f"초기화 오류: {e}")
            self._finish_batch(client, results, aborted=True, reason=str(e))
            return

        for i, item in enumerate(self.pending):
            self._set_item_status(i, "처리중")
            try:
                client.set_fiscal_year(int(item["year"]))
                if item["doc_type"] == "원천징수영수증":
                    client.go_to_withholding_menu()
                    client.set_period_withholding(int(item["year"]))
                    client.select_employee_withholding(item["name"])
                    outcome, actual_email = client.send_withholding_email(item["email"])
                else:
                    client.go_to_payslip_menu()
                    client.set_period_payslip(int(item["year"]), int(item["month"]))
                    client.select_employee_payslip(item["name"])
                    outcome, actual_email = client.send_payslip_email(item["email"])

                if outcome == "sent":
                    self._set_item_status(i, "성공")
                    results.append((item, "성공", ""))
                else:  # mismatch
                    note = f"이메일 불일치 — 신청: {item['email']} / Wehago 등록: {actual_email}"
                    self._set_item_status(i, "확인 필요", note)
                    results.append((item, "확인 필요", note))

            except EmployeeNotFoundError as e:
                self._set_item_status(i, "실패", str(e))
                results.append((item, "실패", str(e)))
            except Exception as e:
                self._set_item_status(i, "실패", f"오류: {e}")
                results.append((item, "실패", f"오류: {e}"))

        self._finish_batch(client, results, aborted=False)

    def _finish_batch(self, client, results, aborted, reason=""):
        if client:
            client.close()
        if results:
            self._write_log(results)

        def _done():
            self._batch_running = False
            self.run_btn.configure(state="normal")
            self.status_var.set("처리 완료")
            if aborted:
                messagebox.showerror(
                    "처리 중단", f"로그인/초기화 단계에서 중단되었습니다:\n{reason}", parent=self
                )
                return
            success = sum(1 for _, s, _ in results if s == "성공")
            need_check = sum(1 for _, s, _ in results if s == "확인 필요")
            failed = sum(1 for _, s, _ in results if s == "실패")
            messagebox.showinfo(
                "처리 결과",
                f"완료: {success}건 성공 / {need_check}건 확인 필요 / {failed}건 실패\n\n"
                f"자세한 내용은 logs 폴더의 로그 파일을 확인하세요.",
                parent=self,
            )

        self.after(0, _done)

    def _write_log(self, results):
        os.makedirs(LOG_DIR, exist_ok=True)
        path = os.path.join(LOG_DIR, f"증명서발급_{datetime.now():%Y%m%d}.csv")
        is_new = not os.path.exists(path)
        with open(path, "a", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            if is_new:
                writer.writerow(["처리일시", "사원명", "문서종류", "대상기간", "신청자이메일", "결과", "비고"])
            for item, status, note in results:
                writer.writerow([
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    item["name"], item["doc_type"], item["period"], item["email"], status, note,
                ])


def open_window(parent):
    win = App(parent)
    return win
