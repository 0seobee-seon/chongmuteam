"""
계약서 색인표 자동 업데이트 — 허브용 모듈
핵심 로직(파싱/엑셀 반영)은 원본 스크립트를 그대로 불러와 사용한다.
원본: desktop wrok/BOOKING/docs/계약서_색인_자동업데이트.py

연도는 대상 폴더명(예: "2027")에서 자동 인식되므로, 해가 바뀌어도
이 모듈이나 원본 스크립트를 고칠 필요 없이 새 연도 폴더만 선택하면 된다.
"""

import os
import io
import contextlib
import importlib.util
from datetime import date

import tkinter as tk
from tkinter import filedialog, messagebox

_CORE_SCRIPT = (
    r"C:\Users\user\Desktop\AI 영섭 작업기록\desktop wrok\BOOKING\docs"
    r"\계약서_색인_자동업데이트.py"
)
_BASE_FOLDER = r"D:\D\김영섭\계약서\부서별계약서"


def _load_core():
    spec = importlib.util.spec_from_file_location("계약서색인_core", _CORE_SCRIPT)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


class App(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("계약서 색인표 자동 업데이트")
        self.resizable(False, False)
        self.geometry("640x520")
        self._core = None
        self._build_ui()
        self._load_core_module()

    def _load_core_module(self):
        try:
            self._core = _load_core()
        except Exception as e:
            messagebox.showerror(
                "오류",
                f"핵심 스크립트를 불러올 수 없습니다:\n{_CORE_SCRIPT}\n\n{e}",
            )

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        frm_folder = tk.LabelFrame(self, text="대상 폴더 (연도 폴더)", font=("맑은 고딕", 10, "bold"))
        frm_folder.pack(fill="x", **pad)

        default_folder = os.path.join(_BASE_FOLDER, str(date.today().year))
        self.folder_var = tk.StringVar(value=default_folder)
        tk.Entry(frm_folder, textvariable=self.folder_var, width=56).pack(side="left", padx=6, pady=6)
        tk.Button(frm_folder, text="찾아보기", command=self._browse, width=10).pack(side="left", padx=4)

        tk.Label(
            self,
            text="※ 폴더명이 4자리 연도(예: 2027)면 그 연도로 자동 인식합니다. 새해에는 새 연도 폴더만 선택하세요.",
            font=("맑은 고딕", 9), fg="gray", justify="left",
        ).pack(anchor="w", padx=14)

        frm_btn = tk.Frame(self)
        frm_btn.pack(pady=10)
        tk.Button(
            frm_btn, text="미리보기 (Dry-run)", command=self._run_dry,
            font=("맑은 고딕", 11, "bold"), width=16, height=2,
        ).pack(side="left", padx=6)
        tk.Button(
            frm_btn, text="실제 반영", command=self._run_real,
            font=("맑은 고딕", 11, "bold"), bg="#1F497D", fg="white", width=12, height=2,
        ).pack(side="left", padx=6)
        tk.Button(
            frm_btn, text="닫기", command=self.destroy,
            font=("맑은 고딕", 11), width=8, height=2,
        ).pack(side="left", padx=6)

        frm_result = tk.LabelFrame(self, text="결과", font=("맑은 고딕", 10, "bold"))
        frm_result.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.text = tk.Text(frm_result, font=("Consolas", 10), wrap="word", state="disabled")
        self.text.pack(fill="both", expand=True, padx=6, pady=6)

    def _browse(self):
        path = filedialog.askdirectory(title="계약서 연도 폴더 선택", initialdir=self.folder_var.get())
        if path:
            self.folder_var.set(path)

    def _log(self, text, clear=False):
        self.text.config(state="normal")
        if clear:
            self.text.delete("1.0", "end")
        self.text.insert("end", text + "\n")
        self.text.see("end")
        self.text.config(state="disabled")
        self.update_idletasks()

    def _resolve(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("알림", "올바른 폴더를 선택하세요.")
            return None
        if self._core is None:
            self._load_core_module()
        if self._core is None:
            return None
        year = self._core.resolve_year(folder)
        yy = str(year)[2:]
        file_pairs = self._core.build_file_pairs(year)
        return folder, year, yy, file_pairs

    def _run_dry(self):
        resolved = self._resolve()
        if not resolved:
            return
        folder, year, yy, file_pairs = resolved

        self._log(f"폴더: {folder}", clear=True)
        self._log(f"대상 연도: {year} (계약일자 접두 '{yy}.')")
        self._log("-" * 56)

        try:
            results = self._core.collect_dry_run(folder, file_pairs, yy)
        except Exception as e:
            self._log(f"[오류] {e}")
            messagebox.showerror("오류", f"미리보기 중 오류가 발생했습니다:\n{e}")
            return

        total_new = 0
        for r in results:
            self._log(f"\n📋 {r['xlsx_name']}")
            if r["error"]:
                self._log(f"   ⚠️  {r['error']}")
                continue

            new = r["new_contracts"]
            total_new += len(new)
            if not new:
                self._log(f"   HWP {r['hwp_count']}건 / Excel {r['excel_max']}번까지 입력 → 신규 0건 추가 예정")
            else:
                self._log(f"   HWP {r['hwp_count']}건 / Excel {r['excel_max']}번까지 입력 → 신규 {len(new)}건 추가 예정")
                for c in new[:3]:
                    self._log(f"     [{c['seq']}] {c['용역명'][:30]}")
                if len(new) > 3:
                    self._log(f"     ... 외 {len(new) - 3}건")

        self._log("-" * 56)
        self._log(f"미리보기 완료 — 총 신규 예정 {total_new}건")

    def _run_real(self):
        resolved = self._resolve()
        if not resolved:
            return
        folder, year, yy, file_pairs = resolved

        if not messagebox.askyesno(
            "실행 확인",
            f"'{folder}' 폴더의 {year}년 계약을 Excel에 실제로 반영합니다.\n"
            "(반영 전 Excel 파일은 자동 백업됩니다)\n\n계속할까요?",
        ):
            return

        self._log(f"폴더: {folder}", clear=True)
        self._log(f"대상 연도: {year} (계약일자 접두 '{yy}.')")
        self._log("-" * 56)

        buf = io.StringIO()
        try:
            with contextlib.redirect_stdout(buf):
                total = self._core.run_all(folder, file_pairs, yy)
        except Exception as e:
            self._log(buf.getvalue())
            self._log(f"[오류] {e}")
            messagebox.showerror("오류", f"실행 중 오류가 발생했습니다:\n{e}")
            return

        self._log(buf.getvalue())
        self._log("-" * 56)
        self._log(f"완료 — 총 {total}건 추가")
        messagebox.showinfo("완료", f"총 {total}건이 Excel에 추가되었습니다.")


def open_window(parent):
    win = App(parent)
    return win
