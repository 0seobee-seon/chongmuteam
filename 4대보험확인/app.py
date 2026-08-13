"""4대보험료 확인 프로그램 — 독립 실행용 UI.

기존 '총무팀 자동화 허브'의 mod_*.py 모듈들과 같은 스타일(tkinter)로 작성했다.
검증이 끝나면 mod_4대보험.py로 옮겨 hub.py의 PROGRAMS 목록에 편입할 예정이다.
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from matcher import INSURANCE_ITEMS, classify_errors, merge_records
from parsers import (
    parse_employment_insurance,
    parse_health_insurance,
    parse_national_pension,
    parse_payroll,
)
from report import build_report

# (표시 이름, 내부 키, 필수 여부). 급여대장은 비교 기준이 되므로 필수이고,
# 나머지 공단 자료는 아직 발급되지 않은 달에도 부분적으로 확인할 수 있도록 선택사항이다.
FILE_SLOTS = [
    ("급여대장", "payroll", True),
    ("건강보험공단", "health", False),
    ("국민연금공단", "pension", False),
    ("고용보험(근로복지공단)", "employment", False),
]

# 공단 파일 키 -> 그 파일이 있어야 확인할 수 있는 보험 항목.
FILE_TO_ITEMS = {
    "health": ["건강보험", "장기요양보험"],
    "pension": ["국민연금"],
    "employment": ["고용보험"],
}


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("4대보험료 확인")
        self.geometry("640x520")

        self.file_paths = {}
        self.path_labels = {}
        self.merged_records = []
        self.all_errors = []
        self.귀속년월 = None

        self._build_file_selectors()
        self._build_run_button()
        self._build_result_area()
        self._build_save_button()

    def _build_file_selectors(self):
        frame = tk.Frame(self, padx=12, pady=12)
        frame.pack(fill="x")

        for label_text, key, required in FILE_SLOTS:
            suffix = " (필수)" if required else " (선택)"
            row = tk.Frame(frame)
            row.pack(fill="x", pady=4)

            tk.Label(row, text=label_text + suffix, width=26, anchor="w").pack(side="left")
            path_label = tk.Label(row, text="(선택 안 됨)", anchor="w", fg="#888")
            path_label.pack(side="left", fill="x", expand=True)
            self.path_labels[key] = path_label

            button = tk.Button(row, text="파일 선택", command=lambda k=key: self._select_file(k))
            button.pack(side="right")

    def _select_file(self, key):
        path = filedialog.askopenfilename(filetypes=[("Excel 파일", "*.xlsx")])
        if not path:
            return
        self.file_paths[key] = path
        self.path_labels[key].config(text=os.path.basename(path), fg="black")

    def _build_run_button(self):
        tk.Button(self, text="확인 실행", command=self._run_check, height=2).pack(fill="x", padx=12, pady=8)

    def _build_result_area(self):
        frame = tk.Frame(self, padx=12)
        frame.pack(fill="both", expand=True)

        self.summary_label = tk.Label(frame, text="아직 실행하지 않았습니다.", anchor="w", justify="left")
        self.summary_label.pack(fill="x")

        columns = ("사원번호", "성명", "보험종류", "급여대장금액", "공단금액", "오류유형")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", height=12)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=90)
        self.tree.pack(fill="both", expand=True, pady=8)

    def _build_save_button(self):
        self.save_button = tk.Button(self, text="엑셀로 저장", command=self._save_report, state="disabled")
        self.save_button.pack(fill="x", padx=12, pady=(0, 12))

    def _run_check(self):
        missing_required = [
            label for label, key, required in FILE_SLOTS
            if required and key not in self.file_paths
        ]
        if missing_required:
            messagebox.showerror("오류", f"다음 파일을 선택해주세요:\n{', '.join(missing_required)}")
            return

        try:
            self.귀속년월, payroll_records = parse_payroll(self.file_paths["payroll"])

            health_records = []
            pension_records = []
            employment_records = []
            active_items = []

            if "health" in self.file_paths:
                health_records = parse_health_insurance(self.file_paths["health"])
                active_items += FILE_TO_ITEMS["health"]
            if "pension" in self.file_paths:
                pension_records = parse_national_pension(self.file_paths["pension"])
                active_items += FILE_TO_ITEMS["pension"]
            if "employment" in self.file_paths:
                employment_records = parse_employment_insurance(self.file_paths["employment"])
                active_items += FILE_TO_ITEMS["employment"]
        except Exception as e:
            messagebox.showerror("오류", f"파일을 읽는 중 문제가 발생했습니다:\n{e}")
            return

        if not active_items:
            messagebox.showwarning("안내", "공단 자료를 하나도 선택하지 않아 비교할 항목이 없습니다.")

        self.merged_records = merge_records(
            payroll_records, health_records, pension_records, employment_records,
        )
        self.all_errors = [
            error
            for record in self.merged_records
            for error in classify_errors(record, self.귀속년월, items=active_items)
        ]

        self._render_results(active_items)
        self.save_button.config(state="normal")

    def _render_results(self, active_items):
        self.tree.delete(*self.tree.get_children())
        for error in self.all_errors:
            self.tree.insert("", "end", values=(
                error["사원번호"], error["성명"], error["보험종류"],
                error["급여대장금액"], error["공단금액"], error["오류유형"],
            ))

        skipped_items = [item for item in INSURANCE_ITEMS if item not in active_items]
        skipped_note = f" / 미확인 항목(자료 없음): {', '.join(skipped_items)}" if skipped_items else ""

        self.summary_label.config(
            text=(
                f"귀속년월: {self.귀속년월} / 전체 인원: {len(self.merged_records)}명 / "
                f"오류 건수: {len(self.all_errors)}건{skipped_note}"
            )
        )

    def _save_report(self):
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"4대보험_확인결과_{self.귀속년월}.xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
        )
        if not output_path:
            return

        build_report(self.merged_records, self.all_errors, self.귀속년월, output_path)
        messagebox.showinfo("완료", f"리포트를 저장했습니다:\n{output_path}")


if __name__ == "__main__":
    App().mainloop()
