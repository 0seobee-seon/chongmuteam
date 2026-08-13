"""일용직 근로내용확인신고서 생성 — 허브용 모듈.

일용직 급여자료를 읽어 근로복지공단 전자신고용 서식을 현장별로 채운다.
신고는 사업장(현장) 단위이므로, 현장이 섞인 자료를 넣어도 현장별로 파일이 나뉜다.

핵심 로직은 mod_근로내용신고_core.py 에 있다.
"""

import json
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from mod_근로내용신고_core import (
    DEFAULT_PARAMS,
    INSURANCE_OPTIONS,
    RESIGN_OPTIONS,
    build_reports,
    group_by_site,
    load_people,
    summarize_sites,
)

TEMPLATE_ASSET = os.path.join("assets", "근로내용확인신고_전자신고용 양식.xlsx")
CONFIG_NAME = "config_근로내용신고.json"


def _bundle_dir():
    """PyInstaller onefile 실행 중이면 압축 해제 폴더, 아니면 이 파일이 있는 폴더."""
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))


def _settings_dir():
    """설정을 저장할 폴더. EXE로 실행 중이면 EXE가 있는 폴더에 남겨야 재실행 시 살아남는다."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def bundled_template():
    """함께 배포된 신고서 양식 경로. 없으면 None."""
    path = os.path.join(_bundle_dir(), TEMPLATE_ASSET)
    return path if os.path.exists(path) else None


def load_settings():
    path = os.path.join(_settings_dir(), CONFIG_NAME)
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}


def save_settings(settings):
    path = os.path.join(_settings_dir(), CONFIG_NAME)
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except OSError:
        pass    # 설정 저장 실패가 신고서 생성을 막을 이유는 없다.


class App(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("일용직 근로내용확인신고서 생성")
        self.geometry("760x680")

        self.ledger_paths = []
        self.settings = load_settings()

        self._build_file_area()
        self._build_param_area()
        self._build_output_area()
        self._build_buttons()
        self._build_result_area()
        self._enable_drop()

    # ------------------------------------------------------------------ UI

    def _build_file_area(self):
        frame = tk.LabelFrame(self, text="일용직 급여자료", font=("맑은 고딕", 10, "bold"))
        frame.pack(fill="both", expand=True, padx=12, pady=(12, 6))

        tk.Label(
            frame,
            text="[일용직 급여 지급명세서] 형식(1~31일 ● 출근표 포함)을 현장별로 넣으세요. "
                 "여러 개를 한 번에 처리합니다.",
            font=("맑은 고딕", 9), fg="gray", justify="left", wraplength=700,
        ).pack(anchor="w", padx=8, pady=(6, 4))

        list_row = tk.Frame(frame)
        list_row.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.file_list = tk.Listbox(list_row, height=6, font=("맑은 고딕", 9))
        self.file_list.pack(side="left", fill="both", expand=True)

        scroll = tk.Scrollbar(list_row, orient="vertical", command=self.file_list.yview)
        scroll.pack(side="left", fill="y")
        self.file_list.config(yscrollcommand=scroll.set)

        buttons = tk.Frame(list_row)
        buttons.pack(side="left", fill="y", padx=(8, 0))
        tk.Button(buttons, text="파일 추가", width=10, command=self._add_files).pack(pady=2)
        tk.Button(buttons, text="선택 제거", width=10, command=self._remove_selected).pack(pady=2)
        tk.Button(buttons, text="전체 지우기", width=10, command=self._clear_files).pack(pady=2)

    def _build_param_area(self):
        frame = tk.LabelFrame(self, text="신고 항목", font=("맑은 고딕", 10, "bold"))
        frame.pack(fill="x", padx=12, pady=6)

        saved = self.settings.get("params", {})

        def value_of(key):
            return saved.get(key, DEFAULT_PARAMS[key])

        row1 = tk.Frame(frame)
        row1.pack(fill="x", padx=8, pady=(8, 4))

        tk.Label(row1, text="직종코드", width=10, anchor="w").pack(side="left")
        self.jikjong_var = tk.StringVar(value=str(value_of("jikjong")))
        tk.Entry(row1, textvariable=self.jikjong_var, width=8).pack(side="left")

        tk.Label(row1, text="보험구분", width=10, anchor="w").pack(side="left", padx=(16, 0))
        self.insurance_var = tk.StringVar(value=self._label_for(INSURANCE_OPTIONS, value_of("insurance")))
        ttk.Combobox(
            row1, textvariable=self.insurance_var, width=22, state="readonly",
            values=[f"{code} · {name}" for code, name in INSURANCE_OPTIONS],
        ).pack(side="left")

        tk.Label(row1, text="일평균근로시간", anchor="w").pack(side="left", padx=(16, 4))
        self.hours_var = tk.StringVar(value=str(value_of("avg_hours")))
        tk.Entry(row1, textvariable=self.hours_var, width=5).pack(side="left")

        row2 = tk.Frame(frame)
        row2.pack(fill="x", padx=8, pady=(0, 8))

        tk.Label(row2, text="이직사유", width=10, anchor="w").pack(side="left")
        self.resign_var = tk.StringVar(value=self._label_for(RESIGN_OPTIONS, value_of("resign")))
        ttk.Combobox(
            row2, textvariable=self.resign_var, width=40, state="readonly",
            values=[f"{code} · {name}" for code, name in RESIGN_OPTIONS],
        ).pack(side="left")

        self.nts_var = tk.BooleanVar(value=bool(value_of("nts")))
        tk.Checkbutton(
            row2, text="국세청 일용근로소득도 함께 신고 (지급월·소득세 기재)",
            variable=self.nts_var, font=("맑은 고딕", 9),
        ).pack(side="left", padx=(16, 0))

    def _build_output_area(self):
        frame = tk.LabelFrame(self, text="신고서 양식 · 저장 위치", font=("맑은 고딕", 10, "bold"))
        frame.pack(fill="x", padx=12, pady=6)

        row1 = tk.Frame(frame)
        row1.pack(fill="x", padx=8, pady=(8, 4))
        tk.Label(row1, text="양식 파일", width=10, anchor="w").pack(side="left")
        self.template_var = tk.StringVar(value=self.settings.get("template") or bundled_template() or "")
        tk.Entry(row1, textvariable=self.template_var).pack(side="left", fill="x", expand=True)
        tk.Button(row1, text="찾아보기", width=10, command=self._browse_template).pack(side="left", padx=(6, 0))

        row2 = tk.Frame(frame)
        row2.pack(fill="x", padx=8, pady=(0, 8))
        tk.Label(row2, text="저장 폴더", width=10, anchor="w").pack(side="left")
        self.outdir_var = tk.StringVar(value=self.settings.get("outdir", ""))
        tk.Entry(row2, textvariable=self.outdir_var).pack(side="left", fill="x", expand=True)
        tk.Button(row2, text="찾아보기", width=10, command=self._browse_outdir).pack(side="left", padx=(6, 0))

        tk.Label(
            frame,
            text="※ 저장 폴더를 비워두면 급여자료가 있는 폴더에 저장합니다. "
                 "파일명은 근로내용확인신고_{현장명}_{YYYYMM}.xlsx",
            font=("맑은 고딕", 9), fg="gray", justify="left", wraplength=700,
        ).pack(anchor="w", padx=8, pady=(0, 8))

    def _build_buttons(self):
        frame = tk.Frame(self)
        frame.pack(pady=(0, 8))
        tk.Button(
            frame, text="미리보기", command=self._run_preview,
            font=("맑은 고딕", 11, "bold"), width=14, height=2,
        ).pack(side="left", padx=6)
        tk.Button(
            frame, text="신고서 생성", command=self._run_build,
            font=("맑은 고딕", 11, "bold"), bg="#1F497D", fg="white", width=14, height=2,
        ).pack(side="left", padx=6)
        tk.Button(
            frame, text="닫기", command=self.destroy,
            font=("맑은 고딕", 11), width=8, height=2,
        ).pack(side="left", padx=6)

    def _build_result_area(self):
        frame = tk.LabelFrame(self, text="결과", font=("맑은 고딕", 10, "bold"))
        frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.text = tk.Text(frame, font=("Consolas", 10), wrap="word", height=10, state="disabled")
        self.text.pack(fill="both", expand=True, padx=6, pady=6)

    def _enable_drop(self):
        """tkinterdnd2가 있으면 파일 목록에 끌어다 놓을 수 있게 한다."""
        try:
            from tkinterdnd2 import DND_FILES
            self.file_list.drop_target_register(DND_FILES)
            self.file_list.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            pass

    def _on_drop(self, event):
        paths = [p for p in self.tk.splitlist(event.data) if p.lower().endswith((".xlsx", ".xls"))]
        self._append_paths(paths)

    # -------------------------------------------------------------- 파일 관리

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="일용직 급여자료 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xls")],
        )
        self._append_paths(paths)

    def _append_paths(self, paths):
        added = 0
        for path in paths:
            if path not in self.ledger_paths:
                self.ledger_paths.append(path)
                self.file_list.insert("end", os.path.basename(path))
                added += 1
        if added and not self.outdir_var.get().strip():
            self.outdir_var.set(os.path.dirname(self.ledger_paths[0]))

    def _remove_selected(self):
        for index in sorted(self.file_list.curselection(), reverse=True):
            self.file_list.delete(index)
            del self.ledger_paths[index]

    def _clear_files(self):
        self.file_list.delete(0, "end")
        self.ledger_paths.clear()

    def _browse_template(self):
        path = filedialog.askopenfilename(
            title="근로내용확인신고 전자신고용 양식 선택",
            filetypes=[("Excel 파일", "*.xlsx")],
        )
        if path:
            self.template_var.set(path)

    def _browse_outdir(self):
        path = filedialog.askdirectory(title="저장 폴더 선택", initialdir=self.outdir_var.get() or None)
        if path:
            self.outdir_var.set(path)

    # ------------------------------------------------------------------ 실행

    @staticmethod
    def _label_for(options, code):
        code = str(code)
        return next((f"{c} · {n}" for c, n in options if c == code), f"{options[0][0]} · {options[0][1]}")

    @staticmethod
    def _code_from(label):
        return label.split("·")[0].strip()

    def _log(self, text, clear=False):
        self.text.config(state="normal")
        if clear:
            self.text.delete("1.0", "end")
        self.text.insert("end", text + "\n")
        self.text.see("end")
        self.text.config(state="disabled")
        self.update_idletasks()

    def _collect_params(self):
        hours_text = self.hours_var.get().strip()
        if not hours_text.isdigit() or not 1 <= int(hours_text) <= 24:
            messagebox.showwarning("알림", "일평균근로시간은 1~24 사이의 숫자로 입력하세요.")
            return None

        jikjong = self.jikjong_var.get().strip()
        if jikjong and not jikjong.isdigit():
            messagebox.showwarning("알림", "직종코드는 숫자로 입력하세요. (비워두면 공란으로 신고됩니다)")
            return None

        return {
            "jikjong": jikjong,
            "insurance": self._code_from(self.insurance_var.get()),
            "resign": self._code_from(self.resign_var.get()),
            "avg_hours": int(hours_text),
            "nts": self.nts_var.get(),
        }

    def _validate_inputs(self):
        if not self.ledger_paths:
            messagebox.showwarning("알림", "일용직 급여자료를 먼저 추가하세요.")
            return None

        template = self.template_var.get().strip()
        if not template or not os.path.isfile(template):
            messagebox.showwarning("알림", "신고서 양식 파일을 선택하세요.")
            return None

        params = self._collect_params()
        if params is None:
            return None

        outdir = self.outdir_var.get().strip() or os.path.dirname(self.ledger_paths[0])
        return template, outdir, params

    def _report_warnings(self, summary):
        if summary["missing_days"]:
            names = ", ".join(summary["missing_days"][:5])
            more = f" 외 {len(summary['missing_days']) - 5}명" if len(summary["missing_days"]) > 5 else ""
            self._log(f"  [확인필요] 1~31일 출근 칸 비어 있음 {len(summary['missing_days'])}명 — {names}{more}")
            self._log("             [일용직 급여 지급명세서] 형식 파일을 넣어야 일자가 채워집니다.")
        if summary["day_mismatch"]:
            self._log(f"  [확인필요] 일자표시 개수 != 근로일수: {', '.join(summary['day_mismatch'])}")
        if summary["foreigners"]:
            self._log(f"  [확인필요] 외국인 {len(summary['foreigners'])}명 — 국적코드·체류자격코드 수기 입력:")
            self._log(f"             {', '.join(summary['foreigners'])}")

    def _run_preview(self):
        resolved = self._validate_inputs()
        if not resolved:
            return
        _, outdir, params = resolved

        self._log("미리보기 — 파일은 만들지 않습니다.", clear=True)
        self._log(
            f"  적용값: 보험구분 {params['insurance']} / 직종 {params['jikjong'] or '공란'} / "
            f"이직사유 {params['resign']} / 일평균 {params['avg_hours']}시간 / "
            f"국세청 {'Y' if params['nts'] else '미기재'}"
        )
        try:
            people, loaded, duplicates = load_people(self.ledger_paths)
        except Exception as e:
            self._log(f"[오류] {e}")
            messagebox.showerror("오류", f"자료를 읽는 중 문제가 발생했습니다:\n{e}")
            return

        for entry in loaded:
            self._log(f"  [읽음] {os.path.basename(entry['path'])} — 형식 {entry['format']}, {entry['count']}명")
        if duplicates:
            self._log(f"  [안내] 같은 현장·같은 사람 {duplicates}건이 여러 파일에 겹쳐 하나로 합쳤습니다.")

        if not people:
            self._log("[오류] 읽어들인 인원이 없습니다.")
            return

        self._log("-" * 60)
        summaries = summarize_sites(group_by_site(people))
        for summary in summaries:
            self._log(f"{summary['site']} ({summary['paymonth'] or '연월미상'})")
            self._log(f"  인원 {summary['count']}명 / 근로일수 {summary['days']}일 / "
                      f"임금총액 {summary['wage']:,}원")
            self._report_warnings(summary)
            self._log("")

        self._log("-" * 60)
        self._log(f"현장 {len(summaries)}곳 / 총 {sum(s['count'] for s in summaries)}명 / "
                  f"{sum(s['wage'] for s in summaries):,}원")
        self._log(f"생성 시 저장 위치: {outdir}")

    def _run_build(self):
        resolved = self._validate_inputs()
        if not resolved:
            return
        template, outdir, params = resolved

        if not messagebox.askyesno(
            "생성 확인",
            f"현장별 신고서를 만듭니다.\n\n저장 위치: {outdir}\n\n"
            "같은 이름의 파일이 있으면 덮어씁니다. 계속할까요?",
        ):
            return

        self._log(f"저장 위치: {outdir}", clear=True)
        self._log("-" * 60)

        try:
            results, loaded, duplicates = build_reports(self.ledger_paths, template, outdir, params)
        except Exception as e:
            self._log(f"[오류] {e}")
            messagebox.showerror("오류", f"신고서 생성 중 문제가 발생했습니다:\n{e}")
            return

        for entry in loaded:
            self._log(f"  [읽음] {os.path.basename(entry['path'])} — 형식 {entry['format']}, {entry['count']}명")
        if duplicates:
            self._log(f"  [안내] 같은 현장·같은 사람 {duplicates}건이 여러 파일에 겹쳐 하나로 합쳤습니다.")
        self._log("-" * 60)

        for result in results:
            self._log(f"{result['site']} ({result['paymonth'] or '연월미상'})")
            self._log(f"  인원 {result['count']}명 / 근로일수 {result['days']}일 / "
                      f"임금총액 {result['wage']:,}원")
            self._log(f"  저장: {os.path.basename(result['path'])}")
            self._report_warnings(result)
            self._log("")

        self._log("-" * 60)
        self._log(f"완료 — 현장 {len(results)}곳, 총 {sum(r['count'] for r in results)}명")

        self.settings["params"] = params
        self.settings["template"] = template
        self.settings["outdir"] = outdir
        save_settings(self.settings)

        needs_check = sum(len(r["foreigners"]) for r in results)
        extra = f"\n\n외국인 {needs_check}명은 국적코드·체류자격코드를 직접 채워야 합니다." if needs_check else ""
        messagebox.showinfo(
            "완료",
            f"현장 {len(results)}곳의 신고서를 만들었습니다.\n{outdir}{extra}",
        )


def open_window(parent):
    win = App(parent)
    return win
