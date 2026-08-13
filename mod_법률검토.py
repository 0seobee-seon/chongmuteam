"""
법률 검토 도우미 — 허브용 모듈
국가법령정보 공동활용 API로 법령을 검색하고 조문을 확인한다.
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

import law_api
from law_config import load_oc, save_oc

ACCENT = "#1F497D"

SHORTCUTS = {
    "건강보험": ["국민건강보험법", "국민건강보험법 시행령"],
    "국민연금": ["국민연금법", "국민연금법 시행령"],
    "고용노동": ["근로기준법", "근로자퇴직급여 보장법", "최저임금법", "고용보험법", "산업재해보상보험법"],
    "부동산": ["부동산등기법", "공인중개사법", "주택임대차보호법"],
}

MAX_RECENT = 10


class App(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("법률 검토 도우미")
        self.resizable(False, False)
        self.geometry("760x520")

        self._oc = load_oc()
        self._results = []   # 최근 검색 결과 (mst -> 표시용 데이터)
        self._recent = []    # 최근 조회한 (mst, 법령명) 목록, 세션 내 최대 MAX_RECENT개

        self._build_ui()

        if not self._oc:
            self.after(100, self._prompt_oc)

    def _prompt_oc(self):
        oc = simpledialog.askstring(
            "API 키 입력",
            "국가법령정보 API 키(OC)를 입력하세요.\n(open.law.go.kr에서 발급받은 이메일 아이디)",
            parent=self,
        )
        if oc:
            oc = oc.strip()
            save_oc(oc)
            self._oc = oc
        else:
            messagebox.showwarning("알림", "OC 키가 없으면 검색을 사용할 수 없습니다.")

    def _build_ui(self):
        frm_short = tk.LabelFrame(self, text="업무 영역 바로가기", font=("맑은 고딕", 10, "bold"))
        frm_short.pack(fill="x", padx=10, pady=(10, 4))

        for domain in SHORTCUTS:
            tk.Button(
                frm_short, text=domain, font=("맑은 고딕", 10),
                command=lambda d=domain: self._show_shortcut_menu(d),
            ).pack(side="left", padx=6, pady=8)

        frm_search = tk.LabelFrame(self, text="키워드 검색", font=("맑은 고딕", 10, "bold"))
        frm_search.pack(fill="x", padx=10, pady=4)

        self.keyword_var = tk.StringVar()
        entry = tk.Entry(frm_search, textvariable=self.keyword_var, width=40, font=("맑은 고딕", 10))
        entry.pack(side="left", padx=8, pady=8)
        entry.bind("<Return>", lambda e: self._run_search(self.keyword_var.get().strip()))
        tk.Button(
            frm_search, text="검색", command=lambda: self._run_search(self.keyword_var.get().strip()),
            bg=ACCENT, fg="white", font=("맑은 고딕", 10, "bold"),
        ).pack(side="left", padx=4)

        frm_body = tk.Frame(self)
        frm_body.pack(fill="both", expand=True, padx=10, pady=4)

        frm_list = tk.LabelFrame(frm_body, text="검색 결과 (더블클릭으로 조문 보기)", font=("맑은 고딕", 10, "bold"))
        frm_list.pack(side="left", fill="both", expand=True)

        columns = ("법령명", "공포일자", "시행일자", "소관부처")
        self.tree = ttk.Treeview(frm_list, columns=columns, show="headings", height=14)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=160 if col == "법령명" else 90)
        self.tree.pack(fill="both", expand=True, padx=6, pady=6)
        self.tree.bind("<Double-1>", lambda e: self._open_selected())

        frm_recent = tk.LabelFrame(frm_body, text="최근 조회", font=("맑은 고딕", 10, "bold"), width=180)
        frm_recent.pack(side="left", fill="y", padx=(8, 0))
        frm_recent.pack_propagate(False)

        self.recent_list = tk.Listbox(frm_recent, font=("맑은 고딕", 9))
        self.recent_list.pack(fill="both", expand=True, padx=6, pady=6)
        self.recent_list.bind("<Double-1>", self._open_recent)

        self.status_var = tk.StringVar(value="검색어를 입력하거나 업무 영역 바로가기를 눌러주세요.")
        tk.Label(self, textvariable=self.status_var, font=("맑은 고딕", 9), fg="#555").pack(pady=(0, 8))

    def _show_shortcut_menu(self, domain):
        menu = tk.Menu(self, tearoff=0)
        for law_name in SHORTCUTS[domain]:
            menu.add_command(label=law_name, command=lambda n=law_name: self._run_search(n))
        menu.tk_popup(self.winfo_pointerx(), self.winfo_pointery())

    def _run_search(self, keyword):
        if not keyword:
            messagebox.showwarning("알림", "검색어를 입력하세요.")
            return
        if not self._oc:
            self._prompt_oc()
            if not self._oc:
                return

        self.status_var.set("검색 중...")
        self.update()
        try:
            self._results = law_api.search_law(keyword, self._oc)
        except Exception as e:
            messagebox.showerror("오류", f"법제처 서버에 연결할 수 없습니다.\n{e}")
            self.status_var.set("검색 실패")
            return

        self.tree.delete(*self.tree.get_children())
        if not self._results:
            self.status_var.set("검색 결과가 없습니다.")
            return

        for law in self._results:
            self.tree.insert(
                "", "end", iid=law["mst"],
                values=(law["법령명"], law["공포일자"], law["시행일자"], law["소관부처"]),
            )
        self.status_var.set(f"{len(self._results)}건 검색됨")

    def _open_selected(self):
        selection = self.tree.selection()
        if not selection:
            return
        mst = selection[0]
        law_name = self.tree.set(mst, "법령명")
        self._open_law_detail(mst, law_name)

    def _open_recent(self, event):
        selection = self.recent_list.curselection()
        if not selection:
            return
        mst, law_name = self._recent[selection[0]]
        self._open_law_detail(mst, law_name)

    def _open_law_detail(self, mst, law_name):
        if not self._oc:
            self._prompt_oc()
            if not self._oc:
                return
        try:
            detail = law_api.get_law_text(mst, self._oc)
        except Exception as e:
            messagebox.showerror("오류", f"법제처 서버에 연결할 수 없습니다.\n{e}")
            return

        self._add_recent(mst, law_name)
        LawDetailWindow(self, detail)

    def _add_recent(self, mst, law_name):
        self._recent = [(m, n) for m, n in self._recent if m != mst]
        self._recent.insert(0, (mst, law_name))
        self._recent = self._recent[:MAX_RECENT]
        self.recent_list.delete(0, tk.END)
        for m, n in self._recent:
            self.recent_list.insert(tk.END, n)


class LawDetailWindow(tk.Toplevel):
    def __init__(self, parent, detail):
        super().__init__(parent)
        self.title(detail["법령명"] or "조문 보기")
        self.geometry("640x560")
        self._search_start = "1.0"
        self._build_ui(detail)

    def _build_ui(self, detail):
        header = f"{detail['법령명']}  (공포 {detail['공포일자']} / 시행 {detail['시행일자']})"
        tk.Label(self, text=header, font=("맑은 고딕", 11, "bold"), fg=ACCENT).pack(anchor="w", padx=10, pady=(10, 4))

        frm_find = tk.Frame(self)
        frm_find.pack(fill="x", padx=10)
        self.find_var = tk.StringVar()
        tk.Entry(frm_find, textvariable=self.find_var, width=30).pack(side="left", padx=(0, 6))
        tk.Button(frm_find, text="찾기", command=self._find_next).pack(side="left")
        tk.Button(frm_find, text="선택 복사", command=self._copy_selection).pack(side="left", padx=(6, 0))

        self.text = tk.Text(self, wrap="word", font=("맑은 고딕", 10))
        self.text.pack(fill="both", expand=True, padx=10, pady=8)

        body_lines = []
        for art in detail["조문목록"]:
            title = f"제{art['조문번호']}조" + (f"({art['조문제목']})" if art["조문제목"] else "")
            body_lines.append(art["조문내용"] or title)
        self.text.insert("1.0", "\n\n".join(body_lines) if body_lines else "본문 내용이 없습니다.")
        self.text.config(state="disabled")

    def _find_next(self):
        keyword = self.find_var.get().strip()
        if not keyword:
            return
        self.text.tag_remove("found", "1.0", tk.END)
        pos = self.text.search(keyword, self._search_start, tk.END)
        if not pos:
            self._search_start = "1.0"
            pos = self.text.search(keyword, self._search_start, tk.END)
        if not pos:
            messagebox.showinfo("찾기", "찾는 내용이 없습니다.")
            return
        end_pos = f"{pos}+{len(keyword)}c"
        self.text.tag_add("found", pos, end_pos)
        self.text.tag_config("found", background="yellow")
        self.text.see(pos)
        self._search_start = end_pos

    def _copy_selection(self):
        try:
            selected = self.text.get("sel.first", "sel.last")
        except tk.TclError:
            selected = self.text.get("1.0", tk.END)
        self.clipboard_clear()
        self.clipboard_append(selected)


def open_window(parent):
    win = App(parent)
    return win
