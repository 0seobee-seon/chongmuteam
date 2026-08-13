"""
총무팀 자동화 허브
"""

import tkinter as tk
from tkinter import messagebox

try:
    from tkinterdnd2 import TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

ACCENT   = "#1F497D"
BTN_BG   = "#EEF3FA"
BTN_HVR  = "#D6E4F7"
BTN_BRD  = "#A8C4E0"
FG_TITLE = "#1F497D"
FG_DESC  = "#555555"
CAT_FG   = "#1F497D"

# 카테고리 순서. 새 프로그램을 추가할 때는 이 중 하나의 category 값을 쓰거나,
# 새 카테고리가 필요하면 여기에 이름을 추가하면 된다.
CATEGORY_ORDER = ["급여·근태", "인사·보험", "법무·계약", "재무·경비"]

PROGRAMS = [
    {
        "key":      "야근수당",
        "icon":     "📋",
        "title":    "야근수당 내역 생성",
        "desc":     "1번 야근현황 파일에서 수당 내역을 자동 생성합니다. 여러 부서 파일을 한 번에 배치 처리할 수 있습니다.",
        "category": "급여·근태",
    },
    {
        "key":      "야근현황",
        "icon":     "🕐",
        "title":    "출퇴근 기록 입력",
        "desc":     "출퇴근 기록부를 읽어 야근현황 양식에 자동으로 입력하거나 새 양식을 생성합니다.",
        "category": "급여·근태",
    },
    {
        "key":      "4대보험",
        "icon":     "🏥",
        "title":    "4대보험료 확인",
        "desc":     "급여대장과 공단 자료를 대조해 4대보험료 입력 오류와 퇴사자 부과 여부를 확인합니다.",
        "category": "인사·보험",
    },
    {
        "key":      "증명서발급",
        "icon":     "📨",
        "title":    "Wehago 증명서 자동발급",
        "desc":     "원천징수영수증·급여명세서 신청 정보를 여러 건 입력하면 Wehago에 로그인해 조회하고 이메일로 자동 발송합니다.",
        "category": "인사·보험",
    },
    {
        "key":      "근로내용신고",
        "icon":     "🧾",
        "title":    "일용직 근로내용확인신고서 생성",
        "desc":     "일용직 급여자료를 읽어 근로복지공단 전자신고용 서식을 현장별로 채웁니다. 매월 15일까지 전월분을 신고합니다.",
        "category": "인사·보험",
    },
    {
        "key":      "법률검토",
        "icon":     "⚖",
        "title":    "법률 검토 도우미",
        "desc":     "건강보험·국민연금·고용노동·부동산 관련 법령을 검색하고 조문을 바로 확인합니다.",
        "category": "법무·계약",
    },
    {
        "key":      "계약서색인",
        "icon":     "📑",
        "title":    "계약서 색인표 자동 업데이트",
        "desc":     "HWP 색인표의 신규 계약을 부서별 Excel 파일에 자동 반영합니다. 연도 폴더만 선택하면 매년 그대로 사용할 수 있습니다.",
        "category": "법무·계약",
    },
    {
        "key":      "현장경비",
        "icon":     "💰",
        "title":    "현장경비 입금내역 조회",
        "desc":     "현장경비 입금 내역 파일에서 기간·현장별로 조회하고 엑셀로 저장합니다.",
        "category": "재무·경비",
    },
    {
        "key":      "보증금현황",
        "icon":     "🏠",
        "title":    "숙소보증금 현황표 생성",
        "desc":     "임차보증금 계정별 원장을 읽어 숙소보증금 현황표를 자동으로 생성합니다.",
        "category": "재무·경비",
    },
    {
        "key":      "전산기기",
        "icon":     "💻",
        "title":    "전산기기 지급 내역 관리",
        "desc":     "현장별 전산기기·소모품 지급 및 신청 내역을 엑셀 파일로 관리합니다. 조회·추가·수정·삭제·저장이 가능합니다.",
        "category": "재무·경비",
    },
]

_open_windows = {}


def _open_program(root, key):
    win = _open_windows.get(key)
    if win is not None:
        try:
            if win.winfo_exists():
                win.lift()
                win.focus_force()
                root.iconify()
                return
        except Exception:
            pass

    try:
        if key == "야근수당":
            from mod_야근수당 import open_window
        elif key == "야근현황":
            from mod_야근현황 import open_window
        elif key == "현장경비":
            from mod_현장경비 import open_window
        elif key == "법률검토":
            from mod_법률검토 import open_window
        elif key == "4대보험":
            from mod_4대보험 import open_window
        elif key == "증명서발급":
            from mod_증명서발급 import open_window
        elif key == "계약서색인":
            from mod_계약서색인 import open_window
        elif key == "보증금현황":
            from mod_보증금현황 import open_window
        elif key == "전산기기":
            from mod_전산기기 import open_window
        elif key == "근로내용신고":
            from mod_근로내용신고 import open_window
        else:
            return
        win = open_window(root)
        _open_windows[key] = win

        def _on_child_close():
            win.destroy()
            root.deiconify()
            root.lift()

        win.protocol("WM_DELETE_WINDOW", _on_child_close)
        root.iconify()
    except Exception as e:
        messagebox.showerror("오류", f"프로그램을 열 수 없습니다:\n{e}")


class HubApp:
    def __init__(self, root):
        self.root = root
        root.title("총무팀 자동화 프로그램")
        root.resizable(False, False)
        root.configure(bg="white")
        root.geometry("720x680")
        self._row = 0
        self._build()
        self._center()

    def _center(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"+{x}+{y}")

    def _build(self):
        # 상단 헤더
        hdr = tk.Frame(self.root, bg=ACCENT, height=70)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        tk.Label(
            hdr,
            text="총무팀 자동화 프로그램",
            font=("맑은 고딕", 18, "bold"),
            bg=ACCENT, fg="white",
        ).pack(expand=True)

        # 스크롤 가능한 본문 — 프로그램이 계속 추가돼도 창 크기는 고정되고 목록만 늘어난다
        container = tk.Frame(self.root, bg="white")
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, bg="white", highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        body = tk.Frame(canvas, bg="white")

        body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_window = canvas.create_window((0, 0), window=body, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        grid = tk.Frame(body, bg="white")
        grid.pack(fill="both", expand=True, padx=20, pady=16)
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)

        for cat in CATEGORY_ORDER:
            progs = [p for p in PROGRAMS if p["category"] == cat]
            if progs:
                self._make_category(grid, cat, progs)

        # 하단 (스크롤 영역 밖에 고정)
        footer = tk.Frame(self.root, bg="#F0F0F0", height=32)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)
        tk.Label(
            footer,
            text="선엔지니어링종합건축사사무소 총무팀",
            font=("맑은 고딕", 9),
            bg="#F0F0F0", fg="#888",
        ).pack(expand=True)

    def _make_category(self, parent, title, progs):
        header = tk.Label(
            parent, text=title, font=("맑은 고딕", 12, "bold"),
            bg="white", fg=CAT_FG, anchor="w",
        )
        header.grid(row=self._row, column=0, columnspan=2, sticky="w", pady=(14 if self._row else 0, 6))
        self._row += 1

        for idx, prog in enumerate(progs):
            r = self._row + idx // 2
            c = idx % 2
            self._make_card(parent, prog, r, c)
        self._row += (len(progs) + 1) // 2

    def _make_card(self, parent, prog, row, col):
        key = prog["key"]

        outer = tk.Frame(parent, bg=BTN_BRD, bd=0)
        outer.grid(row=row, column=col, padx=6, pady=6, sticky="nsew")

        card = tk.Frame(outer, bg=BTN_BG, cursor="hand2", padx=14, pady=10)
        card.pack(fill="both", expand=True, padx=1, pady=1)

        left = tk.Frame(card, bg=BTN_BG, width=44)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        tk.Label(left, text=prog["icon"], font=("맑은 고딕", 22), bg=BTN_BG).pack(expand=True)

        right = tk.Frame(card, bg=BTN_BG)
        right.pack(side="left", fill="both", expand=True, padx=(8, 0))

        tk.Label(
            right,
            text=prog["title"],
            font=("맑은 고딕", 11, "bold"),
            bg=BTN_BG, fg=FG_TITLE,
            anchor="w",
        ).pack(fill="x")

        tk.Label(
            right,
            text=prog["desc"],
            font=("맑은 고딕", 8),
            bg=BTN_BG, fg=FG_DESC,
            justify="left",
            anchor="w",
            wraplength=230,
        ).pack(fill="x", pady=(4, 0))

        arrow = tk.Label(card, text="▶", font=("맑은 고딕", 11), bg=BTN_BG, fg=ACCENT)
        arrow.pack(side="right", anchor="n", padx=(4, 0))

        widgets = [card, left, right, arrow] + list(card.winfo_children()) + list(right.winfo_children()) + list(left.winfo_children())

        def on_enter(e, w=outer):
            _set_bg(card, BTN_HVR)

        def on_leave(e, w=outer):
            _set_bg(card, BTN_BG)

        def on_click(e, k=key):
            _open_program(self.root, k)

        for w in [card, left, right, arrow]:
            w.bind("<Enter>", on_enter)
            w.bind("<Leave>", on_leave)
            w.bind("<Button-1>", on_click)

        for child in _all_children(card):
            child.bind("<Enter>", on_enter)
            child.bind("<Leave>", on_leave)
            child.bind("<Button-1>", on_click)


def _set_bg(widget, color):
    try:
        widget.config(bg=color)
    except Exception:
        pass
    for child in _all_children(widget):
        try:
            child.config(bg=color)
        except Exception:
            pass


def _all_children(widget):
    result = []
    for child in widget.winfo_children():
        result.append(child)
        result.extend(_all_children(child))
    return result


if __name__ == "__main__":
    root = TkinterDnD.Tk() if HAS_DND else tk.Tk()
    HubApp(root)
    root.mainloop()
