# 釣蝦活動報到程式
# - 刷卡機（鍵盤模擬）、輸入工號或輸入姓名找到同仁那一組
# - 畫面列出本人 + 眷屬，勾選實際到場的人後確認（Enter 或再刷一次）
# - 每位參加人各自記錄報到狀態，寫回報名表最後面的新增欄位
# - 每筆報到即時寫入 CSV 紀錄，Excel 被開啟無法寫入時自動重試
# - 不開放現場加入；查無資料記入異常紀錄
import csv
import os
import sys
import time
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox

from checkin_core import PLAN_COLOR, RESULT_COLS, CheckinLog, Roster, arrived

try:
    import winsound
except ImportError:  # 非 Windows 環境不發聲
    winsound = None

try:  # 高解析度螢幕不模糊、不被裁切
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

APP_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
DEFAULT_XLSX = os.path.join(APP_DIR, "釣蝦報名表.xlsx")
LOG_CSV = os.path.join(APP_DIR, "報到紀錄.csv")

EVENT_TITLE = "MOXA 釣蝦活動 報到"
RESET_SECONDS = 10         # 結果畫面停留秒數
PICK_TIMEOUT = 20          # 勾選畫面無人操作幾秒後自動確認
DOUBLE_SWIPE_SECONDS = 5   # 確認後幾秒內同一組再刷，視為多刷一次，不開新勾選
CARD_HINT_SECONDS = 90     # 刷卡失敗後，多久內改用工號報到會標註「卡號不符」
SAVE_RETRY_SECONDS = 10    # Excel 被開啟而寫入失敗時，每隔幾秒重試

# ---------- 配色 ----------
BG = "#0B1E33"
PANEL = "#12304D"
TEXT = "#FFFFFF"
SUB = "#9FB6CC"
ACCENT = "#2E9CCA"
STATE_COLOR = {"ok": "#1E8449", "dup": "#B9770E", "err": "#B03A2E",
               "pick": "#1F5F8B", "idle": PANEL}
FONT = "Microsoft JhengHei UI"
NUMPAD_KEYCODES = {96 + i: i for i in range(10)}   # Windows 數字鍵盤 0–9


# =====================================================================
# 畫面
# =====================================================================
class App(tk.Tk):
    def __init__(self, roster, log):
        super().__init__()
        self.roster, self.log = roster, log
        self.reset_job = None
        self.save_pending = False      # 寫回 Excel 失敗、等待重試
        self.last_bad_card = None      # (卡號, 時間)：刷卡失敗後改用工號時標註
        self.recent = []               # 最近報到清單 (時間, 組, 本次人數)
        # 勾選中的組：{"party", "how", "checks": [bool], "note", "text"}
        self.pick = None
        self.choices = None            # 姓名搜尋有多組符合時的候選清單

        self.title(EVENT_TITLE)
        self.configure(bg=BG)
        self.geometry("1280x800")
        self.minsize(1000, 680)
        self.attributes("-fullscreen", True)
        self.bind("<F11>", lambda e: self.attributes("-fullscreen",
                                                     not self.attributes("-fullscreen")))
        self.bind("<Escape>", self.on_escape)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self._build()
        self.refresh_stats()
        self.show_idle()
        self.tick()

    # ---------- 版面 ----------
    def _build(self):
        # 頂端標題列
        top = tk.Frame(self, bg=BG)
        top.pack(fill="x", padx=40, pady=(24, 8))
        tk.Label(top, text="🦐  " + EVENT_TITLE, font=(FONT, 30, "bold"),
                 fg=TEXT, bg=BG).pack(side="left")
        self.clock = tk.Label(top, font=(FONT, 24), fg=SUB, bg=BG)
        self.clock.pack(side="right")

        # 先保留底部空間（管理按鈕、統計列），結果卡內容再多也不會把它們擠出畫面
        tools = tk.Frame(self, bg=BG)
        tools.pack(side="bottom", fill="x", padx=40, pady=(0, 14))
        bottom = tk.Frame(self, bg=BG)
        bottom.pack(side="bottom", fill="x", padx=40, pady=(8, 10))

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=40, pady=8)

        # 左：輸入 + 結果卡
        left = tk.Frame(body, bg=BG)
        left.pack(side="left", fill="both", expand=True)

        box = tk.Frame(left, bg=PANEL, padx=28, pady=16)
        box.pack(fill="x")
        head = "請刷卡，或輸入工號 / 姓名後按 Enter" if self.roster.has_card \
            else "請輸入工號或姓名後按 Enter"
        self.hint_lbl = tk.Label(box, text=head, font=(FONT, 18), fg=SUB, bg=PANEL)
        self.hint_lbl.pack(anchor="w")
        self.entry = tk.Entry(box, font=(FONT, 32, "bold"), bg="#FFFFFF", fg="#0B1E33",
                              relief="flat", justify="center", insertwidth=4)
        self.entry.pack(fill="x", pady=(8, 0), ipady=8)
        self.entry.bind("<Return>", self.on_submit)
        self.entry.bind("<KP_Enter>", self.on_submit)
        self.entry.bind("<KeyPress>", self.on_key)

        # 結果卡：上方狀態列，下方左側人員資訊、右側桌號
        self.card = tk.Frame(left, bg=PANEL)
        self.card.pack(fill="both", expand=True, pady=(18, 0))
        self.state_lbl = tk.Label(self.card, font=(FONT, 28, "bold"), fg=TEXT, bg=PANEL)
        self.state_lbl.pack(pady=(20, 0))

        self.content = tk.Frame(self.card, bg=PANEL)
        self.content.pack(fill="both", expand=True, padx=36, pady=(8, 16))

        self.table_box = tk.Frame(self.content, bg="#FFFFFF", width=300)
        self.table_box.pack_propagate(False)
        tk.Label(self.table_box, text="桌號", font=(FONT, 26, "bold"),
                 fg="#0B1E33", bg="#FFFFFF").pack(pady=(26, 0))
        self.table_lbl = tk.Label(self.table_box, font=(FONT, 110, "bold"),
                                  fg="#0B1E33", bg="#FFFFFF")
        self.table_lbl.pack(expand=True)

        self.info = tk.Frame(self.content, bg=PANEL)
        self.info.pack(side="left", fill="both", expand=True)
        self.name_lbl = tk.Label(self.info, font=(FONT, 54, "bold"), fg=TEXT, bg=PANEL)
        self.name_lbl.pack(anchor="center")
        self.sub_lbl = tk.Label(self.info, font=(FONT, 20), fg=TEXT, bg=PANEL)
        self.sub_lbl.pack(anchor="center", pady=(0, 10))
        self.members = tk.Frame(self.info, bg=PANEL)
        self.note_lbl = tk.Label(self.info, font=(FONT, 16), fg="#FDEBD0", bg=PANEL)
        self.help_lbl = tk.Label(self.card, font=(FONT, 15, "bold"), fg="#0B1E33",
                                 bg="#F7DC6F", padx=16, pady=6)

        # 右：最近報到
        right = tk.Frame(body, bg=PANEL, width=340, padx=20, pady=16)
        right.pack(side="right", fill="y", padx=(20, 0))
        right.pack_propagate(False)
        tk.Label(right, text="最近報到", font=(FONT, 18, "bold"), fg=TEXT,
                 bg=PANEL).pack(anchor="w")
        self.recent_box = tk.Frame(right, bg=PANEL)
        self.recent_box.pack(fill="both", expand=True, pady=(10, 0))

        # 底部統計列：各方案 + 合計
        self.stat_lbls = {}
        for g in self.roster.plans() + ["合計"]:
            f = tk.Frame(bottom, bg=PANEL, padx=18, pady=10)
            f.pack(side="left", fill="x", expand=True, padx=(0, 12))
            tk.Frame(f, bg=PLAN_COLOR.get(g, ACCENT), width=8).pack(side="left", fill="y",
                                                                     padx=(0, 14))
            inner = tk.Frame(f, bg=PANEL)
            inner.pack(side="left")
            tk.Label(inner, text=g, font=(FONT, 15), fg=SUB, bg=PANEL).pack(anchor="w")
            v = tk.Label(inner, font=(FONT, 26, "bold"), fg=TEXT, bg=PANEL)
            v.pack(anchor="w")
            s = tk.Label(inner, font=(FONT, 12), fg=SUB, bg=PANEL)
            s.pack(anchor="w")
            self.stat_lbls[g] = (v, s)

        # 最底部管理按鈕（小字，避免誤按）
        for txt, cmd in (("未報到名單", self.show_absent),
                         ("異常紀錄", self.show_errors),
                         ("立即寫回 Excel", lambda: self.save_excel(manual=True))):
            tk.Button(tools, text=txt, command=cmd, font=(FONT, 11), fg=SUB, bg=BG,
                      activebackground=PANEL, activeforeground=TEXT, relief="flat",
                      bd=0, cursor="hand2", takefocus=0).pack(side="right", padx=8)
        tk.Label(tools, text="F11 全螢幕切換", font=(FONT, 11),
                 fg="#5D7A96", bg=BG).pack(side="left")
        # Excel 寫回狀態
        self.save_lbl = tk.Label(tools, text=f"Excel：{os.path.basename(self.roster.path)}",
                                 font=(FONT, 11), fg="#5D7A96", bg=BG)
        self.save_lbl.pack(side="left", padx=(24, 0))

        self.keep_focus()

    def keep_focus(self):
        """隨時把游標拉回輸入框，刷卡不用先點滑鼠"""
        w = self.focus_get()
        if w is None or isinstance(w, tk.Tk):
            self.entry.focus_force()
        elif w.winfo_toplevel() is self:
            self.entry.focus_set()
        self.after(500, self.keep_focus)

    def tick(self):
        self.clock.config(text=datetime.now().strftime("%m/%d  %H:%M:%S"))
        self.after(1000, self.tick)

    # ---------- 結果卡 ----------
    def paint(self, state):
        color = STATE_COLOR[state]
        for w in (self.card, self.state_lbl, self.content, self.info, self.name_lbl,
                  self.sub_lbl, self.members, self.note_lbl):
            w.config(bg=color)
        return color

    def clear_card(self):
        for w in (self.table_box, self.members, self.note_lbl, self.help_lbl):
            w.pack_forget()
        for w in self.members.winfo_children():
            w.destroy()
        # 搜尋畫面會隱藏姓名列，這裡依原順序重新放回
        for w in (self.name_lbl, self.sub_lbl):
            w.pack_forget()
        self.name_lbl.config(font=(FONT, 54, "bold"))
        self.sub_lbl.config(font=(FONT, 20))
        self.name_lbl.pack(anchor="center")
        self.sub_lbl.pack(anchor="center", pady=(0, 10))

    def show_idle(self):
        self.pick = self.choices = None
        self.clear_card()
        self.paint("idle")
        self.state_lbl.config(text="歡迎參加！", fg=SUB)
        self.name_lbl.config(text="請刷卡報到" if self.roster.has_card else "請輸入工號報到",
                             font=(FONT, 54, "bold"))
        self.sub_lbl.config(text="報到後會顯示您的桌號", fg=SUB)

    def show_party(self, state, title, p, checks=None, note="", help_text=""):
        """顯示一組人員；checks 有值時為勾選模式"""
        self.clear_card()
        color = self.paint(state)
        self.state_lbl.config(text=title, fg=TEXT)

        # 右側桌號大字
        self.table_box.pack(side="right", fill="y", padx=(24, 0))
        if p["桌號"]:
            self.table_lbl.config(text=p["桌號"], font=(FONT, 110, "bold"), fg="#0B1E33")
        else:
            self.table_lbl.config(text="未安排\n請洽承辦人", font=(FONT, 26, "bold"),
                                  fg="#B03A2E")

        n = len(p["成員"])
        # 人數多時縮小字體，5–6 位也能完整顯示；勾選模式底下有提示列，門檻較低
        if n <= (2 if help_text else 3):
            f_title, f_name, f_rel, f_plan, f_rod, gap = 54, 28, 22, 21, 20, 4
        elif n <= (3 if help_text else 4):
            f_title, f_name, f_rel, f_plan, f_rod, gap = 40, 22, 18, 17, 16, 2
        else:
            f_title, f_name, f_rel, f_plan, f_rod, gap = 32, 18, 15, 14, 14, 0
        self.name_lbl.config(text=p["員工姓名"], font=(FONT, f_title, "bold"))
        got = len(arrived(p))
        self.sub_lbl.config(text=f"工號 {p['工號']}　｜　已到 {got} / {n} 位", fg=TEXT,
                            font=(FONT, 20 if f_title > 32 else 16))

        # 參加人明細：[勾選] 編號 / 姓名 / 關係 / 方案 / 釣竿或報到時間
        self.members.pack(fill="x", pady=(4, 0))
        for c, w in enumerate((1, 1, 5, 3, 4, 3)):
            self.members.grid_columnconfigure(c, weight=w)
        for i, m in enumerate(p["成員"]):
            if m["報到時間"]:
                mark, mark_fg = "✔", "#ABEBC6"
            elif checks is not None:
                mark, mark_fg = ("☑", "#FFFFFF") if checks[i] else ("☐", "#AEB6BF")
            else:
                mark, mark_fg = "✖", "#F5B7B1"
            dim = checks is not None and not m["報到時間"] and not checks[i]
            fg = "#AEB6BF" if dim else TEXT
            cells = [
                tk.Label(self.members, text=mark, font=(FONT, f_name + 2, "bold"),
                         fg=mark_fg, bg=color),
                tk.Label(self.members, text=str(i + 1) if checks is not None else "",
                         font=(FONT, f_rel, "bold"), fg="#F7DC6F", bg=color),
                tk.Label(self.members, text=m["姓名"], font=(FONT, f_name, "bold"),
                         fg=fg, bg=color, anchor="w"),
                tk.Label(self.members, text=m["關係"], font=(FONT, f_rel), fg=fg,
                         bg=color, anchor="w"),
                tk.Label(self.members, text=m["方案"] or "—", font=(FONT, f_plan, "bold"),
                         fg=PLAN_COLOR.get(m["方案"], ACCENT), bg="#FFFFFF", padx=10),
            ]
            if m["報到時間"] and checks is not None:
                last = tk.Label(self.members, text=f"{m['報到時間']:%H:%M} 已到",
                                font=(FONT, f_rod), fg="#ABEBC6", bg=color, anchor="w")
            else:
                rod = m["釣竿"]
                last = tk.Label(self.members, text=rod,
                                font=(FONT, f_rod, "bold" if "借" in rod else "normal"),
                                fg="#FFF3B0" if "借" in rod else fg, bg=color, anchor="w")
            cells.append(last)
            for c, w in enumerate(cells):
                w.grid(row=i, column=c, sticky="w" if c in (4,) else "we",
                       pady=gap, padx=(0, 12) if c == 4 else 0)
                # 勾選模式下，用滑鼠點該列也能切換
                if checks is not None and not m["報到時間"]:
                    w.bind("<Button-1>", lambda e, k=i: self.toggle(k))
                    w.config(cursor="hand2")
        if note:
            self.note_lbl.config(text=note)
            self.note_lbl.pack(pady=(10, 0))
        if help_text:
            self.show_help(help_text)

    def show_help(self, text):
        """結果卡底部的黃色操作提示，先保留位置，不會被人員清單擠掉"""
        self.help_lbl.config(text=text)
        self.help_lbl.pack(side="bottom", pady=(0, 14), before=self.content)

    def show_error(self, title, msg):
        self.pick = self.choices = None
        self.clear_card()
        self.paint("err")
        self.state_lbl.config(text="✖  " + title, fg=TEXT)
        self.name_lbl.config(text=msg, font=(FONT, 54, "bold"))
        self.sub_lbl.config(text="", fg=TEXT)

    def show_choices(self, parties):
        """姓名搜尋有多組符合：列出讓報到人員選"""
        self.pick = None
        self.choices = parties[:8]
        self.clear_card()
        color = self.paint("pick")
        more = "（只列前 8 組，請輸入完整姓名）" if len(parties) > 8 else ""
        self.state_lbl.config(text=f"🔍  找到 {len(parties)} 組，請選擇{more}", fg=TEXT)
        self.name_lbl.pack_forget()
        self.sub_lbl.pack_forget()
        self.members.pack(fill="x", pady=(10, 0))
        k = len(self.choices)
        size = 20 if k <= 5 else (17 if k <= 7 else 15)
        for i, p in enumerate(self.choices):
            fam = "、".join(m["姓名"] for m in p["成員"][1:])
            txt = f"{p['員工姓名']}（{p['工號']}）　{p['桌號'] or '-'} 桌" + (f"　眷屬：{fam}" if fam else "")
            tk.Label(self.members, text=str(i + 1), font=(FONT, size + 2, "bold"), fg="#F7DC6F",
                     bg=color).grid(row=i, column=0, padx=(0, 16), pady=1)
            lbl = tk.Label(self.members, text=txt, font=(FONT, size), fg=TEXT, bg=color,
                           anchor="w", cursor="hand2")
            lbl.grid(row=i, column=1, sticky="w")
            lbl.bind("<Button-1>", lambda e, k=i: self.choose(k))
        self.show_help("按數字選擇　｜　Esc 取消")
        self.schedule_reset(PICK_TIMEOUT)

    def schedule_reset(self, seconds=RESET_SECONDS, action=None):
        if self.reset_job:
            self.after_cancel(self.reset_job)
        self.reset_job = self.after(seconds * 1000, action or self.show_idle)

    # ---------- 勾選流程 ----------
    def start_pick(self, p, how, text, note=""):
        """開始勾選：第一次報到預設全勾；晚到補報到時，未到的人預設不勾"""
        first = not arrived(p)
        checks = [first and not m["報到時間"] for m in p["成員"]]
        self.choices = None
        self.pick = {"party": p, "how": how, "checks": checks, "note": note, "text": text}
        self.render_pick()
        self.beep(True)

    def render_pick(self):
        pk = self.pick
        p = pk["party"]
        n_sel = sum(pk["checks"])
        title = "請確認到場人員" if not arrived(p) else "補報到：請勾選剛到的人"
        confirm = "再刷一次卡 或 Enter 確認" if pk["how"] == "card" else "Enter 確認"
        self.show_party("pick", f"👥  {title}（已勾 {n_sel} 位）", p, pk["checks"], pk["note"],
                        f"數字鍵盤：直接按切換勾選（上排數字：數字＋Enter）　｜　{confirm}　｜　Esc 取消")
        self.schedule_reset(PICK_TIMEOUT, self.confirm_pick)

    def toggle(self, i):
        pk = self.pick
        if not pk or i >= len(pk["checks"]) or pk["party"]["成員"][i]["報到時間"]:
            return
        pk["checks"][i] = not pk["checks"][i]
        self.render_pick()

    def choose(self, i):
        if self.choices and i < len(self.choices):
            p = self.choices[i]
            self.handle_party(p, "姓名", p["員工姓名"])

    def confirm_pick(self):
        pk, self.pick = self.pick, None
        if not pk:
            return
        p = pk["party"]
        now = datetime.now().replace(microsecond=0)
        how = {"card": "刷卡", "id": "工號"}.get(pk["how"], pk["how"])
        new = [m for m, c in zip(p["成員"], pk["checks"]) if c and not m["報到時間"]]
        if not new:
            self.show_error("未勾選任何人", "已取消，請重新報到")
            self.beep(False)
            self.schedule_reset()
            return
        for m in new:
            m["報到時間"], m["報到方式"] = now, how
        p["last_confirm"] = time.time()
        absent = [m["姓名"] for m in p["成員"] if not m["報到時間"]]
        self.log.write(輸入內容=pk["text"], 工號=p["工號"], 姓名=p["員工姓名"], 桌號=p["桌號"],
                       到場成員="、".join(m["姓名"] for m in new), 結果="報到成功",
                       附註=how + (f"；{pk['note']}" if pk["note"] else "")
                       + (f"；未到：{'、'.join(absent)}" if absent else ""))
        note = pk["note"] or (f"尚未到：{'、'.join(absent)}（晚到時再刷卡 / 輸入工號補報到）"
                              if absent else "")
        self.show_party("ok", f"✔  報到成功　本次 {len(new)} 位", p, note=note)
        self.recent.insert(0, (now, p, len(new)))
        self.render_recent()
        self.refresh_stats()
        self.beep(True)
        self.schedule_reset()
        self.update_idletasks()    # 先把結果畫面畫出來，再寫 Excel
        self.save_excel()

    # ---------- 輸入處理 ----------
    def on_key(self, e):
        """數字鍵盤按鍵：勾選模式下直接切換，搜尋模式下直接選擇"""
        if e.keycode in NUMPAD_KEYCODES and not self.entry.get():
            k = NUMPAD_KEYCODES[e.keycode] - 1
            if self.pick and k >= 0:
                self.toggle(k)
                return "break"
            if self.choices and k >= 0:
                self.choose(k)
                return "break"

    def on_escape(self, _=None):
        if self.pick or self.choices:
            self.entry.delete(0, "end")
            self.show_idle()
        else:
            self.attributes("-fullscreen", False)

    def on_submit(self, _=None):
        text = self.entry.get().strip()
        self.entry.delete(0, "end")

        # 勾選 / 選擇模式下：空白 Enter 確認、單一數字切換或選擇
        if self.pick:
            if not text:
                self.confirm_pick()
                return
            if text.isdigit() and len(text) == 1:
                self.toggle(int(text) - 1)
                return
        if self.choices and text.isdigit() and len(text) == 1:
            self.choose(int(text) - 1)
            return
        if not text:
            return

        p, how = self.roster.find(text)
        if self.pick:
            if p is self.pick["party"]:      # 同一組再刷一次 = 確認
                self.confirm_pick()
                return
            self.confirm_pick()              # 下一位已經刷卡：先把上一組確認掉
        if p is None and not text.isdigit():
            found = self.roster.search_name(text)
            if len(found) == 1:
                p, how = found[0], "姓名"
            elif len(found) > 1:
                self.show_choices(found)
                return

        if p is None:
            # 有卡號資料時，看起來像卡號（比工號長）就提示改輸入工號
            is_card = self.roster.has_card and text.isdigit() and len(text) > 8
            self.log.write(輸入內容=text, 結果="查無資料", 附註="卡號" if is_card else "工號 / 姓名")
            if is_card:
                self.last_bad_card = (text, time.time())
                self.show_error("查無此卡號", "請改輸入工號")
            else:
                self.show_error("查無報名資料", "請洽承辦人")
            self.beep(False)
            self.schedule_reset()
            return
        self.handle_party(p, how, text)

    def handle_party(self, p, how, text):
        if len(arrived(p)) == len(p["成員"]) or \
                time.time() - p["last_confirm"] < DOUBLE_SWIPE_SECONDS:
            # 全組已到齊，或剛確認完又多刷一次
            t = max(m["報到時間"] for m in arrived(p)).strftime("%H:%M")
            self.log.write(輸入內容=text, 工號=p["工號"], 姓名=p["員工姓名"], 桌號=p["桌號"],
                           結果="重複報到", 附註=f"最後報到 {t}")
            self.pick = self.choices = None
            self.show_party("dup", f"⚠  已於 {t} 報到過", p)
            self.beep(False)
            self.schedule_reset()
            return
        note = ""
        if how == "id" and self.last_bad_card and \
                time.time() - self.last_bad_card[1] < CARD_HINT_SECONDS:
            note = f"提醒：剛才刷的卡號 {self.last_bad_card[0]} 與報名資料不符"
        self.last_bad_card = None
        self.start_pick(p, how, text, note)

    def beep(self, ok):
        if winsound:
            winsound.MessageBeep(winsound.MB_OK if ok else winsound.MB_ICONHAND)

    def render_recent(self):
        self.recent = self.recent[:12]
        for w in self.recent_box.winfo_children():
            w.destroy()
        for t, q, k in self.recent:
            row = tk.Frame(self.recent_box, bg=PANEL)
            row.pack(fill="x", pady=3)
            plan = q["成員"][0]["方案"] if q["成員"] else ""
            tk.Frame(row, bg=PLAN_COLOR.get(plan, ACCENT), width=6).pack(side="left", fill="y")
            tk.Label(row, text=t.strftime("%H:%M"), font=(FONT, 13),
                     fg=SUB, bg=PANEL).pack(side="left", padx=(10, 8))
            tk.Label(row, text=f"{q['員工姓名']} {k}位", font=(FONT, 15, "bold"),
                     fg=TEXT, bg=PANEL).pack(side="left")
            tk.Label(row, text=f"{q['桌號'] or '-'} 桌", font=(FONT, 13),
                     fg=SUB, bg=PANEL).pack(side="right")

    def refresh_stats(self):
        s = self.roster.stats()
        for g in self.roster.plans():
            v, sub = self.stat_lbls[g]
            got, tot = s.get(g, [0, 0])
            v.config(text=f"{got} / {tot}")
            sub.config(text="人（含眷屬）")
        parties = self.roster.parties
        all_got = sum(len(arrived(p)) for p in parties)
        all_tot = sum(len(p["成員"]) for p in parties)
        grp_got = sum(1 for p in parties if arrived(p))
        v, sub = self.stat_lbls["合計"]
        v.config(text=f"{all_got} / {all_tot}")
        sub.config(text=f"已報到 {grp_got} / {len(parties)} 組")

    # ---------- 管理功能 ----------
    def list_window(self, title, header, rows):
        win = tk.Toplevel(self, bg=BG)
        win.title(title)
        win.geometry("1000x640")
        tk.Label(win, text=f"{title}（{len(rows)} 筆）", font=(FONT, 16, "bold"),
                 fg=TEXT, bg=BG).pack(anchor="w", padx=16, pady=10)
        txt = tk.Text(win, font=(FONT, 13), bg=PANEL, fg=TEXT, relief="flat", padx=10)
        txt.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        txt.insert("end", header + "\n" + "─" * 40 + "\n")
        txt.insert("end", "\n".join(rows) if rows else "（無）")
        txt.config(state="disabled")
        win.bind("<Escape>", lambda e: win.destroy())

    def show_absent(self):
        def key(p):
            t = p["桌號"]
            return (0, int(t)) if t.isdigit() else (1, t)
        rows = []
        for p in sorted(self.roster.parties, key=key):
            miss = [m for m in p["成員"] if not m["報到時間"]]
            if not miss:
                continue
            state = "全組未到" if len(miss) == len(p["成員"]) else "部分未到"
            names = "、".join(f"{m['姓名']}({m['關係']})" for m in miss)
            rows.append(f"{p['桌號'] or '未安排':>4} 桌\t{p['工號']}\t{p['員工姓名']}\t{state}\t{names}")
        self.list_window("未報到名單", "桌號\t工號\t\t同仁\t狀態\t\t未到人員", rows)

    def show_errors(self):
        with open(self.log.path, newline="", encoding="utf-8-sig") as f:
            rows = [f"{r['時間'][11:]}\t{r['結果']}\t{r['輸入內容']}\t{r['姓名']}\t{r['附註']}"
                    for r in csv.DictReader(f) if r["結果"] != "報到成功"]
        self.list_window("異常紀錄", "時間\t\t結果\t輸入內容\t姓名\t附註", rows)

    def save_excel(self, manual=False):
        """寫回報名表；失敗（多半是檔案被 Excel 開著）就稍後自動重試，資料已在 CSV 不會遺失"""
        try:
            self.roster.save_back()
        except Exception as e:
            first = not self.save_pending
            self.save_pending = True
            self.save_lbl.config(text="⚠ Excel 寫入失敗：請關閉報名表檔案，程式會自動重試"
                                      "（報到資料已存於報到紀錄.csv）", fg="#F5B041")
            if first:
                self.after(SAVE_RETRY_SECONDS * 1000, self.retry_save)
            if manual:
                messagebox.showerror("寫入失敗", f"{e}\n\n請確認報名表沒有在 Excel 中開啟。",
                                     parent=self)
            return False
        self.save_pending = False
        self.save_lbl.config(text=f"✔ 已寫回 Excel {datetime.now():%H:%M:%S}", fg="#58D68D")
        if manual:
            messagebox.showinfo("寫入完成", f"已寫回：\n{self.roster.path}", parent=self)
        return True

    def retry_save(self):
        if self.save_pending and not self.save_excel():
            self.after(SAVE_RETRY_SECONDS * 1000, self.retry_save)

    def on_close(self):
        if not messagebox.askyesno("結束報到", "要結束報到程式嗎？", parent=self):
            return
        if self.pick:
            self.confirm_pick()
        if self.save_excel() or messagebox.askyesno(
                "Excel 尚未寫入",
                "報名表目前無法寫入（可能被開啟）。\n報到資料已保存在「報到紀錄.csv」，"
                "下次開啟程式會自動補寫。\n\n仍要結束嗎？", parent=self):
            self.destroy()


def main():
    root = tk.Tk()
    root.withdraw()
    path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_XLSX
    if not os.path.exists(path):
        path = filedialog.askopenfilename(title="選擇報名表 Excel",
                                          filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
    try:
        roster = Roster(path)
    except Exception as e:
        messagebox.showerror("讀取報名表失敗", str(e))
        return
    try:
        bak = roster.backup(os.path.join(APP_DIR, "備份"))
    except Exception as e:
        messagebox.showerror("備份失敗", f"無法備份報名表，請確認檔案沒有被開啟：\n{e}")
        return
    log = CheckinLog(LOG_CSV)
    restored = log.restore(roster)
    total = sum(len(p["成員"]) for p in roster.parties)
    msg = [f"已載入 {len(roster.parties)} 位同仁，共 {total} 位參加人",
           f"報到結果會寫回報名表最後面的「{'、'.join(RESULT_COLS)}」欄",
           f"原始檔已備份：備份\\{os.path.basename(bak)}"]
    if not roster.has_card:
        msg.append("報名表沒有「卡號」欄位，目前只能用工號或姓名報到")
    if restored:
        msg.append(f"從紀錄檔補回 {restored} 位已報到")
    if roster.warnings:
        msg.append("\n資料提醒：\n" + "\n".join(roster.warnings[:15]))
    messagebox.showinfo("報名表載入完成", "\n".join(msg))
    root.destroy()
    app = App(roster, log)
    if restored:      # CSV 裡有、Excel 還沒寫到的，開啟時補寫
        app.save_excel()
    app.mainloop()


if __name__ == "__main__":
    main()
