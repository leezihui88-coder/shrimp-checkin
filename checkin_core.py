# 釣蝦活動報到：共用資料層（桌面版 checkin_app.py 與網頁版 checkin_web.py 共用）
# - 讀報名表（每位參加人一列，同一工號為一組）
# - 每位參加人各自記錄報到狀態，寫回報名表最後面的新增欄位
# - CSV 報到 / 異常紀錄
import csv
import os
import shutil
import time
from copy import copy
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

LOG_FIELDS = ["時間", "輸入內容", "工號", "姓名", "桌號", "到場成員", "結果", "附註"]

# 報名表欄位：程式用「表頭包含的關鍵字」找欄位，欄位順序不影響
COLUMNS = {
    "工號": "工號",
    "員工姓名": "報名員工",
    "參加人": "參加人姓名",
    "桌號": "桌號",
    "方案": "方案名稱",
    "關係": "關係",
    "釣竿": "釣桿",
    "蝦池": "蝦池",          # 選填：每位參加人的蝦池；填「取消」代表此人取消
    "卡號": "卡號",          # 選填：有這欄才能用卡號比對
}
CANCEL = "取消"              # 桌號或蝦池填這個字，代表該參加人已取消
NO_ID_PREFIX = "無工號-"     # 沒有工號的參加人（例如來賓），用姓名當作組別代號
REQUIRED = ("工號", "參加人", "桌號", "方案")
# 報到結果寫回報名表最後面的新增欄位（已存在就沿用同一欄）
RESULT_COLS = ("報到狀態", "報到時間", "報到方式")

PLAN_COLOR = {"釣蝦比賽": "#E74C3C", "陪釣": "#F39C12", "休閒釣蝦": "#27AE60"}


# =====================================================================
# 資料層
# =====================================================================
def clean(v):
    """儲存格轉字串；Excel 數字 12.0 轉成 12"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def norm_card(v):
    """卡號正規化：去空白、去開頭 0"""
    s = clean(v)
    return s.lstrip("0") or s


def norm_id(v):
    """工號正規化：去空白、轉大寫"""
    return clean(v).upper()


def arrived(p):
    return [m for m in p["成員"] if m["報到時間"]]


def active(p):
    """未取消的參加人"""
    return [m for m in p["成員"] if not m["取消"]]


def all_arrived(p):
    """未取消的人都到齊了（全組取消的不算）"""
    act = active(p)
    return bool(act) and all(m["報到時間"] for m in act)


def show_id(p):
    """畫面顯示用的工號；沒有工號的顯示「無工號」"""
    return "無工號" if p["工號"].startswith(NO_ID_PREFIX) else p["工號"]


def ponds(members):
    """成員的蝦池（去重複、依數字排序），陪釣沒有蝦池"""
    s = {m["蝦池"] for m in members if m["蝦池"] and m["蝦池"] != CANCEL}
    return sorted(s, key=lambda x: (0, int(x)) if x.isdigit() else (1, x))


def rod_count(members):
    """(需借釣竿數, 自備釣竿數)"""
    return (sum(1 for m in members if "借" in m["釣竿"]),
            sum(1 for m in members if "自備" in m["釣竿"]))


class Roster:
    """報名名單（以工號為一組，組內每位參加人各自記錄報到）"""

    def __init__(self, xlsx_path):
        self.path = xlsx_path
        self.parties = []         # 每位同仁一組：本人 + 眷屬
        self.by_id = {}
        self.by_card = {}
        self.warnings = []
        self._load()

    def _load(self):
        wb = load_workbook(self.path, data_only=True)
        ws = wb["報到"] if "報到" in wb.sheetnames else wb.active
        self.sheet_name = ws.title
        # 自動尋找表頭列（前 10 列內含「工號」的那列）
        head_row, header = None, []
        for r in range(1, 11):
            header = [clean(c.value) for c in ws[r]]
            if any("工號" in h for h in header):
                head_row = r
                break
        if head_row is None:
            raise ValueError("找不到表頭：前 10 列沒有「工號」欄位")
        cols = {}
        for key, kw in COLUMNS.items():
            for i, h in enumerate(header):
                if kw in h:
                    cols[key] = i
                    break
        missing = [k for k in REQUIRED if k not in cols]
        if missing:
            raise ValueError("缺少必要欄位：" + "、".join(missing))
        # 報到結果欄：已存在就沿用，否則接在最後一個有表頭的欄位後面（1-based 欄號）
        last = max(i for i, h in enumerate(header) if h and h not in RESULT_COLS) + 1
        self.result_cols = {}
        for name in RESULT_COLS:
            if name in header:
                self.result_cols[name] = header.index(name) + 1
            else:
                last += 1
                self.result_cols[name] = last
        self.head_row, self.cols = head_row, cols

        def get(row, key):
            i = cols.get(key)
            return clean(row[i]) if i is not None and i < len(row) else ""

        def raw(row, name):
            i = self.result_cols[name] - 1
            return row[i] if i < len(row) else None

        for r_idx, row in enumerate(ws.iter_rows(min_row=head_row + 1, values_only=True),
                                    start=head_row + 1):
            emp = norm_id(get(row, "工號"))
            if not emp:
                # 沒有工號的參加人（例如來賓）：用報名人姓名當組別代號，只能用姓名搜尋報到
                who = get(row, "員工姓名") or get(row, "參加人")
                if not who:
                    continue
                emp = NO_ID_PREFIX + who
            party = self.by_id.get(emp)
            if party is None:
                party = {"工號": emp, "員工姓名": get(row, "員工姓名"), "桌號": "",
                         "卡號": "", "成員": [], "last_confirm": 0}
                self.by_id[emp] = party
                self.parties.append(party)
            table = get(row, "桌號")
            pond = get(row, "蝦池")
            member = {"row": r_idx, "姓名": get(row, "參加人") or get(row, "員工姓名"),
                      "關係": get(row, "關係") or "本人", "方案": get(row, "方案"),
                      "釣竿": get(row, "釣竿"), "蝦池": pond,
                      "取消": CANCEL in (table, pond), "報到時間": None, "報到方式": ""}
            # Excel 裡已經寫過「已報到」的（例如程式重開），直接還原
            if clean(raw(row, "報到狀態")) == "已報到":
                t = raw(row, "報到時間")
                if isinstance(t, str):
                    try:
                        t = datetime.strptime(t.strip(), "%Y/%m/%d %H:%M:%S")
                    except ValueError:
                        t = None
                member["報到時間"] = t if isinstance(t, datetime) else \
                    datetime.now().replace(microsecond=0)
                member["報到方式"] = clean(raw(row, "報到方式"))
            # 本人排第一位
            if member["關係"] == "本人":
                party["成員"].insert(0, member)
                party["員工姓名"] = member["姓名"] or party["員工姓名"]
            else:
                party["成員"].append(member)
            if table and table != CANCEL:
                if not party["桌號"]:
                    party["桌號"] = table
                elif table != party["桌號"]:
                    self.warnings.append(f"工號 {emp} 同組桌號不一致（第 {r_idx} 列）")
            card = norm_card(get(row, "卡號"))
            if card and not party["卡號"]:
                party["卡號"] = card

        for p in self.parties:
            p["取消"] = not active(p)          # 全組都取消
        n_cancel = sum(1 for p in self.parties for m in p["成員"] if m["取消"])
        if n_cancel:
            self.warnings.append(f"{n_cancel} 位參加人已取消（不列入統計，報到時會提示）")
        no_id = [p for p in self.parties if p["工號"].startswith(NO_ID_PREFIX)]
        if no_id:
            self.warnings.append(f"{len(no_id)} 位沒有工號：" + "、".join(p["員工姓名"] for p in no_id)
                                 + "（請用姓名報到）")
        for p in self.parties:
            if p["卡號"]:
                if p["卡號"] in self.by_card:
                    self.warnings.append(f"卡號重複：{p['卡號']}（工號 {p['工號']}）")
                else:
                    self.by_card[p["卡號"]] = p
        no_table = [p for p in self.parties if not p["桌號"] and not p["取消"]]
        if no_table:
            self.warnings.append(f"{len(no_table)} 位同仁尚未安排桌號")
        self.has_card = bool(self.by_card)

    def find(self, text):
        """卡號 / 工號查詢，回傳 (組, 方式)；方式為 'card' / 'id' / None"""
        c = norm_card(text)
        if c and c in self.by_card:
            return self.by_card[c], "card"
        e = norm_id(text)
        if e in self.by_id:
            return self.by_id[e], "id"
        return None, None

    def search_name(self, text):
        """姓名查詢（同仁或眷屬，部分符合即可，英文不分大小寫），回傳符合的組"""
        key = text.strip().lower()
        return [p for p in self.parties
                if key in p["員工姓名"].lower() or any(key in m["姓名"].lower() for m in p["成員"])]

    def plans(self):
        """方案清單，統計列用"""
        seen = []
        for p in self.parties:
            for m in p["成員"]:
                if m["方案"] and m["方案"] not in seen:
                    seen.append(m["方案"])
        order = list(PLAN_COLOR)
        return sorted(seen, key=lambda x: order.index(x) if x in order else 99)

    def stats(self):
        """依方案統計參加人數：{方案: [已到, 報名]}（已取消的不算）"""
        s = {}
        for p in self.parties:
            for m in active(p):
                d = s.setdefault(m["方案"], [0, 0])
                d[1] += 1
                if m["報到時間"]:
                    d[0] += 1
        return s

    def confirm(self, p, checks, how):
        """依勾選結果報到，回傳本次新報到的成員（已到過的人不會重複記錄）"""
        now = datetime.now().replace(microsecond=0)
        new = [m for m, c in zip(p["成員"], checks)
               if c and not m["報到時間"] and not m["取消"]]
        for m in new:
            m["報到時間"], m["報到方式"] = now, how
        if new:
            p["last_confirm"] = time.time()
        return new

    def backup(self, folder):
        """開始報到前先備份原始報名表"""
        os.makedirs(folder, exist_ok=True)
        name = os.path.splitext(os.path.basename(self.path))[0]
        dst = os.path.join(folder, f"{name}_備份_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        shutil.copy2(self.path, dst)
        return dst

    def save_back(self, out_path=None):
        """把每位參加人的報到結果寫回報名表最後面的新增欄位（預設寫回原檔）"""
        wb = load_workbook(self.path)
        ws = wb[self.sheet_name]
        # 新增欄位的表頭沿用原表頭最後一格的樣式
        ref = ws.cell(row=self.head_row, column=max(self.cols.values()) + 1)
        green = PatternFill("solid", fgColor="D5F5E3")
        for name, c in self.result_cols.items():
            h = ws.cell(row=self.head_row, column=c, value=name)
            if ref.has_style:
                h.font, h.fill = copy(ref.font), copy(ref.fill)
                h.border, h.alignment = copy(ref.border), copy(ref.alignment)
            ws.column_dimensions[h.column_letter].width = 20 if name == "報到時間" else 11
        c_state, c_time, c_how = (self.result_cols[n] for n in RESULT_COLS)
        for p in self.parties:
            for m in p["成員"]:
                done = bool(m["報到時間"])
                state = "已報到" if done else ("已取消" if m["取消"] else "未報到")
                s = ws.cell(row=m["row"], column=c_state, value=state)
                s.fill = green if done else PatternFill()
                t = ws.cell(row=m["row"], column=c_time, value=m["報到時間"])
                t.number_format = "yyyy/mm/dd hh:mm:ss"
                ws.cell(row=m["row"], column=c_how, value=m["報到方式"] if done else None)
        wb.save(out_path or self.path)


class CheckinLog:
    """CSV 報到 / 異常紀錄，每筆立即寫檔"""

    def __init__(self, path):
        self.path = path
        if not os.path.exists(path):
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                csv.DictWriter(f, LOG_FIELDS).writeheader()

    def write(self, **kw):
        kw.setdefault("時間", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        with open(self.path, "a", newline="", encoding="utf-8-sig") as f:
            csv.DictWriter(f, LOG_FIELDS).writerow({k: kw.get(k, "") for k in LOG_FIELDS})

    def restore(self, roster):
        """重開程式時，從紀錄檔還原 Excel 還沒寫到的報到（依到場成員姓名）"""
        n = 0
        with open(self.path, newline="", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                if row.get("結果") != "報到成功":
                    continue
                p = roster.by_id.get(norm_id(row.get("工號")))
                if not p:
                    continue
                names = set((row.get("到場成員") or "").split("、"))
                t = datetime.strptime(row["時間"], "%Y-%m-%d %H:%M:%S")
                how = (row.get("附註") or "工號").split("；")[0]
                for m in p["成員"]:
                    if m["姓名"] in names and not m["報到時間"]:
                        m["報到時間"], m["報到方式"] = t, how
                        n += 1
        return n
