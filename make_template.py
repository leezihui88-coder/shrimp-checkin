# 產生釣蝦活動報名 / 報到 Excel 範本（美化版）
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment

FONT = "Microsoft JhengHei"
OUT = r"C:\Users\dana_lee\Projects\shrimp_checkin\釣蝦活動報名表.xlsx"
MAX_ROW = 500          # 預留填寫列數
HEAD = 3               # 表頭所在列（第 1 列標題、第 2 列副標）
FIRST = HEAD + 1       # 資料起始列
LAST = HEAD + MAX_ROW  # 資料結束列

# 色票：海洋藍主色 + 三組別代表色
NAVY, TEAL, SKY = "0B3C5D", "1D7A8C", "E8F4F8"
GRAY_H, GRAY_BG, LINE = "8C9AA5", "F4F6F8", "D9E1E7"
GROUP_COLOR = {"比賽組": ("C0392B", "FDECEA"),   # 紅
               "陪釣組": ("D68910", "FEF5E7"),   # 橘
               "休閒組": ("1E8449", "E9F7EF")}   # 綠

def fill(c): return PatternFill("solid", fgColor=c)
def font(**kw): return Font(name=FONT, **kw)

line = Side(style="thin", color=LINE)
border = Border(left=line, right=line, top=line, bottom=line)
bottom_only = Border(bottom=Side(style="medium", color=TEAL))
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", indent=1)

def banner(ws, title, subtitle, last_col):
    """工作表頂端的標題橫幅"""
    ws.merge_cells(f"A1:{last_col}1")
    ws.merge_cells(f"A2:{last_col}2")
    ws["A1"] = title
    ws["A1"].font = font(size=18, bold=True, color="FFFFFF")
    ws["A1"].fill = fill(NAVY)
    ws["A1"].alignment = left
    ws["A2"] = subtitle
    ws["A2"].font = font(size=10, color="FFFFFF")
    ws["A2"].fill = fill(TEAL)
    ws["A2"].alignment = left
    ws.row_dimensions[1].height = 40
    ws.row_dimensions[2].height = 22
    ws.sheet_view.showGridLines = False

wb = Workbook()

# ================= Sheet：報名名單 =================
ws = wb.active
ws.title = "報名名單"
headers = [
    # (欄名, 寬度, 是否自動欄, 說明)
    ("工號", 12, False, "必填，例如 M12345"),
    ("姓名", 12, False, "必填"),
    ("部門", 14, False, "選填"),
    ("卡號", 16, False, "必填，文字格式，保留開頭的 0"),
    ("組別", 11, False, "下拉選單：比賽組 / 陪釣組 / 休閒組"),
    ("眷屬姓名", 26, False, "多位眷屬用頓號「、」分隔；沒有就留白"),
    ("眷屬人數", 10, True, "公式自動計算，請勿填寫"),
    ("備註", 20, False, "選填，例如素食、需停車位"),
    ("報到狀態", 11, True, "報到程式自動寫入，請勿填寫"),
    ("報到時間", 18, True, "報到程式自動寫入，請勿填寫"),
]
banner(ws, "  🦐 釣蝦活動　報名 / 報到名單",
       "  深藍表頭＝請填寫　｜　灰色表頭＝自動欄位，請勿填寫　｜　第 4–6 列為範例，正式使用前請刪除", "J")

for c, (name, width, auto, note) in enumerate(headers, start=1):
    cell = ws.cell(row=HEAD, column=c, value=name)
    cell.font = font(bold=True, color="FFFFFF", size=11)
    cell.fill = fill(GRAY_H if auto else NAVY)
    cell.alignment = center
    cell.border = Border(left=Side(style="thin", color="FFFFFF"),
                         right=Side(style="thin", color="FFFFFF"))
    cell.comment = Comment(note, "範本")
    ws.column_dimensions[cell.column_letter].width = width
ws.row_dimensions[HEAD].height = 28
ws.freeze_panes = f"C{FIRST}"   # 凍結表頭與工號、姓名欄

samples = [
    ("M12345", "王小明", "GTS", "0012345678", "比賽組", "王太太、王小寶", "範例，請刪除"),
    ("M23456", "陳美玲", "RMA", "0023456789", "陪釣組", "", "範例，請刪除"),
    ("M34567", "林大同", "TS", "0034567890", "休閒組", "林媽媽", "範例，請刪除"),
]
for i, (eid, name, dept, card, grp, fam, memo) in enumerate(samples):
    r = FIRST + i
    for c, v in zip((1, 2, 3, 4, 5, 6, 8), (eid, name, dept, card, grp, fam, memo)):
        ws.cell(row=r, column=c, value=v)

for r in range(FIRST, LAST + 1):
    ws.cell(row=r, column=7,
            value=f'=IF(TRIM(F{r})="",0,LEN(TRIM(F{r}))-LEN(SUBSTITUTE(TRIM(F{r}),"、",""))+1)')
    ws.row_dimensions[r].height = 20
    for c in range(1, 11):
        cell = ws.cell(row=r, column=c)
        cell.font = font(size=10.5, italic=(r < FIRST + len(samples)),
                         color="7F8C8D" if r < FIRST + len(samples) else "1C2833")
        cell.border = border
        cell.alignment = center if c in (1, 4, 5, 7, 9, 10) else left
        if headers[c - 1][2]:
            cell.fill = fill(GRAY_BG)
    ws.cell(row=r, column=1).number_format = "@"
    ws.cell(row=r, column=4).number_format = "@"   # 卡號當文字，避免吃掉開頭 0
    ws.cell(row=r, column=10).number_format = "yyyy/mm/dd hh:mm"

# 條件格式：組別上色、報到狀態上色、重複工號 / 卡號標紅
data = f"A{FIRST}:J{LAST}"
for g, (fg, bg) in GROUP_COLOR.items():
    ws.conditional_formatting.add(f"E{FIRST}:E{LAST}", FormulaRule(
        formula=[f'$E{FIRST}="{g}"'], font=Font(name=FONT, bold=True, color=fg), fill=fill(bg)))
ws.conditional_formatting.add(f"I{FIRST}:J{LAST}", FormulaRule(
    formula=[f'$I{FIRST}="已報到"'], font=Font(name=FONT, bold=True, color="1E8449"), fill=fill("D5F5E3")))
for col in ("A", "D"):
    ws.conditional_formatting.add(f"{col}{FIRST}:{col}{LAST}", FormulaRule(
        formula=[f'AND({col}{FIRST}<>"",COUNTIF(${col}${FIRST}:${col}${LAST},{col}{FIRST})>1)'],
        font=Font(name=FONT, bold=True, color="FFFFFF"), fill=fill("E74C3C")))
# 已報到整列淡綠
ws.conditional_formatting.add(f"A{FIRST}:H{LAST}", FormulaRule(
    formula=[f'$I{FIRST}="已報到"'], fill=fill("F0FBF4")))

dv = DataValidation(type="list", formula1='"比賽組,陪釣組,休閒組"', allow_blank=True,
                    showErrorMessage=True, errorTitle="組別錯誤",
                    error="請從下拉選單選擇：比賽組 / 陪釣組 / 休閒組")
ws.add_data_validation(dv)
dv.add(f"E{FIRST}:E{LAST}")
ws.auto_filter.ref = f"A{HEAD}:J{LAST}"

# 列印設定：橫向、一頁寬、每頁重複表頭
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = f"{HEAD}:{HEAD}"
ws.print_options.horizontalCentered = True

# ================= Sheet：統計 =================
st = wb.create_sheet("統計")
banner(st, "  📊 報到統計", "  數字依報名名單自動計算；已報到人數在活動當天由報到程式寫入後更新", "H")
st.column_dimensions["A"].width = 2
for col in "BCDEFGH":
    st.column_dimensions[col].width = 14
rng = lambda col: f"報名名單!${col}${FIRST}:${col}${LAST}"

# KPI 卡片：總報名、總人數（含眷屬）、已報到、報到率
kpis = [("B", "C", "報名同仁", f'=COUNTIF({rng("E")},"?*")'),
        ("D", "E", "含眷屬總人數", "=E14"),
        ("F", "G", "已報到同仁", "=F14"),
        ("H", "H", "報到率", "=H14")]
for c1, c2, label, f in kpis:
    st.merge_cells(f"{c1}4:{c2}4")
    st.merge_cells(f"{c1}5:{c2}6")
    st[f"{c1}4"] = label
    st[f"{c1}4"].font = font(size=10, color="FFFFFF", bold=True)
    st[f"{c1}4"].fill = fill(TEAL)
    st[f"{c1}4"].alignment = center
    st[f"{c1}5"] = f
    st[f"{c1}5"].font = font(size=26, bold=True, color=NAVY)
    st[f"{c1}5"].fill = fill(SKY)
    st[f"{c1}5"].alignment = center
    st[f"{c1}5"].number_format = "0%" if label == "報到率" else "#,##0"
st.row_dimensions[5].height = 26
st.row_dimensions[6].height = 26

# 分組明細表
st["B8"] = "分組明細"
st["B8"].font = font(size=13, bold=True, color=NAVY)
st["B8"].border = bottom_only
for col in "CDEFGH":
    st[f"{col}8"].border = bottom_only
st_headers = ["組別", "報名人數", "眷屬人數", "總人數", "已報到", "未報到", "報到率"]
for i, h in enumerate(st_headers):
    cell = st.cell(row=10, column=2 + i, value=h)
    cell.font = font(bold=True, color="FFFFFF")
    cell.fill = fill(NAVY)
    cell.alignment = center
st.row_dimensions[10].height = 24

for i, (g, (fg, bg)) in enumerate(GROUP_COLOR.items()):
    r = 11 + i
    vals = [g, f'=COUNTIF({rng("E")},B{r})',
            f'=SUMIF({rng("E")},B{r},{rng("G")})', f"=C{r}+D{r}",
            f'=COUNTIFS({rng("E")},B{r},{rng("I")},"已報到")', f"=C{r}-F{r}",
            f"=IF(C{r}=0,0,F{r}/C{r})"]
    for j, v in enumerate(vals):
        cell = st.cell(row=r, column=2 + j, value=v)
        cell.font = font(bold=(j == 0), color=fg if j == 0 else "1C2833")
        cell.fill = fill(bg) if j == 0 else fill("FFFFFF")
        cell.alignment = center
        cell.border = border
    st.cell(row=r, column=8).number_format = "0%"
    st.row_dimensions[r].height = 22

tr = 14
st.cell(row=tr, column=2, value="合計")
for j, col in enumerate("CDEFG"):
    st.cell(row=tr, column=3 + j, value=f"=SUM({col}11:{col}13)")
st.cell(row=tr, column=8, value=f"=IF(C{tr}=0,0,F{tr}/C{tr})")
for c in range(2, 9):
    cell = st.cell(row=tr, column=c)
    cell.font = font(bold=True, color="FFFFFF")
    cell.fill = fill(TEAL)
    cell.alignment = center
st.cell(row=tr, column=8).number_format = "0%"
st.row_dimensions[tr].height = 24

# 報到率資料橫條
from openpyxl.formatting.rule import DataBarRule
st.conditional_formatting.add("H11:H13", DataBarRule(
    start_type="num", start_value=0, end_type="num", end_value=1, color="5DADE2"))

# ================= Sheet：填寫說明 =================
hp = wb.create_sheet("填寫說明")
banner(hp, "  📝 填寫說明", "  請先閱讀再開始填寫報名名單", "C")
hp.column_dimensions["A"].width = 2
hp.column_dimensions["B"].width = 16
hp.column_dimensions["C"].width = 72
rows = [
    ("項目", "說明"),
    ("深藍色表頭", "需要填寫：工號、姓名、部門、卡號、組別、眷屬姓名、備註"),
    ("灰色表頭", "自動欄位，請勿填寫：眷屬人數（公式）、報到狀態 / 報到時間（報到程式寫入）"),
    ("工號", "必填，報到時可直接輸入工號"),
    ("卡號", "必填，報到時刷卡用；欄位為文字格式，開頭的 0 會保留。從其他檔案貼上請用「選擇性貼上 → 值」"),
    ("組別", "下拉選單：比賽組（紅）/ 陪釣組（橘）/ 休閒組（綠），選完會自動上色"),
    ("眷屬姓名", "多位眷屬用頓號「、」分隔，例如：王太太、王小寶；沒有眷屬就留白"),
    ("重複檢查", "工號或卡號重複時，該格會自動變紅，請修正"),
    ("範例資料", "報名名單第 4–6 列灰色斜體為範例，正式填寫前請刪除"),
    ("未報名者", "報到台不開放現場加入，查無資料者請洽承辦人"),
]
for i, (a, b) in enumerate(rows):
    r = 4 + i
    for c, v in ((2, a), (3, b)):
        cell = hp.cell(row=r, column=c, value=v)
        cell.border = border
        cell.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        if i == 0:
            cell.font = font(bold=True, color="FFFFFF")
            cell.fill = fill(NAVY)
        else:
            cell.font = font(bold=(c == 2), color=NAVY if c == 2 else "1C2833")
            cell.fill = fill(SKY if c == 2 else ("FFFFFF" if i % 2 else GRAY_BG))
    hp.row_dimensions[r].height = 30 if i else 24

# 工作表分頁顏色與順序
ws.sheet_properties.tabColor = NAVY
st.sheet_properties.tabColor = TEAL
hp.sheet_properties.tabColor = "5DADE2"
wb.move_sheet("填寫說明", offset=-2)
wb.active = 1
wb.save(OUT)
print("saved")
