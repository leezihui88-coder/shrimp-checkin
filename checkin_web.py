# 釣蝦活動報到：網頁版（Streamlit）
# - 刷卡機（鍵盤模擬）、輸入工號或輸入姓名找到同仁那一組
# - 勾選實際到場的人後確認（Enter、再刷一次卡或按「確認報到」）
# - 多台電腦可同時開網頁報到，共用同一份名單
# - 本機模式：程式資料夾有「釣蝦報名表.xlsx」就直接讀寫它
# - 雲端模式：沒有本機檔案時，從側欄上傳報名表，結果用下載按鈕取回
# 啟動：streamlit run checkin_web.py
import html
import os
import threading
import time
from datetime import datetime

import streamlit as st

from checkin_core import (PLAN_COLOR, CheckinLog, Roster, active, all_arrived, arrived,
                          ponds, rod_count, show_id)

APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOCAL_XLSX = os.path.join(APP_DIR, "釣蝦報名表.xlsx")      # 本機模式
WORK_DIR = os.path.join(APP_DIR, "data")                     # 雲端模式的工作資料夾
UPLOAD_XLSX = os.path.join(WORK_DIR, "報名表.xlsx")

EVENT_TITLE = "MOXA 釣蝦活動 報到"
RESET_SECONDS = 10          # 結果畫面停留秒數
CHOICE_TIMEOUT = 30         # 姓名搜尋清單停留秒數
DOUBLE_SWIPE_SECONDS = 5    # 確認後幾秒內同一組再刷，視為多刷一次
CARD_HINT_SECONDS = 90      # 刷卡失敗後，多久內改用工號報到會標註「卡號不符」
INPUT_LABEL = "報到輸入"
MAX_CHOICES = 6            # 姓名搜尋最多列幾組（字放大後一頁放得下）
CARD_VH = 60               # 結果卡高度（螢幕高度的百分比），整頁不需捲動

STATE_COLOR = {"ok": "#1E8449", "dup": "#B9770E", "err": "#B03A2E",
               "pick": "#1F5F8B", "choices": "#1F5F8B", "cancel": "#5D6D7E",
               "edit": "#6C3483",
               "idle": "#12304D"}
BADGE = {"釣蝦比賽": "red", "陪釣": "orange", "休閒釣蝦": "green"}

st.set_page_config(page_title="釣蝦報到", page_icon="🦐", layout="wide",
                   initial_sidebar_state="collapsed")


# =====================================================================
# 共用狀態：所有瀏覽器分頁共用同一份名單
# =====================================================================
@st.cache_resource
def registry():
    return {"state": None, "lock": threading.Lock()}


def load_state(path, log_path):
    """載入報名表；本機模式先備份原始檔"""
    reg = registry()
    with reg["lock"]:
        roster = Roster(path)
        bak = roster.backup(os.path.join(os.path.dirname(path), "備份"))
        log = CheckinLog(log_path)
        restored = log.restore(roster)
        reg["state"] = {"roster": roster, "log": log, "path": path, "backup": bak,
                        "restored": restored, "save_ok": None, "save_msg": ""}
        if restored:
            save_excel(reg["state"])
    return reg["state"]


def save_excel(state):
    """寫回報名表；失敗（多半是檔案被 Excel 開著）時記下狀態，下次報到會再寫"""
    try:
        state["roster"].save_back()
        state["save_ok"], state["save_msg"] = True, f"已寫回 Excel {datetime.now():%H:%M:%S}"
    except Exception as e:
        state["save_ok"] = False
        state["save_msg"] = f"Excel 寫入失敗（{type(e).__name__}），請關閉報名表檔案；報到資料已存於 CSV"


def current_state():
    """決定資料來源：本機檔案優先，否則使用上傳的檔案"""
    reg = registry()
    if os.path.exists(LOCAL_XLSX):
        path, log_path = LOCAL_XLSX, os.path.join(APP_DIR, "報到紀錄.csv")
    elif os.path.exists(UPLOAD_XLSX):
        path, log_path = UPLOAD_XLSX, os.path.join(WORK_DIR, "報到紀錄.csv")
    else:
        return None
    s = reg["state"]
    if s is None or s["path"] != path:
        s = load_state(path, log_path)
    return s


# =====================================================================
# 存取控制（選用）：在 .streamlit/secrets.toml 設 password = "..."
# =====================================================================
def check_password():
    try:
        pw = st.secrets.get("password")
    except Exception:
        pw = None
    if not pw or st.session_state.get("authed"):
        return True
    st.title("🦐 " + EVENT_TITLE)
    v = st.text_input("請輸入密碼", type="password")
    if v:
        if v == pw:
            st.session_state.authed = True
            st.rerun()
        st.error("密碼錯誤")
    return False


# =====================================================================
# 報到流程（每個瀏覽器分頁各自的畫面狀態放在 session_state.view）
# =====================================================================
def view():
    return st.session_state.setdefault("view", {"kind": "idle", "ts": time.time()})


def set_view(**kw):
    kw.setdefault("ts", time.time())
    st.session_state.view = kw


def chk_key(v, i):
    return f"chk_{v['emp']}_{v['nonce']}_{i}"


def start_pick(state, p, how, text, note=""):
    """開始勾選：第一次報到預設全勾；晚到補報到時，未到的人預設不勾"""
    first = not arrived(p)
    nonce = st.session_state.get("nonce", 0) + 1
    st.session_state.nonce = nonce
    v = {"kind": "pick", "emp": p["工號"], "how": how, "text": text,
         "note": note, "nonce": nonce}
    for i, m in enumerate(p["成員"]):
        if not m["報到時間"] and not m["取消"]:
            st.session_state[chk_key(v, i)] = first
    set_view(**v)
    beep("ok")


def toggle(state, i):
    v = view()
    p = state["roster"].by_id[v["emp"]]
    if not 0 <= i < len(p["成員"]):
        return
    m = p["成員"][i]
    # 勾選模式只能勾還沒到的人；修改模式已到的人也能取消
    if m["取消"] or (m["報到時間"] and v["kind"] != "edit"):
        return
    k = chk_key(v, i)
    st.session_state[k] = not st.session_state.get(k, False)


def start_edit(state, emp):
    """修改報到：列出全組（已取消的除外），目前已到的人預設打勾"""
    p = state["roster"].by_id[emp]
    nonce = st.session_state.get("nonce", 0) + 1
    st.session_state.nonce = nonce
    v = {"kind": "edit", "emp": emp, "nonce": nonce}
    for i, m in enumerate(p["成員"]):
        if not m["取消"]:
            st.session_state[chk_key(v, i)] = bool(m["報到時間"])
    set_view(**v)


def save_edit(state):
    v = view()
    if v["kind"] != "edit":
        return
    roster, log = state["roster"], state["log"]
    p = roster.by_id[v["emp"]]
    checks = [st.session_state.get(chk_key(v, i), False) for i in range(len(p["成員"]))]
    with registry()["lock"]:
        added, removed = roster.update(p, checks, "修改")
        if not added and not removed:
            set_view(kind="ok", emp=p["工號"], title="沒有變更", note="")
            return
        changes = "；".join(x for x in (
            "新增：" + "、".join(m["姓名"] for m in added) if added else "",
            "取消：" + "、".join(m["姓名"] for m in removed) if removed else "") if x)
        # 到場成員記「修改後的完整名單」，重開程式時才能正確重播
        log.write(輸入內容="修改報到", 工號=p["工號"], 姓名=p["員工姓名"], 桌號=p["桌號"],
                  到場成員="、".join(m["姓名"] for m in arrived(p)), 結果="修改報到",
                  附註="修改；" + changes)
        save_excel(state)
    set_view(kind="ok", emp=p["工號"], title="已修改報到", note=changes)
    beep("ok")


def confirm_pick(state):
    v = view()
    if v["kind"] != "pick":
        return
    roster, log = state["roster"], state["log"]
    p = roster.by_id[v["emp"]]
    how = {"card": "刷卡", "id": "工號"}.get(v["how"], v["how"])
    checks = [st.session_state.get(chk_key(v, i), False) for i in range(len(p["成員"]))]
    with registry()["lock"]:
        new = roster.confirm(p, checks, how)
        if not new:
            set_view(kind="err", title="未勾選任何人", msg="已取消，請重新報到")
            beep("err")
            return
        absent = [m["姓名"] for m in active(p) if not m["報到時間"]]
        log.write(輸入內容=v["text"], 工號=p["工號"], 姓名=p["員工姓名"], 桌號=p["桌號"],
                  到場成員="、".join(m["姓名"] for m in new), 結果="報到成功",
                  附註=how + (f"；{v['note']}" if v["note"] else "")
                  + (f"；未到：{'、'.join(absent)}" if absent else ""))
        save_excel(state)
    note = v["note"] or (f"尚未到：{'、'.join(absent)}（晚到時再刷卡 / 輸入工號補報到）"
                         if absent else "")
    set_view(kind="ok", emp=p["工號"], title=f"報到成功　本次 {len(new)} 位", note=note)
    beep("ok")


def handle_party(state, p, how, text):
    log = state["log"]
    if p["取消"]:
        # 全組已取消報名：不開放報到，請報到人員引導洽承辦人
        log.write(輸入內容=text, 工號=p["工號"], 姓名=p["員工姓名"], 結果="已取消報名")
        set_view(kind="cancel", emp=p["工號"])
        beep("err")
        return
    if all_arrived(p) or time.time() - p["last_confirm"] < DOUBLE_SWIPE_SECONDS:
        t = max(m["報到時間"] for m in arrived(p)).strftime("%H:%M")
        log.write(輸入內容=text, 工號=p["工號"], 姓名=p["員工姓名"], 桌號=p["桌號"],
                  結果="重複報到", 附註=f"最後報到 {t}")
        set_view(kind="dup", emp=p["工號"], t=t)
        beep("err")
        return
    note = ""
    bad = st.session_state.get("last_bad_card")
    if how == "id" and bad and time.time() - bad[1] < CARD_HINT_SECONDS:
        note = f"提醒：剛才刷的卡號 {bad[0]} 與報名資料不符"
    st.session_state.last_bad_card = None
    start_pick(state, p, how, text, note)


def process(state, text):
    """處理一次輸入（刷卡 / 工號 / 姓名 / 數字 / 空白 Enter）"""
    text = (text or "").strip()
    v = view()
    roster = state["roster"]

    # 勾選 / 修改 / 選擇模式下：空白 Enter 確認、單一數字切換或選擇
    if v["kind"] in ("pick", "edit"):
        if not text:
            (confirm_pick if v["kind"] == "pick" else save_edit)(state)
            return
        if text.isdigit() and len(text) == 1:
            toggle(state, int(text) - 1)
            return
    if v["kind"] == "choices" and text.isdigit() and len(text) == 1:
        choose(state, int(text) - 1)
        return
    if not text:
        return

    p, how = roster.find(text)
    if v["kind"] == "edit":
        # 修改中又輸入別人：修改要明確按「儲存修改」才算，這裡直接放棄修改
        set_view(kind="idle")
    if v["kind"] == "pick":
        if p is not None and p["工號"] == v["emp"]:   # 同一組再刷一次 = 確認
            confirm_pick(state)
            return
        confirm_pick(state)                             # 下一位已經刷卡：先確認上一組
    if p is None and not text.isdigit():
        found = roster.search_name(text)
        if len(found) == 1:
            p, how = found[0], "姓名"
        elif len(found) > 1:
            set_view(kind="choices", emps=[x["工號"] for x in found[:MAX_CHOICES]], total=len(found))
            return
    if p is None:
        is_card = roster.has_card and text.isdigit() and len(text) > 8
        state["log"].write(輸入內容=text, 結果="查無資料", 附註="卡號" if is_card else "工號 / 姓名")
        if is_card:
            st.session_state.last_bad_card = (text, time.time())
            set_view(kind="err", title="查無此卡號", msg="請改輸入工號")
        else:
            set_view(kind="err", title="查無報名資料", msg="請洽承辦人")
        beep("err")
        return
    handle_party(state, p, how, text)


def choose(state, i):
    v = view()
    emps = v.get("emps", [])
    if 0 <= i < len(emps):
        p = state["roster"].by_id[emps[i]]
        handle_party(state, p, "姓名", p["員工姓名"])


def cancel():
    set_view(kind="idle")


# =====================================================================
# 畫面元件
# =====================================================================
def beep(kind):
    """排一個提示音，在本次畫面輸出時播放"""
    st.session_state.beep = (kind, time.time())


def inject_css(card_color):
    """整頁固定在一個畫面內：尺寸都用螢幕高度（vh）計算，內容太多時只在結果卡內捲動"""
    c = CARD_VH
    st.markdown(f"""
<style>
html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"] {{ overflow:hidden !important; }}
.stApp {{ background:#0B1E33; }}
header[data-testid="stHeader"] {{ background:transparent; }}
.block-container {{ padding:.8vh 1.6vw 0 3vw !important; max-width:100% !important; }}
[data-testid="stMainBlockContainer"] [data-testid="stVerticalBlock"] {{ gap:1vh; }}
[data-testid="stMarkdownContainer"] {{ margin-bottom:0 !important; }}
.element-container:has(iframe), [data-testid="stElementContainer"]:has(iframe) {{
    position:absolute; width:0; height:0; overflow:hidden; }}
h1.title {{ color:#fff; font-size:4.6vh; margin:0; padding:0; line-height:1.25; }}
.clock {{ color:#9FB6CC; font-size:3.4vh; text-align:right; line-height:1.6; }}
.hint {{ color:#9FB6CC; font-size:2.4vh; margin-bottom:.3vh; }}
div[data-testid="stForm"] {{ background:#12304D; border:none; border-radius:10px; padding:1.2vh 1.5vw; }}
div[data-testid="stForm"] input {{ font-size:4.2vh !important; font-weight:700; text-align:center;
    height:6.8vh; background:#fff; color:#0B1E33; }}
div[data-testid="stFormSubmitButton"] {{ display:none; }}
div[data-testid="InputInstructions"] {{ display:none; }}
.st-key-card {{ background:{card_color}; border-radius:12px; padding:1.4vh 1.4vw;
    height:{c}vh; max-height:{c}vh; flex:0 0 auto; overflow:hidden; }}
.st-key-card [data-testid="stVerticalBlock"] {{ gap:.6vh; }}
.st-key-cardleft, [data-testid="stLayoutWrapper"]:has(> .st-key-cardleft) {{
    height:calc({c}vh - 4vh) !important; flex:0 0 auto !important; }}
.st-key-cardleft {{ display:flex; flex-direction:column; }}
.st-key-cardleft > [data-testid="stLayoutWrapper"]:has(> .st-key-actions),
.st-key-cardleft > div:has(> .st-key-actions) {{ margin-top:auto !important; }}
.st-key-actions {{ gap:.5vh !important; opacity:.92; }}
.st-key-actions .help {{ font-size:1.8vh !important; padding:.4vh 1vw; }}
.st-key-actions .stButton button, .st-key-actions .stButton button p {{ font-size:2vh !important; }}
.state {{ color:#fff; font-size:3.9vh; font-weight:800; text-align:center; margin-bottom:.3vh; }}
.pname {{ color:#fff; font-weight:800; text-align:center; line-height:1.15; }}
.psub {{ color:#fff; font-size:2.6vh; text-align:center; margin-bottom:1vh; }}
.tbl {{ background:#fff; border-radius:8px; text-align:center; padding:1.2vh .5vw;
    height:calc({c}vh - 4vh); box-sizing:border-box; display:flex; flex-direction:column;
    justify-content:space-evenly; }}
.tbl .cap {{ color:#0B1E33; font-size:3.8vh; font-weight:800; }}
.tbl .num {{ color:#0B1E33; font-size:11vh; font-weight:900; line-height:1.05; white-space:nowrap; }}
.tbl .none {{ color:#B03A2E; font-size:2.9vh; font-weight:800; }}
.tbl .sep {{ border-top:2px dashed #AEB6BF; margin:0 .6vw; }}
.tbl .pnum {{ color:#1F5F8B; }}
.tbl .pnone {{ color:#7F8C8D; font-size:2.4vh; font-weight:700; }}
.mrow {{ display:flex; align-items:center; gap:1vw; color:#fff; flex-wrap:wrap; line-height:1.5; }}
.mrow .mk {{ width:3vh; font-weight:800; }}
.mrow .nm {{ font-weight:800; min-width:6em; }}
.mrow .nm.gone {{ text-decoration:line-through; color:#AEB6BF; }}
.mrow .rel {{ min-width:4em; }}
.mrow .cancel {{ color:#AEB6BF; font-weight:800; }}
.badge {{ background:#fff; padding:0 .5vw; border-radius:4px; font-weight:800; }}
.pond {{ background:#0B1E33; color:#fff; border:2px solid #fff; padding:0 .5vw; border-radius:4px; font-weight:800; }}
.pond.none {{ background:transparent; border-color:transparent; color:#AEB6BF; }}
.rod {{ padding:0 .5vw; border-radius:4px; font-weight:800; white-space:nowrap; }}
.rod.borrow {{ background:#F7DC6F; color:#0B1E33; }}
.rod.own {{ border:2px solid #fff; color:#fff; }}
.note {{ color:#FDEBD0; font-size:2.4vh; text-align:center; margin-top:.6vh; }}
.help {{ background:#F7DC6F; color:#0B1E33; font-weight:800; font-size:2.3vh; text-align:center;
         border-radius:6px; padding:.6vh 1vw; }}
.st-key-card [data-testid="stCheckbox"] {{ zoom:1.3; }}
.st-key-card [data-testid="stCheckbox"] p {{ color:#fff; font-weight:800; font-size:2.6vh; }}
.st-key-card .stButton button {{ font-size:2.4vh; font-weight:800; min-height:0; padding:.4vh 1vw; line-height:1.4; }}
.st-key-card .stButton button p {{ font-size:2.4vh; }}
.stat {{ background:#12304D; border-radius:10px; padding:.7vh 1vw; border-left:8px solid; }}
.stat .t {{ color:#9FB6CC; font-size:2.1vh; }}
.stat .v {{ color:#fff; font-size:4.4vh; font-weight:800; line-height:1.15; }}
.stat .s {{ color:#9FB6CC; font-size:1.8vh; }}
.recent {{ background:#12304D; border-radius:10px; padding:1.4vh 1vw;
    height:calc({c}vh + 16.6vh); overflow-y:auto; box-sizing:border-box; }}
.recent h3 {{ color:#fff; font-size:2.7vh; margin:0 0 1vh 0; padding:0; }}
.recent .r {{ color:#fff; font-size:2.4vh; padding:.2vh 0 .2vh .6vw; line-height:1.4;
              border-left:5px solid; margin-bottom:.8vh; }}
.recent .r .tm {{ color:#9FB6CC; margin-right:.6vw; }}
.recent .r .tb {{ color:#9FB6CC; font-size:2.1vh; }}
.savebar {{ font-size:1.9vh; text-align:right; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
</style>""", unsafe_allow_html=True)


def keep_focus_and_beep():
    """在瀏覽器端：游標隨時拉回輸入框（刷卡不用先點滑鼠）、播放提示音"""
    b = st.session_state.pop("beep", None)
    sound = ""
    if b:
        f1, f2 = (880, 1320) if b[0] == "ok" else (300, 220)
        sound = f"""
try {{
  const C = new (window.AudioContext || window.webkitAudioContext)();
  [[{f1},0],[{f2},0.12]].forEach(([f,t]) => {{
    const o = C.createOscillator(), g = C.createGain();
    o.frequency.value = f; o.connect(g); g.connect(C.destination);
    g.gain.setValueAtTime(0.2, C.currentTime + t);
    o.start(C.currentTime + t); o.stop(C.currentTime + t + 0.11);
  }});
}} catch (e) {{}}  // {b[1]}"""
    st.iframe(f"""
<script>
const doc = window.parent.document;
function focusScan() {{
  const inp = doc.querySelector('input[aria-label="{INPUT_LABEL}"]');
  const a = doc.activeElement;
  const typing = a && (a.tagName === 'TEXTAREA' ||
                       (a.tagName === 'INPUT' && a.type !== 'checkbox' && a !== inp));
  if (inp && a !== inp && !typing) inp.focus();
}}
focusScan();
if (!window.__scanTimer) window.__scanTimer = setInterval(focusScan, 600);
{sound}
</script>""", height=1)


def plan_badge(plan, size):
    color = PLAN_COLOR.get(plan, "#2E9CCA")
    return f'<span class="badge" style="color:{color};font-size:{size}vh">{html.escape(plan or "—")}</span>'


def pond_badge(m, size):
    """每位參加人的蝦池；陪釣沒有蝦池"""
    if not m["蝦池"] or m["取消"]:
        return f'<span class="pond none" style="font-size:{size}vh">—</span>'
    return f'<span class="pond" style="font-size:{size}vh">{html.escape(m["蝦池"])} 號池</span>'


def rod_badge(m, size):
    """釣竿：需借（黃底）、自備（白框）"""
    rod = m["釣竿"]
    if not rod:
        return ""
    cls = "rod borrow" if "借" in rod else "rod own"
    return f'<span class="{cls}" style="font-size:{size}vh">{html.escape(rod)}</span>'


def sizes(n, pick):
    """人數多時縮小字體，5–6 位也能完整顯示（勾選畫面有按鈕與提示列，門檻較低）"""
    if n <= (2 if pick else 4):
        return 6.4, 3.6
    if n <= (3 if pick else 6):
        return 5.2, 3.1
    return 4.2, 2.55


def render_party_header(p, title, n_title=None):
    n = len(p["成員"])
    title_size, _ = sizes(n, n_title == "pick")
    act = active(p)
    got = len(arrived(p))
    borrow, own = rod_count(act)
    rods = []
    if borrow:
        rods.append(f'<span class="rod borrow">需借釣竿 {borrow} 支</span>')
    if own:
        rods.append(f'<span class="rod own">自備 {own} 支</span>')
    st.markdown(f'<div class="state">{title}</div>'
                f'<div class="pname" style="font-size:{title_size}vh">{html.escape(p["員工姓名"])}</div>'
                f'<div class="psub">工號 {html.escape(show_id(p))}　｜　已到 {got} / {len(act)} 位'
                f'{"　｜　" + " ".join(rods) if rods else ""}</div>',
                unsafe_allow_html=True)


def render_table(p):
    """右側白框：桌號（大字）+ 蝦池"""
    if p["桌號"]:
        tbl = f'<div class="num">{html.escape(p["桌號"])}</div>'
    else:
        tbl = '<div class="none">未安排<br>請洽承辦人</div>'
    ps = ponds(active(p))
    if ps:
        # 蝦池數字與桌號同大小（兩個池時用「・」分隔）
        pond = f'<div class="num pnum">{"・".join(map(html.escape, ps))}</div>'
    else:
        pond = '<div class="pnone">陪釣（無蝦池）</div>'
    st.markdown(f'<div class="tbl"><div class="cap">桌號</div>{tbl}'
                f'<div class="sep"></div><div class="cap">蝦池</div>{pond}</div>',
                unsafe_allow_html=True)


def member_cells(m, fs):
    """關係 / 方案 / 蝦池 / 釣竿；已取消的人只顯示「已取消」"""
    if m["取消"]:
        return f'<span class="rel">{m["關係"]}</span><span class="cancel">已取消</span>'
    return (f'<span class="rel">{m["關係"]}</span>{plan_badge(m["方案"], fs * .8)}'
            f'{pond_badge(m, fs * .8)}{rod_badge(m, fs * .75)}')


def render_members_static(p):
    _, fs = sizes(len(p["成員"]), False)
    rows = []
    for m in p["成員"]:
        if m["取消"]:
            mk = '<span style="color:#AEB6BF">－</span>'
        elif m["報到時間"]:
            mk = '<span style="color:#ABEBC6">✔</span>'
        else:
            mk = '<span style="color:#F5B7B1">✖</span>'
        name_cls = "nm gone" if m["取消"] else "nm"
        rows.append(f'<div class="mrow" style="font-size:{fs}vh"><span class="mk">{mk}</span>'
                    f'<span class="{name_cls}">{html.escape(m["姓名"])}</span>{member_cells(m, fs)}</div>')
    st.markdown("".join(rows), unsafe_allow_html=True)


def render_members_pick(state, p, v):
    _, fs = sizes(len(p["成員"]), True)
    edit = v["kind"] == "edit"
    for i, m in enumerate(p["成員"]):
        c1, c2 = st.columns([1.3, 3], vertical_alignment="center")
        # 勾選模式：已到的人鎖定；修改模式：已到的人也可以取消勾選
        if m["取消"] or (m["報到時間"] and not edit):
            mark = '<span class="mk" style="color:#ABEBC6">✔</span>' if m["報到時間"] else \
                '<span class="mk" style="color:#AEB6BF">－</span>'
            name_cls = "nm gone" if m["取消"] else "nm"
            c1.markdown(f'<div class="mrow" style="font-size:{fs}vh">{mark}'
                        f'<span class="{name_cls}">{i + 1}　{html.escape(m["姓名"])}</span></div>',
                        unsafe_allow_html=True)
        else:
            c1.checkbox(f"{i + 1}　{m['姓名']}", key=chk_key(v, i))
        cells = member_cells(m, fs)
        if m["報到時間"]:
            cells += f'<span style="color:#ABEBC6">{m["報到時間"]:%H:%M} 已到</span>'
        c2.markdown(f'<div class="mrow" style="font-size:{fs * .9}vh">{cells}</div>',
                    unsafe_allow_html=True)


def render_card(state):
    v = view()
    roster = state["roster"]
    kind = v["kind"]
    inject_css(STATE_COLOR[kind])
    with st.container(key="card"):
        if kind == "idle":
            st.markdown('<div class="state" style="color:#9FB6CC">歡迎參加！</div>'
                        f'<div class="pname" style="font-size:7.5vh;margin-top:4vh">'
                        f'{"請刷卡報到" if roster.has_card else "請輸入工號報到"}</div>'
                        '<div class="psub" style="color:#9FB6CC;margin-top:2vh">報到後會顯示您的桌號與蝦池</div>',
                        unsafe_allow_html=True)
            return
        if kind == "err":
            st.markdown(f'<div class="state">✖　{v["title"]}</div>'
                        f'<div class="pname" style="font-size:7.5vh;margin-top:4vh">{v["msg"]}</div>',
                        unsafe_allow_html=True)
            return
        if kind == "choices":
            more = f"（只列前 {MAX_CHOICES} 組，請輸入完整姓名）" if v["total"] > MAX_CHOICES else ""
            st.markdown(f'<div class="state">🔍　找到 {v["total"]} 組，請選擇{more}</div>',
                        unsafe_allow_html=True)
            for i, emp in enumerate(v["emps"]):
                p = roster.by_id[emp]
                fam = "、".join(m["姓名"] for m in p["成員"][1:])
                where = "已取消" if p["取消"] else f"{p['桌號'] or '-'} 桌"
                label = f"{i + 1}　{p['員工姓名']}（{show_id(p)}）　{where}" + \
                        (f"　眷屬：{fam}" if fam else "")
                st.button(label, key=f"choose_{i}", on_click=choose, args=(state, i))
            st.markdown('<div class="help">輸入數字＋Enter 或點選　｜　按「取消」回到首頁</div>',
                        unsafe_allow_html=True)
            st.button("取消", on_click=cancel)
            return

        p = roster.by_id[v["emp"]]
        if kind == "cancel":
            st.markdown('<div class="state">✖　已取消報名</div>'
                        f'<div class="pname" style="font-size:5vh">{html.escape(p["員工姓名"])}</div>'
                        f'<div class="psub">工號 {html.escape(show_id(p))}　｜　請洽承辦人</div>',
                        unsafe_allow_html=True)
            render_members_static(p)
            return
        left, right = st.columns([3, 1.1])
        with right:
            render_table(p)
        # 左半部：上面是人員資訊（重點），按鈕與操作提示固定在卡片最下方
        with left, st.container(key="cardleft"):
            if kind == "pick":
                n_sel = sum(1 for i, m in enumerate(p["成員"])
                            if not m["報到時間"] and not m["取消"]
                            and st.session_state.get(chk_key(v, i)))
                t = "請確認到場人員" if not arrived(p) else "補報到：請勾選剛到的人"
                render_party_header(p, f"👥　{t}（已勾 {n_sel} 位）", "pick")
                render_members_pick(state, p, v)
                if v.get("note"):
                    st.markdown(f'<div class="note">{v["note"]}</div>', unsafe_allow_html=True)
                with st.container(key="actions"):
                    if arrived(p):   # 補報到時，也可以改已到的人
                        b1, b2, b3, _ = st.columns([1.1, 0.7, 1.4, 0.8])
                        b3.button("📝 修改已報到的人", on_click=start_edit, args=(state, p["工號"]),
                                  width="stretch")
                    else:
                        b1, b2, _ = st.columns([1.1, 0.7, 2.2])
                    b1.button("✔ 確認報到", type="primary", on_click=confirm_pick, args=(state,),
                              width="stretch")
                    b2.button("取消", on_click=cancel, width="stretch")
                    confirm = "再刷一次卡 或 Enter 確認" if v["how"] == "card" else "空白 Enter 確認"
                    st.markdown(f'<div class="help">勾選：點名字，或輸入編號＋Enter　｜　{confirm}</div>',
                                unsafe_allow_html=True)
            elif kind == "edit":
                n_sel = sum(1 for i, m in enumerate(p["成員"])
                            if not m["取消"] and st.session_state.get(chk_key(v, i)))
                render_party_header(p, f"📝　修改報到：勾選實際到場的人（已勾 {n_sel} 位）", "pick")
                render_members_pick(state, p, v)
                with st.container(key="actions"):
                    b1, b2, _ = st.columns([1.1, 0.7, 2.2])
                    b1.button("✔ 儲存修改", type="primary", on_click=save_edit, args=(state,),
                              width="stretch")
                    b2.button("取消", on_click=cancel, width="stretch")
                    st.markdown('<div class="help">取消勾選＝撤銷報到　｜　勾選＝補上報到　｜　空白 Enter 儲存</div>',
                                unsafe_allow_html=True)
            elif kind in ("ok", "dup"):
                title = f"✔　{v['title']}" if kind == "ok" else f"⚠　已於 {v['t']} 報到過"
                render_party_header(p, title)
                render_members_static(p)
                if v.get("note"):
                    st.markdown(f'<div class="note">{v["note"]}</div>', unsafe_allow_html=True)
                with st.container(key="actions"):
                    st.button("📝 修改報到", key="edit_btn", on_click=start_edit, args=(state, p["工號"]))


def render_recent(state):
    roster = state["roster"]
    rows = []
    for t, p, k in roster.recent(300):   # 直接從名單算，多台電腦、修改報到都會即時反映
        color = PLAN_COLOR.get(p["成員"][0]["方案"] if p["成員"] else "", "#2E9CCA")
        # 每筆兩行：時間＋姓名＋人數 / 桌號＋蝦池，長名字也不會擠掉桌號
        ps = ponds(active(p))
        where = f'{p["桌號"] or "-"} 桌' + (f'　{"・".join(ps)} 號池' if ps else "")
        rows.append(f'<div class="r" style="border-color:{color}"><div><span class="tm">{t:%H:%M}</span>'
                    f'<b>{html.escape(p["員工姓名"])}　{k}位</b></div>'
                    f'<div class="tb">{where}</div></div>')
    st.markdown(f'<div class="recent"><h3>最近報到</h3>{"".join(rows)}</div>',
                unsafe_allow_html=True)


def render_stats(state):
    roster = state["roster"]
    s = roster.stats()
    plans = roster.plans()
    cols = st.columns(len(plans) + 1)
    for c, g in zip(cols, plans):
        got, tot = s.get(g, [0, 0])
        c.markdown(f'<div class="stat" style="border-color:{PLAN_COLOR.get(g, "#2E9CCA")}">'
                   f'<div class="t">{g}</div><div class="v">{got} / {tot}</div>'
                   f'<div class="s">人（含眷屬）</div></div>', unsafe_allow_html=True)
    parties = [p for p in roster.parties if not p["取消"]]
    all_got = sum(len(arrived(p)) for p in parties)
    all_tot = sum(len(active(p)) for p in parties)
    grp = sum(1 for p in parties if arrived(p))
    cols[-1].markdown(f'<div class="stat" style="border-color:#2E9CCA"><div class="t">合計</div>'
                      f'<div class="v">{all_got} / {all_tot}</div>'
                      f'<div class="s">已報到 {grp} / {len(parties)} 組</div></div>',
                      unsafe_allow_html=True)


# =====================================================================
# 側欄：資料來源、下載結果、未報到名單、異常紀錄
# =====================================================================
def sidebar(state):
    sb = st.sidebar
    sb.header("管理")
    local = os.path.exists(LOCAL_XLSX)
    if state:
        sb.caption(("本機模式：" if local else "上傳模式：") + os.path.basename(state["path"]))
        if state["save_ok"] is True:
            sb.success(state["save_msg"])
        elif state["save_ok"] is False:
            sb.warning(state["save_msg"])
            if sb.button("重新寫回 Excel"):
                with registry()["lock"]:
                    save_excel(state)
                st.rerun()
        with open(state["path"], "rb") as f:
            sb.download_button("⬇ 下載報到結果 Excel", f.read(),
                               file_name=f"報到結果_{datetime.now():%Y%m%d_%H%M}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               width="stretch")
    if not local:
        sb.divider()
        sb.subheader("上傳報名表")
        sb.caption("檔案只存在執行中的伺服器上；App 重新啟動後需重新上傳。"
                   "上傳已下載的報到結果 Excel，可接續之前的報到狀態。")
        up = sb.file_uploader("選擇 .xlsx", type=["xlsx"], label_visibility="collapsed")
        if up is not None and sb.button("使用這份報名表", type="primary"):
            os.makedirs(WORK_DIR, exist_ok=True)
            with open(UPLOAD_XLSX, "wb") as f:
                f.write(up.getbuffer())
            try:
                load_state(UPLOAD_XLSX, os.path.join(WORK_DIR, "報到紀錄.csv"))
                set_view(kind="idle")
                st.rerun()
            except Exception as e:
                sb.error(f"讀取失敗：{e}")
    if not state:
        return
    sb.divider()
    roster = state["roster"]
    with sb.expander("未報到名單"):
        def key(p):
            t = p["桌號"]
            return (0, int(t)) if t.isdigit() else (1, t)
        rows = []
        for p in sorted(roster.parties, key=key):
            miss = [m for m in active(p) if not m["報到時間"]]
            if miss:
                rows.append({"桌號": p["桌號"] or "未安排", "工號": show_id(p), "同仁": p["員工姓名"],
                             "蝦池": "・".join(ponds(miss)) or "—",
                             "狀態": "全組未到" if len(miss) == len(active(p)) else "部分未到",
                             "未到人員": "、".join(f"{m['姓名']}({m['關係']})" for m in miss)})
        st.caption(f"{len(rows)} 組")
        st.dataframe(rows, hide_index=True, width="stretch")
    with sb.expander("異常紀錄"):
        rows = [{k: r.get(k, "") for k in ("時間", "結果", "輸入內容", "姓名", "附註")}
                for r in state["log"].rows() if r.get("結果") not in ("報到成功", "修改報到")]
        st.caption(f"{len(rows)} 筆")
        st.dataframe(rows, hide_index=True, width="stretch")
    if roster.warnings:
        with sb.expander("資料提醒"):
            for w in roster.warnings:
                st.write("・" + w)
    with sb.expander("重設報到（測試用）"):
        st.caption("清除所有報到紀錄（CSV 與 Excel 的報到欄），活動當天請勿使用。")
        if st.checkbox("我確定要清除") and st.button("清除所有報到", type="primary"):
            with registry()["lock"]:
                for p in roster.parties:
                    p["last_confirm"] = 0
                    for m in p["成員"]:
                        m["報到時間"], m["報到方式"] = None, ""
                os.remove(state["log"].path)
                state["log"] = CheckinLog(state["log"].path)
                save_excel(state)
            set_view(kind="idle")
            st.rerun()


# =====================================================================
# 主畫面
# =====================================================================
@st.fragment(run_every=2)
def clock_and_timeout():
    """右上角時鐘；結果畫面停留一段時間後自動回到首頁"""
    st.markdown(f'<div class="clock">{datetime.now():%m/%d  %H:%M:%S}</div>',
                unsafe_allow_html=True)
    v = view()
    limit = CHOICE_TIMEOUT if v["kind"] == "choices" else RESET_SECONDS
    if v["kind"] in ("ok", "dup", "err", "choices") and time.time() - v["ts"] > limit:
        set_view(kind="idle")
        st.rerun(scope="app")


def main():
    if not check_password():
        return
    state = current_state()
    sidebar(state)

    h1, h_save, h2 = st.columns([2.2, 1.8, 1], vertical_alignment="center")
    h1.markdown(f'<h1 class="title">🦐 {EVENT_TITLE}</h1>', unsafe_allow_html=True)
    save_ph = h_save.empty()          # Excel 寫回狀態，報到處理完才填入
    with h2:
        clock_and_timeout()

    if state is None:
        inject_css(STATE_COLOR["idle"])
        st.info("尚未載入報名表：請展開左側「管理」側欄（左上角 >），上傳報名表 Excel。")
        return

    roster = state["roster"]
    main_col, side_col = st.columns([3.2, 1])
    with main_col:
        hint = "請刷卡，或輸入工號 / 姓名後按 Enter" if roster.has_card \
            else "請輸入工號或姓名後按 Enter"
        with st.form("scan", clear_on_submit=True, border=False):
            st.markdown(f'<div class="hint">{hint}</div>', unsafe_allow_html=True)
            text = st.text_input(INPUT_LABEL, label_visibility="collapsed")
            submitted = st.form_submit_button("送出")
        if submitted:
            process(state, text)
        render_card(state)
    with side_col:
        render_recent(state)
    render_stats(state)
    if state["save_ok"] is False:
        save_ph.markdown(f'<div class="savebar" style="color:#F5B041">⚠ {state["save_msg"]}</div>',
                         unsafe_allow_html=True)
    elif state["save_ok"]:
        save_ph.markdown(f'<div class="savebar" style="color:#58D68D">✔ {state["save_msg"]}</div>',
                         unsafe_allow_html=True)
    keep_focus_and_beep()


main()
