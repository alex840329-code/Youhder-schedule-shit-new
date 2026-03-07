import streamlit as st
import pandas as pd
import numpy as np
import json
import os
import calendar
import io
import collections
import random
import requests
from datetime import datetime, date, timedelta

# 嘗試匯入專業表格套件
try:
    from st_aggrid import AgGrid, GridOptionsBuilder, JsCode, GridUpdateMode
    HAS_AGGRID = True
except ImportError:
    HAS_AGGRID = False

# 嘗試匯入 AI 模組
try:
    import google.generativeai as genai
    HAS_AI_LIB = True
except ImportError:
    HAS_AI_LIB = False

# --- 頁面設定 ---
st.set_page_config(page_title="祐德牙醫排班系統 v20.4 (全面強健穩定版)", layout="wide", page_icon="🦷")
CONFIG_FILE = 'yude_config_v11.json'

# --- 阻擋未安裝套件的狀態 ---
if not HAS_AGGRID:
    st.error("🚨 **系統升級通知：需要安裝專業表格套件！** 🚨")
    st.markdown("""
    👉 **請確認您的 `requirements.txt` 中包含以下內容：**
    ```text
    streamlit
    pandas
    streamlit-aggrid
    requests
    ```
    安裝完成後，請重新整理此網頁或重新執行程式，即可看見完美的排班表格！
    """)
    st.stop()

# === 外觀鎖定區 (AgGrid 雙層漸層樣式) ===
st.markdown("""
<style>
    .ag-header-group-cell-label { justify-content: center !important; font-size: 15px !important; font-weight: bold !important; color: #000 !important; }
    .ag-header-cell-label { justify-content: center !important; font-size: 13px !important; font-weight: bold !important; color: #000 !important; }
    .header-odd { background-color: #FFD966 !important; border-right: 2px solid #333 !important; border-bottom: 1px solid #333 !important;}
    .header-even { background-color: #9DC3E6 !important; border-right: 2px solid #333 !important; border-bottom: 1px solid #333 !important;}
    .header-disabled { background-color: #e0e0e0 !important; border-right: 2px solid #333 !important; }
    .ag-cell-wrapper { display: flex; justify-content: center; align-items: center; height: 100%; width: 100%; }
</style>
""", unsafe_allow_html=True)

cell_style_js = JsCode("""
function(params) {
    var shift = params.colDef.headerName;
    var cellClass = params.colDef.cellClass || '';
    var isOdd = cellClass === 'is_odd';
    var val = params.value;
    var style = { 
        'textAlign': 'center', 'borderRight': '1px solid #d3d3d3', 'borderBottom': '1px solid #d3d3d3',
        'color': '#000', 'fontWeight': 'bold', 'display': 'flex', 'alignItems': 'center', 'justifyContent': 'center'
    };
    if (shift === '晚') style['borderRight'] = '2px solid #333'; 
    if (val === '-' || val === '⬛') {
        style['backgroundColor'] = '#f0f0f0'; style['color'] = '#ccc'; return style;
    }
    if (isOdd) {
        if (shift === '早') style['backgroundColor'] = '#FDE9D9';
        if (shift === '午') style['backgroundColor'] = '#FCD5B4';
        if (shift === '晚') style['backgroundColor'] = '#FABF8F';
    } else {
        if (shift === '早') style['backgroundColor'] = '#DDEBF7';
        if (shift === '午') style['backgroundColor'] = '#BDD7EE';
        if (shift === '晚') style['backgroundColor'] = '#9DC3E6';
    }
    return style;
}
""")

# --- 全域成功提示系統 ---
if "sys_msg" in st.session_state:
    st.success(st.session_state["sys_msg"])
    del st.session_state["sys_msg"]

# --- 1. 核心資料結構與初始化 ---
def get_default_config():
    return {
        "api_key": "", "is_locked": False, 
        "doctors_struct": [
            {"order": 1, "name": "郭長熀醫師", "nick": "郭", "active": True},
            {"order": 2, "name": "陳冰沁醫師", "nick": "沁", "active": True},
            {"order": 3, "name": "陳志鈴醫師", "nick": "鈴", "active": True},
            {"order": 4, "name": "陳哲毓醫師", "nick": "毓", "active": True},
            {"order": 5, "name": "陳奕安醫師", "nick": "安", "active": True},
            {"order": 6, "name": "吳峻豪醫師", "nick": "吳", "active": True},
            {"order": 7, "name": "蔡尚妤醫師", "nick": "蔡", "active": True},
            {"order": 8, "name": "陳貞羽醫師", "nick": "貞", "active": True},
            {"order": 9, "name": "吳麗君醫師", "nick": "麗", "active": True},
            {"order": 10, "name": "魏大鈞醫師", "nick": "魏", "active": True},
            {"order": 11, "name": "郭燿東醫師", "nick": "東", "active": True}
        ],
        "assistants_struct": [
            {"name": "雯萱", "nick": "萱", "active": True, "type": "全職", "custom_max": None, "pref": "normal", "is_main_counter": True},
            {"name": "小瑜", "nick": "瑜", "active": True, "type": "兼職", "custom_max": 20, "pref": "normal", "is_main_counter": True},
            {"name": "欣霓", "nick": "霓", "active": True, "type": "兼職", "custom_max": 15, "pref": "normal", "is_main_counter": True},
            {"name": "昀霏", "nick": "霏", "active": True, "type": "全職", "custom_max": None, "pref": "normal", "is_main_counter": False},
            {"name": "湘婷", "nick": "湘", "active": True, "type": "全職", "custom_max": None, "pref": "normal", "is_main_counter": False},
            {"name": "怡安", "nick": "怡", "active": True, "type": "全職", "custom_max": None, "pref": "normal", "is_main_counter": False},
            {"name": "嘉宜", "nick": "宜", "active": True, "type": "全職", "custom_max": None, "pref": "low", "is_main_counter": False},
            {"name": "芷瑜", "nick": "芷", "active": True, "type": "全職", "custom_max": None, "pref": "normal", "is_main_counter": False},
            {"name": "佳臻", "nick": "臻", "active": True, "type": "全職", "custom_max": None, "pref": "normal", "is_main_counter": False},
            {"name": "紫心", "nick": "紫", "active": True, "type": "全職", "custom_max": None, "pref": "normal", "is_main_counter": False},
            {"name": "又嘉", "nick": "又", "active": True, "type": "全職", "custom_max": None, "pref": "low", "is_main_counter": False},
            {"name": "佳萱", "nick": "佳", "active": True, "type": "全職", "custom_max": None, "pref": "normal", "is_main_counter": False},
            {"name": "紫媛", "nick": "媛", "active": True, "type": "全職", "custom_max": None, "pref": "normal", "is_main_counter": False},
            {"name": "暐貽", "nick": "貽", "active": True, "type": "兼職", "custom_max": 18, "pref": "normal", "is_main_counter": False}
        ],
        "pairing_matrix": {}, "adv_rules": {}, "template_odd": {}, "template_even": {},
        "year": datetime.today().year, "month": datetime.today().month % 12 + 1,
        "manual_schedule": [], "leaves": {}, "saved_result": {}
    }

def load_config():
    defaults = get_default_config()
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                for k, v in defaults.items():
                    if k not in data: data[k] = v
                return data
        except Exception: return defaults
    return defaults

def save_config(config):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except Exception as e: st.error(f"存檔發生錯誤: {e}")

if 'config' not in st.session_state:
    st.session_state.config = load_config()
    if st.session_state.config.get("saved_result"):
        st.session_state.result = st.session_state.config["saved_result"]

# --- 2. 日期與輔助函式 ---
def get_active_doctors():
    raw_docs = st.session_state.config.get("doctors_struct")
    if raw_docs is None: raw_docs = []
    # 增加強健性：過濾非 dict 項目並處理 None
    docs = sorted([d for d in raw_docs if isinstance(d, dict)], key=lambda x: x.get("order", 99))
    return [d for d in docs if d.get("active", True)]

def get_active_assistants():
    raw_assts = st.session_state.config.get("assistants_struct")
    if raw_assts is None: raw_assts = []
    # 增加強健性：過濾非 dict 項目
    return [a for a in raw_assts if isinstance(a, dict) and a.get("active", True)]

def generate_month_dates(year, month):
    num_days = calendar.monthrange(year, month)[1]
    return [date(year, month, d) for d in range(1, num_days + 1) if date(year, month, d).weekday() != 6]

def get_padded_weeks(year, month):
    first_day = date(year, month, 1)
    last_day = date(year, month, calendar.monthrange(year, month)[1])
    curr = first_day - timedelta(days=first_day.weekday()) 
    weeks = []
    while curr <= last_day or curr.weekday() != 0:
        if curr > last_day and curr.weekday() == 0: break
        week_dates = []
        for _ in range(7):
            if curr.weekday() != 6:
                is_curr = (curr.month == month)
                week_dates.append({"date": curr, "is_curr": is_curr, "str": str(curr),
                    "disp": f"{curr.month}/{curr.day}({['一','二','三','四','五','六'][curr.weekday()]})" if is_curr else f"⬛ {curr.month}/{curr.day}"})
            curr += timedelta(days=1)
        weeks.append(week_dates)
    return weeks

def calculate_shift_limits(year, month):
    dates = generate_month_dates(year, month)
    max_s = len(dates) * 2
    return max_s - 8, max_s

def parse_slot_string(text, is_fixed=False):
    wd_map = {"一":0, "二":1, "三":2, "四":3, "五":4, "六":5}
    sh_map = {"早":"早", "午":"午", "晚":"晚"}
    role_map = {"櫃":"counter", "流":"floater", "看":"look", "跟":"doctor", "行":"look"}
    if not text: return {} if is_fixed else set()
    items = [x.strip() for x in text.replace("、", ",").split(",") if x.strip()]
    if is_fixed:
        res = {}
        for item in items:
            if len(item) < 3: continue
            wd = wd_map.get(item[0]); sh = sh_map.get(item[1]); rl = role_map.get(item[2])
            if wd is not None and sh is not None and rl is not None: res[(wd, sh)] = rl
        return res
    res_set = set()
    for item in items:
        if len(item) < 2: continue
        wd = wd_map.get(item[0]); sh = sh_map.get(item[1])
        if wd is not None and sh is not None: res_set.add((wd, sh))
    return res_set

# --- 3. 核心排班演算法 (強化流動平均分配與兼職優先) ---
def run_auto_schedule(manual_schedule, leaves, pairing_matrix, adv_rules, ctr_count, flt_count):
    assts = get_active_assistants(); docs = get_active_doctors()
    year = st.session_state.config.get("year", datetime.today().year)
    month = st.session_state.config.get("month", datetime.today().month % 12 + 1)
    dates = generate_month_dates(year, month)
    std_min, std_max = calculate_shift_limits(year, month)
    
    p_targets = {}; p_limits = {}
    for a in assts:
        nm = a["name"]
        if a.get("type") == "全職":
            p_limits[nm] = std_max
            p_targets[nm] = std_min + 1 if a.get("pref") == "low" else std_max
        else:
            lim = a.get("custom_max") if a.get("custom_max") else 15
            p_limits[nm] = lim; p_targets[nm] = lim
    if "又嘉" in p_targets: p_targets["又嘉"] = max(0, std_max - 3)

    p_counts = {a["name"]: 0 for a in assts}
    p_floater_counts = {a["name"]: 0 for a in assts} # 追蹤流動診次
    p_daily = {a["name"]: collections.defaultdict(set) for a in assts}
    slots = sorted(list(set([f"{x['Date']}_{x['Shift']}" for x in manual_schedule])), 
                   key=lambda x: (x.split("_")[0], {"早":1,"午":2,"晚":3}.get(x.split("_")[1], 9)))
    result = {s: {"doctors": {}, "counter": [], "floater": [], "look": []} for s in slots}
    
    parsed_fixed = {}; parsed_admin = {}
    for name, r in adv_rules.items():
        if r.get("fixed_slots"): parsed_fixed[name] = parse_slot_string(r["fixed_slots"], is_fixed=True)
        if r.get("admin_slots"): parsed_admin[name] = parse_slot_string(r["admin_slots"], is_fixed=False)

    # 1. 填入固定班
    for slot in slots:
        dt_str, sh = slot.split("_"); wd = datetime.strptime(dt_str, "%Y-%m-%d").date().weekday()
        for name, fix_map in parsed_fixed.items():
            if (wd, sh) in fix_map:
                role = fix_map[(wd, sh)]
                if role in ["look", "counter", "floater"]:
                    result[slot][role].append(name); p_counts[name] += 1; p_daily[name][dt_str].add(sh)
                    if role == "floater": p_floater_counts[name] += 1
        for name, admin_set in parsed_admin.items():
            if (wd, sh) in admin_set: p_counts[name] += 1; p_daily[name][dt_str].add(sh)

    # 2. 自動演算
    for slot in slots:
        dt_str, sh = slot.split("_"); curr_dt = datetime.strptime(dt_str, "%Y-%m-%d").date(); wd = curr_dt.weekday()
        duty_docs = [x["Doctor"] for x in manual_schedule if x["Date"]==dt_str and x["Shift"]==sh]
        slot_res = result[slot]
        
        def assigned_in_slot(name):
            is_admin = (wd, sh) in parsed_admin.get(name, set())
            return name in slot_res["counter"] or name in slot_res["floater"] or name in slot_res["look"] or name in slot_res["doctors"].values() or is_admin

        def can_assign(name, role):
            if assigned_in_slot(name): return False
            if f"{name}_{dt_str}_{sh}" in leaves: return False
            if p_counts[name] >= p_limits[name]: return False 
            rule = adv_rules.get(name, {})
            r_lim = rule.get("role_limit", "無限制")
            if r_lim == "僅櫃台" and role != "counter": return False
            if r_lim == "僅流動" and role != "floater": return False
            if r_lim == "僅跟診" and role != "doctor": return False
            s_lim = rule.get("shift_limit", "無限制")
            if s_lim == "僅早班" and sh != "早": return False
            if s_lim == "僅午班" and sh != "午": return False
            if s_lim == "僅晚班" and sh != "晚": return False
            s_wl = parse_slot_string(rule.get("slot_whitelist", ""), is_fixed=False)
            if s_wl and (wd, sh) not in s_wl: return False
            
            # 互斥邏輯：僅針對櫃檯
            if role == "counter":
                avoids = [x.strip() for x in rule.get("avoid", "").split(",") if x.strip()]
                for av in avoids:
                    if av in slot_res["counter"]: return False

            today = p_daily[name][dt_str]
            if sh == "晚" and "早" in today and "午" not in today: return False 
            if sh == "晚" and "早" in today and "午" in today:
                yesterday = str(curr_dt - timedelta(days=1))
                if len(p_daily[name].get(yesterday, set())) == 3: return False 
            return True

        def calculate_priority(candidates, r_type):
            scored = []
            for c in candidates:
                if not can_assign(c, r_type): continue
                asst_info = next((a for a in assts if a["name"] == c), {})
                rule = adv_rules.get(c, {})
                
                score = (p_targets[c] - p_counts[c]) * 10
                
                # --- 特殊邏輯：櫃檯兼職優先 ---
                if r_type == "counter":
                    if asst_info.get("type") == "兼職":
                        score += 2000 
                    elif asst_info.get("is_main_counter"):
                        score += 500  
                
                # --- 特殊邏輯：流動診平均分配 ---
                if r_type == "floater":
                    if asst_info.get("is_main_counter"):
                        score -= 500 # 專職櫃檯不排流動
                    else:
                        score += (20 - p_floater_counts[c]) * 5 
                
                # 兼職與白名單優先級
                s_wl = parse_slot_string(rule.get("slot_whitelist", ""), is_fixed=False)
                if s_wl: score += (100 / max(1, len(s_wl)))
                
                if wd == 5: # 週六邏輯
                    sat_dates = [str(dt) for dt in dates if dt.weekday() == 5]
                    sat_nites = sum(1 for d in sat_dates if "晚" in p_daily[c][d])
                    if sh == "晚" and sat_nites < 2: score += 50
                    elif sh == "晚" and sat_nites >= 2: score -= 1000
                
                scored.append((c, score + random.random()))
            scored.sort(key=lambda x: x[1], reverse=True)
            return [x[0] for x in scored]

        cand_pool = [a["name"] for a in assts]
        # 醫師跟診
        for d_name in duty_docs:
            picked = None; targets = [pairing_matrix.get(d_name, {}).get(k) for k in ["1","2","3"]]
            for t in [x for x in targets if x]:
                if can_assign(t, "doctor"): picked = t; break
            if not picked:
                for c in calculate_priority(cand_pool, "doctor"): picked = c; break
            if picked: slot_res["doctors"][d_name] = picked; p_counts[picked] += 1; p_daily[picked][dt_str].add(sh)
            
        # 櫃檯
        needed_ctr = ctr_count - len(slot_res["counter"])
        for c in calculate_priority(cand_pool, "counter"):
            if needed_ctr <= 0: break
            slot_res["counter"].append(c); p_counts[c] += 1; p_daily[c][dt_str].add(sh); needed_ctr -= 1
            
        # 流動
        needed_flt = flt_count - len(slot_res["floater"])
        for c in calculate_priority(cand_pool, "floater"):
            if needed_flt <= 0: break
            slot_res["floater"].append(c); p_counts[c] += 1; p_daily[c][dt_str].add(sh); p_floater_counts[c] += 1; needed_flt -= 1
            
    return result

# --- 4. Excel 輸出 ---
def to_excel_master(schedule_result, year, month, docs, assts):
    output = io.BytesIO(); writer = pd.ExcelWriter(output, engine='xlsxwriter'); workbook = writer.book
    fmts = get_excel_formats(workbook); sheet = workbook.add_worksheet("總班表")
    dates = generate_month_dates(year, month); weeks = collections.defaultdict(list)
    for dt in dates: weeks[dt.isocalendar()[1]].append(dt)
    row = 0
    for wk_id, w_dates in enumerate(weeks.values()):
        sheet.merge_range(row, 0, row, len(w_dates)*3, f"祐德牙醫 {year}年{month}月 - 第 {wk_id+1} 週", fmts['h_title']); row += 1
        sheet.write(row, 0, "日期", fmts['h_col']); col = 1
        for dt in w_dates:
            sheet.merge_range(row, col, row, col+2, f"{dt.month}/{dt.day} ({['一','二','三','四','五','六'][dt.weekday()]})", fmts['h_col']); col += 3
        row += 1; sheet.write(row, 0, "時段", fmts['h_col']); col = 1
        for dt in w_dates:
            f = fmts['c_wknd'] if dt.weekday() == 5 else fmts['c_norm']
            for s in ["早","午","晚"]: sheet.write(row, col, s, f); col += 1
        row += 1
        for doc in docs:
            sheet.write(row, 0, doc["nick"], fmts['h_col']); col = 1
            for dt in w_dates:
                f = fmts['c_wknd'] if dt.weekday() == 5 else fmts['c_norm']
                for s in ["早","午","晚"]:
                    k = f"{dt}_{s}"; anm = schedule_result.get(k, {}).get("doctors", {}).get(doc["name"], "")
                    sheet.write(row, col, next((a["nick"] for a in assts if a["name"]==anm), ""), f); col += 1
            row += 1
        for rnm, rk, ri in [("櫃台1","counter",0), ("櫃台2","counter",1), ("流動","floater",0), ("看/行","look",0)]:
            sheet.write(row, 0, rnm, fmts['h_col']); col = 1
            for dt in w_dates:
                f = fmts['c_wknd'] if dt.weekday() == 5 else fmts['c_norm']
                for s in ["早","午","晚"]:
                    lst = schedule_result.get(f"{dt}_{s}", {}).get(rk, [])
                    sheet.write(row, col, next((a["nick"] for a in assts if a["name"]==lst[ri]), "") if ri < len(lst) else "", f); col += 1
            row += 1
        row += 2
    writer.close(); output.seek(0); return output

def get_excel_formats(workbook):
    return {
        'h_title': workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#D9E1F2', 'border': 1}),
        'h_col': workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#E0E0E0', 'border': 1}),
        'c_norm': workbook.add_format({'align': 'center', 'border': 1}),
        'c_wknd': workbook.add_format({'align': 'center', 'border': 1, 'bg_color': '#FFF2CC'})
    }

def to_excel_individual(schedule_result, year, month, assts, docs):
    output = io.BytesIO(); writer = pd.ExcelWriter(output, engine='xlsxwriter'); workbook = writer.book
    fmts = get_excel_formats(workbook); dates = generate_month_dates(year, month)
    b_min, b_max = calculate_shift_limits(year, month)
    for a in assts:
        s = workbook.add_worksheet(a["nick"]); anm = a["name"]; act = 0
        for k, v in schedule_result.items():
            if anm in (list(v["doctors"].values()) + v["counter"] + v["floater"] + v["look"]): act += 1
        s.write(0, 0, f"{anm} - {year}/{month}", fmts['h_title']); s.write(0, 8, f"上限: {a['custom_max'] or b_max}", fmts['c_norm']); s.write(1, 8, f"實排: {act}", fmts['c_norm'])
        for i, h in enumerate(["日期","星期","早","午","晚"]):
            s.write(2, i, h, fmts['h_col']); s.write(2, i+6, h, fmts['h_col'])
        mid = (len(dates)+1)//2
        for r, dt in enumerate(dates):
            col_off = 0 if r < mid else 6; row_off = r if r < mid else r - mid
            s.write(row_off+3, col_off, f"{dt.month}/{dt.day}", fmts['c_norm'])
            s.write(row_off+3, col_off+1, ['一','二','三','四','五','六'][dt.weekday()], fmts['c_norm'])
            for ci, sh in enumerate(["早","午","晚"]):
                v = ""; data = schedule_result.get(f"{dt}_{sh}", {})
                if anm in data.get("look", []): v="看"
                elif anm in data["floater"]: v="流"
                elif anm in data["counter"]: v="櫃"
                else:
                    for dn, asn in data.get("doctors", {}).items():
                        if asn == anm: v = next((d["nick"] for d in docs if d["name"]==dn), dn)
                s.write(row_off+3, col_off+2+ci, v, fmts['c_norm'])
    writer.close(); output.seek(0); return output

def to_excel_doctor_confirmed(manual_schedule, year, month, doc_name):
    output = io.BytesIO(); writer = pd.ExcelWriter(output, engine='xlsxwriter'); workbook = writer.book
    fmts = get_excel_formats(workbook); dates = generate_month_dates(year, month)
    s = workbook.add_worksheet("確認班表")
    s.merge_range(0, 0, 0, 4, f"{doc_name} - {year}/{month} 班表", fmts['h_title'])
    for i, h in enumerate(["日期","星期","早","午","晚"]): s.write(2, i, h, fmts['h_col'])
    for r, dt in enumerate(dates):
        s.write(r+3, 0, f"{dt.month}/{dt.day}", fmts['c_norm']); s.write(r+3, 1, ['一','二','三','四','五','六'][dt.weekday()], fmts['c_norm'])
        for ci, sh in enumerate(["早","午","晚"]):
            is_wk = any(x for x in manual_schedule if x["Date"]==str(dt) and x["Shift"]==sh and x["Doctor"]==doc_name)
            s.write(r+3, 2+ci, "看診" if is_wk else "", fmts['c_norm'])
    writer.close(); output.seek(0); return output

# --- 7. UI 介面 ---
is_locked_system = st.session_state.config.get("is_locked", False)

with st.sidebar:
    st.subheader("⚙️ 系統管理")
    lock_val = st.toggle("🔒 鎖定前台", value=is_locked_system)
    if lock_val != is_locked_system: st.session_state.config["is_locked"] = lock_val; save_config(st.session_state.config); st.rerun()

    st.divider()
    st.subheader("💾 資料備份與還原")
    y_cfg = st.session_state.config.get("year"); m_cfg = st.session_state.config.get("month")
    
    t_logic, t_month = st.tabs(["⚙️ 基本邏輯", "📅 當月班表"])
    with t_logic:
        logic_keys = ["doctors_struct", "assistants_struct", "pairing_matrix", "adv_rules", "template_odd", "template_even"]
        logic_data = {k: st.session_state.config.get(k) for k in logic_keys}
        st.download_button("📥 下載基本邏輯", json.dumps(logic_data, ensure_ascii=False, indent=4), f"yude_logic_config_{datetime.now().strftime('%Y%m%d')}.json", "application/json", use_container_width=True)
        up_logic = st.file_uploader("📤 還原邏輯", type="json", key="ul")
        if up_logic and st.button("確認還原邏輯", use_container_width=True):
            try:
                new_logic = json.load(up_logic)
                df_cfg = get_default_config()
                for k in logic_keys: 
                    val = new_logic.get(k)
                    st.session_state.config[k] = val if val is not None else df_cfg.get(k)
                save_config(st.session_state.config); st.session_state["sys_msg"]="✅ 邏輯還原成功！"; st.rerun()
            except: st.error("還原失敗")
            
    with t_month:
        month_keys = ["year", "month", "manual_schedule", "leaves", "saved_result"]
        if st.session_state.get("result"): st.session_state.config["saved_result"] = st.session_state.result
        month_data = {k: st.session_state.config.get(k) for k in month_keys}
        st.download_button("📥 下載當月班表", json.dumps(month_data, ensure_ascii=False, indent=4), f"yude_month_{y_cfg}_{m_cfg}_backup_{datetime.now().strftime('%m%d_%H%M')}.json", "application/json", use_container_width=True)
        up_month = st.file_uploader("📤 還原班表", type="json", key="um")
        if up_month and st.button("確認還原班表", use_container_width=True):
            try:
                new_month = json.load(up_month)
                df_cfg = get_default_config()
                for k in month_keys: 
                    val = new_month.get(k)
                    st.session_state.config[k] = val if val is not None else df_cfg.get(k)
                if st.session_state.config.get("saved_result"): 
                    st.session_state.result = st.session_state.config["saved_result"]
                save_config(st.session_state.config); st.session_state["sys_msg"]="✅ 班表還原成功！"; st.rerun()
            except: st.error("還原失敗")

step = st.sidebar.radio("導覽", ["1. 人員設定", "2. 跟診配對", "3. 進階限制", "4. 班表生成", "5. 醫師入口", "6. 助理入口", "7. 排班微調", "8. 報表下載"])

def safe_index(options, value, default=0):
    try: return options.index(value)
    except (ValueError, TypeError): return default

if step == "1. 人員設定":
    st.header("人員設定")
    y = st.session_state.config.get("year"); m = st.session_state.config.get("month")
    min_s, max_s = calculate_shift_limits(y, m)
    st.info(f"📅 {y}/{m} ｜ 全職標準：上限 {max_s}，基本 {min_s}")
    c1, c2 = st.columns(2)
    with c1:
        ed_doc = st.data_editor(pd.DataFrame(st.session_state.config.get("doctors_struct", [])), use_container_width=True, hide_index=True, num_rows="dynamic")
        if st.button("存醫師"): st.session_state.config["doctors_struct"] = ed_doc.to_dict('records'); save_config(st.session_state.config); st.rerun()
    with c2:
        ed_asst = st.data_editor(pd.DataFrame(st.session_state.config.get("assistants_struct", [])), column_config={"type": st.column_config.SelectboxColumn(options=["全職","兼職"]), "pref": st.column_config.SelectboxColumn(options=["high","normal","low"])}, use_container_width=True, hide_index=True, num_rows="dynamic")
        if st.button("存助理"): st.session_state.config["assistants_struct"] = ed_asst.replace({np.nan: None}).to_dict('records'); save_config(st.session_state.config); st.rerun()

elif step == "2. 跟診配對":
    st.header("跟診指定順位表")
    docs = get_active_doctors(); assts = [""] + [a["name"] for a in get_active_assistants()]
    matrix_data = []
    curr = st.session_state.config.get("pairing_matrix", {})
    for d in docs:
        row = {"醫師": d["name"]}; p = curr.get(d["name"], {})
        row["第一順位"] = p.get("1",""); row["第二順位"] = p.get("2",""); row["第三順位"] = p.get("3","")
        matrix_data.append(row)
    ed_mat = st.data_editor(pd.DataFrame(matrix_data), column_config={k: st.column_config.SelectboxColumn(options=assts) for k in ["第一順位","第二順位","第三順位"]}, use_container_width=True, hide_index=True)
    if st.button("儲存配對"):
        new_mat = {r["醫師"]: {"1":r["第一順位"],"2":r["第二順位"],"3":r["第三順位"]} for i, r in ed_mat.iterrows()}
        st.session_state.config["pairing_matrix"] = new_mat; save_config(st.session_state.config); st.rerun()

elif step == "3. 進階限制":
    st.header("🛡️ 助理進階動態鎖定")
    assts = get_active_assistants(); curr_rules = st.session_state.config.get("adv_rules", {})
    st.info("💡 **自動連動：** A 避開 B，儲存後 B 也會自動避開 A。行政診時段排班會自動跳過。")
    new_rules = {}
    
    role_options = ["無限制", "僅櫃台", "僅行政", "僅流動", "僅跟診"]
    shift_options = ["無限制", "僅早班", "僅午班", "僅晚班"]
    
    for a in assts:
        nm = a["name"]; r = curr_rules.get(nm, {})
        c1, c2, c3, c4, c5, c6 = st.columns([1, 1.5, 1.5, 2.5, 1.5, 1.5])
        c1.markdown(f"**{nm}**")
        
        rv = c2.selectbox("職位", role_options, index=safe_index(role_options, r.get("role_limit", "無限制")), key=f"r_{nm}", label_visibility="collapsed")
        sv = c3.selectbox("班別", shift_options, index=safe_index(shift_options, r.get("shift_limit", "無限制")), key=f"s_{nm}", label_visibility="collapsed")
        
        others = [x["name"] for x in assts if x["name"] != nm]
        av = c4.multiselect("避開", others, default=[x.strip() for x in r.get("avoid","").split(",") if x.strip() in others], key=f"v_{nm}", label_visibility="collapsed")
        with c5.popover("📅 白名單"):
            wl_s = parse_slot_string(r.get("slot_whitelist",""))
            wl_grid = []; days_map = ["一","二","三","四","五","六"]
            for d_idx in range(6):
                gc = st.columns(4); gc[0].write(days_map[d_idx])
                for si, sn in enumerate(["早","午","晚"]):
                    if gc[si+1].checkbox("", value=(d_idx, sn) in wl_s, key=f"wl_{nm}_{d_idx}_{sn}"): wl_grid.append(f"{days_map[d_idx]}{sn}")
        with c6.popover("💼 行政診"):
            ad_s = parse_slot_string(r.get("admin_slots",""))
            ad_grid = []
            for d_idx in range(6):
                gc = st.columns(4); gc[0].write(days_map[d_idx])
                for si, sn in enumerate(["早","午","晚"]):
                    if gc[si+1].checkbox("", value=(d_idx, sn) in ad_s, key=f"ad_{nm}_{d_idx}_{sn}"): ad_grid.append(f"{days_map[d_idx]}{sn}")
        new_rules[nm] = {"role_limit":rv, "shift_limit":sv, "avoid":",".join(av), "slot_whitelist":",".join(wl_grid), "admin_slots":",".join(ad_grid), "fixed_slots":r.get("fixed_slots","")}
        st.divider()
    if st.button("💾 儲存進階限制", type="primary"):
        for n, rl in new_rules.items():
            for t in [x.strip() for x in rl["avoid"].split(",") if x.strip()]:
                if t in new_rules:
                    t_av = [x.strip() for x in new_rules[t]["avoid"].split(",") if x.strip()]
                    if n not in t_av: t_av.append(n); new_rules[t]["avoid"] = ",".join(t_av)
        st.session_state.config["adv_rules"] = new_rules; save_config(st.session_state.config); st.rerun()

elif step == "4. 班表生成":
    st.header("醫師班表範本與初始化")
    c1, c2, c3 = st.columns(3)
    y = c1.number_input("年", 2025, 2030, st.session_state.config.get("year"))
    m = c2.number_input("月", 1, 12, st.session_state.config.get("month"))
    fws = c3.radio("第一週設定為：", ["單週","雙週"], horizontal=True)
    st.session_state.config["year"] = y; st.session_state.config["month"] = m
    doc_names = [d["name"] for d in get_active_doctors()]; days = ["一","二","三","四","五","六"]
    def render_ag(key):
        data = st.session_state.config.get(key, {}); rows = []
        for d in doc_names:
            r = {"doctor": f"👨‍⚕️ {d}"}; s = data.get(d, [False]*18)
            for i, dn in enumerate(days):
                for si, sn in enumerate(["早","午","晚"]): r[f"{dn}_{sn}"] = bool(s[i*3+si]) if len(s)==18 else False
            rows.append(r)
        df = pd.DataFrame(rows)
        cd = [{"headerName": "醫師", "field": "doctor", "pinned": "left", "width": 160, "cellStyle": {"fontWeight":"bold","borderRight":"2px solid #333","backgroundColor":"#fff"}}]
        for i, dn in enumerate(days):
            children = []
            for sn in ["早","午","晚"]:
                children.append({"headerName": sn, "field": f"{dn}_{sn}", "editable": True, "cellEditor": "agCheckboxCellEditor", "cellRenderer": "agCheckboxCellRenderer", "cellClass": "is_odd" if i%2==0 else "is_even", "cellStyle": cell_style_js, "width": 55})
            cd.append({"headerName": f"星期{dn}", "children": children, "headerClass": "header-odd" if i%2==0 else "header-even"})
        res = AgGrid(df, gridOptions={"columnDefs": cd, "rowHeight": 45}, allow_unsafe_jscode=True, update_mode=GridUpdateMode.MODEL_CHANGED, theme="alpine", key=f"ag_{key}")
        if res['data'] is not None:
            nr = {}; rd = res['data'].to_dict('records') if isinstance(res['data'], pd.DataFrame) else res['data']
            for row in rd:
                dn_clean = row["doctor"].replace("👨‍⚕️ ","")
                nr[dn_clean] = [bool(row.get(f"{d}_{s}", False)) for d in days for s in ["早","午","晚"]]
            st.session_state.config[key] = nr
    t1, t2 = st.tabs(["單週","雙週"])
    with t1: render_ag("template_odd")
    with t2: render_ag("template_even")
    if st.button("🚀 儲存並套用至本月", type="primary"):
        save_config(st.session_state.config); generated = []; dates = generate_month_dates(y, m)
        weeks = collections.defaultdict(list)
        for dt in dates: weeks[dt.isocalendar()[1]].append(dt)
        for wi, w_dates in enumerate(weeks.values()):
            tmpl = st.session_state.config.get("template_odd" if (wi%2==0 if fws=="單週" else wi%2!=0) else "template_even", {})
            for dt in w_dates:
                base = dt.weekday()*3
                for si, sn in enumerate(["早","午","晚"]):
                    for d in get_active_doctors():
                        if tmpl.get(d["name"]) and tmpl[d["name"]][base+si]: generated.append({"Date": str(dt), "Shift": sn, "Doctor": d["name"]})
        st.session_state.config["manual_schedule"] = generated; save_config(st.session_state.config); st.rerun()

elif step == "5. 醫師入口":
    st.header("👨‍⚕️ 醫師入口")
    docs = get_active_doctors()
    if docs:
        sel_doc = st.selectbox("📌 醫師名字", [d["name"] for d in docs])
        y, m = st.session_state.config.get("year"), st.session_state.config.get("month")
        manual = st.session_state.config.get("manual_schedule", [])
        p_weeks = get_padded_weeks(y, m); col_map = {}
        for wi, w_dates in enumerate(p_weeks):
            st.markdown(f"**第 {wi+1} 週**")
            cols = [d["disp"] for d in w_dates]; rows = []
            for d_info in w_dates:
                if d_info["is_curr"]: col_map[d_info["disp"]] = d_info["str"]
            for sn in ["早","午","晚"]:
                r = {"時段": sn}
                for c in cols: r[c] = any(x for x in manual if x["Date"]==col_map.get(c) and x["Shift"]==sn and x["Doctor"]==sel_doc)
                rows.append(r)
            df = pd.DataFrame(rows).set_index("時段")
            cfg = {c: st.column_config.CheckboxColumn(disabled=(not any(d["is_curr"] for d in w_dates if d["disp"]==c))) for c in cols}
            p_weeks[wi] = st.data_editor(df, column_config=cfg, use_container_width=True, key=f"doc_{sel_doc}_{wi}", disabled=is_locked_system)
        if not is_locked_system and st.button("💾 儲存班表修改", type="primary"):
            new_man = [x for x in manual if x["Doctor"] != sel_doc]
            st.session_state.config["manual_schedule"] = new_man; save_config(st.session_state.config); st.rerun()
        st.download_button("📥 下載確認版班表", to_excel_doctor_confirmed(manual, y, m, sel_doc), f"{sel_doc}_{m}月班表.xlsx", use_container_width=True)

elif step == "6. 助理入口":
    st.header("👩‍⚕️ 助理入口")
    assts = get_active_assistants(); sel_asst = st.selectbox("📌 助理名字", [a["name"] for a in assts])
    y, m = st.session_state.config.get("year"), st.session_state.config.get("month")
    leaves = st.session_state.config.get("leaves", {}); p_weeks = get_padded_weeks(y, m); col_map = {}
    for wi, w_dates in enumerate(p_weeks):
        st.markdown(f"**第 {wi+1} 週**")
        cols = [d["disp"] for d in w_dates]; rows = []
        for d_info in w_dates:
            if d_info["is_curr"]: col_map[d_info["disp"]] = d_info["str"]
        for sn in ["早","午","晚"]:
            r = {"時段": sn}
            for c in cols: r[c] = leaves.get(f"{sel_asst}_{col_map.get(c)}_{sn}", False)
            rows.append(r)
        df = pd.DataFrame(rows).set_index("時段")
        cfg = {c: st.column_config.CheckboxColumn(disabled=(not any(d["is_curr"] for d in w_dates if d["disp"]==c))) for c in cols}
        p_weeks[wi] = st.data_editor(df, column_config=cfg, use_container_width=True, key=f"asst_{sel_asst}_{wi}", disabled=is_locked_system)
    if not is_locked_system and st.button("💾 儲存休假修改", type="primary"):
        st.session_state.config["leaves"] = leaves; save_config(st.session_state.config); st.rerun()

elif step == "7. 排班微調":
    st.header("智慧排班與微調助手")
    
    y, m = st.session_state.config.get("year"), st.session_state.config.get("month")
    dates = generate_month_dates(y, m)
    std_min, std_max = calculate_shift_limits(y, m)
    assts = get_active_assistants()

    # === 側邊欄：即時監控儀表板 (包含流動診統計) ===
    with st.sidebar:
        st.markdown("---")
        st.subheader("📊 總管即時監控")
        if 'result' in st.session_state:
            curr_counts = {a["name"]: 0 for a in assts}
            curr_floaters = {a["name"]: 0 for a in assts}
            sat_dates = [str(dt) for dt in dates if dt.weekday() == 5]
            sat_stats = {a["name"]: {"full_off": 0, "nights": 0} for a in assts}
            daily_p = collections.defaultdict(lambda: collections.defaultdict(set))
            
            for k, v in st.session_state.result.items():
                dt_str, sh = k.split("_")
                ppl = list(v["doctors"].values()) + v["counter"] + v["floater"] + v["look"]
                for p in ppl: 
                    if p:
                        curr_counts[p] += 1
                        daily_p[p][dt_str].add(sh)
                        if p in v["floater"]: curr_floaters[p] += 1
            
            for a in assts:
                nm = a["name"]
                for d in sat_dates:
                    if not daily_p[nm][d]: sat_stats[nm]["full_off"] += 1
                    if "晚" in daily_p[nm][d]: sat_stats[nm]["nights"] += 1
                
                c_val = curr_counts[nm]; f_val = curr_floaters[nm]
                status_color = "green" if std_min <= c_val <= std_max else "red"
                st.markdown(f"**{nm}** ({a['type']})")
                st.markdown(f"- 總診: :{status_color}[{c_val}] | **流: {f_val}**")
                st.caption(f"- 週六: 休{sat_stats[nm]['full_off']} | 晚{sat_stats[nm]['nights']}")
                st.markdown("---")

    c1, c2 = st.columns(2); ctr = c1.slider("櫃台數", 1,3,2); flt = c2.slider("流動數",0,3,1)
    if st.button("🚀 執行自動排班", type="primary"):
        with st.spinner("運算中..."):
            res = run_auto_schedule(st.session_state.config["manual_schedule"], st.session_state.config["leaves"], st.session_state.config.get("pairing_matrix",{}), st.session_state.config.get("adv_rules",{}), ctr, flt)
            st.session_state.result = res; st.session_state.config["saved_result"] = res; save_config(st.session_state.config); st.rerun()
            
    if 'result' in st.session_state:
        # AI 助手
        with st.expander("🤖 Gemini AI 指令助手", expanded=False):
            api_key = st.text_input("Gemini API Key", value=st.session_state.config.get("api_key",""), type="password")
            cmd = st.text_area("口語化指令", placeholder="Ex: 峻豪醫師禮拜四都給昀霏跟診 / 小瑜3/21晚上休假")
            if st.button("讓 AI 調整"):
                if api_key:
                    st.session_state.config["api_key"] = api_key; save_config(st.session_state.config)
                    with st.spinner("AI 解析中..."):
                        try:
                            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-09-2025:generateContent?key={api_key}"
                            docs_str = ",".join([d["name"] for d in get_active_doctors()])
                            asst_str = ",".join([a["name"] for a in get_active_assistants()])
                            prompt = f"年月:{y}/{m}.醫師:{docs_str}.助理:{asst_str}.將指令轉為JSON格式動作清單(assign_assistant_to_doctor, doctor_leave, assistant_leave, assign_admin).指令:{cmd}"
                            payload = {"contents": [{"parts": [{"text": prompt}]}]}
                            resp = requests.post(url, json=payload); raw_txt = resp.json()['candidates'][0]['content']['parts'][0]['text']
                            
                            # 安全處理字串避免被截斷或當機
                            mk_j = "`" * 3 + "json"; mk_p = "`" * 3
                            clean = raw_txt.strip()
                            if clean.startswith(mk_j): clean = clean[len(mk_j):]
                            elif clean.startswith(mk_p): clean = clean[len(mk_p):]
                            if clean.endswith(mk_p): clean = clean[:-len(mk_p)]
                            
                            acts = json.loads(clean.strip())
                            st.session_state["sys_msg"] = f"✅ AI 建議套用 {len(acts)} 項變更 (請點擊執行排班套用)"; st.rerun()
                        except Exception as e: st.error(f"解析失敗: {e}")

        # AgGrid 排班微調
        docs = get_active_doctors(); p_weeks = get_padded_weeks(y, m)
        a_opts = [""] + [a["nick"] for a in assts]; nm2n = {a["name"]: a["nick"] for a in assts}; n2nm = {a["nick"]: a["name"] for a in assts}
        with st.form("adj"):
            all_grids = []
            for wi, w_dates in enumerate(p_weeks):
                st.markdown(f"#### 第 {wi+1} 週")
                rows = []
                for d in docs:
                    r = {"person": f"👨‍⚕️ {d['nick']}", "type":"doc", "name":d["name"]}
                    for dt in w_dates:
                        for s in ["早","午","晚"]:
                            f = f"{dt['str']}_{s}"
                            r[f] = nm2n.get(st.session_state.result.get(f, {}).get("doctors", {}).get(d["name"], ""), "") if dt["is_curr"] else "-"
                    rows.append(r)
                for rnm, rk, ri in [("櫃1","counter",0), ("櫃2","counter",1), ("流","floater",0), ("看","look",0)]:
                    r = {"person": rnm, "type":"role", "key":rk, "idx":ri}
                    for dt in w_dates:
                        for s in ["早","午","晚"]:
                            f = f"{dt['str']}_{s}"
                            if dt["is_curr"]:
                                lst = st.session_state.result.get(f, {}).get(rk, [])
                                r[f] = nm2n.get(lst[ri], "") if ri < len(lst) else ""
                            else: r[f] = "-"
                    rows.append(r)
                df = pd.DataFrame(rows)
                cd = [{"headerName": "人員", "field": "person", "pinned": "left", "width": 160, "editable": False, "cellStyle": {"fontWeight":"bold","borderRight":"2px solid #333","backgroundColor":"#fff"}}]
                for dt in w_dates:
                    child = []
                    for s in ["早","午","晚"]:
                        f = f"{dt['str']}_{s}"
                        child.append({"headerName": s, "field": f, "editable": dt["is_curr"], "cellEditor": "agSelectCellEditor", "cellEditorParams": {"values": a_opts}, "cellClass": "is_odd" if dt["date"].weekday()%2==0 else "is_even", "cellStyle": cell_style_js, "width": 55})
                    cd.append({"headerName": dt["disp"], "children": child, "headerClass": "header-odd" if dt["date"].weekday()%2==0 else "header-even"})
                res = AgGrid(df, gridOptions={"columnDefs": cd, "rowHeight": 40}, allow_unsafe_jscode=True, update_mode=GridUpdateMode.MODEL_CHANGED, theme="alpine", key=f"ag_adj_{wi}")
                all_grids.append((w_dates, res['data']))
            if st.form_submit_button("💾 儲存微調數據"):
                for w_dates, gd in all_grids:
                    if gd is not None:
                        rd = gd.to_dict('records') if isinstance(gd, pd.DataFrame) else gd
                        for row in rd:
                            pt, pn, pk, pi = row.get("type"), row.get("name"), row.get("key"), row.get("idx")
                            for d_info in w_dates:
                                if not d_info["is_curr"]: continue
                                for s in ["早","午","晚"]:
                                    k = f"{d_info['str']}_{s}"; vnm = n2nm.get(row.get(k, ""), "")
                                    if pt == "doc": st.session_state.result[k]["doctors"][pn] = vnm
                                    else:
                                        if pk not in st.session_state.result[k]: st.session_state.result[k][pk] = []
                                        while len(st.session_state.result[k][pk]) <= pi: st.session_state.result[k][pk].append("")
                                        st.session_state.result[k][pk][pi] = vnm
                st.session_state.config["saved_result"] = st.session_state.result; save_config(st.session_state.config); st.rerun()

elif step == "8. 報表下載":
    st.header("下載 Excel 報表")
    if 'result' in st.session_state:
        sch = st.session_state.result; y, m = st.session_state.config["year"], st.session_state.config["month"]
        d = get_active_doctors(); a = get_active_assistants()
        c1, c2 = st.columns(2)
        c1.download_button("📊 總班表", to_excel_master(sch, y, m, d, a), f"祐德總班表_{m}月.xlsx")
        c2.download_button("👤 助理個人表", to_excel_individual(sch, y, m, a, d), f"祐德助理表_{m}月.xlsx")
