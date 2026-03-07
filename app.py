import streamlit as st
import pandas as pd
import numpy as np
import json
import os
import calendar
import io
import collections
import random
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
st.set_page_config(page_title="祐德牙醫排班系統 v20.0 (分離備份與完美保存版)", layout="wide", page_icon="🦷")
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
    google-generativeai
    ```
    安裝完成後，請重新整理此網頁或重新執行程式，即可看見完美的排班表格！
    """)
    st.stop()

# === 外觀鎖定區 (絕對不動) ===
st.markdown("""
<style>
    /* AgGrid 標題置中與樣式 */
    .ag-header-group-cell-label { justify-content: center !important; font-size: 15px !important; font-weight: bold !important; color: #000 !important; }
    .ag-header-cell-label { justify-content: center !important; font-size: 13px !important; font-weight: bold !important; color: #000 !important; }
    
    /* 標題列背景顏色與粗線 (還原 Excel 視覺) */
    .header-odd { background-color: #FFD966 !important; border-right: 2px solid #333 !important; border-bottom: 1px solid #333 !important;}
    .header-even { background-color: #9DC3E6 !important; border-right: 2px solid #333 !important; border-bottom: 1px solid #333 !important;}
    .header-disabled { background-color: #e0e0e0 !important; border-right: 2px solid #333 !important; }
    
    /* 確保打勾框置中 */
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
        'textAlign': 'center', 
        'borderRight': '1px solid #d3d3d3',
        'borderBottom': '1px solid #d3d3d3',
        'color': '#000',
        'fontWeight': 'bold',
        'display': 'flex',
        'alignItems': 'center',
        'justifyContent': 'center'
    };
    
    // 星期之間的粗黑線
    if (shift === '晚') {
        style['borderRight'] = '2px solid #333'; 
    }
    
    // 跨月的反黑無效區
    if (val === '-' || val === '⬛') {
        style['backgroundColor'] = '#f0f0f0';
        style['color'] = '#ccc';
        return style;
    }
    
    // 單數日漸層 (橘黃暖色)
    if (isOdd) {
        if (shift === '早') style['backgroundColor'] = '#FDE9D9';
        if (shift === '午') style['backgroundColor'] = '#FCD5B4';
        if (shift === '晚') style['backgroundColor'] = '#FABF8F';
    } 
    // 偶數日漸層 (藍冷色)
    else {
        if (shift === '早') style['backgroundColor'] = '#DDEBF7';
        if (shift === '午') style['backgroundColor'] = '#BDD7EE';
        if (shift === '晚') style['backgroundColor'] = '#9DC3E6';
    }
    
    return style;
}
""")
# ========================

# --- 全域成功提示系統 ---
if "sys_msg" in st.session_state:
    st.success(st.session_state["sys_msg"])
    del st.session_state["sys_msg"]

# --- 1. 核心資料結構與初始化 ---
def get_default_config():
    return {
        "api_key": "", 
        "is_locked": False, 
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
        "pairing_matrix": {
            "郭長熀醫師": {"1": "又嘉", "2": "紫心", "3": "怡安"},
            "陳冰沁醫師": {"1": "嘉宜", "2": "芷瑜", "3": ""},
            "陳志鈴醫師": {"1": "紫媛", "2": "芷瑜", "3": ""},
            "陳哲毓醫師": {"1": "佳萱", "2": "", "3": ""},
            "陳奕安醫師": {"1": "昀霏", "2": "", "3": ""},
            "吳峻豪醫師": {"1": "湘婷", "2": "", "3": ""},
            "蔡尚妤醫師": {"1": "佳臻", "2": "", "3": ""},
            "陳貞羽醫師": {"1": "怡安", "2": "", "3": ""},
            "吳麗君醫師": {"1": "又嘉", "2": "芷瑜", "3": ""},
            "魏大鈞醫師": {"1": "又嘉", "2": "", "3": ""},
            "郭燿東醫師": {"1": "芷瑜", "2": "嘉宜", "3": "昀霏"}
        },
        "adv_rules": {
            "雯萱": {"role_limit": "無限制", "shift_limit": "無限制", "slot_whitelist": "", "fixed_slots": "一早櫃,一晚櫃,二早櫃,三早櫃,四午櫃,五早櫃,五午櫃,五晚櫃", "avoid": "", "admin_slots": "二午,三午"},
            "小瑜": {"role_limit": "僅櫃台", "shift_limit": "僅晚班", "slot_whitelist": "", "fixed_slots": "", "avoid": "怡安", "admin_slots": ""},
            "欣霓": {"role_limit": "僅櫃台", "shift_limit": "無限制", "slot_whitelist": "一午,二午,四晚", "fixed_slots": "", "avoid": "", "admin_slots": ""},
            "暐貽": {"role_limit": "僅流動", "shift_limit": "無限制", "slot_whitelist": "二晚,三晚,四晚,六午,六晚", "fixed_slots": "", "avoid": "", "admin_slots": ""},
            "怡安": {"role_limit": "無限制", "shift_limit": "無限制", "slot_whitelist": "", "fixed_slots": "", "avoid": "小瑜", "admin_slots": ""},
            "紫媛": {"role_limit": "無限制", "shift_limit": "無限制", "slot_whitelist": "", "fixed_slots": "", "avoid": "昀霏", "admin_slots": ""},
            "昀霏": {"role_limit": "無限制", "shift_limit": "無限制", "slot_whitelist": "", "fixed_slots": "", "avoid": "紫媛", "admin_slots": ""}
        },
        "template_odd": {}, 
        "template_even": {},
        "year": datetime.today().year,
        "month": datetime.today().month % 12 + 1,
        "manual_schedule": [], 
        "clinic_holidays": [], 
        "leaves": {},
        "saved_result": {} # 新增：用來將排班結果持久化存檔
    }

def load_config():
    defaults = get_default_config()
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
                for k, v in defaults.items():
                    if k not in data: data[k] = v
                if "assistants_struct" in data:
                    for a in data["assistants_struct"]:
                        if "pref" not in a: a["pref"] = "normal"
                        if "type" not in a: a["type"] = "全職"
                        if "is_main_counter" not in a: a["is_main_counter"] = False
                if "adv_rules" in data:
                    for k, v in data["adv_rules"].items():
                        if "admin_slots" not in v: v["admin_slots"] = ""
                return data
        except Exception: return defaults
    return defaults

def save_config(config):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except Exception as e: st.error(f"存檔發生錯誤: {e}")

# 初始化載入設定與排班結果
if 'config' not in st.session_state:
    st.session_state.config = load_config()
    # 如果有儲存過的排班微調結果，自動還原到 session_state 避免遺失
    if "saved_result" in st.session_state.config and st.session_state.config["saved_result"]:
        if 'result' not in st.session_state:
            st.session_state.result = st.session_state.config["saved_result"]

# --- 2. 日期與輔助函式 ---
def get_active_doctors():
    docs = sorted(st.session_state.config.get("doctors_struct", []), key=lambda x: x.get("order", 99))
    return [d for d in docs if d.get("active", True)]

def get_active_assistants():
    return [a for a in st.session_state.config.get("assistants_struct", []) if a.get("active", True)]

def generate_month_dates(year, month):
    num_days = calendar.monthrange(year, month)[1]
    dates = []
    for d in range(1, num_days + 1):
        dt = date(year, month, d)
        if dt.weekday() == 6: continue 
        dates.append(dt)
    return dates

def get_padded_weeks(year, month):
    first_day = date(year, month, 1)
    last_day = date(year, month, calendar.monthrange(year, month)[1])
    start_date = first_day - timedelta(days=first_day.weekday()) 
    weeks = []; current_date = start_date
    while current_date <= last_day or current_date.weekday() != 0:
        if current_date > last_day and current_date.weekday() == 0: break
        week_dates = []
        for _ in range(7):
            if current_date.weekday() != 6:
                is_curr_month = (current_date.month == month)
                week_dates.append({
                    "date": current_date, "is_curr": is_curr_month, "str": str(current_date),
                    "disp": f"{current_date.month}/{current_date.day}({['一','二','三','四','五','六'][current_date.weekday()]})" if is_curr_month else f"⬛ {current_date.month}/{current_date.day}"
                })
            current_date += timedelta(days=1)
        weeks.append(week_dates)
    return weeks

def calculate_shift_limits(year, month):
    dates = generate_month_dates(year, month)
    max_s = len(dates) * 2
    return max_s - 8, max_s

def parse_slot_string(text, is_fixed=False):
    wd_map = {"一":0, "二":1, "三":2, "四":3, "五":4, "六":5}
    shift_map = {"早":"早", "午":"午", "晚":"晚"}
    role_map = {"櫃":"counter", "流":"floater", "看":"look", "跟":"doctor", "行":"look"} 
    if not text: return {} if is_fixed else set()
    items = [x.strip() for x in text.replace("、", ",").split(",") if x.strip()]
    if is_fixed:
        res = {}
        for item in items:
            if len(item) < 3: continue
            wd = wd_map.get(item[0]); sh = shift_map.get(item[1]); rl = role_map.get(item[2])
            if wd is not None and sh is not None and rl is not None: res[(wd, sh)] = rl
        return res
    else:
        res = set()
        for item in items:
            if len(item) < 2: continue
            wd = wd_map.get(item[0]); sh = shift_map.get(item[1])
            if wd is not None and sh is not None: res.add((wd, sh))
        return res

# --- 3. 核心排班演算法 ---
def run_auto_schedule(manual_schedule, leaves, pairing_matrix, adv_rules, ctr_count, flt_count):
    assts = get_active_assistants()
    docs = get_active_doctors()
    year = st.session_state.config.get("year", datetime.today().year)
    month = st.session_state.config.get("month", datetime.today().month % 12 + 1)
    dates = generate_month_dates(year, month)
    
    std_min, std_max = calculate_shift_limits(year, month)
    total_sats = len([dt for dt in dates if dt.weekday() == 5])
    main_counters = [a["name"] for a in assts if a.get("is_main_counter", False)]
    
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
    p_daily = {a["name"]: collections.defaultdict(set) for a in assts} 
    
    shift_order = {"早": 1, "午": 2, "晚": 3}
    slots = sorted(list(set([f"{x['Date']}_{x['Shift']}" for x in manual_schedule])), 
                   key=lambda x: (x.split("_")[0], shift_order.get(x.split("_")[1], 99)))
                   
    result = {s: {"doctors": {}, "counter": [], "floater": [], "look": []} for s in slots}
    
    parsed_fixed = {}
    parsed_admin = {}
    for name, r in adv_rules.items():
        if r.get("fixed_slots"): parsed_fixed[name] = parse_slot_string(r["fixed_slots"], is_fixed=True)
        if r.get("admin_slots"): parsed_admin[name] = parse_slot_string(r["admin_slots"], is_fixed=False)

    for slot in slots:
        dt_str, sh = slot.split("_")
        wd = datetime.strptime(dt_str, "%Y-%m-%d").date().weekday()
        for name, fix_map in parsed_fixed.items():
            if (wd, sh) in fix_map:
                role = fix_map[(wd, sh)]
                if role == "look": result[slot]["look"].append(name)
                elif role == "counter": result[slot]["counter"].append(name)
                elif role == "floater": result[slot]["floater"].append(name)
                if role in ["look", "counter", "floater"]:
                    p_counts[name] += 1
                    p_daily[name][dt_str].add(sh)
                    
        for name, admin_set in parsed_admin.items():
            if (wd, sh) in admin_set:
                p_counts[name] += 1
                p_daily[name][dt_str].add(sh)

    for slot in slots:
        dt_str, sh = slot.split("_")
        curr_dt = datetime.strptime(dt_str, "%Y-%m-%d").date()
        wd = curr_dt.weekday()
        
        duty_docs = [x["Doctor"] for x in manual_schedule if x["Date"]==dt_str and x["Shift"]==sh]
        d_order = {d["name"]: d["order"] for d in docs}
        duty_docs.sort(key=lambda x: d_order.get(x, 99))
        slot_res = result[slot]
        
        def assigned_in_slot(name):
            is_admin = (wd, sh) in parsed_admin.get(name, set())
            return name in slot_res["counter"] or name in slot_res["floater"] or name in slot_res["look"] or name in slot_res["doctors"].values() or is_admin

        def can_assign(name, role):
            if assigned_in_slot(name): return False
            if f"{name}_{slot}" in leaves: return False
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
            
            s_wl_str = rule.get("slot_whitelist", "")
            if s_wl_str:
                if (wd, sh) not in parse_slot_string(s_wl_str, is_fixed=False): return False

            avoid_str = rule.get("avoid", "")
            if avoid_str:
                for av in [x.strip() for x in avoid_str.split(",") if x.strip()]:
                    if assigned_in_slot(av): return False 

            today_shifts = p_daily[name][dt_str]
            if sh == "晚" and "早" in today_shifts and "午" not in today_shifts: return False 
            if sh == "晚" and "早" in today_shifts and "午" in today_shifts:
                yesterday_str = str(curr_dt - timedelta(days=1))
                if len(p_daily[name].get(yesterday_str, set())) == 3: return False 

            return True

        def calculate_priority(candidates, curr_wd, curr_sh, curr_dt_str):
            scored = []
            for c in candidates:
                if not can_assign(c, "floater"): continue 
                
                gap = p_targets[c] - p_counts[c]
                score = gap * 10 + random.random() * 2 
                
                if curr_wd == 5:
                    is_ft = next((a["type"] == "全職" for a in assts if a["name"] == c), False)
                    if is_ft:
                        sat_dates = [str(dt) for dt in dates if dt.weekday() == 5]
                        sats_worked = sum(1 for d in sat_dates if p_daily[c][d])
                        sat_nights = sum(1 for d in sat_dates if "晚" in p_daily[c][d])
                        working_today = bool(p_daily[c][curr_dt_str])
                        
                        if curr_sh == "晚":
                            if sat_nights >= 2: score -= 1000 
                            elif sat_nights < 2: score += 50  
                        else:
                            if not working_today and sats_worked >= total_sats - 1:
                                score -= 1000 
                                
                scored.append((c, score))
            scored.sort(key=lambda x: x[1], reverse=True)
            return [x[0] for x in scored]

        candidates_pool = [a["name"] for a in assts]
        
        for doc_name in duty_docs:
            if doc_name in slot_res["doctors"] and slot_res["doctors"][doc_name]: continue 
            picked = None
            prefs = pairing_matrix.get(doc_name, {})
            targets = [t for t in [prefs.get("1"), prefs.get("2"), prefs.get("3")] if t]
            
            for t in targets:
                if can_assign(t, "doctor"): picked = t; break
            
            if not picked:
                for c in calculate_priority(candidates_pool, wd, sh, dt_str):
                    if can_assign(c, "doctor"): picked = c; break
                    
            if picked:
                slot_res["doctors"][doc_name] = picked
                p_counts[picked] += 1
                p_daily[picked][dt_str].add(sh)
            else: slot_res["doctors"][doc_name] = ""

        needed_ctr = ctr_count - len(slot_res["counter"])
        if needed_ctr > 0:
            has_main = any(c in main_counters for c in slot_res["counter"])
            if not has_main:
                for c in calculate_priority(main_counters, wd, sh, dt_str):
                    if can_assign(c, "counter"):
                        slot_res["counter"].append(c)
                        p_counts[c] += 1
                        p_daily[c][dt_str].add(sh)
                        needed_ctr -= 1
                        break
            for c in calculate_priority(candidates_pool, wd, sh, dt_str):
                if needed_ctr <= 0: break
                if can_assign(c, "counter"):
                    slot_res["counter"].append(c)
                    p_counts[c] += 1
                    p_daily[c][dt_str].add(sh)
                    needed_ctr -= 1
        
        needed_flt = flt_count - len(slot_res["floater"])
        if needed_flt > 0:
            for c in calculate_priority(candidates_pool, wd, sh, dt_str):
                if needed_flt <= 0: break
                if can_assign(c, "floater"):
                    slot_res["floater"].append(c)
                    p_counts[c] += 1
                    p_daily[c][dt_str].add(sh)
                    needed_flt -= 1

        result[slot] = slot_res

    return result, p_counts, std_min, std_max

# --- 4. Excel 輸出 ---
def get_excel_formats(workbook):
    return {
        'h_title': workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 14, 'bg_color': '#D9E1F2', 'border': 1}),
        'h_col': workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#E0E0E0', 'border': 1}),
        'c_norm': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1}),
        'c_wknd': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFF2CC'}),
        'n_note': workbook.add_format({'align': 'left', 'valign': 'top', 'text_wrap': True})
    }

def to_excel_master(schedule_result, year, month, docs, assts):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    fmts = get_excel_formats(workbook)
    sheet = workbook.add_worksheet("總班表")
    
    dates = generate_month_dates(year, month)
    weeks_dict = collections.defaultdict(list)
    for dt in dates: weeks_dict[dt.isocalendar()[1]].append(dt)
    
    current_row = 0
    shifts = ["早", "午", "晚"]
    
    for wk_idx, w_dates in enumerate(weeks_dict.values()):
        sheet.merge_range(current_row, 0, current_row, len(w_dates)*3, f"祐德牙醫 {year}年{month}月 - 第 {wk_idx+1} 週", fmts['h_title'])
        current_row += 1
        
        sheet.write(current_row, 0, "日期", fmts['h_col'])
        col = 1
        for dt in w_dates:
            sheet.merge_range(current_row, col, current_row, col+2, f"{dt.month}/{dt.day} ({['一','二','三','四','五','六'][dt.weekday()]})", fmts['h_col'])
            col += 3
        current_row += 1
        
        sheet.write(current_row, 0, "時段", fmts['h_col'])
        col = 1
        for dt in w_dates:
            fmt = fmts['c_wknd'] if dt.weekday() == 5 else fmts['c_norm']
            for s in shifts:
                sheet.write(current_row, col, s, fmt); col += 1
        current_row += 1
        
        for doc in docs:
            sheet.write(current_row, 0, doc["nick"], fmts['h_col'])
            col = 1
            for dt in w_dates:
                fmt = fmts['c_wknd'] if dt.weekday() == 5 else fmts['c_norm']
                for s in shifts:
                    k = f"{dt}_{s}"; v = ""
                    if k in schedule_result:
                        v = schedule_result[k]["doctors"].get(doc["name"], "")
                        for a in assts: 
                            if a["name"]==v: v=a["nick"]; break
                    sheet.write(current_row, col, v, fmt); col += 1
            current_row += 1
            
        roles = [("櫃台1", "counter", 0), ("櫃台2", "counter", 1), ("流動", "floater", 0), ("看/行", "look", 0)]
        for rname, rkey, ridx in roles:
            sheet.write(current_row, 0, rname, fmts['h_col'])
            col = 1
            for dt in w_dates:
                fmt = fmts['c_wknd'] if dt.weekday() == 5 else fmts['c_norm']
                for s in shifts:
                    k = f"{dt}_{s}"; v = ""
                    if k in schedule_result:
                        lst = schedule_result[k].get(rkey, [])
                        if ridx < len(lst):
                            nm = lst[ridx]
                            for a in assts:
                                if a["name"]==nm: v=a["nick"]; break
                            if not v: v = nm
                    sheet.write(current_row, col, v, fmt); col += 1
            current_row += 1
        current_row += 2 
        
    writer.close()
    output.seek(0)
    return output

def to_excel_individual(schedule_result, year, month, assts, docs):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    fmts = get_excel_formats(workbook)
    
    dates = generate_month_dates(year, month); mid = (len(dates) + 1) // 2
    dL, dR = dates[:mid], dates[mid:]
    b_min, b_max = calculate_shift_limits(year, month)
    note = "註：全診及午晚班有空請輪流抽空吃飯，謹守30分鐘規定。\n1〉早午班 8:30-12:00 13:30-18:00\n2〉午晚班 13:30-22:00\n3〉早晚班 08:00-12:00 18:00-22:00"
    
    for a in assts:
        s = workbook.add_worksheet(a["nick"])
        aname = a["name"]; act = 0
        for k, v in schedule_result.items():
            ppl = list(v["doctors"].values()) + v["counter"] + v["floater"] + v.get("look",[])
            if aname in ppl: act += 1
            
        lim = a["custom_max"] if a["custom_max"] is not None else b_max
        s.merge_range(0, 0, 0, 4, f"{aname} - {year}年{month}月", fmts['h_title'])
        s.write(0, 8, f"應排: {lim}", fmts['c_norm'])
        s.write(1, 8, f"實排: {act}", fmts['c_norm'])
        
        for i, h in enumerate(["日期","星期","早","午","晚"]):
            s.write(2, i, h, fmts['h_col']); s.write(2, i+6, h, fmts['h_col'])
            
        def fill(d_lst, off):
            for r, dt in enumerate(d_lst):
                row = r + 3
                s.write(row, off, f"{dt.month}/{dt.day}", fmts['c_norm'])
                s.write(row, off+1, ['一','二','三','四','五','六'][dt.weekday()], fmts['c_norm'])
                for c, sh in enumerate(["早", "午", "晚"]):
                    k = f"{dt}_{sh}"; v = ""
                    if k in schedule_result:
                        data = schedule_result[k]
                        if aname in data.get("look", []): v="看"
                        elif aname in data["floater"]: v="流"
                        elif aname in data["counter"]: v="櫃"
                        else:
                            for dn, asg in data["doctors"].items():
                                if asg == aname: v = next((d["nick"] for d in docs if d["name"]==dn), dn)
                    s.write(row, off+2+c, v, fmts['c_norm'])
                    
        fill(dL, 0); fill(dR, 6)
        s.merge_range(max(len(dL), len(dR))+5, 0, max(len(dL), len(dR))+10, 10, note, fmts['n_note'])
    writer.close()
    output.seek(0)
    return output

def to_excel_doctor_confirmed(manual_schedule, year, month, doc_name):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    fmts = get_excel_formats(workbook)
    
    dates = generate_month_dates(year, month); mid = (len(dates) + 1) // 2
    dL, dR = dates[:mid], dates[mid:]
    
    s = workbook.add_worksheet("確認班表")
    s.merge_range(0, 0, 0, 4, f"{doc_name} - {year}年{month}月 確認班表", fmts['h_title'])
    
    for i, h in enumerate(["日期","星期","早","午","晚"]):
        s.write(2, i, h, fmts['h_col']); s.write(2, i+6, h, fmts['h_col'])
        
    def fill(d_lst, off):
        for r, dt in enumerate(d_lst):
            row = r + 3
            s.write(row, off, f"{dt.month}/{dt.day}", fmts['c_norm'])
            s.write(row, off+1, ['一','二','三','四','五','六'][dt.weekday()], fmts['c_norm'])
            for c, sh in enumerate(["早", "午", "晚"]):
                is_working = any(x for x in manual_schedule if x["Date"] == str(dt) and x["Shift"] == sh and x["Doctor"] == doc_name)
                s.write(row, off+2+c, "看診" if is_working else "", fmts['c_norm'])
                
    fill(dL, 0); fill(dR, 6)
    writer.close()
    output.seek(0)
    return output

# --- 7. UI 介面 ---
st.title("🦷 祐德牙醫 - 智慧排班系統 v19.4 (備份防呆版)")

is_locked_system = st.session_state.config.get("is_locked", False)

with st.sidebar:
    st.divider()
    st.subheader("⚙️ 系統權限管理")
    new_lock_state = st.toggle("🔒 鎖定前台修改 (Deadline)", value=is_locked_system, help="開啟後，醫師與助理將無法更改假單。")
    if new_lock_state != is_locked_system:
        st.session_state.config["is_locked"] = new_lock_state; save_config(st.session_state.config); st.rerun()

    # === 分離式備份系統 ===
    st.divider()
    st.subheader("💾 分離式資料備份與還原")
    st.info("💡 雲端伺服器重啟會遺失資料。已將備份拆分為「基本邏輯」與「當月班表」，方便跨月沿用設定！")
    
    tab_logic, tab_month = st.tabs(["⚙️ 基本邏輯 (前4項)", "📅 當月班表 (後4項)"])
    
    with tab_logic:
        logic_data = {k: st.session_state.config.get(k) for k in ["doctors_struct", "assistants_struct", "pairing_matrix", "adv_rules", "template_odd", "template_even"]}
        st.download_button("📥 下載【基本邏輯】備份檔", json.dumps(logic_data, ensure_ascii=False, indent=4), f"yude_logic_backup_{datetime.now().strftime('%Y%m%d')}.json", mime="application/json", use_container_width=True)
        
        up_logic = st.file_uploader("📤 上傳還原【基本邏輯】", type="json", key="up_logic")
        if up_logic and st.button("⚠️ 確認還原邏輯", type="primary", use_container_width=True):
            try:
                new_logic = json.load(up_logic)
                for k in ["doctors_struct", "assistants_struct", "pairing_matrix", "adv_rules", "template_odd", "template_even"]:
                    if k in new_logic: st.session_state.config[k] = new_logic[k]
                save_config(st.session_state.config)
                st.session_state["sys_msg"] = "✅ 基本邏輯還原成功！"
                st.rerun()
            except:
                st.error("檔案格式錯誤")
                
    with tab_month:
        if 'result' in st.session_state:
            st.session_state.config['saved_result'] = st.session_state.result
        month_data = {k: st.session_state.config.get(k) for k in ["year", "month", "manual_schedule", "leaves", "saved_result"]}
        st.download_button("📥 下載【當月班表】備份檔", json.dumps(month_data, ensure_ascii=False, indent=4), f"yude_month_backup_{datetime.now().strftime('%Y%m%d')}.json", mime="application/json", use_container_width=True)
        
        up_month = st.file_uploader("📤 上傳還原【當月班表】", type="json", key="up_month")
        if up_month and st.button("⚠️ 確認還原班表", type="primary", use_container_width=True):
            try:
                new_month = json.load(up_month)
                for k in ["year", "month", "manual_schedule", "leaves", "saved_result"]:
                    if k in new_month: st.session_state.config[k] = new_month[k]
                if "saved_result" in new_month and new_month["saved_result"]:
                    st.session_state.result = new_month["saved_result"]
                save_config(st.session_state.config)
                st.session_state["sys_msg"] = "✅ 當月班表還原成功！"
                st.rerun()
            except:
                st.error("檔案格式錯誤")

step = st.sidebar.radio("導覽步驟", [
    "1. 系統與人員設定", "2. 醫師配對順位", "3. 助理進階限制", "4. 醫師範本與生成", 
    "5. 👨‍⚕️ 醫師專屬入口", "6. 👩‍⚕️ 助理專屬入口", "7. 排班與總管微調", "8. 報表下載"
])

if step == "1. 系統與人員設定":
    st.header("人員與權重設定")
    y = st.session_state.config.get("year", datetime.today().year)
    m = st.session_state.config.get("month", datetime.today().month % 12 + 1)
    min_s, max_s = calculate_shift_limits(y, m)
    st.info(f"📅 {y}年{m}月 ｜ 全職標準：上限 {max_s} 診，基本 {min_s} 診\n\n💡 **小提示：** 表格最下方有個 `+` 號，點擊就可以**無限新增**人員。")
    
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("👨‍⚕️ 醫師名單")
        ed_doc = st.data_editor(pd.DataFrame(st.session_state.config.get("doctors_struct", get_default_config()["doctors_struct"])), use_container_width=True, hide_index=True, num_rows="dynamic")
        if st.button("存醫師"): st.session_state.config["doctors_struct"] = ed_doc.to_dict('records'); save_config(st.session_state.config); st.session_state["sys_msg"]="✅ 醫師名單儲存成功！"; st.rerun()
    with c2:
        st.subheader("👩‍⚕️ 助理名單")
        ed_asst = st.data_editor(pd.DataFrame(st.session_state.config.get("assistants_struct", get_default_config()["assistants_struct"])), column_config={
            "type": st.column_config.SelectboxColumn("全/兼職", options=["全職","兼職"]),
            "custom_max": st.column_config.NumberColumn("兼職上限", min_value=0),
            "pref": st.column_config.SelectboxColumn("偏好", options=["high","normal","low"]),
            "is_main_counter": st.column_config.CheckboxColumn("主櫃台?")
        }, use_container_width=True, hide_index=True, num_rows="dynamic")
        if st.button("存助理"): st.session_state.config["assistants_struct"] = ed_asst.replace({np.nan: None}).to_dict('records'); save_config(st.session_state.config); st.session_state["sys_msg"]="✅ 助理名單儲存成功！"; st.rerun()

elif step == "2. 醫師配對順位":
    st.header("跟診指定順位表")
    docs = get_active_doctors(); assts = [""] + [a["name"] for a in get_active_assistants()]
    matrix_data = []
    curr = st.session_state.config.get("pairing_matrix", {})
    for doc in docs:
        row = {"醫師": doc["name"]}; prefs = curr.get(doc["name"], {})
        row["第一順位"] = prefs.get("1", ""); row["第二順位"] = prefs.get("2", ""); row["第三順位"] = prefs.get("3", "")
        matrix_data.append(row)
    ed_mat = st.data_editor(pd.DataFrame(matrix_data), column_config={
        "醫師": st.column_config.TextColumn(disabled=True),
        "第一順位": st.column_config.SelectboxColumn(options=assts),
        "第二順位": st.column_config.SelectboxColumn(options=assts),
        "第三順位": st.column_config.SelectboxColumn(options=assts)
    }, hide_index=True, use_container_width=True)
    if st.button("儲存配對"):
        new_mat = {}
        for idx, row in ed_mat.iterrows(): new_mat[row["醫師"]] = {"1": row["第一順位"], "2": row["第二順位"], "3": row["第三順位"]}
        st.session_state.config["pairing_matrix"] = new_mat; save_config(st.session_state.config); st.session_state["sys_msg"]="✅ 配對順位儲存成功！"; st.rerun()

elif step == "3. 助理進階限制":
    st.header("🛡️ 助理進階動態鎖定")
    st.info("💡 **獨立行政診：** 您可以為每位助理設定特定的「行政診」時段，排班演算法將會避開這些時段！")
    
    assts = get_active_assistants()
    curr_rules = st.session_state.config.get("adv_rules", {})
    
    # 建立表頭
    h1, h2, h3, h4, h5, h6 = st.columns([1.5, 2, 2, 3, 2, 2])
    h1.markdown("**助理**"); h2.markdown("**限定職位**"); h3.markdown("**限定班別**"); h4.markdown("**避開人員**"); h5.markdown("**限定時段(白名單)**"); h6.markdown("**行政診時段**")
    st.markdown("<hr style='margin:0 0 10px 0;'>", unsafe_allow_html=True)
    
    new_rules = {}
    
    for a in assts:
        nm = a["name"]
        r = curr_rules.get(nm, {})
        c1, c2, c3, c4, c5, c6 = st.columns([1.5, 2, 2, 3, 2, 2])
        
        c1.markdown(f"<div style='padding-top:10px; font-weight:bold;'>{nm}</div>", unsafe_allow_html=True)
        
        roles = ["無限制", "僅櫃台", "僅行政", "僅流動", "僅跟診"]
        def_role = r.get("role_limit", "無限制")
        role_val = c2.selectbox("職位", roles, index=roles.index(def_role) if def_role in roles else 0, key=f"role_{nm}", label_visibility="collapsed")
        
        shifts = ["無限制", "僅早班", "僅午班", "僅晚班"]
        def_shift = r.get("shift_limit", "無限制")
        shift_val = c3.selectbox("班別", shifts, index=shifts.index(def_shift) if def_shift in shifts else 0, key=f"shift_{nm}", label_visibility="collapsed")
        
        other_assts = [x["name"] for x in assts if x["name"] != nm]
        curr_avoid = [x.strip() for x in r.get("avoid", "").split(",") if x.strip() in other_assts]
        avoid_val = c4.multiselect("避開", other_assts, default=curr_avoid, key=f"avoid_{nm}", label_visibility="collapsed")
        
        # 白名單網格 Popover
        curr_wl_str = r.get("slot_whitelist", "")
        wl_set = parse_slot_string(curr_wl_str, is_fixed=False)
        wl_grid_vals = []
        
        with c5.popover(f"📅 白名單 ({len(wl_set)})"):
            st.markdown("**勾選允許排班的時段 (若全空則代表無限制)**")
            days_map = {0:"一", 1:"二", 2:"三", 3:"四", 4:"五", 5:"六"}
            gc_h = st.columns([1,1,1,1])
            gc_h[0].write(""); gc_h[1].write("**早**"); gc_h[2].write("**午**"); gc_h[3].write("**晚**")
            
            for wd in range(6):
                gc = st.columns([1,1,1,1])
                gc[0].markdown(f"<div style='padding-top:8px;'>星期{days_map[wd]}</div>", unsafe_allow_html=True)
                m_val = gc[1].checkbox("", value=(wd, "早") in wl_set, key=f"wl_{nm}_{wd}_早")
                a_val = gc[2].checkbox("", value=(wd, "午") in wl_set, key=f"wl_{nm}_{wd}_午")
                e_val = gc[3].checkbox("", value=(wd, "晚") in wl_set, key=f"wl_{nm}_{wd}_晚")
                if m_val: wl_grid_vals.append(f"{days_map[wd]}早")
                if a_val: wl_grid_vals.append(f"{days_map[wd]}午")
                if e_val: wl_grid_vals.append(f"{days_map[wd]}晚")
                
        # 行政診網格 Popover
        curr_admin_str = r.get("admin_slots", "")
        admin_set = parse_slot_string(curr_admin_str, is_fixed=False)
        admin_grid_vals = []
        
        with c6.popover(f"💼 行政診 ({len(admin_set)})"):
            st.markdown("**勾選固定行政診時段 (排班將避開)**")
            agc_h = st.columns([1,1,1,1])
            agc_h[0].write(""); agc_h[1].write("**早**"); agc_h[2].write("**午**"); agc_h[3].write("**晚**")
            for wd in range(6):
                agc = st.columns([1,1,1,1])
                agc[0].markdown(f"<div style='padding-top:8px;'>星期{days_map[wd]}</div>", unsafe_allow_html=True)
                am_val = agc[1].checkbox("", value=(wd, "早") in admin_set, key=f"admin_{nm}_{wd}_早")
                aa_val = agc[2].checkbox("", value=(wd, "午") in admin_set, key=f"admin_{nm}_{wd}_午")
                ae_val = agc[3].checkbox("", value=(wd, "晚") in admin_set, key=f"admin_{nm}_{wd}_晚")
                if am_val: admin_grid_vals.append(f"{days_map[wd]}早")
                if aa_val: admin_grid_vals.append(f"{days_map[wd]}午")
                if ae_val: admin_grid_vals.append(f"{days_map[wd]}晚")
                
        new_rules[nm] = {
            "role_limit": role_val, "shift_limit": shift_val, "avoid": ",".join(avoid_val),
            "slot_whitelist": ",".join(wl_grid_vals), "admin_slots": ",".join(admin_grid_vals),
            "fixed_slots": r.get("fixed_slots", "")
        }
        st.markdown("<hr style='margin:0 0 10px 0; border-color:#f0f2f6;'>", unsafe_allow_html=True)
        
    if st.button("💾 儲存進階限制", type="primary"):
        for nm, rules in new_rules.items():
            avoids = [x.strip() for x in rules["avoid"].split(",") if x.strip()]
            for target in avoids:
                if target in new_rules:
                    target_avoids = [x.strip() for x in new_rules[target]["avoid"].split(",") if x.strip()]
                    if nm not in target_avoids:
                        target_avoids.append(nm)
                        new_rules[target]["avoid"] = ",".join(target_avoids)
        st.session_state.config["adv_rules"] = new_rules; save_config(st.session_state.config); st.session_state["sys_msg"] = "✅ 進階限制儲存成功！(避開人員已自動雙向同步)"; st.rerun()

elif step == "4. 醫師範本與生成":
    st.header("醫師班表範本與初始化")
    
    # 年月設定與儲存
    c1, c2, c3 = st.columns(3)
    y = c1.number_input("年", 2025, 2030, st.session_state.config.get("year", datetime.today().year))
    m = c2.number_input("月", 1, 12, st.session_state.config.get("month", datetime.today().month % 12 + 1))
    first_week_setting = c3.radio("畫面【第 1 週】設定為：", ["單週", "雙週"], horizontal=True)
    is_first_odd = (first_week_setting == "單週")
    
    st.session_state.config["year"] = y
    st.session_state.config["month"] = m
    
    st.info("💡 **全新體驗：** 現在點擊下方【儲存並自動套用】按鈕，系統會將您勾選的範本 **自動覆蓋並寫入本月的醫師班表** 中，您到「步驟 5」就可以直接看到結果囉！")
    
    doc_names = [d["name"] for d in get_active_doctors()]
    days = ["一", "二", "三", "四", "五", "六"]
    
    def render_template_aggrid(key):
        data = st.session_state.config.get(key, {})
        rows = []
        for doc in doc_names:
            row = {"doctor": f"👨‍⚕️ {doc}"}
            sched = data.get(doc, [False]*18)
            for i, d in enumerate(days):
                for s_idx, s in enumerate(["早", "午", "晚"]):
                    row[f"星期{d}_{s}"] = bool(sched[i*3 + s_idx]) if len(sched)==18 else False
            rows.append(row)
        df = pd.DataFrame(rows)
        
        # 欄寬增加為 130，以完整顯示醫師姓名
        col_defs = [{"headerName": "醫師", "field": "doctor", "pinned": "left", "width": 130, "cellStyle": {"fontWeight": "bold", "borderRight": "2px solid #333", "backgroundColor": "#fff"}}]
        
        for i, d in enumerate(days):
            is_odd = (i % 2 == 0)
            children = []
            for s in ["早", "午", "晚"]:
                children.append({
                    "headerName": s, "field": f"星期{d}_{s}", "editable": True,
                    "cellEditor": "agCheckboxCellEditor", "cellRenderer": "agCheckboxCellRenderer",
                    "cellClass": "is_odd" if is_odd else "is_even", "cellStyle": cell_style_js, "width": 50
                })
            col_defs.append({
                "headerName": f"星期{d}", "children": children, "headerClass": "header-odd" if is_odd else "header-even"
            })
            
        go = {"columnDefs": col_defs, "defaultColDef": {"suppressMovable": True}, "rowHeight": 45, "headerHeight": 40}
        
        res = AgGrid(df, gridOptions=go, allow_unsafe_jscode=True, fit_columns_on_grid_load=False, update_mode=GridUpdateMode.MODEL_CHANGED, theme="alpine", key=f"ag_{key}")
        
        grid_data = res['data']
        if grid_data is not None and len(grid_data) > 0:
            records = grid_data.to_dict('records') if isinstance(grid_data, pd.DataFrame) else grid_data
            new_res = {}
            for row in records:
                doc_clean = row["doctor"].replace("👨‍⚕️ ", "")
                vals = []
                for d in days:
                    for s in ["早", "午", "晚"]:
                        val = row.get(f"星期{d}_{s}", False)
                        if isinstance(val, str): val = val.lower() == 'true'
                        vals.append(bool(val))
                new_res[doc_clean] = vals
            st.session_state.config[key] = new_res
        
        return new_res

    t1, t2 = st.tabs(["單週範本", "雙週範本"])
    with t1: render_template_aggrid("template_odd")
    with t2: render_template_aggrid("template_even")
    
    col_btn1, col_btn2 = st.columns(2)
    
    if col_btn1.button("💾 僅儲存範本 (不覆蓋本月)"):
        save_config(st.session_state.config)
        st.session_state["sys_msg"] = "✅ 雙/單週範本已儲存成功！(尚未套用至本月班表)"
        st.rerun()
        
    if col_btn2.button("🚀 儲存範本並自動套用至本月", type="primary"):
        save_config(st.session_state.config)
        
        generated = []
        t_odd = st.session_state.config.get("template_odd", {})
        t_even = st.session_state.config.get("template_even", {})
        
        dates = generate_month_dates(y, m)
        weeks_dict = collections.defaultdict(list)
        for dt in dates: weeks_dict[dt.isocalendar()[1]].append(dt)
        
        for w_idx, w_dates in enumerate(weeks_dict.values()):
            is_odd_week = (w_idx % 2 == 0) if is_first_odd else (w_idx % 2 != 0)
            tmpl = t_odd if is_odd_week else t_even
            
            for dt in w_dates:
                base = dt.weekday()*3
                for s_idx, s in enumerate(["早", "午", "晚"]):
                    idx = base + s_idx
                    for doc in get_active_doctors():
                        dn = doc["name"]
                        if dn in tmpl and idx < len(tmpl[dn]) and tmpl[dn][idx]:
                            generated.append({"Date": str(dt), "Shift": s, "Doctor": dn})
                            
        st.session_state.config["manual_schedule"] = generated
        save_config(st.session_state.config)
        st.session_state["sys_msg"] = "✅ 範本已儲存，並成功為您套用至本月份！請至「步驟 5」查看結果。"
        st.rerun()

elif step == "5. 👨‍⚕️ 醫師專屬入口":
    st.header("👨‍⚕️ 醫師個人班表確認與修改")
    if is_locked_system: st.error("🔒 修改期限已過，目前為唯讀模式。")
    else: st.info("💡 若此處為空，請確認您已在「步驟 4」點擊【儲存範本並自動套用至本月】。\n\n請選擇您的名字。若要請假請將勾選取消；若要加診請打勾。反黑區域 (⬛) 不可點選。")
        
    docs = get_active_doctors()
    if docs:
        selected_doc = st.selectbox("📌 選擇醫師", [d["name"] for d in docs])
        y = st.session_state.config.get("year", datetime.today().year)
        m = st.session_state.config.get("month", datetime.today().month % 12 + 1)
        manual = st.session_state.config.get("manual_schedule", [])
        
        padded_weeks = get_padded_weeks(y, m)
        edited_dfs = {}; col_map = {}
        st.markdown(f"### 📅 {selected_doc} - {m}月 班表")
        
        for w_idx, w_dates in enumerate(padded_weeks):
            st.markdown(f"第 {w_idx+1} 週")
            cols = []
            for d_info in w_dates:
                disp = d_info["disp"]
                cols.append(disp)
                if d_info["is_curr"]: col_map[disp] = d_info["str"]

            rows = []
            for s in ["早", "午", "晚"]:
                row = {"時段": s}
                for c, d_info in zip(cols, w_dates):
                    if d_info["is_curr"]:
                        row[c] = any(x for x in manual if x["Date"] == d_info["str"] and x["Shift"] == s and x["Doctor"] == selected_doc)
                    else:
                        row[c] = False
                rows.append(row)

            df = pd.DataFrame(rows).set_index("時段")
            cfg = {c: st.column_config.CheckboxColumn(c, disabled=(not w_dates[i]["is_curr"])) for i, c in enumerate(cols)}
            edited_dfs[w_idx] = st.data_editor(df, column_config=cfg, key=f"doc_wk_{w_idx}", use_container_width=True, disabled=is_locked_system)
        
        if not is_locked_system and st.button("💾 儲存我的班表修改", type="primary"):
            new_manual = [x for x in manual if x["Doctor"] != selected_doc]
            for iso, df in edited_dfs.items():
                for shift, row in df.iterrows():
                    for c in df.columns:
                        if bool(row[c]) and c in col_map: 
                            new_manual.append({"Date": col_map[c], "Shift": shift, "Doctor": selected_doc})
            st.session_state.config["manual_schedule"] = new_manual
            save_config(st.session_state.config)
            
            if 'result' in st.session_state: del st.session_state['result']
            if 'saved_result' in st.session_state.config: del st.session_state.config['saved_result']
            st.session_state["sys_msg"] = f"✅ {selected_doc} 班表已儲存！"
            st.rerun()
            
        st.divider()
        st.markdown("### 📥 取得確認版班表")
        st.info("儲存完畢後，您可以點擊下方按鈕下載您專屬的確認班表，以防忘記。")
        st.download_button(
            "下載我的 Excel 班表", 
            to_excel_doctor_confirmed(st.session_state.config.get("manual_schedule", []), y, m, selected_doc), 
            file_name=f"{selected_doc}_{m}月班表.xlsx"
        )
    else:
        st.warning("⚠️ 系統內尚未設定任何醫師，請先至「系統與人員設定」新增醫師。")

elif step == "6. 👩‍⚕️ 助理專屬入口":
    st.header("👩‍⚕️ 助理個人休假登記")
    if is_locked_system: st.error("🔒 劃假期限已過，目前為唯讀模式。")
    else: st.info("請選擇名字。在想休假的時段「打勾」。反黑區域 (⬛) 不可點選。")
        
    assts = get_active_assistants()
    if assts:
        selected_asst = st.selectbox("📌 選擇助理", [a["name"] for a in assts])
        y = st.session_state.config.get("year", datetime.today().year)
        m = st.session_state.config.get("month", datetime.today().month % 12 + 1)
        current_leaves = st.session_state.config.get("leaves", {})
        
        padded_weeks = get_padded_weeks(y, m)
        edited_dfs = {}; col_map = {}
        st.markdown(f"### 🌴 {selected_asst} - {m}月 休假表")
        
        for w_idx, w_dates in enumerate(padded_weeks):
            st.markdown(f"第 {w_idx+1} 週")
            cols = []
            for d_info in w_dates:
                disp = d_info["disp"]
                cols.append(disp)
                if d_info["is_curr"]: col_map[disp] = d_info["str"]

            rows = []
            for s in ["早", "午", "晚"]:
                row = {"時段": s}
                for c, d_info in zip(cols, w_dates):
                    if d_info["is_curr"]:
                        row[c] = current_leaves.get(f"{selected_asst}_{d_info['str']}_{s}", False)
                    else:
                        row[c] = False
                rows.append(row)

            df = pd.DataFrame(rows).set_index("時段")
            cfg = {c: st.column_config.CheckboxColumn(c, disabled=(not w_dates[i]["is_curr"])) for i, c in enumerate(cols)}
            edited_dfs[w_idx] = st.data_editor(df, column_config=cfg, key=f"asst_wk_{w_idx}", use_container_width=True, disabled=is_locked_system)
        
        if not is_locked_system and st.button("💾 儲存我的休假", type="primary"):
            new_leaves = {k: v for k, v in current_leaves.items() if not k.startswith(f"{selected_asst}_")}
            for iso, df in edited_dfs.items():
                for shift, row in df.iterrows():
                    for c in df.columns:
                        if bool(row[c]) and c in col_map: 
                            new_leaves[f"{selected_asst}_{col_map[c]}_{shift}"] = True
            st.session_state.config["leaves"] = new_leaves
            save_config(st.session_state.config)
            
            if 'result' in st.session_state: del st.session_state['result']
            if 'saved_result' in st.session_state.config: del st.session_state.config['saved_result']
            st.session_state["sys_msg"] = f"✅ {selected_asst} 休假已儲存！(請至步驟 7 重新執行排班套用最新假單)"
            st.rerun()
    else:
        st.warning("⚠️ 系統內尚未設定任何助理，請先至「系統與人員設定」新增助理。")

elif step == "7. 排班與總管微調":
    st.header("智慧排班與微調面板")
    c1, c2 = st.columns(2)
    ctr = c1.slider("預設櫃台數", 1, 3, 2); flt = c2.slider("預設流動數", 0, 3, 1)
    
    if st.button("🚀 執行自動排班", type="primary"):
        with st.spinner("🧠 演算法激烈運算中，請稍候..."):
            man = st.session_state.config.get("manual_schedule", []); lea = st.session_state.config.get("leaves", {})
            pair = st.session_state.config.get("pairing_matrix", {}); rules = st.session_state.config.get("adv_rules", {}) 
            res, counts, s_min, s_max = run_auto_schedule(man, lea, pair, rules, ctr, flt)
            st.session_state.result = res
            st.session_state.config["saved_result"] = res
            save_config(st.session_state.config)
            st.session_state["sys_msg"] = "✅ 排班演算法執行完成！"
            st.rerun()
    
    if 'result' in st.session_state:
        y = st.session_state.config.get("year", datetime.today().year)
        m = st.session_state.config.get("month", datetime.today().month % 12 + 1)
        
        dates = generate_month_dates(y, m)
        std_min, std_max = calculate_shift_limits(y, m)
        
        # === 側邊欄：即時診數與週六指標 ===
        with st.sidebar:
            st.markdown("---")
            st.subheader("📊 總管監控儀表板")
            
            assts = get_active_assistants()
            curr_counts = {a["name"]: 0 for a in assts}
            sat_dates = [str(dt) for dt in dates if dt.weekday() == 5]
            sat_stats = {a["name"]: {"worked_days": 0, "nights": 0, "no_night_worked": 0, "full_off": 0} for a in assts}
            daily_shifts = collections.defaultdict(lambda: collections.defaultdict(set))
            
            for k, v in st.session_state.result.items():
                dt_str, sh = k.split("_")
                ppl = list(v["doctors"].values()) + v["counter"] + v["floater"] + v.get("look", [])
                for p in ppl: 
                    if p:
                        daily_shifts[p][dt_str].add(sh)
                        curr_counts[p] += 1
                        
            for a in assts:
                nm = a["name"]
                for d in sat_dates:
                    shifts = daily_shifts[nm][d]
                    if shifts:
                        sat_stats[nm]["worked_days"] += 1
                        if "晚" in shifts: sat_stats[nm]["nights"] += 1
                        else: sat_stats[nm]["no_night_worked"] += 1
                    else:
                        sat_stats[nm]["full_off"] += 1

            heaven_earth_warnings = []
            
            c_stats1, c_stats2 = st.columns(2)
            for idx, a in enumerate(assts):
                nm = a["name"]; c_val = curr_counts[nm]
                target_col = c_stats1 if idx % 2 == 0 else c_stats2
                
                for d_str, shifts in daily_shifts[nm].items():
                    if "早" in shifts and "晚" in shifts and "午" not in shifts:
                        heaven_earth_warnings.append(f"{a['nick']} 在 {d_str} 被排了天地班！")
                
                with target_col:
                    if a["type"] == "全職":
                        lim = std_max; target = std_min if a["pref"] == "low" else std_max
                        msg = f"{a['nick']}: {c_val} (標:{target})"
                        if c_val < std_min: st.warning(f"🟡 {msg}")
                        elif c_val > lim: st.error(f"🔴 {msg} 爆")
                        else: st.success(f"🟢 {msg}")
                        
                        s_off = sat_stats[nm]["full_off"]
                        s_non = sat_stats[nm]["no_night_worked"]
                        s_nit = sat_stats[nm]["nights"]
                        sat_ok = (s_off >= 1) and (s_nit == 2) 
                        sat_icon = "✅" if sat_ok else "⚠️"
                        st.caption(f"{sat_icon} 休{s_off} | 日{s_non} | 晚{s_nit}")
                    else:
                        lim = a["custom_max"] if a["custom_max"] else 15
                        msg = f"{a['nick']} (PT): {c_val}/{lim}"
                        if c_val > lim: st.error(f"🔴 {msg}")
                        else: st.info(f"🔵 {msg}")

            if heaven_earth_warnings:
                st.error("🚨 **天地班警告**\n\n" + "\n".join(heaven_earth_warnings))
        
        # === 🤖 Gemini AI 助手區塊 ===
        st.divider()
        st.subheader("🤖 Gemini AI 班表微調助手")
        if not HAS_AI_LIB:
            st.warning("請在終端機執行 `pip install google-generativeai` 即可啟用 AI 助手！")
        else:
            with st.expander("✨ 展開 AI 助手 (使用自然語言調整班表)", expanded=False):
                api_key = st.text_input("Google Gemini API Key (僅儲存於本地)", value=st.session_state.config.get("api_key", ""), type="password")
                ai_cmd = st.text_area("請輸入您的口語指示：", placeholder="例如：\n峻豪醫師禮拜四都給芸霏跟診\n燿東醫師3/14晚上要休假\n小瑜3/21晚上要休假\n紫媛於3/22下午安排行政診")
                
                if st.button("🚀 讓 AI 幫我調整"):
                    if not api_key: st.error("請輸入 API Key！")
                    elif not ai_cmd: st.error("請輸入調整指示！")
                    else:
                        st.session_state.config["api_key"] = api_key
                        save_config(st.session_state.config)
                        
                        with st.spinner("🧠 AI 正在思考與解析您的指示，並連動更新系統..."):
                            try:
                                import google.generativeai as genai
                                genai.configure(api_key=api_key)
                                model = genai.GenerativeModel('gemini-1.5-flash')
                                
                                doc_names_str = ", ".join([d["name"] for d in get_active_doctors()])
                                asst_names_str = ", ".join([a["name"] for a in get_active_assistants()])
                                
                                sys_prompt = f"""
                                你是一個牙醫診所排班系統的 JSON 解析器。目前排班年月為 {y}年{m}月。
                                請將使用者的自然語言指示，轉換成精準的 JSON 陣列，以供系統修改班表。
                                
                                醫師精確名單：{doc_names_str}
                                助理精確名單：{asst_names_str}
                                (請自動將使用者的錯字對應到正確名單上，例如若輸入「芸霏」請修正為「昀霏」)
                                
                                可用的動作 (action) 包含：
                                1. "assign_assistant_to_doctor": 指定跟診助理 (可指定特定日期 date，或特定星期 weekday(1-6代表一到六)，或特定時段 shift(早/午/晚))。
                                2. "doctor_leave": 醫師休假/取消門診 (請假)。
                                3. "assistant_leave": 助理休假 (請假)。
                                4. "assign_admin": 安排助理行政診。

                                嚴格輸出 JSON 格式陣列，不要包含任何 markdown 語法或說明文字。
                                JSON 格式範例：
                                [
                                  {{"action": "assign_assistant_to_doctor", "doctor": "吳峻豪醫師", "assistant": "昀霏", "weekday": 4}},
                                  {{"action": "doctor_leave", "doctor": "郭燿東醫師", "date": "{y}-03-14", "shift": "晚"}},
                                  {{"action": "assistant_leave", "assistant": "小瑜", "date": "{y}-03-21", "shift": "晚"}},
                                  {{"action": "assign_admin", "assistant": "紫媛", "date": "{y}-03-22", "shift": "午"}}
                                ]
                                注意：weekday 為 1代表星期一, 6代表星期六。沒有指定的條件請留空字串或 null。如果日期是例如 3/14，請轉換成 {y}-03-14 格式。
                                """
                                response = model.generate_content(sys_prompt + "\n\n使用者指示：" + ai_cmd)
                                
                                # 安全處理 Markdown 標記
                                mkd_json = "`" * 3 + "json"
                                mkd_plain = "`" * 3
                                cleaned_text = response.text.strip()
                                
                                if cleaned_text.startswith(mkd_json): 
                                    cleaned_text = cleaned_text[len(mkd_json):]
                                elif cleaned_text.startswith(mkd_plain): 
                                    cleaned_text = cleaned_text[len(mkd_plain):]
                                if cleaned_text.endswith(mkd_plain): 
                                    cleaned_text = cleaned_text[:-len(mkd_plain)]
                                
                                actions = json.loads(cleaned_text.strip())
                                
                                edited_res = st.session_state.result.copy()
                                manual = st.session_state.config.get("manual_schedule", [])
                                leaves = st.session_state.config.get("leaves", {})
                                changes_applied = 0
                                
                                for act in actions:
                                    action = act.get("action")
                                    t_date = act.get("date")
                                    t_shift = act.get("shift")
                                    t_wd = act.get("weekday")
                                    d_name = act.get("doctor")
                                    a_name = act.get("assistant")
                                    
                                    for k, v in edited_res.items():
                                        dt_str, sh = k.split("_")
                                        dt_obj = datetime.strptime(dt_str, "%Y-%m-%d").date()
                                        
                                        if t_date and dt_str != t_date: continue
                                        if t_shift and sh != t_shift: continue
                                        if t_wd and dt_obj.weekday() + 1 != t_wd: continue
                                        
                                        if action == "assign_assistant_to_doctor" and d_name and a_name:
                                            if d_name in v["doctors"]:
                                                v["doctors"][d_name] = a_name
                                                v["counter"] = [x if x != a_name else "" for x in v["counter"]]
                                                v["floater"] = [x if x != a_name else "" for x in v["floater"]]
                                                v["look"] = [x if x != a_name else "" for x in v["look"]]
                                                changes_applied += 1
                                                
                                        elif action == "doctor_leave" and d_name:
                                            if d_name in v["doctors"]:
                                                v["doctors"][d_name] = "" 
                                                manual = [m for m in manual if not (m["Date"] == dt_str and m["Shift"] == sh and m["Doctor"] == d_name)]
                                                changes_applied += 1
                                                
                                        elif action == "assistant_leave" and a_name:
                                            for doc_key in v["doctors"]:
                                                if v["doctors"][doc_key] == a_name: v["doctors"][doc_key] = ""
                                            v["counter"] = [x if x != a_name else "" for x in v["counter"]]
                                            v["floater"] = [x if x != a_name else "" for x in v["floater"]]
                                            v["look"] = [x if x != a_name else "" for x in v["look"]]
                                            leaves[f"{a_name}_{dt_str}_{sh}"] = True
                                            changes_applied += 1
                                            
                                        elif action == "assign_admin" and a_name:
                                            for doc_key in v["doctors"]:
                                                if v["doctors"][doc_key] == a_name: v["doctors"][doc_key] = ""
                                            v["counter"] = [x if x != a_name else "" for x in v["counter"]]
                                            v["floater"] = [x if x != a_name else "" for x in v["floater"]]
                                            placed = False
                                            for i in range(len(v["look"])):
                                                if v["look"][i] == "":
                                                    v["look"][i] = a_name
                                                    placed = True
                                                    break
                                            if not placed: v["look"].append(a_name)
                                            changes_applied += 1

                                st.session_state.result = edited_res
                                st.session_state.config["manual_schedule"] = manual
                                st.session_state.config["leaves"] = leaves
                                st.session_state.config["saved_result"] = edited_res
                                save_config(st.session_state.config)
                                
                                st.session_state["sys_msg"] = f"✅ AI 已經根據指示，為您在後台及畫面上套用了 {changes_applied} 項變更！"
                                st.rerun()
                                
                            except Exception as e:
                                st.error(f"⚠️ AI 解析失敗，請換個說法試試看。詳細錯誤：{e}")

        st.divider()
        # === 以下為鎖定的 AgGrid 雙層渲染區域 ===
        padded_weeks = get_padded_weeks(y, m)
        docs = get_active_doctors()
        assts = get_active_assistants()
        asst_opts = [""] + [a["nick"] for a in assts]
        n2nm = {a["nick"]: a["name"] for a in assts}
        nm2n = {a["name"]: a["nick"] for a in assts}
        edited_res = st.session_state.result.copy()
        
        with st.form("schedule_adjust_form"):
            st.subheader("📝 班表微調區 (AgGrid 雙層表頭版)")
            
            all_grids_data = []
            
            for w_idx, w_dates in enumerate(padded_weeks):
                st.markdown(f"#### 第 {w_idx+1} 週")
                
                rows = []
                for doc in docs:
                    row = {"person": f"👨‍⚕️ {doc['nick']}", "type": "doc", "name": doc["name"]}
                    for d_info in w_dates:
                        for s in ["早", "午", "晚"]:
                            f = f"{d_info['str']}_{s}"
                            if d_info["is_curr"]: row[f] = nm2n.get(edited_res.get(f, {}).get("doctors", {}).get(doc["name"], ""), "")
                            else: row[f] = "-" 
                    rows.append(row)
                
                r_defs = [("櫃1", "counter", 0), ("櫃2", "counter", 1), ("流", "floater", 0), ("看/行", "look", 0)]
                for rn, rk, ri in r_defs:
                    row = {"person": rn, "type": "role", "key": rk, "idx": ri}
                    for d_info in w_dates:
                        for s in ["早", "午", "晚"]:
                            f = f"{d_info['str']}_{s}"
                            if d_info["is_curr"]:
                                lst = edited_res.get(f, {}).get(rk, [])
                                row[f] = nm2n.get(lst[ri], "") if ri < len(lst) else ""
                            else: row[f] = "-"
                    rows.append(row)
                    
                df = pd.DataFrame(rows)
                # 欄寬增加為 130，以完整顯示人員姓名
                col_defs = [{"headerName": "人員", "field": "person", "pinned": "left", "width": 130, "editable": False, "cellStyle": {"fontWeight": "bold", "borderRight": "2px solid #333", "backgroundColor": "#fff"}}]
                
                for d_info in w_dates:
                    is_odd = (d_info["date"].weekday() % 2 == 0)
                    h_class = "header-disabled"
                    if d_info["is_curr"]: h_class = "header-odd" if is_odd else "header-even"
                        
                    children = []
                    for s in ["早", "午", "晚"]:
                        f = f"{d_info['str']}_{s}"
                        children.append({
                            "headerName": s, "field": f, "editable": d_info["is_curr"],
                            "cellEditor": "agSelectCellEditor", "cellEditorParams": {"values": asst_opts},
                            "cellClass": "is_odd" if is_odd else "is_even", "cellStyle": cell_style_js, "width": 50
                        })
                    col_defs.append({"headerName": d_info["disp"], "children": children, "headerClass": h_class})
                    
                go = {"columnDefs": col_defs, "defaultColDef": {"suppressMovable": True}, "rowHeight": 40, "headerHeight": 40}
                res = AgGrid(df, gridOptions=go, allow_unsafe_jscode=True, fit_columns_on_grid_load=False, update_mode=GridUpdateMode.MODEL_CHANGED, theme="alpine", key=f"ag_sch_{w_idx}")
                
                all_grids_data.append((w_dates, res['data']))

            if st.form_submit_button("💾 儲存並更新數據", type="primary"):
                for w_dates, grid_data in all_grids_data:
                    if grid_data is not None and len(grid_data) > 0:
                        if isinstance(grid_data, pd.DataFrame): records = grid_data.to_dict('records')
                        else: records = grid_data
                        
                        for row in records:
                            p_type = row.get("type"); p_name = row.get("name")
                            p_key = row.get("key"); p_idx = row.get("idx")
                            
                            for d_info in w_dates:
                                if not d_info["is_curr"]: continue
                                for s in ["早", "午", "晚"]:
                                    field = f"{d_info['str']}_{s}"
                                    val = row.get(field, "")
                                    v_name = n2nm.get(val, "")
                                    k = field
                                    
                                    if p_type == "doc": edited_res[k]["doctors"][p_name] = v_name
                                    elif p_type == "role":
                                        if p_key not in edited_res[k]: edited_res[k][p_key] = []
                                        while len(edited_res[k][p_key]) <= p_idx: edited_res[k][p_key].append("")
                                        edited_res[k][p_key][p_idx] = v_name

                st.session_state.result = edited_res
                st.session_state.config["saved_result"] = edited_res
                save_config(st.session_state.config)
                st.session_state["sys_msg"] = "✅ 總管班表微調已儲存，演算法及防呆指標更新完畢！"
                st.rerun()

elif step == "8. 報表下載":
    st.header("下載 Excel 報表")
    if 'result' in st.session_state:
        sch = st.session_state.result
        y = st.session_state.config.get("year", datetime.today().year)
        m = st.session_state.config.get("month", datetime.today().month % 12 + 1)
        d = get_active_doctors(); a = get_active_assistants()
        c1, c2, c3 = st.columns(3)
        c1.download_button("📊 總班表", to_excel_master(sch, y, m, d, a), f"祐德總班表_{m}月.xlsx")
        c2.download_button("👤 助理個人表", to_excel_individual(sch, y, m, a, d), f"祐德助理表_{m}月.xlsx")
