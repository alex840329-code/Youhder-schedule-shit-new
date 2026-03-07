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

# 嘗試匯入 AI 模組
try:
    import google.generativeai as genai
    HAS_AI_LIB = True
except ImportError:
    HAS_AI_LIB = False

# --- 頁面設定 ---
st.set_page_config(page_title="祐德牙醫排班系統 v18.1 (精準框線與獨立行政版)", layout="wide", page_icon="🦷")
CONFIG_FILE = 'yude_config_v11.json'

# --- 注入自訂 CSS 優化網格視覺、智慧漸層底色與框線 ---
st.markdown("""
<style>
    /* 縮小 selectbox 的高度與字體以符合表格感 */
    div[data-baseweb="select"] > div {
        font-size: 13px !important;
        padding: 0px 2px !important;
        min-height: 32px !important;
    }
    /* 隱藏 selectbox 下方的空白 */
    div[data-testid="stSelectbox"] {
        margin-bottom: -15px !important;
    }
    /* 讓 checkbox 置中對齊 */
    div[data-testid="stCheckbox"] {
        display: flex;
        justify-content: center;
        margin-bottom: -10px !important;
    }
    /* 微調欄位間距，設定 relative，並移除預設 padding 以便畫線 */
    div[data-testid="column"] {
        padding: 0 !important;
        position: relative;
        z-index: 0;
    }
    /* 自訂表頭外觀 */
    .header-tier1 {
        text-align: center; 
        font-weight: bold; 
        border-top: 1px solid #d3d3d3;
        border-bottom: 1px solid #d3d3d3;
        border-left: 2px solid #333; /* 星期左側粗線 */
        border-right: 2px solid #333; /* 星期右側粗線 */
        padding: 6px; 
        margin-bottom: 4px;
        box-sizing: border-box;
    }
    .header-tier2 {
        text-align: center; 
        font-weight: bold; 
        border-top: 1px solid #d3d3d3;
        border-bottom: 1px solid #d3d3d3;
        padding: 4px; 
        font-size: 13px;
        box-sizing: border-box;
    }
    /* 早中晚的左右細線 */
    .border-left-thin { border-left: 1px solid #d3d3d3; }
    .border-right-thin { border-right: 1px solid #d3d3d3; }
    /* 星期的左右粗線 */
    .border-left-thick { border-left: 2px solid #333; }
    .border-right-thick { border-right: 2px solid #333; }
    
    .name-col {
        padding-top: 10px; 
        padding-left: 10px;
        font-weight: bold; 
        font-size: 14px; 
        color: #333;
        border-right: 2px solid #333; /* 人員欄位右側粗線 */
        height: 100%;
        box-sizing: border-box;
    }
    
    /* 填滿整個 Column 的背景色區塊與框線 */
    .bg-fill {
        position: absolute;
        top: 0; left: 0; width: 100%; height: 100%;
        z-index: -1;
        pointer-events: none; /* 點擊穿透 */
        box-sizing: border-box;
    }
    /* 確保元件浮在背景之上 */
    div[data-testid="stCheckbox"] > label, div[data-testid="stSelectbox"] > label {
        position: relative;
        z-index: 1;
        padding: 0 5px; /* 加回一點 padding 避免文字貼邊 */
    }
    /* 隱藏外層容器的 gap 以便框線密合 */
    div[data-testid="stHorizontalBlock"] {
        gap: 0 !important;
    }
</style>
""", unsafe_allow_html=True)

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
        "leaves": {}
    }

def load_config():
    defaults = get_default_config()
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
                # 自動補齊舊版 JSON 缺少的頂層欄位
                for k, v in defaults.items():
                    if k not in data:
                        data[k] = v
                        
                # 確保舊版本助理內部屬性補齊新欄位
                if "assistants_struct" in data:
                    for a in data["assistants_struct"]:
                        if "pref" not in a: a["pref"] = "normal"
                        if "type" not in a: a["type"] = "全職"
                        if "is_main_counter" not in a: a["is_main_counter"] = False
                
                # 確保進階規則有 admin_slots 欄位
                if "adv_rules" in data:
                    for k, v in data["adv_rules"].items():
                        if "admin_slots" not in v:
                            v["admin_slots"] = ""
                return data
        except Exception as e:
            return defaults
    return defaults

def save_config(config):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except Exception as e:
        st.error(f"存檔發生錯誤: {e}")

if 'config' not in st.session_state:
    st.session_state.config = load_config()

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
        if dt.weekday() == 6: continue # 排除星期日
        dates.append(dt)
    return dates

def get_padded_weeks(year, month):
    """產生包含跨月黑方塊的完整週一至週六結構"""
    first_day = date(year, month, 1)
    last_day = date(year, month, calendar.monthrange(year, month)[1])
    
    start_date = first_day - timedelta(days=first_day.weekday()) 
    weeks = []
    current_date = start_date
    
    while current_date <= last_day or current_date.weekday() != 0:
        if current_date > last_day and current_date.weekday() == 0: break
        week_dates = []
        for _ in range(7):
            if current_date.weekday() != 6: # 排除週日
                is_curr_month = (current_date.month == month)
                week_dates.append({
                    "date": current_date,
                    "is_curr": is_curr_month,
                    "str": str(current_date),
                    "disp": f"{current_date.month}/{current_date.day} ({['一','二','三','四','五','六'][current_date.weekday()]})" if is_curr_month else f"⬛ {current_date.month}/{current_date.day}"
                })
            current_date += timedelta(days=1)
        weeks.append(week_dates)
    return weeks

def calculate_shift_limits(year, month):
    dates = generate_month_dates(year, month)
    max_s = len(dates) * 2
    min_s = max_s - 8
    return min_s, max_s

def parse_slot_string(text, is_fixed=False):
    wd_map = {"一":0, "二":1, "三":2, "四":3, "五":4, "六":5}
    shift_map = {"早":"早", "午":"午", "晚":"晚"}
    role_map = {"櫃":"counter", "流":"floater", "看":"look", "跟":"doctor", "行":"look"} # 舊有的固定班轉換，行政轉為look
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

# --- 3. 核心排班演算法 (加入防勞與天地班防禦) ---
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

    # 1. 優先填入絕對固定班
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
                    
        # 處理獨立的行政診
        for name, admin_set in parsed_admin.items():
            if (wd, sh) in admin_set:
                # 假設行政診也算作排班，但不佔用看診/櫃台/流動的扣打（依實際需求可調整，這裡先不排入 result，只阻擋）
                # 這裡設計為：如果有行政診，就不要排其他工作，因此先記上一筆
                p_counts[name] += 1
                p_daily[name][dt_str].add(sh)
                # 可選：在 result 裡標記他在行政，這裡用 look 暫代或獨立一個 admin 列表
                # result[slot]["look"].append(f"{name}(行)")

    # 2. 自動演算填補
    for slot in slots:
        dt_str, sh = slot.split("_")
        curr_dt = datetime.strptime(dt_str, "%Y-%m-%d").date()
        wd = curr_dt.weekday()
        
        duty_docs = [x["Doctor"] for x in manual_schedule if x["Date"]==dt_str and x["Shift"]==sh]
        d_order = {d["name"]: d["order"] for d in docs}
        duty_docs.sort(key=lambda x: d_order.get(x, 99))
        slot_res = result[slot]
        
        def assigned_in_slot(name):
            # 檢查是否已安排其他工作，包含行政診
            is_admin = (wd, sh) in parsed_admin.get(name, set())
            return name in slot_res["counter"] or name in slot_res["floater"] or name in slot_res["look"] or name in slot_res["doctors"].values() or is_admin

        def can_assign(name, role):
            if assigned_in_slot(name): return False
            if f"{name}_{slot}" in leaves: return False
            if p_counts[name] >= p_limits[name]: return False 
            
            rule = adv_rules.get(name, {})
            r_lim = rule.get("role_limit", "無限制")
            
            # 精準角色判斷
            if r_lim == "僅櫃台" and role != "counter": return False
            if r_lim == "僅流動" and role != "floater": return False
            if r_lim == "僅跟診" and role != "doctor": return False
            
            s_lim = rule.get("shift_limit", "無限制")
            if s_lim == "僅早班" and sh != "早": return False
            if s_lim == "僅午班" and sh != "午": return False
            if s_lim == "僅晚班" and sh != "晚": return False
            
            s_wl_str = rule.get("slot_whitelist", "")
            if s_wl_str:
                s_wl = parse_slot_string(s_wl_str, is_fixed=False)
                if (wd, sh) not in s_wl: return False

            avoid_str = rule.get("avoid", "")
            if avoid_str:
                avoids = [x.strip() for x in avoid_str.split(",")]
                for av in avoids:
                    if assigned_in_slot(av): return False 

            # ★ 防護機制：天地班防禦 (早+晚，無午)
            today_shifts = p_daily[name][dt_str]
            if sh == "晚" and "早" in today_shifts and "午" not in today_shifts: return False 
            
            # ★ 防護機制：連三防禦 (不可跨日連續三診滿班)
            if sh == "晚" and "早" in today_shifts and "午" in today_shifts:
                yesterday_str = str(curr_dt - timedelta(days=1))
                if len(p_daily[name].get(yesterday_str, set())) == 3: return False 

            return True

        def calculate_priority(candidates, curr_wd, curr_sh, curr_dt_str):
            scored = []
            for c in candidates:
                if not can_assign(c, "floater"): continue # 快篩
                
                gap = p_targets[c] - p_counts[c]
                score = gap * 10 + random.random() * 2 # 權重放大
                
                # ★ 週六完美邏輯權重調整
                if curr_wd == 5:
                    is_ft = next((a["type"] == "全職" for a in assts if a["name"] == c), False)
                    if is_ft:
                        sat_dates = [str(dt) for dt in dates if dt.weekday() == 5]
                        sats_worked = sum(1 for d in sat_dates if p_daily[c][d])
                        sat_nights = sum(1 for d in sat_dates if "晚" in p_daily[c][d])
                        
                        working_today = bool(p_daily[c][curr_dt_str])
                        
                        if curr_sh == "晚":
                            if sat_nights >= 2: score -= 1000 # 嚴格禁止超過兩個晚班
                            elif sat_nights < 2: score += 50  # 鼓勵補滿兩個晚班
                        else:
                            if not working_today and sats_worked >= total_sats - 1:
                                score -= 1000 # 必須保留一天全休日
                                
                scored.append((c, score))
            scored.sort(key=lambda x: x[1], reverse=True)
            return [x[0] for x in scored]

        candidates_pool = [a["name"] for a in assts]
        
        # 安排跟診
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

        # 安排櫃台 (確保含主櫃台)
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
        
        # 安排流動
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

# --- 4. Excel 輸出 (美化版) ---
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
        current_row += 2 # 空行分隔週次
        
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

# --- 7. UI 介面 ---
st.title("🦷 祐德牙醫 - 智慧排班系統 v18.1 (獨立行政與細緻框線版)")

is_locked_system = st.session_state.config.get("is_locked", False)

with st.sidebar:
    st.divider()
    st.subheader("⚙️ 系統權限管理")
    new_lock_state = st.toggle("🔒 鎖定前台修改 (Deadline)", value=is_locked_system, help="開啟後，醫師與助理將無法更改假單。")
    if new_lock_state != is_locked_system:
        st.session_state.config["is_locked"] = new_lock_state; save_config(st.session_state.config); st.rerun()

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
            days_map = {0:"一", 1:"二", 2:"三", 3:"四", 4:"五", 5:"六"}
            
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
            "role_limit": role_val,
            "shift_limit": shift_val,
            "avoid": ",".join(avoid_val),
            "slot_whitelist": ",".join(wl_grid_vals),
            "admin_slots": ",".join(admin_grid_vals),
            "fixed_slots": r.get("fixed_slots", "")
        }
        st.markdown("<hr style='margin:0 0 10px 0; border-color:#f0f2f6;'>", unsafe_allow_html=True)
        
    if st.button("💾 儲存進階限制", type="primary"):
        # 雙向同步「避開人員」邏輯
        for nm, rules in new_rules.items():
            avoids = [x.strip() for x in rules["avoid"].split(",") if x.strip()]
            for target in avoids:
                if target in new_rules:
                    target_avoids = [x.strip() for x in new_rules[target]["avoid"].split(",") if x.strip()]
                    if nm not in target_avoids:
                        target_avoids.append(nm)
                        new_rules[target]["avoid"] = ",".join(target_avoids)
                        
        st.session_state.config["adv_rules"] = new_rules
        save_config(st.session_state.config)
        st.session_state["sys_msg"] = "✅ 進階限制儲存成功！(避開人員已自動雙向同步)"
        st.rerun()

elif step == "4. 醫師範本與生成":
    st.header("醫師班表範本與初始化")
    st.info("💡 已加入細緻的內外框線，整行皆有背景底色方便辨識。")
    
    doc_names = [d["name"] for d in get_active_doctors()]
    days = ["一", "二", "三", "四", "五", "六"]
    shifts = ["早", "午", "晚"]
    
    # 決定第一層和第二層的寬度比例
    outer_weights = [1.5] + [3] * 6
    
    def render_template_ui(key):
        data = st.session_state.config.get(key, {})
        new_data = {}
        
        # 繪製 Tier 1 (第一層表頭：星期)
        hc1 = st.columns(outer_weights)
        hc1[0].markdown("<div class='name-col'>醫師</div>", unsafe_allow_html=True)
        for i, d in enumerate(days):
            is_odd = (i % 2 == 0) # 0=Mon(Odd), 1=Tue(Even)...
            bg_t1 = "#FFD966" if is_odd else "#9DC3E6"
            hc1[i+1].markdown(f"<div class='header-tier1' style='background-color:{bg_t1};'>星期{d}</div>", unsafe_allow_html=True)
            
        # 繪製 Tier 2 (第二層表頭：早午晚)
        hc2 = st.columns(outer_weights)
        hc2[0].markdown("<div class='name-col'></div>", unsafe_allow_html=True)
        for i in range(6):
            is_odd = (i % 2 == 0)
            bg_morn = "#FFF9C4" if is_odd else "#E6F0FA"
            bg_aft  = "#FFF2CC" if is_odd else "#CCE0F5"
            bg_eve  = "#FFE699" if is_odd else "#B3D1F0"
            
            inner = hc2[i+1].columns(3)
            inner[0].markdown(f"<div class='header-tier2 border-left-thick border-right-thin' style='background-color:{bg_morn};'>早</div>", unsafe_allow_html=True)
            inner[1].markdown(f"<div class='header-tier2 border-right-thin' style='background-color:{bg_aft};'>午</div>", unsafe_allow_html=True)
            inner[2].markdown(f"<div class='header-tier2 border-right-thick' style='background-color:{bg_eve};'>晚</div>", unsafe_allow_html=True)
            
        # 移除水平分隔線，改由 CSS 控制
        
        # 繪製 資料列
        for doc in doc_names:
            rc = st.columns(outer_weights)
            rc[0].markdown(f"<div class='name-col' style='border-top:1px solid #d3d3d3; border-bottom:1px solid #d3d3d3;'>{doc}</div>", unsafe_allow_html=True)
            
            sched = data.get(doc, [False]*18)
            doc_sched = []
            
            for i in range(6):
                is_odd = (i % 2 == 0)
                bg_morn = "#FFF9C4" if is_odd else "#E6F0FA"
                bg_aft  = "#FFF2CC" if is_odd else "#CCE0F5"
                bg_eve  = "#FFE699" if is_odd else "#B3D1F0"
                
                inner = rc[i+1].columns(3)
                
                # 注入背景標籤並加上框線，與內建 checkbox
                inner[0].markdown(f"<div class='bg-fill border-left-thick border-right-thin' style='background-color:{bg_morn}; border-top:1px solid #d3d3d3; border-bottom:1px solid #d3d3d3;'></div>", unsafe_allow_html=True)
                v1 = inner[0].checkbox("", value=bool(sched[i*3]) if len(sched)==18 else False, key=f"{key}_{doc}_{i}_0", label_visibility="collapsed")
                
                inner[1].markdown(f"<div class='bg-fill border-right-thin' style='background-color:{bg_aft}; border-top:1px solid #d3d3d3; border-bottom:1px solid #d3d3d3;'></div>", unsafe_allow_html=True)
                v2 = inner[1].checkbox("", value=bool(sched[i*3+1]) if len(sched)==18 else False, key=f"{key}_{doc}_{i}_1", label_visibility="collapsed")
                
                inner[2].markdown(f"<div class='bg-fill border-right-thick' style='background-color:{bg_eve}; border-top:1px solid #d3d3d3; border-bottom:1px solid #d3d3d3;'></div>", unsafe_allow_html=True)
                v3 = inner[2].checkbox("", value=bool(sched[i*3+2]) if len(sched)==18 else False, key=f"{key}_{doc}_{i}_2", label_visibility="collapsed")
                
                doc_sched.extend([v1, v2, v3])
                
            new_data[doc] = doc_sched
            
        return new_data

    t1, t2 = st.tabs(["單週範本", "雙週範本"])
    with t1: new_odd = render_template_ui("template_odd")
    with t2: new_even = render_template_ui("template_even")
    
    if st.button("💾 存範本", type="primary"):
        st.session_state.config["template_odd"] = new_odd
        st.session_state.config["template_even"] = new_even
        save_config(st.session_state.config)
        st.session_state["sys_msg"] = "✅ 雙/單週範本儲存成功！"
        st.rerun()
        
    st.divider()
    st.subheader("生成本月初始班表")
    c1, c2, c3 = st.columns(3)
    y = c1.number_input("年", 2025, 2030, st.session_state.config.get("year", datetime.today().year))
    m = c2.number_input("月", 1, 12, st.session_state.config.get("month", datetime.today().month % 12 + 1))
    first_week_setting = c3.radio("畫面【第 1 週】設定為：", ["單週", "雙週"])
    is_first_odd = (first_week_setting == "單週")
    
    if st.button("🚀 一鍵生成本月初始班表"):
        st.session_state.config["year"] = y; st.session_state.config["month"] = m
        generated = []
        t_odd = st.session_state.config.get("template_odd", {}); t_even = st.session_state.config.get("template_even", {})
        
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
                            
        st.session_state.config["manual_schedule"] = generated; save_config(st.session_state.config)
        st.session_state["sys_msg"] = f"✅ 初始班表生成完畢！畫面第 1 週已精準套用【{first_week_setting}】範本。"
        st.rerun()

elif step == "5. 👨‍⚕️ 醫師專屬入口":
    st.header("👨‍⚕️ 醫師個人班表確認與修改")
    if is_locked_system: st.error("🔒 修改期限已過，目前為唯讀模式。")
    else: st.info("請選擇名字。若要請假請將勾選取消；若要加診請打勾。反黑區域 (⬛) 不可點選。")
        
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
                        row[c] = False # 非當月預設為 False
                rows.append(row)

            df = pd.DataFrame(rows).set_index("時段")
            # 鎖定非當月日期
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
            
            # 清除步驟7的暫存，強迫重新跑演算法
            if 'result' in st.session_state: del st.session_state['result']
            st.session_state["sys_msg"] = f"✅ {selected_doc} 班表已儲存！(請至步驟 7 重新執行排班套用最新假單)"
            st.rerun()
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
            
            # 清除步驟7的暫存，強迫重新跑演算法
            if 'result' in st.session_state: del st.session_state['result']
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
            st.session_state["sys_msg"] = "✅ 排班演算法執行完成！"
            st.rerun()
    
    if 'result' in st.session_state:
        st.divider()
        st.info("💡 已加入細緻的內外框線，整行皆有背景底色方便辨識。")
        
        y = st.session_state.config.get("year", datetime.today().year)
        m = st.session_state.config.get("month", datetime.today().month % 12 + 1)
        padded_weeks = get_padded_weeks(y, m)
        
        docs = get_active_doctors(); assts = get_active_assistants()
        asst_opts = [""] + [a["nick"] for a in assts]
        n2nm = {a["nick"]: a["name"] for a in assts}; nm2n = {a["name"]: a["nick"] for a in assts}
        
        edited_res = st.session_state.result.copy()
        
        # 建立 Form 以便一次性送出修改
        with st.form("schedule_adjust_form"):
            st.subheader("📝 班表微調區")
            
            for w_idx, w_dates in enumerate(padded_weeks):
                st.markdown(f"#### 第 {w_idx+1} 週")
                
                num_days = len(w_dates)
                outer_weights = [1.5] + [3] * num_days
                
                # 繪製 Tier 1 (第一層表頭：日期)
                hc1 = st.columns(outer_weights)
                hc1[0].markdown("<div class='name-col'>人員</div>", unsafe_allow_html=True)
                
                # 繪製 Tier 2 (第二層表頭：早午晚) 必須同時跟著建構
                hc2 = st.columns(outer_weights)
                hc2[0].markdown("<div class='name-col'></div>", unsafe_allow_html=True)
                
                for i, d_info in enumerate(w_dates):
                    is_odd = (d_info["date"].weekday() % 2 == 0) # 0=Mon(Odd), 1=Tue(Even)...
                    
                    if d_info["is_curr"]:
                        bg_t1 = "#FFD966" if is_odd else "#9DC3E6"
                        bg_morn = "#FFF9C4" if is_odd else "#E6F0FA"
                        bg_aft  = "#FFF2CC" if is_odd else "#CCE0F5"
                        bg_eve  = "#FFE699" if is_odd else "#B3D1F0"
                    else:
                        bg_t1 = "#e0e0e0"
                        bg_morn = bg_aft = bg_eve = "#f0f0f0"
                        
                    hc1[i+1].markdown(f"<div class='header-tier1' style='background-color:{bg_t1};'>{d_info['disp']}</div>", unsafe_allow_html=True)
                    
                    sc = hc2[i+1].columns(3)
                    sc[0].markdown(f"<div class='header-tier2 border-left-thick border-right-thin' style='background-color:{bg_morn};'>早</div>", unsafe_allow_html=True)
                    sc[1].markdown(f"<div class='header-tier2 border-right-thin' style='background-color:{bg_aft};'>午</div>", unsafe_allow_html=True)
                    sc[2].markdown(f"<div class='header-tier2 border-right-thick' style='background-color:{bg_eve};'>晚</div>", unsafe_allow_html=True)
                    
                
                # 準備資料列結構
                r_defs = [("櫃1", "counter", 0), ("櫃2", "counter", 1), ("流", "floater", 0), ("看", "look", 0)]
                all_rows = [{"name": doc["name"], "label": f"👨‍⚕️{doc['nick']}", "type": "doc"} for doc in docs]
                for rn, rk, ri in r_defs:
                    all_rows.append({"name": rn, "label": rn, "type": "role", "key": rk, "idx": ri})
                
                # 繪製 每一個列
                for row_data in all_rows:
                    rc = st.columns(outer_weights)
                    rc[0].markdown(f"<div class='name-col' style='border-top:1px solid #d3d3d3; border-bottom:1px solid #d3d3d3;'>{row_data['label']}</div>", unsafe_allow_html=True)
                    
                    for col_idx, d_info in enumerate(w_dates):
                        is_odd = (d_info["date"].weekday() % 2 == 0)
                        bg_morn = "#FFF9C4" if is_odd else "#E6F0FA"
                        bg_aft  = "#FFF2CC" if is_odd else "#CCE0F5"
                        bg_eve  = "#FFE699" if is_odd else "#B3D1F0"
                        
                        inner = rc[col_idx+1].columns(3)
                        for s_idx, s in enumerate(["早", "午", "晚"]):
                            if not d_info["is_curr"]:
                                # 加入細線與粗線的 class
                                b_cls = ""
                                if s_idx == 0: b_cls = "border-left-thick border-right-thin"
                                elif s_idx == 1: b_cls = "border-right-thin"
                                elif s_idx == 2: b_cls = "border-right-thick"
                                inner[s_idx].markdown(f"<div class='bg-fill {b_cls}' style='background-color:#f0f0f0; border-top:1px solid #d3d3d3; border-bottom:1px solid #d3d3d3;'></div><div style='text-align:center; color:#ccc; padding-top:10px;'>-</div>", unsafe_allow_html=True)
                                continue
                                
                            k = f"{d_info['str']}_{s}"
                            curr_val = ""
                            
                            # 抓取當前值
                            if row_data["type"] == "doc":
                                curr_val = nm2n.get(edited_res.get(k, {}).get("doctors", {}).get(row_data["name"], ""), "")
                            else:
                                lst = edited_res.get(k, {}).get(row_data["key"], [])
                                if row_data["idx"] < len(lst):
                                    curr_val = nm2n.get(lst[row_data["idx"]], "")
                                    
                            # 設定下拉選單與顏色背景及框線
                            try:
                                def_idx = asst_opts.index(curr_val)
                            except ValueError:
                                def_idx = 0
                                
                            bg_color = [bg_morn, bg_aft, bg_eve][s_idx]
                            b_cls = ""
                            if s_idx == 0: b_cls = "border-left-thick border-right-thin"
                            elif s_idx == 1: b_cls = "border-right-thin"
                            elif s_idx == 2: b_cls = "border-right-thick"
                            
                            inner[s_idx].markdown(f"<div class='bg-fill {b_cls}' style='background-color:{bg_color}; border-top:1px solid #d3d3d3; border-bottom:1px solid #d3d3d3;'></div>", unsafe_allow_html=True)
                            
                            new_val = inner[s_idx].selectbox(
                                "", 
                                options=asst_opts, 
                                index=def_idx, 
                                key=f"sel_{w_idx}_{row_data['label']}_{k}",
                                label_visibility="collapsed"
                            )
                            
                            # 更新至 edited_res
                            v_name = n2nm.get(new_val, "")
                            if k not in edited_res:
                                edited_res[k] = {"doctors": {}, "counter": [], "floater": [], "look": []}
                                
                            if row_data["type"] == "doc":
                                edited_res[k]["doctors"][row_data["name"]] = v_name
                            else:
                                rk = row_data["key"]; ri = row_data["idx"]
                                if rk not in edited_res[k]: edited_res[k][rk] = []
                                while len(edited_res[k][rk]) <= ri: edited_res[k][rk].append("")
                                edited_res[k][rk][ri] = v_name

            submitted = st.form_submit_button("💾 儲存並更新數據", type="primary")
            if submitted:
                st.session_state.result = edited_res
                st.session_state["sys_msg"] = "✅ 總管班表微調已儲存，演算法及防呆指標更新完畢！"
                st.rerun()

        # --- 統計面板 ---
        st.subheader("📊 即時診數與週六指標")
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
                    st.caption(f"{sat_icon} 週六: 全休{s_off} | 日班{s_non} | 晚班{s_nit}")
                else:
                    lim = a["custom_max"] if a["custom_max"] else 15
                    msg = f"{a['nick']} (PT): {c_val}/{lim}"
                    if c_val > lim: st.error(f"🔴 {msg}")
                    else: st.info(f"🔵 {msg}")

        if heaven_earth_warnings:
            st.markdown("---")
            st.error("🚨 **天地班警告**\n\n" + "\n".join(heaven_earth_warnings))

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