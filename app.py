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
import time
import re
from datetime import datetime, date, timedelta

# 嘗試匯入專業表格套件
try:
    from st_aggrid import AgGrid, GridOptionsBuilder, JsCode, GridUpdateMode
    HAS_AGGRID = True
except ImportError:
    HAS_AGGRID = False

# --- 頁面設定 ---
st.set_page_config(page_title="祐德牙醫排班系統 v21.15 (精準人力面板與 AI 討論版)", layout="wide", page_icon="🦷")
CONFIG_FILE = 'yude_config_v11.json'

if not HAS_AGGRID:
    st.error("🚨 系統錯誤：請在 requirements.txt 加入 streamlit-aggrid 🚨")
    st.stop()

# === 全域樣式鎖定 ===
st.markdown("""
<style>
    .ag-header-group-cell-label { justify-content: center !important; font-size: 15px !important; font-weight: bold !important; color: #000 !important; }
    .ag-header-cell-label { justify-content: center !important; font-size: 13px !important; font-weight: bold !important; color: #000 !important; }
    .header-odd { background-color: #FFD966 !important; border-right: 2px solid #333 !important; border-bottom: 1px solid #333 !important;}
    .header-even { background-color: #9DC3E6 !important; border-right: 2px solid #333 !important; border-bottom: 1px solid #333 !important;}
    .header-disabled { background-color: #e0e0e0 !important; border-right: 2px solid #333 !important; }
    .ag-cell-wrapper { display: flex; justify-content: center; align-items: center; height: 100%; width: 100%; }
    .off-staff-box { background-color: #ffffff; border: 1px solid #ffcccc; border-left: 5px solid #ff4b4b; padding: 8px; margin-top: 5px; font-size: 11px; border-radius: 4px; color: #333; min-height: 40px; }
    .block-container { padding-top: 1.5rem; padding-bottom: 1.5rem; }
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
    if (val === '-' || val === '⬛') { style['backgroundColor'] = '#f0f0f0'; style['color'] = '#ccc'; return style; }
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

if "sys_msg" in st.session_state:
    st.success(st.session_state["sys_msg"])
    del st.session_state["sys_msg"]

# --- 1. 核心資料結構 ---
def get_default_config():
    return {
        "api_key": "", "is_locked": False, 
        "doctors_struct": [], "assistants_struct": [],
        "pairing_matrix": {}, "adv_rules": {}, "template_odd": {}, "template_even": {},
        "year": datetime.today().year, "month": datetime.today().month % 12 + 1,
        "manual_schedule": [], "leaves": {}, "saved_result": {}, "forced_assigns": {}
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
        except: return defaults
    return defaults

def save_config(config):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except Exception as e: st.error(f"存檔出錯: {e}")

if 'config' not in st.session_state:
    st.session_state.config = load_config()
    if st.session_state.config.get("saved_result"):
        st.session_state.result = st.session_state.config["saved_result"]

# --- 2. 輔助函數 ---
def get_active_doctors():
    raw = st.session_state.config.get("doctors_struct") or []
    return [d for d in sorted([x for x in raw if isinstance(x, dict)], key=lambda x: x.get("order", 99)) if d.get("active", True)]

def get_active_assistants():
    raw = st.session_state.config.get("assistants_struct") or []
    return [a for a in raw if isinstance(a, dict) and a.get("active", True)]

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
            if curr.weekday() != 6: # 排除週日
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
    wd_map = {"一":0, "二":1, "三":2, "四":3, "五":4, "六":5, "日":6}
    sh_map = {"早":"早", "午":"午", "晚":"晚"}
    role_map = {"櫃":"counter", "流":"floater", "看":"look", "跟":"doctor", "行":"look"}
    if not text or not isinstance(text, str): return {} if is_fixed else set()
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

# --- 3. 核心排班演算法 (雙階段填洞版 + 動態流動2) ---
def run_auto_schedule(manual_schedule, leaves, pairing_matrix, adv_rules, ctr_count, flt_count, forced_assigns):
    assts = get_active_assistants(); docs = get_active_doctors()
    year = st.session_state.config.get("year", datetime.today().year)
    month = st.session_state.config.get("month", datetime.today().month % 12 + 1)
    dates = generate_month_dates(year, month)
    std_min, std_max = calculate_shift_limits(year, month)
    sat_dates = [str(dt) for dt in dates if dt.weekday() == 5]
    
    p_targets = {}; p_limits = {}
    for a in assts:
        nm = a["name"]
        p_limits[nm] = std_max if a.get("type") == "全職" else (a.get("custom_max") or 15)
        p_targets[nm] = std_max if a.get("type") == "全職" else (a.get("custom_max") or 15)
            
    p_counts = {a["name"]: 0 for a in assts}
    p_floater_counts = {a["name"]: 0 for a in assts} 
    p_daily = {a["name"]: collections.defaultdict(set) for a in assts}
    
    def slot_sort_key(x):
        dt_str, sh = x.split("_"); wd = datetime.strptime(dt_str, "%Y-%m-%d").date().weekday()
        return (0 if wd == 5 else 1, dt_str, {"早":1,"午":2,"晚":3}.get(sh, 9))
    
    slots = sorted(list(set([f"{x['Date']}_{x['Shift']}" for x in manual_schedule])), key=slot_sort_key)
    result = {s: {"doctors": {}, "counter": [], "floater": [], "look": []} for s in slots}
    
    parsed_fixed = {}; parsed_admin = {}
    for name, r in adv_rules.items():
        parsed_fixed[name] = parse_slot_string(r.get("fixed_slots"), is_fixed=True)
        parsed_admin[name] = parse_slot_string(r.get("admin_slots"), is_fixed=False)

    # 1. 處理固定班與行政診
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

    # 1.5. 處理 AI/手動 指令強制排班
    for slot in slots:
        dt_str, sh = slot.split("_")
        f_assign = forced_assigns.get(slot, {})
        duty_docs = [x["Doctor"] for x in manual_schedule if x["Date"]==dt_str and x["Shift"]==sh]
        
        for d_name, a_name in f_assign.get("doctors", {}).items():
            if d_name in duty_docs and a_name:
                result[slot]["doctors"][d_name] = a_name
                p_counts[a_name] += 1
                p_daily[a_name][dt_str].add(sh)
                
        for a_name in f_assign.get("counter", []):
            if a_name and a_name not in result[slot]["counter"]:
                result[slot]["counter"].append(a_name)
                p_counts[a_name] += 1
                p_daily[a_name][dt_str].add(sh)
                
        for a_name in f_assign.get("floater", []):
            if a_name and a_name not in result[slot]["floater"]:
                result[slot]["floater"].append(a_name)
                p_counts[a_name] += 1
                p_daily[a_name][dt_str].add(sh)

    # 2. 自動排班 (嚴格階段)
    for slot in slots:
        dt_str, sh = slot.split("_"); curr_dt = datetime.strptime(dt_str, "%Y-%m-%d").date(); wd = curr_dt.weekday()
        duty_docs = [x["Doctor"] for x in manual_schedule if x["Date"]==dt_str and x["Shift"]==sh]
        slot_res = result[slot]
        
        # --- 動態判斷是否需要流動 2 ---
        current_flt_count = flt_count
        if len(duty_docs) >= 4:
            current_flt_count = max(flt_count, 2)
        
        def assigned_in_slot(name):
            is_admin = (wd, sh) in parsed_admin.get(name, set())
            return name in slot_res["counter"] or name in slot_res["floater"] or name in slot_res["look"] or name in slot_res["doctors"].values() or is_admin

        def can_assign_strict(name, role):
            if assigned_in_slot(name): return False
            if f"{name}_{dt_str}_{sh}" in leaves: return False
            
            asst_info = next((a for a in assts if a["name"] == name), {})
            
            # --- 絕對白名單 ---
            rule = adv_rules.get(name, {})
            s_wl = parse_slot_string(rule.get("slot_whitelist", ""), is_fixed=False)
            if s_wl and (wd, sh) not in s_wl: return False
            
            # --- 嚴格鐵律 ---
            if wd == 5 and asst_info.get("type") == "全職":
                sat_nites = sum(1 for d in sat_dates if "晚" in p_daily[name][d])
                if sh == "晚" and sat_nites >= 1: return False 
                has_off_sat = any(not p_daily[name][sd] for sd in sat_dates if sd != dt_str)
                if not has_off_sat and dt_str == sat_dates[-1]: return False 
            
            if p_counts[name] >= p_limits[name] and wd != 5: return False 
            
            if rule.get("role_limit") == "僅櫃台" and role != "counter": return False
            if rule.get("role_limit") == "僅流動" and role != "floater": return False
            if rule.get("role_limit") == "僅跟診" and role != "doctor": return False
            if rule.get("shift_limit") == "僅晚班" and sh != "晚": return False
            if rule.get("shift_limit") == "僅早班" and sh != "早": return False
            if rule.get("shift_limit") == "僅午班" and sh != "午": return False
            
            if role == "counter":
                for av in [x.strip() for x in rule.get("avoid", "").split(",") if x.strip()]:
                    if av in slot_res["counter"]: return False
            return True

        def calculate_priority(candidates, r_type, strict=True):
            scored = []
            for c in candidates:
                if strict and not can_assign_strict(c, r_type): continue
                if not strict and not can_assign_relaxed(c, r_type): continue
                
                asst_info = next((a for a in assts if a["name"] == c), {})
                rule = adv_rules.get(c, {})
                gap = p_targets[c] - p_counts[c]
                score = gap * 2000 
                
                s_wl = parse_slot_string(rule.get("slot_whitelist", ""), is_fixed=False)
                if s_wl and (wd, sh) in s_wl: score += 500000 
                
                if r_type == "counter":
                    if asst_info.get("is_main_counter"): score += 50000 
                    if asst_info.get("type") == "兼職": score += 20000
                    if rule.get("shift_limit") == "僅晚班" and sh == "晚": score += 200000
                
                if r_type == "floater":
                    if asst_info.get("is_main_counter"): score -= 100000 
                    else: score += (30 - p_floater_counts[c]) * 500 
                
                if wd == 5 and asst_info.get("type") == "全職": score += 15000
                
                scored.append((c, score + random.random()))
            scored.sort(key=lambda x: x[1], reverse=True)
            return [x[0] for x in scored]

        cand_pool = [a["name"] for a in assts]
        
        for d_name in duty_docs:
            if d_name in slot_res["doctors"]: continue
            picked = None; targets = [pairing_matrix.get(d_name, {}).get(k) for k in ["1","2","3"]]
            for t in [x for x in targets if x]:
                if can_assign_strict(t, "doctor"): picked = t; break
            if not picked:
                for c in calculate_priority(cand_pool, "doctor", strict=True): picked = c; break
            if picked: slot_res["doctors"][d_name] = picked; p_counts[picked] += 1; p_daily[picked][dt_str].add(sh)
            
        needed_ctr = ctr_count - len(slot_res["counter"])
        for c in calculate_priority(cand_pool, "counter", strict=True):
            if needed_ctr <= 0: break
            slot_res["counter"].append(c); p_counts[c] += 1; p_daily[c][dt_str].add(sh); needed_ctr -= 1
            
        needed_flt = current_flt_count - len(slot_res["floater"])
        for c in calculate_priority(cand_pool, "floater", strict=True):
            if needed_flt <= 0: break
            slot_res["floater"].append(c); p_counts[c] += 1; p_daily[c][dt_str].add(sh); p_floater_counts[c] += 1; needed_flt -= 1

    # 3. 自動排班 (填洞救援階段)
    for slot in slots:
        dt_str, sh = slot.split("_"); curr_dt = datetime.strptime(dt_str, "%Y-%m-%d").date(); wd = curr_dt.weekday()
        slot_res = result[slot]
        duty_docs = [x["Doctor"] for x in manual_schedule if x["Date"]==dt_str and x["Shift"]==sh]
        
        current_flt_count = flt_count
        if len(duty_docs) >= 4:
            current_flt_count = max(flt_count, 2)
            
        needed_ctr = ctr_count - len(slot_res["counter"])
        needed_flt = current_flt_count - len(slot_res["floater"])
        
        if needed_ctr <= 0 and needed_flt <= 0: continue
        
        def can_assign_relaxed(name, role):
            if name in slot_res["counter"] or name in slot_res["floater"] or name in slot_res["look"] or name in slot_res["doctors"].values(): return False
            if f"{name}_{dt_str}_{sh}" in leaves: return False
            
            rule = adv_rules.get(name, {})
            s_wl = parse_slot_string(rule.get("slot_whitelist", ""), is_fixed=False)
            if s_wl and (wd, sh) not in s_wl: return False
            
            if rule.get("role_limit") == "僅櫃台" and role != "counter": return False
            if rule.get("role_limit") == "僅流動" and role != "floater": return False
            if rule.get("role_limit") == "僅跟診" and role != "doctor": return False
            if rule.get("shift_limit") == "僅晚班" and sh != "晚": return False
            if rule.get("shift_limit") == "僅早班" and sh != "早": return False
            if rule.get("shift_limit") == "僅午班" and sh != "午": return False
            
            return True

        cand_pool = [a["name"] for a in assts]
        
        if needed_ctr > 0:
            for c in calculate_priority(cand_pool, "counter", strict=False):
                if needed_ctr <= 0: break
                slot_res["counter"].append(c); p_counts[c] += 1; p_daily[c][dt_str].add(sh); needed_ctr -= 1
                
        if needed_flt > 0:
            for c in calculate_priority(cand_pool, "floater", strict=False):
                if needed_flt <= 0: break
                slot_res["floater"].append(c); p_counts[c] += 1; p_daily[c][dt_str].add(sh); p_floater_counts[c] += 1; needed_flt -= 1

    return result

# --- 4. 本地關鍵字解析與 API 雙引擎 ---
def fuzzy_match_person(name_str, lst):
    clean = name_str.replace("醫師", "").strip()
    
    # 【重大修正 1】優先執行：精確比對 (避免將芷瑜誤判為瑜)
    for item in lst:
        if clean == item["name"] or (item.get("nick") and clean == item["nick"]):
            return item["name"]
            
    # 【重大修正 2】執行字串包含檢查，但只回傳「最長配對」的名稱
    best_match = None
    max_overlap = 0
    for item in lst:
        nm = item["name"]
        nk = item.get("nick", "")
        
        if nm in clean and len(nm) > max_overlap:
            best_match = nm; max_overlap = len(nm)
        elif clean in nm and len(clean) > max_overlap:
            best_match = nm; max_overlap = len(clean)
            
        if nk:
            if nk in clean and len(nk) > max_overlap:
                best_match = nm; max_overlap = len(nk)
            elif clean in nk and len(clean) > max_overlap:
                best_match = nm; max_overlap = len(clean)
                
    if best_match: 
        return best_match

    # 如果都沒配到，加上醫師後綴看看
    return clean + "醫師" if any("醫師" in d["name"] for d in lst) else clean

def parse_command_local(cmd, year, month, docs, assts):
    acts = []; wd_map = {"一":1, "二":2, "三":3, "四":4, "五":5, "六":6, "日":7, "天":7, "1":1, "2":2, "3":3, "4":4, "5":5, "6":6, "7":7}
    lines = cmd.replace("，", "\n").replace("、", "\n").split("\n")
    
    for line in lines:
        line = line.strip()
        if not line: 
            continue
        
        # Rule 1: 醫師給助理跟診 (加強彈性支援)
        m1 = re.search(r'([^\s\d\(\)]+?)(?:醫師)?\s*(?:禮拜|星期|週|周)([一二三四五六日天1-7])\s*(整天|早上|下午|晚上|早午晚|早午|午晚|早晚|早|午|晚)?\s*(?:給|讓|由|指定)?\s*([^\s\d\(\)]+?)\s*(?:跟|上)', line)
        if m1:
            doc = fuzzy_match_person(m1.group(1), docs)
            wd = wd_map.get(m1.group(2))
            sh_str = m1.group(3) or "整天"
            asst = fuzzy_match_person(m1.group(4), assts)
            shift = None
            if "早" in sh_str and "午" not in sh_str: shift = "早"
            elif "午" in sh_str or "下午" in sh_str: shift = "午"
            elif "晚" in sh_str: shift = "晚"
            acts.append({"action": "assign_assistant_to_doctor", "doctor": doc, "assistant": asst, "weekday": wd, "shift": shift})
            continue
            
        # Rule 2: 特定第幾個星期幾上班/休假 (加強彈性支援)
        m2 = re.search(r'([^\s\d\(\)]+?)(?:醫師)?\s*第\s*(\d+)\s*[個]*\s*(?:星期|禮拜|週|周)([一二三四五六日天1-7])\s*(整天|早上|下午|晚上|早午晚|早午|午晚|早晚|早|午|晚)?\s*(?:要|想)?(?:休假|請假|排班|上班|休息)', line)
        if m2:
            person = fuzzy_match_person(m2.group(1), assts + docs)
            w_num = int(m2.group(2)); wd = wd_map.get(m2.group(3))
            sh_str = m2.group(4) or "整天"
            # 以整行內容檢查動作
            act_type = "leave" if any(x in line for x in ["休", "請", "息"]) else "force_assign"
            shifts = []
            if "早" in sh_str or "整" in sh_str: shifts.append("早")
            if "午" in sh_str or "整" in sh_str or "下" in sh_str: shifts.append("午")
            if "晚" in sh_str or "整" in sh_str: shifts.append("晚")
            if not shifts: shifts = [None]
            for s in shifts:
                if "醫師" in person: acts.append({"action": "doctor_leave", "doctor": person, "weekday": wd, "week_number": w_num, "shift": s})
                else: acts.append({"action": act_type, "assistant": person, "weekday": wd, "week_number": w_num, "shift": s})
            continue

        # Rule 3: 終極升級版！完全容忍符號、括號與前後順序對調
        m3 = re.search(r'([^\s\d\(\)]+?)(?:醫師)?\s*(?:於)?\s*(\d+)[月/.-]\s*(\d+)[號日]?\s*(?:\(?\s*(?:星期|禮拜|週|周)?\s*([一二三四五六日天1-7])\s*\)?)?\s*(整天|早上|下午|晚上|早午晚|早午|午晚|早晚|早|午|晚)?\s*(?:要|想)?(休假|請假|排班|上班|休息)', line)
        m3_rev = re.search(r'(\d+)[月/.-]\s*(\d+)[號日]?\s*(?:\(?\s*(?:星期|禮拜|週|周)?\s*([一二三四五六日天1-7])\s*\)?)?\s*(?:由|讓|是)?\s*([^\s\d\(\)]+?)(?:醫師)?\s*(整天|早上|下午|晚上|早午晚|早午|午晚|早晚|早|午|晚)?\s*(?:要|想)?(休假|請假|排班|上班|休息)', line)
        
        if m3 or m3_rev:
            if m3:
                person_str = m3.group(1); m_str = m3.group(2); d_str = m3.group(3); sh_str = m3.group(5) or "整天"; act_str = m3.group(6)
            else:
                m_str = m3_rev.group(1); d_str = m3_rev.group(2); person_str = m3_rev.group(4); sh_str = m3_rev.group(5) or "整天"; act_str = m3_rev.group(6)
                
            person = fuzzy_match_person(person_str, assts + docs)
            date_str = f"{year}-{int(m_str):02d}-{int(d_str):02d}"
            act_type = "leave" if any(x in act_str for x in ["休", "請", "息"]) else "force_assign"
            shifts = []
            if "早" in sh_str or "整" in sh_str: shifts.append("早")
            if "午" in sh_str or "整" in sh_str or "下" in sh_str: shifts.append("午")
            if "晚" in sh_str or "整" in sh_str: shifts.append("晚")
            if not shifts: shifts = [None]
            for s in shifts:
                if "醫師" in person: acts.append({"action": "doctor_leave", "doctor": person, "date": date_str, "shift": s})
                else: acts.append({"action": act_type, "assistant": person, "date": date_str, "shift": s})
            continue
            
    return acts

def call_gemini_api(api_key, prompt):
    models = ["gemini-1.5-flash-latest", "gemini-1.5-flash", "gemini-pro"]
    for model in models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
        payload = {"contents": [{"parts": [{"text": prompt}]}]}
        for delay in [3, 8]:
            try:
                resp = requests.post(url, json=payload, timeout=20)
                if resp.status_code == 200 and 'candidates' in resp.json():
                    return resp.json()['candidates'][0]['content']['parts'][0]['text']
                elif resp.status_code == 429: time.sleep(delay); continue
                elif resp.status_code == 404: break 
            except: time.sleep(delay); continue
    return "ERROR: Google API 額度耗盡或無可用模型，請改用「本地關鍵字解析」！"

# --- 5. Excel 輸出 ---
def get_excel_formats(workbook):
    return {
        'h_main_title': workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 18}),
        'h_name': workbook.add_format({'bold': True, 'align': 'left', 'valign': 'vcenter', 'font_size': 14}),
        'h_col': workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#E0E0E0', 'border': 1}),
        'c_norm': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1}),
        'c_wknd': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFF2CC'}),
        
        # 對應 AgGrid 奇數天(一三五) / 偶數天(二四六) / 灰階(非本月) 的顏色
        'h_odd': workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFD966', 'border': 1}),
        'h_even': workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#9DC3E6', 'border': 1}),
        'h_gray': workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#E0E0E0', 'border': 1, 'font_color': '#808080'}),
        
        'c_odd_1': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FDE9D9'}),
        'c_odd_2': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FCD5B4'}),
        'c_odd_3': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FABF8F', 'right': 2}),
        
        'c_even_1': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#DDEBF7'}),
        'c_even_2': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#BDD7EE'}),
        'c_even_3': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#9DC3E6', 'right': 2}),
        
        'c_gray': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#F0F0F0', 'font_color': '#A0A0A0'}),
        'c_gray_edge': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#F0F0F0', 'font_color': '#A0A0A0', 'right': 2}),
        
        'note_style': workbook.add_format({'align': 'left', 'valign': 'vcenter', 'font_size': 11, 'bold': False})
    }

def to_excel_master(schedule_result, year, month, docs, assts):
    output = io.BytesIO(); writer = pd.ExcelWriter(output, engine='xlsxwriter'); workbook = writer.book
    fmts = get_excel_formats(workbook); sheet = workbook.add_worksheet("總班表")
    
    # 標題單一行置頂，放大並跨欄合併
    sheet.merge_range(0, 0, 0, 18, f"祐德牙醫診所 {month}月班表", fmts['h_main_title'])
    sheet.set_row(0, 30)
    
    p_weeks = get_padded_weeks(year, month)
    row = 2
    
    for w_dates in p_weeks:
        # 1. 星期/日期標題列
        sheet.write(row, 0, "日期", fmts['h_col']); col = 1
        for dt in w_dates:
            is_even = dt["date"].weekday() % 2 != 0
            f_head = fmts['h_even'] if is_even else fmts['h_odd']
            if not dt["is_curr"]: f_head = fmts['h_gray']
            
            disp_text = f"{dt['date'].month}/{dt['date'].day} ({['一','二','三','四','五','六'][dt['date'].weekday()]})" if dt["is_curr"] else f"非本月 {dt['date'].month}/{dt['date'].day}"
            sheet.merge_range(row, col, row, col+2, disp_text, f_head)
            col += 3
        row += 1
        
        # 2. 早午晚時段標題列
        sheet.write(row, 0, "時段", fmts['h_col']); col = 1
        for dt in w_dates:
            is_even = dt["date"].weekday() % 2 != 0
            for i, s in enumerate(["早","午","晚"]):
                if not dt["is_curr"]: f_cell = fmts['c_gray_edge'] if i == 2 else fmts['c_gray']
                else: f_cell = [fmts['c_even_1'], fmts['c_even_2'], fmts['c_even_3']][i] if is_even else [fmts['c_odd_1'], fmts['c_odd_2'], fmts['c_odd_3']][i]
                sheet.write(row, col, s, f_cell)
                col += 1
        row += 1
        
        # 3. 醫師與助理排班列
        for doc in docs:
            sheet.write(row, 0, doc["nick"], fmts['h_col']); col = 1
            for dt in w_dates:
                is_even = dt["date"].weekday() % 2 != 0
                for i, s in enumerate(["早","午","晚"]):
                    if not dt["is_curr"]:
                        f_cell = fmts['c_gray_edge'] if i == 2 else fmts['c_gray']
                        sheet.write(row, col, "-", f_cell)
                    else:
                        f_cell = [fmts['c_even_1'], fmts['c_even_2'], fmts['c_even_3']][i] if is_even else [fmts['c_odd_1'], fmts['c_odd_2'], fmts['c_odd_3']][i]
                        k = f"{dt['str']}_{s}"
                        anm = schedule_result.get(k, {}).get("doctors", {}).get(doc["name"], "")
                        sheet.write(row, col, next((a["nick"] for a in assts if a["name"]==anm), ""), f_cell)
                    col += 1
            row += 1
            
        # 4. 行政櫃台流動角色列
        for rnm, rk, ri in [("櫃台1","counter",0), ("櫃台2","counter",1), ("流動","floater",0), ("流動2","floater",1), ("看/行","look",0)]:
            sheet.write(row, 0, rnm, fmts['h_col']); col = 1
            for dt in w_dates:
                is_even = dt["date"].weekday() % 2 != 0
                for i, s in enumerate(["早","午","晚"]):
                    if not dt["is_curr"]:
                        f_cell = fmts['c_gray_edge'] if i == 2 else fmts['c_gray']
                        sheet.write(row, col, "-", f_cell)
                    else:
                        f_cell = [fmts['c_even_1'], fmts['c_even_2'], fmts['c_even_3']][i] if is_even else [fmts['c_odd_1'], fmts['c_odd_2'], fmts['c_odd_3']][i]
                        lst = schedule_result.get(f"{dt['str']}_{s}", {}).get(rk, [])
                        val = next((a["nick"] for a in assts if a["name"]==lst[ri]), "") if ri < len(lst) else ""
                        sheet.write(row, col, val, f_cell)
                    col += 1
            row += 1
        row += 2 
        
    writer.close(); output.seek(0); return output

def to_excel_individual(schedule_result, year, month, assts, docs):
    output = io.BytesIO(); writer = pd.ExcelWriter(output, engine='xlsxwriter'); workbook = writer.book
    fmts = get_excel_formats(workbook); dates = generate_month_dates(year, month)
    b_min, b_max = calculate_shift_limits(year, month)
    
    adv_rules = st.session_state.config.get("adv_rules", {})
    parsed_admin = {n: parse_slot_string(r.get("admin_slots", ""), is_fixed=False) for n, r in adv_rules.items()}
    
    for a in assts:
        s = workbook.add_worksheet(a["nick"]); anm = a["name"]; act = 0
        for k, v in schedule_result.items():
            dt_str, sh = k.split("_"); dt_obj = datetime.strptime(dt_str, "%Y-%m-%d").date()
            if anm in (list(v["doctors"].values()) + v["counter"] + v["floater"] + v["look"]): 
                act += 1
            elif (dt_obj.weekday(), sh) in parsed_admin.get(anm, set()):
                act += 1
                
        # 個人班表標題 - 擴展兩行避免吃字
        s.merge_range(0, 0, 0, 10, f"祐德牙醫診所 {month}月班表", fmts['h_main_title'])
        s.set_row(0, 30)
        s.merge_range(1, 0, 1, 10, f"姓名：{anm}    (實排: {act} / 上限: {a['custom_max'] or b_max})", fmts['h_name'])
        s.set_row(1, 25)
        
        for i, h in enumerate(["日期","星期","早","午","晚"]):
            s.write(3, i, h, fmts['h_col']); s.write(3, i+6, h, fmts['h_col'])
            
        mid = (len(dates)+1)//2
        for r, dt in enumerate(dates):
            col_off = 0 if r < mid else 6; row_off = r if r < mid else r - mid
            is_wknd = dt.weekday() == 5
            f_cell = fmts['c_wknd'] if is_wknd else fmts['c_norm']
            
            s.write(row_off+4, col_off, f"{dt.month}/{dt.day}", f_cell)
            s.write(row_off+4, col_off+1, ['一','二','三','四','五','六'][dt.weekday()], f_cell)
            for ci, sh in enumerate(["早","午","晚"]):
                v = ""; data = schedule_result.get(f"{dt}_{sh}", {})
                
                if (dt.weekday(), sh) in parsed_admin.get(anm, set()):
                    v="行"
                elif anm in data.get("look", []): v="看"
                elif anm in data["floater"]: v="流"
                elif anm in data["counter"]: v="櫃"
                else:
                    for dn, asn in data.get("doctors", {}).items():
                        if asn == anm: v = next((d["nick"] for d in docs if d["name"]==dn), dn)
                s.write(row_off+4, col_off+2+ci, v, f_cell)
                
        # 底部加入排班註記與工時
        last_row = mid + 6
        notes = [
            "註：全診及午晚班有空請輪流抽空吃飯，謹守30分鐘規定，以免影響其他助理。",
            "1〉早午班 8:30AM~ 12:00AM 1:30PM~ 6:00PM。　2〉午晚班 1:30PM ~10:00PM。",
            "3〉早晚班 8:00AM~ 12:00AM 6:00PM~ 10:00PM。 4〉一個早班 8:00AM~12:00AM。",
            "5〉一個午班 2:00PM ~ 6:00PM。　6〉一個晚班 6:00PM ~10:00PM。",
            "7〉全診〈早中晚〉8:00AM~ 12:00AM 1:30PM~ 10:00PM。"
        ]
        for i, note in enumerate(notes):
            s.merge_range(last_row + i, 0, last_row + i, 11, note, fmts['note_style'])
            
    writer.close(); output.seek(0); return output


# --- UI 介面 ---
is_locked_system = st.session_state.config.get("is_locked", False)

with st.sidebar:
    st.subheader("⚙️ 系統管理")
    lock_val = st.toggle("🔒 鎖定前台修改", value=is_locked_system)
    if lock_val != is_locked_system: st.session_state.config["is_locked"] = lock_val; save_config(st.session_state.config); st.rerun()

    st.divider()
    st.subheader("💾 備份與還原")
    y_cfg = st.session_state.config.get("year"); m_cfg = st.session_state.config.get("month")
    t_logic, t_month = st.tabs(["⚙️ 邏輯", "📅 班表"])
    with t_logic:
        logic_keys = ["api_key", "doctors_struct", "assistants_struct", "pairing_matrix", "adv_rules", "template_odd", "template_even", "forced_assigns"]
        st.download_button("📥 下載基本邏輯", json.dumps({k:st.session_state.config.get(k) for k in logic_keys}, ensure_ascii=False, indent=4), f"yude_logic_{datetime.now().strftime('%Y%m%d')}.json", "application/json", use_container_width=True)
        ul = st.file_uploader("📤 還原邏輯", type="json", key="ulogic")
        if ul and st.button("確認還原邏輯", use_container_width=True):
            try:
                new = json.load(ul); d_cfg = get_default_config()
                for k in logic_keys: 
                    v = new.get(k)
                    if k in ["doctors_struct", "assistants_struct"] and isinstance(v, list): v = [x for x in v if isinstance(x, dict) and x is not None]
                    st.session_state.config[k] = v if v is not None else d_cfg.get(k)
                save_config(st.session_state.config); st.rerun()
            except: st.error("還原失敗")
            
    with t_month:
        month_keys = ["year", "month", "manual_schedule", "leaves", "saved_result", "forced_assigns"]
        if st.session_state.get("result"): st.session_state.config["saved_result"] = st.session_state.result
        st.download_button("📥 下載當月班表", json.dumps({k:st.session_state.config.get(k) for k in month_keys}, ensure_ascii=False, indent=4), f"yude_month_{y_cfg}_{m_cfg}_backup.json", "application/json", use_container_width=True)
        um = st.file_uploader("📤 還原班表", type="json", key="umonth")
        if um and st.button("確認還原班表", use_container_width=True):
            try:
                new = json.load(um); d_cfg = get_default_config()
                for k in month_keys: st.session_state.config[k] = new.get(k) if new.get(k) is not None else d_cfg.get(k)
                if st.session_state.config.get("saved_result"): st.session_state.result = st.session_state.config["saved_result"]
                save_config(st.session_state.config); st.rerun()
            except: st.error("還原失敗")

step = st.sidebar.radio("導覽", ["1. 人員設定", "2. 跟診配對", "3. 進階限制", "4. 班表生成", "5. 醫師入口", "6. 助理入口", "7. 排班微調", "8. 報表下載"])

def safe_index(options, value, default=0):
    try: return options.index(value)
    except: return default

if step == "1. 人員設定":
    st.header("人員設定")
    y, m = st.session_state.config.get("year"), st.session_state.config.get("month")
    min_s, max_s = calculate_shift_limits(y, m)
    st.info(f"📅 {y}/{m} ｜ 正職上限：{max_s}")
    c1, c2 = st.columns(2)
    with c1:
        ed_doc = st.data_editor(pd.DataFrame(st.session_state.config.get("doctors_struct", [])), use_container_width=True, hide_index=True, num_rows="dynamic")
        if st.button("存醫師"): st.session_state.config["doctors_struct"] = ed_doc.to_dict('records'); save_config(st.session_state.config); st.rerun()
    with c2:
        ed_asst = st.data_editor(pd.DataFrame(st.session_state.config.get("assistants_struct", [])), column_config={"type": st.column_config.SelectboxColumn(options=["全職","兼職"]), "pref": st.column_config.SelectboxColumn(options=["high","normal","low"]), "is_main_counter": st.column_config.CheckboxColumn("專職櫃台?")}, use_container_width=True, hide_index=True, num_rows="dynamic")
        if st.button("存助理"): st.session_state.config["assistants_struct"] = ed_asst.replace({np.nan: None}).to_dict('records'); save_config(st.session_state.config); st.rerun()

elif step == "2. 跟診配對":
    st.header("跟診指定順位表")
    docs = get_active_doctors(); assts = [""] + [a["name"] for a in get_active_assistants()]
    matrix_data = [{"醫師": d["name"], "第一順位": st.session_state.config.get("pairing_matrix", {}).get(d["name"], {}).get("1",""), "第二順位": st.session_state.config.get("pairing_matrix", {}).get(d["name"], {}).get("2",""), "第三順位": st.session_state.config.get("pairing_matrix", {}).get(d["name"], {}).get("3","")} for d in docs]
    ed_mat = st.data_editor(pd.DataFrame(matrix_data), column_config={k: st.column_config.SelectboxColumn(options=assts) for k in ["第一順位","第二順位","第三順位"]}, use_container_width=True, hide_index=True)
    if st.button("儲存配對"):
        st.session_state.config["pairing_matrix"] = {r["醫師"]: {"1":r["第一順位"],"2":r["第二順位"],"3":r["第三順位"]} for i, r in ed_mat.iterrows()}; save_config(st.session_state.config); st.rerun()

elif step == "3. 進階限制":
    st.header("🛡️ 助理進階限制")
    assts = get_active_assistants(); curr_rules = st.session_state.config.get("adv_rules", {})
    new_rules = {}
    r_opts = ["無限制", "僅櫃台", "僅行政", "僅流動", "僅跟診"]; s_opts = ["無限制", "僅早班", "僅午班", "僅晚班"]
    for a in assts:
        nm = a["name"]; r = curr_rules.get(nm, {})
        c1, c2, c3, c4, c5, c6 = st.columns([1, 1.5, 1.5, 2.5, 1.5, 1.5])
        c1.markdown(f"**{nm}**")
        rv = c2.selectbox("職位", r_opts, index=safe_index(r_opts, r.get("role_limit", "無限制")), key=f"r_{nm}", label_visibility="collapsed")
        sv = c3.selectbox("班別", s_opts, index=safe_index(s_opts, r.get("shift_limit", "無限制")), key=f"s_{nm}", label_visibility="collapsed")
        av = c4.multiselect("避開", [x["name"] for x in assts if x["name"] != nm], default=[x.strip() for x in r.get("avoid","").split(",") if x.strip() in [x["name"] for x in assts]], key=f"v_{nm}", label_visibility="collapsed")
        with c5.popover("📅 白名單"):
            wl_s = parse_slot_string(r.get("slot_whitelist","")); wl_grid = []; days = ["一","二","三","四","五","六"]
            
            head_cols = st.columns(4)
            head_cols[0].write("")
            head_cols[1].write("**早**")
            head_cols[2].write("**午**")
            head_cols[3].write("**晚**")
            
            for di in range(6):
                gc = st.columns(4); gc[0].write(days[di])
                for si, sn in enumerate(["早","午","晚"]):
                    if gc[si+1].checkbox("", value=(di, sn) in wl_s, key=f"wl_{nm}_{di}_{sn}"): wl_grid.append(f"{days[di]}{sn}")
        with c6.popover("💼 行政"):
            ad_s = parse_slot_string(r.get("admin_slots","")); ad_grid = []
            
            head_cols = st.columns(4)
            head_cols[0].write("")
            head_cols[1].write("**早**")
            head_cols[2].write("**午**")
            head_cols[3].write("**晚**")
            
            for di in range(6):
                gc = st.columns(4); gc[0].write(days[di])
                for si, sn in enumerate(["早","午","晚"]):
                    if gc[si+1].checkbox("", value=(di, sn) in ad_s, key=f"ad_{nm}_{di}_{sn}"): ad_grid.append(f"{days[di]}{sn}")
        new_rules[nm] = {"role_limit":rv, "shift_limit":sv, "avoid":",".join(av), "slot_whitelist":",".join(wl_grid), "admin_slots":",".join(ad_grid), "fixed_slots":r.get("fixed_slots","")}
        st.divider()
    if st.button("💾 儲存進階限制", type="primary"):
        st.session_state.config["adv_rules"] = new_rules; save_config(st.session_state.config); st.rerun()

elif step == "4. 班表生成":
    st.header("醫師班表範本與初始化")
    c1, c2, c3 = st.columns(3)
    y, m = c1.number_input("年", 2025, 2030, st.session_state.config.get("year")), c2.number_input("月", 1, 12, st.session_state.config.get("month"))
    fws = c3.radio("第一週設定：", ["單週","雙週"], horizontal=True)
    st.session_state.config["year"], st.session_state.config["month"] = y, m
    doc_names = [d["name"] for d in get_active_doctors()]; days = ["一","二","三","四","五","六"]
    def render_ag(key):
        data = st.session_state.config.get(key, {}); rows = []
        for d in doc_names:
            r = {"doctor": f"👨‍⚕️ {d}"}; s = data.get(d, [False]*18)
            for i, dn in enumerate(days):
                for si, sn in enumerate(["早","午","晚"]): r[f"{dn}_{sn}"] = bool(s[i*3+si]) if len(s)==18 else False
            rows.append(r)
        cd = [{"headerName": "醫師", "field": "doctor", "pinned": "left", "width": 80, "cellStyle": {"fontWeight":"bold","borderRight":"2px solid #333","backgroundColor":"#fff"}}]
        for i, dn in enumerate(days):
            child = [{"headerName": sn, "field": f"{dn}_{sn}", "editable": True, "cellEditor": "agCheckboxCellEditor", "cellRenderer": "agCheckboxCellRenderer", "cellClass": "is_odd" if i%2==0 else "is_even", "cellStyle": cell_style_js, "width": 55} for sn in ["早","午","晚"]]
            cd.append({"headerName": f"星期{dn}", "children": child, "headerClass": "header-odd" if i%2==0 else "header-even"})
        res = AgGrid(pd.DataFrame(rows), gridOptions={"columnDefs": cd, "rowHeight": 45, "domLayout": 'autoHeight'}, allow_unsafe_jscode=True, theme="alpine", key=f"ag_{key}")
        if res['data'] is not None:
            nr = {}; rd = res['data'].to_dict('records') if isinstance(res['data'], pd.DataFrame) else res['data']
            for row in rd:
                dn_clean = row["doctor"].replace("👨‍⚕️ ",""); nr[dn_clean] = [bool(row.get(f"{d}_{s}", False)) for d in days for s in ["早","午","晚"]]
            st.session_state.config[key] = nr
    t1, t2 = st.tabs(["單週","雙週"])
    with t1: render_ag("template_odd")
    with t2: render_ag("template_even")
    if st.button("🚀 儲存並套用至本月", type="primary"):
        save_config(st.session_state.config); generated = []; dates = generate_month_dates(y, m); weeks = collections.defaultdict(list)
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
        p_weeks = get_padded_weeks(y, m); eds = []
        for wi, w_dates in enumerate(p_weeks):
            st.markdown(f"**第 {wi+1} 週**")
            cols = [d["disp"] for d in w_dates]; cmap = {d["disp"]: d["str"] for d in w_dates if d["is_curr"]}; rows = []
            for sn in ["早","午","晚"]:
                r = {"時段": sn}
                for c in cols: r[c] = any(x for x in manual if x["Date"]==cmap.get(c) and x["Shift"]==sn and x["Doctor"]==sel_doc)
                rows.append(r)
            cfg = {c: st.column_config.CheckboxColumn(disabled=(not any(d["is_curr"] for d in w_dates if d["disp"]==c))) for c in cols}
            eds.append((w_dates, st.data_editor(pd.DataFrame(rows).set_index("時段"), column_config=cfg, use_container_width=True, key=f"doc_{sel_doc}_{wi}", disabled=is_locked_system)))
        if not is_locked_system and st.button("💾 儲存修改", type="primary"):
            new_man = [x for x in manual if x["Doctor"] != sel_doc]
            for w_dates, df_out in eds:
                cmap = {d["disp"]: d["str"] for d in w_dates if d["is_curr"]}
                for sn in ["早","午","晚"]:
                    for disp, dt_str in cmap.items():
                        if df_out.at[sn, disp]: new_man.append({"Date": dt_str, "Shift": sn, "Doctor": sel_doc})
            st.session_state.config["manual_schedule"] = new_man; save_config(st.session_state.config); st.rerun()

elif step == "6. 助理入口":
    st.header("👩‍⚕️ 助理入口")
    assts = get_active_assistants(); sel_asst = st.selectbox("📌 助理名字", [a["name"] for a in assts])
    y, m = st.session_state.config.get("year"), st.session_state.config.get("month")
    leaves = st.session_state.config.get("leaves", {}); p_weeks = get_padded_weeks(y, m); eds = []
    for wi, w_dates in enumerate(p_weeks):
        st.markdown(f"**第 {wi+1} 週**")
        cols = [d["disp"] for d in w_dates]; cmap = {d["disp"]: d["str"] for d in w_dates if d["is_curr"]}; rows = []
        for sn in ["早","午","晚"]:
            r = {"時段": sn}
            for c in cols: r[c] = leaves.get(f"{sel_asst}_{cmap.get(c)}_{sn}", False)
            rows.append(r)
        cfg = {c: st.column_config.CheckboxColumn(disabled=(not any(d["is_curr"] for d in w_dates if d["disp"]==c))) for c in cols}
        eds.append((w_dates, st.data_editor(pd.DataFrame(rows).set_index("時段"), column_config=cfg, use_container_width=True, key=f"asst_{sel_asst}_{wi}", disabled=is_locked_system)))
    if not is_locked_system and st.button("💾 儲存休假登記", type="primary"):
        new_leaves = {k: v for k, v in leaves.items() if not k.startswith(f"{sel_asst}_")}
        for w_dates, df_out in eds:
            cmap = {d["disp"]: d["str"] for d in w_dates if d["is_curr"]}
            for sn in ["早","午","晚"]:
                for disp, dt_str in cmap.items():
                    if df_out.at[sn, disp]: new_leaves[f"{sel_asst}_{dt_str}_{sn}"] = True
        st.session_state.config["leaves"] = new_leaves; save_config(st.session_state.config); st.rerun()

elif step == "7. 排班微調":
    y, m = st.session_state.config.get("year"), st.session_state.config.get("month"); dates = generate_month_dates(y, m)
    std_min, std_max = calculate_shift_limits(y, m); assts = get_active_assistants(); sat_dates = [str(dt) for dt in dates if dt.weekday() == 5]
    docs = get_active_doctors()
    st.title(f"📅 {y} 年 {m} 月 班表總管")
    
    with st.sidebar:
        st.markdown("---")
        st.subheader("📊 即時監控 (已包含行政診)")
        if 'result' in st.session_state:
            curr_counts = {a["name"]: 0 for a in assts}; curr_floaters = {a["name"]: 0 for a in assts}
            daily_p = collections.defaultdict(lambda: collections.defaultdict(set))
            
            adv_rules = st.session_state.config.get("adv_rules", {})
            parsed_admin = {n: parse_slot_string(r.get("admin_slots", ""), is_fixed=False) for n, r in adv_rules.items()}
            
            for k, v in st.session_state.result.items():
                dt_str, sh = k.split("_"); dt_obj = datetime.strptime(dt_str, "%Y-%m-%d").date()
                for a in assts:
                    nm = a["name"]
                    in_grid = nm in (list(v["doctors"].values()) + v["counter"] + v["floater"] + v["look"])
                    is_admin = (dt_obj.weekday(), sh) in parsed_admin.get(nm, set())
                    if in_grid or is_admin:
                        curr_counts[nm] += 1; daily_p[nm][dt_str].add(sh)
                    if nm in v["floater"]: curr_floaters[nm] += 1

            for a in assts:
                nm = a["name"]; status_color = "green" if curr_counts[nm] >= std_min else "red"
                triples = sum(1 for s_set in daily_p[nm].values() if len(s_set) == 3)
                heaven_earth = sum(1 for s_set in daily_p[nm].values() if "早" in s_set and "晚" in s_set and "午" not in s_set)
                
                s_off, s_nite, s_day = 0, 0, 0
                for d in sat_dates:
                    s_set = daily_p[nm].get(d, set())
                    if not s_set: s_off += 1
                    elif "晚" in s_set: s_nite += 1
                    else: s_day += 1
                        
                st.markdown(f"**{nm}** ({a['type']})\n- 總診: :{status_color}[{curr_counts[nm]}] | **流: {curr_floaters[nm]}**")
                s_status = "✅" if s_off >= 1 and s_nite <= 1 else "⚠️"
                if a["type"] == "兼職": s_status = "🆗(PT)"
                st.caption(f"- {s_status} 週六: 休{s_off}|早午{s_day}|晚{s_nite}")
                if triples or heaven_earth: st.markdown(f"- 🚩 :orange[全:{triples}]|:red[天:{heaven_earth}]")
                st.markdown("---")

    c1, c2 = st.columns(2); ctr = c1.slider("櫃台人數", 1,3,2); flt = c2.slider("流動人數",0,3,1)
    if st.button("🚀 執行自動排班演算法", type="primary"):
        with st.spinner("雙階段演算法運算中..."):
            if 'result' in st.session_state: del st.session_state['result']
            if 'saved_result' in st.session_state.config: del st.session_state.config['saved_result']
            res = run_auto_schedule(st.session_state.config["manual_schedule"], st.session_state.config["leaves"], st.session_state.config.get("pairing_matrix",{}), st.session_state.config.get("adv_rules",{}), ctr, flt, st.session_state.config.get("forced_assigns", {}))
            st.session_state.result = res; st.session_state.config["saved_result"] = res; save_config(st.session_state.config); st.rerun()
            
    if 'result' in st.session_state:
        st.divider()
        c_mode1, c_mode2 = st.columns([7, 3])
        with c_mode1:
            mode = st.radio("選擇快速調整模式", ["🔧 本地關鍵字解析 (無須 API Key，極速推薦)", "🤖 Google AI 語意解析 (需 API Key)"], horizontal=True)
        with c_mode2:
            if st.button("🧹 清除所有 AI/手動強制指定"):
                st.session_state.config["forced_assigns"] = {}; save_config(st.session_state.config); st.rerun()
                
        if "本地關鍵字" in mode:
            st.info("💡 **支援的關鍵字句型：**\n1. `XX醫師禮拜X[早上/下午/晚上/整天]給YY跟`\n2. `XX第N個星期X[早午/晚/整天]上班/休假`\n3. `XX[於]M月D日[星期X][早上/下午/晚上/整天]請假/上班` (例: 欣霓4/11星期六早午上班)")
            cmd = st.text_area("請輸入指令 (可換行輸入多筆)", placeholder="峻豪醫師禮拜四整天給昀霏跟\n欣霓第2個星期六早午上班\n雯萱第3個星期六休假\n燿東醫師4/10號要請假")
            if st.button("執行本地調整"):
                if cmd:
                    with st.spinner("系統極速解析中..."):
                        acts = parse_command_local(cmd, y, m, docs, assts)
                        if acts:
                            forced = st.session_state.config.get("forced_assigns", {})
                            leaves = st.session_state.config.get("leaves", {})
                            manual = st.session_state.config.get("manual_schedule", [])
                            apply_count = 0
                            
                            def get_target_dates(act, year, month):
                                targets = []
                                if act.get("date"):
                                    targets.append(act["date"])
                                elif act.get("weekday"):
                                    wd_target = act["weekday"] - 1 
                                    count = 0
                                    for d in range(1, calendar.monthrange(year, month)[1] + 1):
                                        dt_obj = date(year, month, d)
                                        if dt_obj.weekday() == wd_target:
                                            count += 1
                                            if act.get("week_number") and count != act["week_number"]: continue
                                            targets.append(str(dt_obj))
                                return targets
                            
                            for act in acts:
                                targets = get_target_dates(act, y, m)
                                anm = act.get("assistant"); doc_name = act.get("doctor"); action_type = act.get("action")
                                shifts = ["早","午","晚"] if not act.get("shift") else [act["shift"]]
                                
                                for dt_str in targets:
                                    for sh in shifts:
                                        k = f"{dt_str}_{sh}"
                                        if k not in forced: forced[k] = {"doctors": {}, "counter": [], "floater": []}
                                        
                                        if action_type == "leave":
                                            leaves[f"{anm}_{dt_str}_{sh}"] = True
                                            for d_key, a in list(forced[k]["doctors"].items()):
                                                if a == anm: forced[k]["doctors"].pop(d_key, None)
                                            if anm in forced[k]["counter"]: forced[k]["counter"].remove(anm)
                                            if anm in forced[k]["floater"]: forced[k]["floater"].remove(anm)
                                            apply_count += 1
                                            
                                        elif action_type == "doctor_leave":
                                            manual = [m_s for m_s in manual if not (m_s["Date"] == dt_str and m_s["Shift"] == sh and m_s["Doctor"] == doc_name)]
                                            apply_count += 1
                                            
                                        elif action_type == "assign_assistant_to_doctor":
                                            for d_key, a in list(forced[k]["doctors"].items()):
                                                if a == anm: forced[k]["doctors"].pop(d_key, None)
                                            forced[k]["doctors"][doc_name] = anm
                                            if anm in forced[k]["counter"]: forced[k]["counter"].remove(anm)
                                            if anm in forced[k]["floater"]: forced[k]["floater"].remove(anm)
                                            leaves.pop(f"{anm}_{dt_str}_{sh}", None)
                                            apply_count += 1
                                            
                                        elif action_type == "force_assign":
                                            is_assigned = anm in forced[k]["doctors"].values() or anm in forced[k]["counter"] or anm in forced[k]["floater"]
                                            if not is_assigned:
                                                asst_rules = st.session_state.config.get("adv_rules", {}).get(anm, {})
                                                target_ky = "counter" if asst_rules.get("role_limit") == "僅櫃台" else "floater"
                                                if anm not in forced[k][target_ky]: forced[k][target_ky].append(anm)
                                            leaves.pop(f"{anm}_{dt_str}_{sh}", None)
                                            apply_count += 1
                                            
                            st.session_state.config["forced_assigns"] = forced
                            st.session_state.config["leaves"] = leaves
                            st.session_state.config["manual_schedule"] = manual
                            save_config(st.session_state.config)
                            
                            st.session_state.result = run_auto_schedule(manual, leaves, st.session_state.config.get("pairing_matrix",{}), st.session_state.config.get("adv_rules",{}), ctr, flt, forced)
                            st.session_state.config["saved_result"] = st.session_state.result
                            save_config(st.session_state.config)
                            
                            st.success(f"✅ 本地解析成功！寫入 {apply_count} 筆限制，並已連動完成排班。")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error("未能辨識任何指令，請確認格式是否正確。")
        else:
            api_key = st.text_input("API Key", value=st.session_state.config.get("api_key",""), type="password")
            cmd = st.text_area("口語指令", placeholder="Ex: 峻豪醫師禮拜四整天昀霏跟診")
            if st.button("執行 AI 調整"):
                if api_key and cmd:
                    st.session_state.config["api_key"] = api_key; save_config(st.session_state.config)
                    with st.spinner("AI 思考中..."):
                        docs_str = ",".join([d["name"] for d in get_active_doctors()]); asst_str = ",".join([a["name"] for a in get_active_assistants()])
                        prompt = f"牙醫排班年月:{y}年{m}月。醫師:{docs_str}。助理:{asst_str}。轉JSON動作:[{{'action': 'assign_assistant_to_doctor'|'leave'|'force_assign', 'doctor': 'NAME', 'assistant': 'NAME', 'weekday': 1-6, 'week_number': 1-5, 'date': 'YYYY-MM-DD', 'shift': '早/午/晚/null'}}].指令:{cmd[:500]}"
                        raw = call_gemini_api(api_key, prompt)
                        if raw.startswith("ERROR:"): st.error(f"AI 異常: {raw[6:]}")
                        else:
                            try:
                                clean_json = re.sub(r'```json\s*|\s*```', '', raw).strip(); acts = json.loads(clean_json)
                                forced = st.session_state.config.get("forced_assigns", {}); leaves = st.session_state.config.get("leaves", {}); manual = st.session_state.config.get("manual_schedule", [])
                                apply_count = 0
                                
                                def get_target_dates(act, year, month):
                                    targets = []
                                    if act.get("date"):
                                        targets.append(act["date"])
                                    elif act.get("weekday"):
                                        wd_target = act["weekday"] - 1
                                        count = 0
                                        for d in range(1, calendar.monthrange(year, month)[1] + 1):
                                            dt_obj = date(year, month, d)
                                            if dt_obj.weekday() == wd_target:
                                                count += 1
                                                if act.get("week_number") and count != act["week_number"]: continue
                                                targets.append(str(dt_obj))
                                    return targets

                                for act in acts:
                                    targets = get_target_dates(act, y, m)
                                    anm = act.get("assistant"); doc_name = act.get("doctor"); action_type = act.get("action")
                                    shifts = ["早","午","晚"] if not act.get("shift") else [act["shift"]]
                                    
                                    for dt_str in targets:
                                        for sh in shifts:
                                            k = f"{dt_str}_{sh}"
                                            if k not in forced: forced[k] = {"doctors": {}, "counter": [], "floater": []}
                                            
                                            if action_type == "leave":
                                                leaves[f"{anm}_{dt_str}_{sh}"] = True
                                                for d_key, a in list(forced[k]["doctors"].items()):
                                                    if a == anm: forced[k]["doctors"].pop(d_key, None)
                                                if anm in forced[k]["counter"]: forced[k]["counter"].remove(anm)
                                                if anm in forced[k]["floater"]: forced[k]["floater"].remove(anm)
                                                apply_count += 1
                                                
                                            elif action_type == "doctor_leave":
                                                manual = [m_s for m_s in manual if not (m_s["Date"] == dt_str and m_s["Shift"] == sh and m_s["Doctor"] == doc_name)]
                                                apply_count += 1
                                                
                                            elif action_type == "assign_assistant_to_doctor":
                                                for d_key, a in list(forced[k]["doctors"].items()):
                                                    if a == anm: forced[k]["doctors"].pop(d_key, None)
                                                forced[k]["doctors"][doc_name] = anm
                                                if anm in forced[k]["counter"]: forced[k]["counter"].remove(anm)
                                                if anm in forced[k]["floater"]: forced[k]["floater"].remove(anm)
                                                leaves.pop(f"{anm}_{dt_str}_{sh}", None)
                                                apply_count += 1
                                                
                                            elif action_type == "force_assign":
                                                is_assigned = anm in forced[k]["doctors"].values() or anm in forced[k]["counter"] or anm in forced[k]["floater"]
                                                if not is_assigned:
                                                    asst_rules = st.session_state.config.get("adv_rules", {}).get(anm, {})
                                                    target_ky = "counter" if asst_rules.get("role_limit") == "僅櫃台" else "floater"
                                                    if anm not in forced[k][target_ky]: forced[k][target_ky].append(anm)
                                                leaves.pop(f"{anm}_{dt_str}_{sh}", None)
                                                apply_count += 1
                                                
                                st.session_state.config["forced_assigns"] = forced
                                st.session_state.config["leaves"] = leaves
                                st.session_state.config["manual_schedule"] = manual
                                save_config(st.session_state.config)
                                
                                st.session_state.result = run_auto_schedule(manual, leaves, st.session_state.config.get("pairing_matrix",{}), st.session_state.config.get("adv_rules",{}), ctr, flt, forced)
                                st.session_state.config["saved_result"] = st.session_state.result
                                save_config(st.session_state.config)
                                
                                st.success(f"✅ AI 解析成功！寫入 {apply_count} 筆限制，並已連動完成排班。")
                                time.sleep(1)
                                st.rerun()
                            except Exception as e: st.error(f"AI 解析失敗，請換個說法。詳細錯誤：{e}")

        p_weeks = get_padded_weeks(y, m); a_opts = [""] + [a["nick"] for a in get_active_assistants()]
        nm2n = {a["name"]: a["nick"] for a in get_active_assistants()}; n2nm = {a["nick"]: a["name"] for a in get_active_assistants()}
        leaves_data = st.session_state.config.get("leaves", {})
        
        with st.form("adj_form"):
            for wi, w_dates in enumerate(p_weeks):
                # 改為以「早/午/晚」為單位的空閒人力陣列
                shift_off_staff = {}
                for dt in w_dates:
                    if not dt["is_curr"]: continue
                    d_str = dt["str"]
                    for sh in ["早","午","晚"]:
                        data = st.session_state.result.get(f"{d_str}_{sh}", {})
                        working = set(list(data.get("doctors", {}).values()) + data.get("counter", []) + data.get("floater", []) + data.get("look", []))
                        shift_off_staff[f"{d_str}_{sh}"] = [
                            a["name"] for a in get_active_assistants() 
                            if a["name"] not in working and not leaves_data.get(f"{a['name']}_{d_str}_{sh}")
                        ]

                rows = []
                for d in docs:
                    r = {"person": f"👨‍⚕️ {d['nick']}", "type":"doc", "name":d["name"]}
                    for dt in w_dates:
                        for s in ["早","午","晚"]:
                            f = f"{dt['str']}_{s}"; r[f] = nm2n.get(st.session_state.result.get(f, {}).get("doctors", {}).get(d["name"], ""), "") if dt["is_curr"] else "-"
                    rows.append(r)
                for rnm, rk, ri in [("櫃1","counter",0), ("櫃2","counter",1), ("流","floater",0), ("流2","floater",1), ("看/行","look",0)]:
                    r = {"person": rnm, "type":"role", "key":rk, "idx":ri}
                    for dt in w_dates:
                        for s in ["早","午","晚"]:
                            f = f"{dt['str']}_{s}"
                            if dt["is_curr"]:
                                lst = st.session_state.result.get(f, {}).get(rk, [])
                                r[f] = nm2n.get(lst[ri], "") if ri < len(lst) else ""
                            else: r[f] = "-"
                    rows.append(r)
                
                cd = [{"headerName": "人員", "field": "person", "pinned": "left", "width": 80, "editable": False, "cellStyle": {"fontWeight":"bold","borderRight":"2px solid #333","backgroundColor":"#fff"}}]
                for dt in w_dates:
                    child = [{"headerName": s, "field": f"{dt['str']}_{s}", "editable": dt["is_curr"], "cellEditor": "agSelectCellEditor", "cellEditorParams": {"values": a_opts}, "cellClass": "is_odd" if dt["date"].weekday()%2==0 else "is_even", "cellStyle": cell_style_js, "width": 55} for s in ["早","午","晚"]]
                    cd.append({"headerName": dt["disp"], "children": child, "headerClass": "header-odd" if dt["date"].weekday()%2==0 else "header-even"})
                AgGrid(pd.DataFrame(rows), gridOptions={"columnDefs": cd, "rowHeight": 40, "domLayout": 'autoHeight'}, allow_unsafe_jscode=True, theme="alpine", key=f"ag_final_{wi}")
                
                # 獨立繪製每一天的早、午、晚空閒名單
                off_cols = st.columns(len(w_dates))
                for idx, dt in enumerate(w_dates):
                    if dt["is_curr"]:
                        d_str = dt["str"]
                        html_parts = []
                        for sh in ["早","午","晚"]:
                            off_nicks = [nm2n.get(n, n) for n in shift_off_staff.get(f"{d_str}_{sh}", []) if n]
                            html_parts.append(f"<div style='margin-bottom:2px;'><b>{sh}:</b> {','.join(off_nicks) if off_nicks else '<span style=\"color:#ccc\">無</span>'}</div>")
                        
                        off_cols[idx].markdown(f"""<div class="off-staff-box" style="padding:5px;">{''.join(html_parts)}</div>""", unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)

            if st.form_submit_button("💾 同步更新與儲存"):
                st.session_state.config["saved_result"] = st.session_state.result; save_config(st.session_state.config); st.rerun()
                
        # 專屬 AI 匯出 JSON 功能區塊
        st.divider()
        with st.expander("💬 匯出當前班表 (與 AI 討論專用)"):
            st.info("點擊下方內容並複製，您可以直接貼給 ChatGPT 或 Claude 等 AI，請它幫忙找出班表盲點或檢查人力是否分配不均。")
            export_payload = {
                "month": f"{y}-{m}",
                "schedule": st.session_state.result,
                "rules": st.session_state.config.get("adv_rules", {}),
                "leaves": leaves_data
            }
            st.text_area("請複製以下 JSON 格式資料：", json.dumps(export_payload, ensure_ascii=False), height=200)

elif step == "8. 報表下載":
    st.header("下載 Excel 報表")
    if 'result' in st.session_state:
        sch = st.session_state.result; y, m = st.session_state.config["year"], st.session_state.config["month"]
        d = get_active_doctors(); a = get_active_assistants()
        c1, c2 = st.columns(2)
        c1.download_button("📊 總班表", to_excel_master(sch, y, m, d, a), f"祐德總班表_{m}月.xlsx")
        c2.download_button("👤 助理個人表", to_excel_individual(sch, y, m, a, d), f"祐德助理表_{m}月.xlsx")
