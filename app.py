# --- 5. Excel 輸出 ---
def to_excel_master(schedule_result, year, month, docs, assts):
    output = io.BytesIO(); writer = pd.ExcelWriter(output, engine='xlsxwriter'); workbook = writer.book
    fmts = get_excel_formats(workbook); sheet = workbook.add_worksheet("總班表")
    
    padded_weeks = get_padded_weeks(year, month)
    
    # 標題只寫一次在最上方，並跨越所有欄位 (7天 * 3時段 = 21, 加上名稱欄 = 22)
    sheet.merge_range(0, 0, 0, 21, f"祐德牙醫診所 {month}月 班表", fmts['h_title_big'])
    sheet.set_row(0, 40) # 加高標題行
    sheet.set_column(0, 0, 11) # 人員名稱欄寬
    for c in range(1, 22): sheet.set_column(c, c, 4.5) # 班表時段欄寬
    
    row = 1
    for w_dates in padded_weeks:
        # 寫入日期列
        sheet.write(row, 0, "日期", fmts['name_col']); col = 1
        for dt in w_dates:
            # 判斷是否為單數日或非本月
            is_odd = dt['date'].weekday() % 2 == 0
            h_fmt = fmts['head_odd'] if is_odd else fmts['head_even']
            if not dt['is_curr']: h_fmt = fmts['gray_head']
            
            sheet.merge_range(row, col, row, col+2, dt['disp'], h_fmt)
            col += 3
        row += 1
        
        # 寫入早中晚列
        sheet.write(row, 0, "時段", fmts['name_col']); col = 1
        for dt in w_dates:
            is_odd = dt['date'].weekday() % 2 == 0
            f_m = fmts['morn_odd'] if is_odd else fmts['morn_even']
            f_a = fmts['aft_odd'] if is_odd else fmts['aft_even']
            f_e = fmts['eve_odd'] if is_odd else fmts['eve_even']
            
            # 非本月的格子變灰
            if not dt['is_curr']: 
                f_m = f_a = fmts['gray']
                f_e = fmts['gray_eve']
                
            sheet.write(row, col, "早", f_m)
            sheet.write(row, col+1, "午", f_a)
            sheet.write(row, col+2, "晚", f_e)
            col += 3
        row += 1
        
        # 寫入醫師班表
        for doc in docs:
            sheet.write(row, 0, doc["nick"], fmts['name_col']); col = 1
            for dt in w_dates:
                is_odd = dt['date'].weekday() % 2 == 0
                f_m = fmts['morn_odd'] if is_odd else fmts['morn_even']
                f_a = fmts['aft_odd'] if is_odd else fmts['aft_even']
                f_e = fmts['eve_odd'] if is_odd else fmts['eve_even']
                
                if not dt['is_curr']: 
                    f_m = f_a = fmts['gray']
                    f_e = fmts['gray_eve']
                
                for i, sh in enumerate(["早", "午", "晚"]):
                    fmt = [f_m, f_a, f_e][i]
                    val = ""
                    if dt['is_curr']:
                        k = f"{dt['str']}_{sh}"
                        anm = schedule_result.get(k, {}).get("doctors", {}).get(doc["name"], "")
                        val = next((a["nick"] for a in assts if a["name"]==anm), "")
                    sheet.write(row, col+i, val, fmt)
                col += 3
            row += 1
            
        # 寫入助理班表
        for rnm, rk, ri in [("櫃1","counter",0), ("櫃2","counter",1), ("流動","floater",0), ("流動2","floater",1), ("看/行","look",0)]:
            sheet.write(row, 0, rnm, fmts['name_col']); col = 1
            for dt in w_dates:
                is_odd = dt['date'].weekday() % 2 == 0
                f_m = fmts['morn_odd'] if is_odd else fmts['morn_even']
                f_a = fmts['aft_odd'] if is_odd else fmts['aft_even']
                f_e = fmts['eve_odd'] if is_odd else fmts['eve_even']
                
                if not dt['is_curr']: 
                    f_m = f_a = fmts['gray']
                    f_e = fmts['gray_eve']
                
                for i, sh in enumerate(["早", "午", "晚"]):
                    fmt = [f_m, f_a, f_e][i]
                    val = ""
                    if dt['is_curr']:
                        k = f"{dt['str']}_{sh}"
                        lst = schedule_result.get(k, {}).get(rk, [])
                        val = next((a["nick"] for a in assts if a["name"]==lst[ri]), "") if ri < len(lst) else ""
                    sheet.write(row, col+i, val, fmt)
                col += 3
            row += 1
            
        row += 1 # 每一週之間空一行
        
    writer.close(); output.seek(0); return output

def get_excel_formats(workbook):
    return {
        'h_title_big': workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 20, 'bg_color': '#D9E1F2', 'border': 1}),
        'h_name': workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 18, 'border': 1, 'bg_color': '#f8f9fa'}),
        'head_odd': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True, 'bg_color': '#FFD966', 'border': 1, 'right': 2}),
        'head_even': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True, 'bg_color': '#9DC3E6', 'border': 1, 'right': 2}),
        'morn_odd': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FDE9D9'}),
        'aft_odd': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FCD5B4'}),
        'eve_odd': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FABF8F', 'right': 2}),
        'morn_even': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#DDEBF7'}),
        'aft_even': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#BDD7EE'}),
        'eve_even': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#9DC3E6', 'right': 2}),
        'gray': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#f0f0f0', 'font_color': '#cccccc'}),
        'gray_eve': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#f0f0f0', 'font_color': '#cccccc', 'right': 2}),
        'gray_head': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#e0e0e0', 'font_color': '#888888', 'right': 2}),
        'name_col': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True, 'bg_color': '#ffffff', 'border': 1, 'right': 2}),
        'h_col': workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#E0E0E0', 'border': 1}),
        'c_norm': workbook.add_format({'align': 'center', 'border': 1}),
        'note_fmt': workbook.add_format({'font_size': 11, 'valign': 'vcenter'})
    }

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
                
        # 標題與名字加大合併
        s.merge_range(0, 0, 0, 10, f"祐德牙醫診所 {month}月 班表", fmts['h_title_big']); s.set_row(0, 30)
        s.merge_range(1, 0, 1, 10, f"{anm}", fmts['h_name']); s.set_row(1, 25)
        s.write(0, 12, f"上限: {a['custom_max'] or b_max}", fmts['c_norm'])
        s.write(1, 12, f"實排: {act}", fmts['c_norm'])
        
        # 設定欄寬避免超出
        for c in range(12): s.set_column(c, c, 5)
        s.set_column(0, 0, 6); s.set_column(6, 6, 6)
        
        for i, h in enumerate(["日期","星期","早","午","晚"]):
            s.write(3, i, h, fmts['h_col']); s.write(3, i+6, h, fmts['h_col'])
            
        mid = (len(dates)+1)//2
        for r, dt in enumerate(dates):
            col_off = 0 if r < mid else 6; row_off = r if r < mid else r - mid
            tr = row_off + 4
            s.write(tr, col_off, f"{dt.month}/{dt.day}", fmts['c_norm'])
            s.write(tr, col_off+1, ['一','二','三','四','五','六'][dt.weekday()], fmts['c_norm'])
            for ci, sh in enumerate(["早","午","晚"]):
                v = ""; data = schedule_result.get(f"{dt}_{sh}", {})
                if (dt.weekday(), sh) in parsed_admin.get(anm, set()): v="行"
                elif anm in data.get("look", []): v="看"
                elif anm in data["floater"]: v="流"
                elif anm in data["counter"]: v="櫃"
                else:
                    for dn, asn in data.get("doctors", {}).items():
                        if asn == anm: v = next((d["nick"] for d in docs if d["name"]==dn), dn)
                s.write(tr, col_off+2+ci, v, fmts['c_norm'])
                
        # 底部加上備註文字
        bottom_row = 4 + mid + 2
        notes = [
            "註：全診及午晚班有空請輪流抽空吃飯，謹守30分鐘規定，以免影響其他助理。",
            "1〉早午班 8:30AM~ 12:00AM 1:30PM~ 6:00PM。　2〉午晚班 1:30PM ~10:00PM。",
            "3〉早晚班 8:00AM~ 12:00AM 6:00PM~ 10:00PM。 4〉一個早班 8:00AM~12:00AM。",
            "5〉一個午班 2:00PM ~ 6:00PM。  6〉一個晚班 6:00PM ~10:00PM。",
            "7〉全診〈早中晚〉8:00AM~ 12:00AM 1:30PM~ 10:00PM。"
        ]
        for i, note in enumerate(notes):
            s.write(bottom_row + i, 0, note, fmts['note_fmt'])
            
    writer.close(); output.seek(0); return output

def to_excel_doctor_confirmed(manual_schedule, year, month, doc_name):
