import re
import difflib

with open("server.py", "r", encoding="utf-8") as f:
    text = f.read()

# 1. Ensure difflib is imported
if "import difflib" not in text:
    text = text.replace("import openpyxl", "import openpyxl\nimport difflib")

# 2. Add sync helpers before save_quick_subtask_toggle
sync_helpers_code = '''
def extract_kks(text):
    if not text:
        return ""
    matches = re.findall(r'[0-9]{1,2}[A-Z]{3}[0-9]{2}[A-Z]{2}[0-9]{3}', str(text))
    if matches:
        return matches[0]
    matches2 = re.findall(r'[0-9]{1,2}[A-Z]{3}[0-9]{2}[A-Z]{1,2}[0-9]{1,3}', str(text))
    if matches2:
        return matches2[0]
    return ""

def is_same_component(comp_desc, comp_kks, st_desc):
    comp_kks = str(comp_kks or '').strip().upper()
    st_desc_upper = str(st_desc or '').strip().upper()
    comp_desc_upper = str(comp_desc or '').strip().upper()
    
    st_kks = extract_kks(st_desc_upper)
    
    # 1. KKS Match
    if comp_kks and st_kks:
        if comp_kks == st_kks or comp_kks[2:] == st_kks[2:]:
            return True
    if comp_kks and len(comp_kks) >= 6 and comp_kks in st_desc_upper:
        return True
        
    # 2. Text Match
    if comp_desc_upper and (comp_desc_upper in st_desc_upper or st_desc_upper in comp_desc_upper):
        return True
        
    norm_comp = re.sub(r'^(ACTUATOR|VALVE|MOV|AOV|UNIT \\d+|MSW)\\s*', '', comp_desc_upper).strip()
    norm_st = re.sub(r'^(ACTUATOR|VALVE|MOV|AOV|UNIT \\d+|MSW|[0-9]{1,2}[A-Z]{3}[0-9]{2}[A-Z]{1,2}[0-9]{1,3}[:\\s]*)\\s*', '', st_desc_upper).strip()
    
    if norm_comp and norm_st and norm_comp == norm_st:
        return True
        
    if len(norm_comp) > 10 and len(norm_st) > 10:
        ratio = difflib.SequenceMatcher(None, norm_comp, norm_st).ratio()
        if ratio > 0.85:
            return True
            
    return False

def sync_subtask_to_components(wb, no_wo, subtask_desc, is_done, date_str):
    today_str = date_str or datetime.date.today().strftime("%d/%m/%Y")
    
    # 1. Check ActuatorValve
    if "ActuatorValve" in wb.sheetnames:
        ws_act = wb["ActuatorValve"]
        for row in ws_act.iter_rows(min_row=2):
            eq_id = str(row[0].value or '').strip()
            desc = str(row[2].value or '').strip()
            kks = str(row[3].value or '').strip()
            if is_same_component(desc, kks or eq_id, subtask_desc):
                row[9].value = is_done
                row[10].value = is_done
                row[7].value = 100 if is_done else 0
                row[6].value = "FINISH" if is_done else "SCHED-OK"
                row[8].value = today_str if is_done else None

    # 2. Check Instruments
    inst_sheets = [
        "Instrument_PressureTX",
        "Instrument_TemperatureTX",
        "Instrument_PressureSwitch"
    ]
    for sname in inst_sheets:
        if sname in wb.sheetnames:
            ws_inst = wb[sname]
            for row in ws_inst.iter_rows(min_row=2):
                desc = str(row[2].value or '').strip()
                kks = str(row[3].value or '').strip()
                if is_same_component(desc, kks, subtask_desc):
                    if sname == "Instrument_PressureSwitch":
                        row[13].value = is_done
                        row[14].value = today_str if is_done else None
                        row[15].value = today_str if is_done else None
                        if len(row) > 20: row[20].value = is_done
                    else:
                        row[7].value = is_done
                        row[6].value = today_str if is_done else None
                        if len(row) > 12: row[12].value = is_done

def sync_actuator_to_subtasks(wb, eq_id, kks, desc, is_done, date_str):
    today_str = date_str or datetime.date.today().strftime("%d/%m/%Y")
    affected_wos = set()
    
    if "WorkOrder_Checklist" in wb.sheetnames:
        ws_chk = wb["WorkOrder_Checklist"]
        for row in ws_chk.iter_rows(min_row=2):
            no_wo = str(row[0].value or '').strip()
            st_desc = str(row[1].value or '').strip()
            if is_same_component(desc, kks or eq_id, st_desc):
                row[4].value = is_done
                row[2].value = today_str if is_done else None
                if no_wo:
                    affected_wos.add(no_wo)
                    
    recalculate_wos_progress(wb, affected_wos, today_str)

def sync_instrument_to_subtasks(wb, inst_type, kks, no_val, desc, is_done, date_str):
    today_str = date_str or datetime.date.today().strftime("%d/%m/%Y")
    affected_wos = set()
    
    if "WorkOrder_Checklist" in wb.sheetnames:
        ws_chk = wb["WorkOrder_Checklist"]
        for row in ws_chk.iter_rows(min_row=2):
            no_wo = str(row[0].value or '').strip()
            st_desc = str(row[1].value or '').strip()
            if is_same_component(desc, kks or no_val, st_desc):
                row[4].value = is_done
                row[2].value = today_str if is_done else None
                if no_wo:
                    affected_wos.add(no_wo)
                    
    recalculate_wos_progress(wb, affected_wos, today_str)

def recalculate_wos_progress(wb, affected_wos, date_str):
    if not affected_wos:
        return
    if "WorkOrder_Checklist" not in wb.sheetnames or "WorkOrder" not in wb.sheetnames:
        return
    
    ws_chk = wb["WorkOrder_Checklist"]
    ws_wo = wb["WorkOrder"]
    
    wo_stats = {wn: {'total': 0, 'done': 0} for wn in affected_wos}
    for row in ws_chk.iter_rows(min_row=2):
        wn = str(row[0].value or '').strip()
        if wn in wo_stats:
            wo_stats[wn]['total'] += 1
            if row[4].value == True:
                wo_stats[wn]['done'] += 1
                
    for row in ws_wo.iter_rows(min_row=2):
        wn = str(row[1].value or '').strip()
        if wn in wo_stats:
            total_st = wo_stats[wn]['total']
            done_st = wo_stats[wn]['done']
            row[10].value = total_st
            pct = round((done_st / total_st) * 100, 1) if total_st > 0 else 0.0
            row[11].value = pct
            if done_st == total_st and total_st > 0:
                row[8].value = "FINISH"
                if not row[7].value:
                    row[7].value = date_str
            elif done_st > 0:
                row[8].value = "IN PROGRESS"
                if row[8].value == "FINISH":
                    row[7].value = None
            else:
                row[8].value = "SCHED-OK"
                row[7].value = None

def load_master_components(unit):
    path = get_excel_path(unit)
    if not os.path.exists(path):
        return {"actuators": [], "instruments": []}
    
    with FILE_LOCK:
        wb = openpyxl.load_workbook(path, data_only=True)
        actuators = []
        if "ActuatorValve" in wb.sheetnames:
            ws = wb["ActuatorValve"]
            for r in list(ws.iter_rows(values_only=True))[1:]:
                if r[0] or r[2]:
                    actuators.append({
                        "equipment_id": clean_val(r[0]),
                        "area": clean_val(r[1]) or "BOILER",
                        "equipment_description": clean_val(r[2]),
                        "kks": clean_val(r[3]),
                        "pic": normalize_pic(clean_val(r[5])),
                        "status": clean_val(r[6]) or "SCHED-OK",
                        "persen_progress": clean_val(r[7]) or 0
                    })
                    
        instruments = []
        inst_sheets = [
            ("Instrument_PressureTX", "pressure_tx", "Pressure Transmitter (PT)"),
            ("Instrument_TemperatureTX", "temperature_tx", "Temperature Transmitter (TT)"),
            ("Instrument_PressureSwitch", "pressure_switch", "Pressure Switch (PS)")
        ]
        for sname, itype, ilabel in inst_sheets:
            if sname in wb.sheetnames:
                ws = wb[sname]
                for r in list(ws.iter_rows(values_only=True))[1:]:
                    if r[2]:
                        instruments.append({
                            "no": clean_val(r[0]),
                            "area": clean_val(r[1]) or "GENERAL",
                            "equipment": clean_val(r[2]),
                            "kks": clean_val(r[3]),
                            "type": itype,
                            "type_label": ilabel,
                            "range": clean_val(r[5]) if len(r)>5 else "",
                            "status_wdone": bool(r[13] if sname=="Instrument_PressureSwitch" else r[7])
                        })
                        
    return {"actuators": actuators, "instruments": instruments}
'''

# Check if sync_helpers_code already in text
if "def is_same_component" not in text:
    text = text.replace("def save_quick_subtask_toggle(data):", sync_helpers_code + "\ndef save_quick_subtask_toggle(data):")

# 3. Update save_quick_subtask_toggle to trigger sync
old_toggle = '''                    if (sub_task and str(r[1].value).strip() == sub_task) or (sub_idx is not None and match_idx == int(sub_idx)):
                        r[4].value = selesai
                        if selesai:
                            if not r[2].value:
                                r[2].value = today_str
                        else:
                            r[2].value = None
                        break'''

new_toggle = '''                    if (sub_task and str(r[1].value).strip() == sub_task) or (sub_idx is not None and match_idx == int(sub_idx)):
                        r[4].value = selesai
                        actual_sub_desc = str(r[1].value).strip()
                        if selesai:
                            if not r[2].value:
                                r[2].value = today_str
                        else:
                            r[2].value = None
                        # Auto-sync to Actuators and Instruments
                        sync_subtask_to_components(wb, no_wo, actual_sub_desc, selesai, today_str)
                        break'''

if old_toggle in text:
    text = text.replace(old_toggle, new_toggle)

# 4. Update save_batch_subtask_toggle to trigger sync
old_batch = '''            for r in ws_chk.iter_rows(min_row=2):
                if r[0].value and str(r[0].value).strip() == no_wo:
                    total_cnt += 1
                    r[4].value = is_done
                    r[2].value = today_str'''

new_batch = '''            for r in ws_chk.iter_rows(min_row=2):
                if r[0].value and str(r[0].value).strip() == no_wo:
                    total_cnt += 1
                    r[4].value = is_done
                    r[2].value = today_str
                    actual_sub_desc = str(r[1].value or '').strip()
                    if actual_sub_desc:
                        sync_subtask_to_components(wb, no_wo, actual_sub_desc, is_done, today_str)'''

if old_batch in text:
    text = text.replace(old_batch, new_batch)

# 5. Update save_quick_actuator_toggle to trigger sync
old_act_toggle = '''                    new_pct = r[7].value
                    new_status = r[6].value
                    break'''

new_act_toggle = '''                    new_pct = r[7].value
                    new_status = r[6].value
                    act_desc = str(r[2].value or '').strip()
                    act_kks = str(r[3].value or '').strip()
                    is_finish = (new_status == "FINISH" or new_pct == 100)
                    sync_actuator_to_subtasks(wb, eq_id, act_kks, act_desc, is_finish, today_str)
                    break'''

if old_act_toggle in text:
    text = text.replace(old_act_toggle, new_act_toggle)

# 6. Update save_quick_instrument_toggle to trigger sync
old_inst_toggle = '''                    break
        ok, err = safe_save_workbook(wb, path)'''

new_inst_toggle = '''                    inst_desc = str(row[2].value or '').strip()
                    inst_kks = str(row[3].value or '').strip()
                    inst_no = str(row[0].value or '').strip()
                    sync_instrument_to_subtasks(wb, inst_type, inst_kks, inst_no, inst_desc, status_wdone, today_str)
                    break
        ok, err = safe_save_workbook(wb, path)'''

if old_inst_toggle in text:
    text = text.replace(old_inst_toggle, new_inst_toggle, 1)

# 7. Add GET /api/master_components route
if 'elif path == "/api/master_components":' not in text:
    target_route = '''        elif path == "/api/actuator_matrix":
            matrix = load_actuator_matrix()
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps(matrix, ensure_ascii=False).encode('utf-8'))
            return'''
            
    expanded_route = target_route + '''

        elif path == "/api/master_components":
            unit = int(query.get("unit", [1])[0])
            data = load_master_components(unit)
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))
            return'''
    text = text.replace(target_route, expanded_route)

with open("server.py", "w", encoding="utf-8") as f:
    f.write(text)

print("Backend sync successfully patched into server.py!")
