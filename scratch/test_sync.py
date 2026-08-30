import openpyxl
import datetime
import re
import difflib

def extract_kks(text):
    if not text:
        return ""
    matches = re.findall(r'[0-9]{1,2}[A-Z]{3}[0-9]{2}[A-Z]{2}[0-9]{3}', text)
    if matches:
        return matches[0]
    matches2 = re.findall(r'[0-9]{1,2}[A-Z]{3}[0-9]{2}[A-Z]{1,2}[0-9]{1,3}', text)
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
    if comp_desc_upper and comp_desc_upper in st_desc_upper:
        return True
        
    norm_comp = re.sub(r'^(ACTUATOR|VALVE|MOV|AOV|UNIT \d+|MSW)\s*', '', comp_desc_upper).strip()
    norm_st = re.sub(r'^(ACTUATOR|VALVE|MOV|AOV|UNIT \d+|MSW|[0-9]{1,2}[A-Z]{3}[0-9]{2}[A-Z]{1,2}[0-9]{1,3}[:\s]*)\s*', '', st_desc_upper).strip()
    
    if norm_comp and norm_st and norm_comp == norm_st:
        return True
        
    if len(norm_comp) > 10 and len(norm_st) > 10:
        ratio = difflib.SequenceMatcher(None, norm_comp, norm_st).ratio()
        if ratio > 0.88:
            return True
            
    return False

print("Testing matching:")
print("1. Actuator Inlet ID Fan 1 MOV 10HNA61AA001 vs UNIT 1 INDUCED DRAUGHT FAN 1 INLET GATE ACTUATOR MOV 20HNA61AA001:", 
      is_same_component("ACTUATOR INLET ID FAN 1 MOV 10HNA61AA001", "10HNA61AA001", "UNIT 1 INDUCED DRAUGHT FAN 1 INLET GATE ACTUATOR MOV 20HNA61AA001"))
print("2. 10MAA10CP001 vs PRESSURE BEHIND TURBINE CONTROL:",
      is_same_component("PRESSURE BEHIND TURBINE CONTROL", "10MAA10CP001", "10MAA10CP001: PRESSURE BEHIND TURBINE CONTROL"))
