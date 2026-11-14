import re
import math
import datetime
import tkinter as tk
import json
import os
import sys
import threading
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
import customtkinter as ctk

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)


try:
    from openpyxl import load_workbook
except ModuleNotFoundError:
    load_workbook = None

try:
    import xlrd
except ModuleNotFoundError:
    xlrd = None

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
except Exception:
    TkinterDnD = None
    DND_FILES = None

CONFIG_FILE = "config.json"

WEIGHT_GUIDE_TEXT = """
워크 패키지 분석기 – 가중치(Manhour Factors) 계산 기준

목적
- WP(Work Package) 분석 결과에 "추가 Man-hour(M/H)"를 더해 TAT(일수)를 보정합니다.

최종 산출 로직
- Total M/H = MPD 합계 + AD/SB/EO 합계 + 가중치 합계
- TAT(일) 계산: (Total M/H / 100) 결과가 정수면 그대로 일수로 사용하고,
            소수점이 나오면 "반나절" 개념이 없으므로 다음 날(정수)로 계산합니다.

가중치 적용 항목 및 값
1) 중정비(C-Check 단계) – 파일명 기준
- 조건: Work Package 파일명에 "4C" 이상 포함 (예: 4C, 6C, 8C)
- 적용: +30 M/H (파일명에 4C 이상이 포함되면 1회 적용)

2) 노후 항공기 – 기령 기준
- 조건: Work Package 파일명에서 HL 기번(HL####) 추출
        Aircraft Status 파일에서 해당 기번을 C열에서 찾고 제작일(W열) 확인
        현재 기준 기령이 15년 이상
- 적용: +30 M/H

3) Wet Area Access Task
- 조건: MPD 시트 C열에 "53-380-00-01" 포함
        (단, 글씨/채움이 적색이거나 취소선이면 제외)
- 적용: +300 M/H
- 추가 규칙: Wet Area가 포함되어 있는데 계산 결과가 12일 미만이면, 예측 TAT는 12일로 표시합니다.

4) 보정 작업 항목(990 계열)
- 조건: MPD 시트 C열에 아래 중 하나라도 포함
  - 53-990-00-30J
  - 53-990-00-50J
- 적용: +10 M/H
- 적용 사유: FWD GALLEY 장탈이 필요한 작업이 포함될 수 있어 보정치로 보완합니다.
- 비고: 둘 다 발견되어도 +10 M/H 1회만 적용됩니다.

5) Wing Task (지정 Wing 카드 목록)
- 조건: MPD 시트 C열에 아래 코드들 중 하나라도 포함 (단, 적색이면 제외)
  - 57-020-00-01
  - 57-170-01-01
  - 57-822-01-01
  - 57-170-02-01
  - 57-892-02-01
  - 57-180-01-01
  - 57-826-01-01
  - 57-180-02-01
  - 57-896-02-01
  - 57-190-01-01
  - 57-830-01-01
  - 57-190-02-01
  - 57-900-02-01
- 적용: +40 M/H (여러 개가 발견되어도 최대 40 M/H 1회만 적용)
- 적용 사유: 일부 Wing Task는 BOEING STANDARD 공수(예: 6.0 M/H)만으로는 Seal/Gasket 점검·교환 및 Manual Torque 준수 작업까지 반영하기 어려워 보정치로 보완합니다.

6) 보정 작업 항목(개별 코드)
- 조건: MPD 시트 C열에 "27-108-00-01" 포함 → +50 M/H
- 적용 사유: 27-108-00-01(Remove the stabilizer trim actuator for restoration)은 BOEING STANDARD MANHOUR(6.0)만으로는 실제 작업 반영이 부족하여 보정치로 보완합니다.
- 조건: MPD 시트 C열에 "47-310-00-01" 포함 → +10 M/H
- 적용 사유: 47-310-00-01(Restore (clean) the Nitrogen Generation System (NGS) heat exchanger (off aircraft))은 BOEING STANDARD MANHOUR(0.5)만으로는 실제 작업 반영이 부족하여 보정치로 보완합니다.

시트/열 기준(합계 계산)
- MPD 합계:
  - 시트 이름에 "MPD" 포함
  - O열 값을 합산
  - 단, A열이 비어있으면 제외
  - 단, O열이 적색이면 제외

- AD/SB/EO 합계:
  - 시트 이름에 "AD SB EO" 포함
  - D열 값을 합산
  - 단, A열이 비어있으면 제외
  - 단, D열이 적색이면 제외
""".strip()

TASK_DESCRIPTIONS = {
    "53-380-00-01": "WET AREA ACCESS",
    "53-990-00-30J": "FWD GALLEY REMOVAL (Potential)",
    "53-990-00-50J": "FWD GALLEY REMOVAL (Potential)",
    "27-108-00-01": "Remove the stabilizer trim actuator for restoration",
    "47-310-00-01": "Restore (clean) the Nitrogen Generation System (NGS) heat exchanger (off aircraft)",
    "57-020-00-01": "Wing Leading Edge",
    "57-170-01-01": "Wing Fuel System - Left",
    "57-822-01-01": "Wing Fuel System - Left Check",
    "57-170-02-01": "Wing Fuel System - Right",
    "57-892-02-01": "Wing Fuel System - Right Check",
    "57-180-01-01": "Wing Leading Edge Systems - Left",
    "57-826-01-01": "Wing Systems Check - Left",
    "57-180-02-01": "Wing Leading Edge Systems - Right",
    "57-896-02-01": "Wing Systems Check - Right",
    "57-190-01-01": "Wing Tip Assembly - Left",
    "57-830-01-01": "Wing Tip Systems Check - Left",
    "57-190-02-01": "Wing Tip Assembly - Right",
    "57-900-02-01": "Wing Tip Systems Check - Right",
}

def _format_task_code_with_desc(code, found_desc_map=None):
    found_desc = None
    if isinstance(found_desc_map, dict):
        found_desc = found_desc_map.get(code)
    desc = found_desc or TASK_DESCRIPTIONS.get(code)
    return f"{code} ({desc})" if desc else f"{code} (Description 미등록)"

def _load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
    return {}

def _save_config(data):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except:
        pass

# --- Utilities ---

def _is_red_openpyxl(cell):
    try:
        c = cell.font.color
        if c and getattr(c, "rgb", None) and "FF0000" in str(c.rgb).upper():
            return True
        f = cell.fill.fgColor
        if f and getattr(f, "rgb", None) and "FF0000" in str(f.rgb).upper():
            return True
        return False
    except:
        return False

def _is_strike_openpyxl(cell):
    try:
        return bool(getattr(cell.font, "strike", False))
    except:
        return False

def _to_number(value):
    if value is None: return None
    if isinstance(value, (int, float)): return float(value)
    s = str(value).strip().replace(",", "")
    try: return float(s)
    except: return None

def _extract_hl(filename):
    m = re.search(r"(HL\d{4})", filename.upper())
    return m.group(1) if m else None

def _extract_c_level(filename):
    vals = [int(x) for x in re.findall(r"(\d+)C", filename.upper())]
    return max(vals) if vals else None

def _years_from(date_value):
    now = datetime.datetime.now().date()
    d = None
    if isinstance(date_value, datetime.datetime): d = date_value.date()
    elif isinstance(date_value, datetime.date): d = date_value
    elif isinstance(date_value, (int, float)):
        try: d = datetime.date(int(date_value), 1, 1)
        except: return None
    elif isinstance(date_value, str):
        s = date_value.strip()
        for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y-%m", "%Y/%m", "%Y.%m", "%Y"]:
            try:
                d = datetime.datetime.strptime(s, fmt).date()
                break
            except: pass
    return (now - d).days / 365.25 if d else None

# --- Main Analysis Logic ---

def _analyze_wp_openpyxl(path, progress_callback=None):
    wb = load_workbook(path, data_only=True)
    results = {
        "mpd_sum": 0.0, "ad_sum": 0.0, "found_codes": set(),
        "task_desc": {},
        "found_53_380": False, "found_53_990_30j": False, "found_53_990_50j": False,
        "found_wing": False, "found_wing_codes": [], "found_27_108": False, "found_47_310": False
    }
    
    wing_set = {
        "57-020-00-01","57-170-01-01","57-822-01-01","57-170-02-01","57-892-02-01",
        "57-180-01-01","57-826-01-01","57-180-02-01","57-896-02-01","57-190-01-01",
        "57-830-01-01","57-190-02-01","57-900-02-01"
    }

    mpd_ws = None
    for n in wb.sheetnames:
        if "MPD" in n.upper():
            mpd_ws = wb[n]
            break
    
    if mpd_ws:
        total_rows = mpd_ws.max_row
        for i, row in enumerate(mpd_ws.iter_rows()):
            if progress_callback:
                progress_callback((i / total_rows) * 50, f"MPD 데이터 분석 중... ({i + 1}/{total_rows}행)", is_progress=True)
            
            a_val = row[0].value
            c_val = str(row[2].value or "")
            
            # Optimization: Extract row description only if needed
            row_desc = None
            def get_row_desc():
                nonlocal row_desc
                if row_desc is not None: return row_desc
                for desc_cell in row[3:10]:
                    v = desc_cell.value
                    if isinstance(v, str) and v.strip():
                        row_desc = v.strip().replace("\r", "").replace("\n", " ")
                        return row_desc
                return ""

            # Check 53-380-00-01 (Wet Area)
            if "53-380-00-01" in c_val:
                c_cell = row[2]
                if not _is_red_openpyxl(c_cell) and not _is_strike_openpyxl(c_cell):
                    results["found_53_380"] = True
                    desc = get_row_desc()
                    if desc and "53-380-00-01" not in results["task_desc"]:
                        results["task_desc"]["53-380-00-01"] = desc

            # Check 990 Series
            if "53-990-00-30J" in c_val: results["found_53_990_30j"] = True
            if "53-990-00-50J" in c_val: results["found_53_990_50j"] = True
            
            # Check Wing Tasks
            for w in wing_set:
                if w in c_val:
                    c_cell = row[2]
                    if not _is_red_openpyxl(c_cell):
                        results["found_wing"] = True
                        if w not in results["found_wing_codes"]:
                            results["found_wing_codes"].append(w)
            
            # Check individual codes
            if "27-108-00-01" in c_val: results["found_27_108"] = True
            if "47-310-00-01" in c_val: results["found_47_310"] = True

            # Populate task descriptions for found codes
            for code in ("53-990-00-30J", "53-990-00-50J", "27-108-00-01", "47-310-00-01"):
                if code in c_val and code not in results["task_desc"]:
                    desc = get_row_desc()
                    if desc: results["task_desc"][code] = desc
            
            # Man-hour Sum (MPD)
            if a_val and str(a_val).strip():
                o_cell = row[14]
                if not _is_red_openpyxl(o_cell):
                    num = _to_number(o_cell.value)
                    if num: results["mpd_sum"] += num

    ad_ws = None
    for n in wb.sheetnames:
        if "AD SB EO" in n.upper():
            ad_ws = wb[n]
            break
    
    if ad_ws:
        total_rows = ad_ws.max_row
        for i, row in enumerate(ad_ws.iter_rows()):
            if progress_callback:
                progress_callback(50 + (i / total_rows) * 50, f"AD/SB/EO 데이터 분석 중... ({i + 1}/{total_rows}행)", is_progress=True)
                
            a_val = row[0].value
            if a_val and str(a_val).strip():
                d_cell = row[3]
                if not _is_red_openpyxl(d_cell):
                    num = _to_number(d_cell.value)
                    if num: results["ad_sum"] += num
                    
    return results

def _analyze_wp_com(path, progress_callback=None):
    import win32com.client
    excel = win32com.client.Dispatch("Excel.Application")
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
    except:
        pass
    
    norm_path = os.path.normpath(os.path.abspath(path))
    try:
        wb = excel.Workbooks.Open(norm_path, ReadOnly=True)
    except Exception as e:
        try: excel.Quit()
        except: pass
        raise e
    results = {
        "mpd_sum": 0.0, "ad_sum": 0.0,
        "task_desc": {},
        "found_53_380": False, "found_53_990_30j": False, "found_53_990_50j": False,
        "found_wing": False, "found_wing_codes": [], "found_27_108": False, "found_47_310": False
    }
    wing_set = {
        "57-020-00-01","57-170-01-01","57-822-01-01","57-170-02-01","57-892-02-01",
        "57-180-01-01","57-826-01-01","57-180-02-01","57-896-02-01","57-190-01-01",
        "57-830-01-01","57-190-02-01","57-900-02-01"
    }

    def is_red_com(cell):
        try:
            if cell.Font.Color == 255: return True
            if cell.Interior.Color == 255: return True
            return False
        except: return False

    try:
        mpd_ws = None
        for i in range(1, wb.Sheets.Count + 1):
            if "MPD" in wb.Sheets(i).Name.upper():
                mpd_ws = wb.Sheets(i)
                break
        
        if mpd_ws:
            data = mpd_ws.UsedRange.Value
            if data:
                total_rows = len(data)
                for r_idx, row in enumerate(data):
                    if progress_callback:
                        progress_callback((r_idx / total_rows) * 50, f"MPD 데이터 분석 중(정밀 모드)... ({r_idx + 1}/{total_rows}행)", is_progress=True)
                        
                    if not row or len(row) < 15: continue
                    a_val, c_val, o_val = row[0], str(row[2] or ""), row[14]
                    
                    # Optimization: Extract row description only if needed
                    row_desc = None
                    def get_row_desc_com():
                        nonlocal row_desc
                        if row_desc is not None: return row_desc
                        for v in row[3:10]:
                            if isinstance(v, str) and v.strip():
                                row_desc = v.strip().replace("\r", "").replace("\n", " ")
                                return row_desc
                        return ""

                    # Check 53-380-00-01 (Wet Area)
                    if "53-380-00-01" in c_val:
                        cell_c = mpd_ws.Cells(r_idx + 1, 3)
                        if not is_red_com(cell_c) and not cell_c.Font.Strikethrough:
                            results["found_53_380"] = True
                            desc = get_row_desc_com()
                            if desc and "53-380-00-01" not in results["task_desc"]:
                                results["task_desc"]["53-380-00-01"] = desc

                    # Check 990 Series
                    if "53-990-00-30J" in c_val: results["found_53_990_30j"] = True
                    if "53-990-00-50J" in c_val: results["found_53_990_50j"] = True
                    
                    # Check Wing Tasks
                    for w in wing_set:
                        if w in c_val:
                            cell_c = mpd_ws.Cells(r_idx + 1, 3)
                            if not is_red_com(cell_c):
                                results["found_wing"] = True
                                if w not in results["found_wing_codes"]:
                                    results["found_wing_codes"].append(w)
                    
                    # Check individual codes
                    if "27-108-00-01" in c_val: results["found_27_108"] = True
                    if "47-310-00-01" in c_val: results["found_47_310"] = True

                    # Populate task descriptions for found codes
                    for code in ("53-990-00-30J", "53-990-00-50J", "27-108-00-01", "47-310-00-01"):
                        if code in c_val and code not in results["task_desc"]:
                            desc = get_row_desc_com()
                            if desc: results["task_desc"][code] = desc
                    
                    # Man-hour Sum (MPD)
                    if a_val and str(a_val).strip():
                        cell_o = mpd_ws.Cells(r_idx + 1, 15)
                        if not is_red_com(cell_o):
                            num = _to_number(o_val)
                            if num: results["mpd_sum"] += num

        ad_ws = None
        for i in range(1, wb.Sheets.Count + 1):
            if "AD SB EO" in wb.Sheets(i).Name.upper():
                ad_ws = wb.Sheets(i)
                break
        if ad_ws:
            data = ad_ws.UsedRange.Value
            if data:
                total_rows = len(data)
                for r_idx, row in enumerate(data):
                    if progress_callback:
                        progress_callback(50 + (r_idx / total_rows) * 50, f"AD/SB/EO 데이터 분석 중(정밀 모드)... ({r_idx + 1}/{total_rows}행)", is_progress=True)
                        
                    if not row or len(row) < 4: continue
                    a_val, d_val = row[0], row[3]
                    if a_val and str(a_val).strip() and d_val:
                        cell_d = ad_ws.Cells(r_idx + 1, 4)
                        if not is_red_com(cell_d):
                            num = _to_number(d_val)
                            if num: results["ad_sum"] += num
    finally:
        try: wb.Close(False)
        except: pass
        try: excel.Quit()
        except: pass
    return results

def _get_mfg_date(path, hl_code):
    norm_path = os.path.normpath(os.path.abspath(path))
    
    # 1. xlrd 우선 시도 (.xls 파일에 최적화)
    if xlrd and (norm_path.lower().endswith(".xls") or not norm_path.lower().endswith(".xlsx")):
        try:
            wb = xlrd.open_workbook(norm_path)
            ws = wb.sheet_by_index(0)
            for r in range(ws.nrows):
                row = ws.row_values(r)
                if len(row) >= 23 and str(row[2]).strip().upper() == hl_code:
                    val = row[22]
                    if isinstance(val, float) and val > 0:
                        try:
                            date_tuple = xlrd.xldate_as_tuple(val, wb.datemode)
                            return datetime.date(date_tuple[0], date_tuple[1], date_tuple[2])
                        except Exception:
                            pass
                    return val
            return None
        except Exception:
            pass

    # 2. openpyxl 시도 (.xlsx 파일)
    if load_workbook:
        try:
            wb = load_workbook(norm_path, data_only=True, read_only=True)
            ws = wb.active
            for row in ws.iter_rows(values_only=True):
                if len(row) >= 23 and str(row[2]).strip().upper() == hl_code:
                    return row[22]
            return None
        except Exception:
            pass

    # 3. win32com Fallback
    excel = None
    wb = None
    try:
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(norm_path, ReadOnly=True)
        ws = wb.ActiveSheet
        data = ws.UsedRange.Value
        if data:
            for row in data:
                if row and len(row) >= 23 and str(row[2]).strip().upper() == hl_code:
                    return row[22]
        return None
    except Exception:
        return None
    finally:
        if wb:
            try: wb.Close(False)
            except: pass
        if excel:
            try: excel.Quit()
            except: pass

def analyze(wp_path, st_path, progress_callback=None):
    reasons, weights = [], 0
    filename = os.path.basename(wp_path)
    
    if progress_callback: progress_callback(5, "파일명 기반 중정비 단계 분석 중...")
    
    c_level = _extract_c_level(filename)
    if c_level and c_level >= 4:
        weights += 30
        reasons.append(f"• [중정비] {c_level}C 단계 확인 (+30 M/H)\n  - 사유: 4C 이상의 고단계 중정비는 Structure 결함 발생 빈도가 높음")
    elif c_level:
        reasons.append(f"• [중정비] {c_level}C 단계 (4C 미만 가중치 미적용)")
    else:
        reasons.append("• [중정비] 단계 확인 불가 (가중치 미적용)")
    
    if progress_callback: progress_callback(10, "항공기 기령 확인 및 노후 항공기 분석 중...")
    
    hl = _extract_hl(filename)
    if hl:
        mfg = _get_mfg_date(st_path, hl)
        yrs = _years_from(mfg)
        if yrs and yrs >= 15:
            weights += 30
            reasons.append(f"• [노후 항공기] {hl} 기령 {yrs:.1f}년 확인 (+30 M/H)\n  - 사유: 15년 이상의 기령은 균열과 부식에 취약함")
        else:
            reasons.append(f"• [노후 항공기] 해당 없음 ({hl}, 기령: {yrs:.1f}년)" if yrs else f"• [노후 항공기] {hl} 제작일 정보 없음")
    else:
        reasons.append("• [노후 항공기] 파일명에서 HL기번을 찾을 수 없음")

    if progress_callback: progress_callback(15, "엑셀 시트 데이터 추출 및 Man-hour 합산 시작...")

    try:
        def openpyxl_progress_mapper(internal_val, internal_msg=None, is_progress=False):
            if progress_callback:
                scaled_progress = 15 + (internal_val * (50 - 15) / 100)
                progress_callback(scaled_progress, internal_msg, is_progress=is_progress)

        res = _analyze_wp_openpyxl(wp_path, openpyxl_progress_mapper)
    except Exception as e_openpyxl:
        print(f"[DEBUG] Openpyxl fallback triggered: {e_openpyxl}")
        if progress_callback: progress_callback(50, "외부 생성 파일 형식 감지: Excel 정밀 호환 모드로 전환하여 분석합니다...")
        try:
            def com_progress_mapper(internal_val, internal_msg=None, is_progress=False):
                if progress_callback:
                    scaled_progress = 50 + (internal_val * (90 - 50) / 100)
                    progress_callback(scaled_progress, internal_msg, is_progress=is_progress)
            res = _analyze_wp_com(wp_path, com_progress_mapper)
        except Exception as e_com: return {"error": f"파일 분석 실패: {e_com}"}

    if progress_callback: progress_callback(95, "가중치 최종 계산 및 TAT 산출 중...")

    if res["found_53_380"]:
        weights += 300
        reasons.append(f"• [Wet Area] {_format_task_code_with_desc('53-380-00-01', res.get('task_desc'))} 포함 확인 (+300 M/H)\n  - 사유: Wet Area Access 작업 포함")
    
    codes_990 = []
    if res["found_53_990_30j"]: codes_990.append("53-990-00-30J")
    if res["found_53_990_50j"]: codes_990.append("53-990-00-50J")
    if codes_990:
        weights += 10
        code_str = ", ".join(_format_task_code_with_desc(c, res.get("task_desc")) for c in codes_990)
        reasons.append(f"• [보정 작업 항목] {code_str} 포함 확인 (+10 M/H)\n  - 사유: FWD GALLEY 장탈이 필요한 작업이 포함될 수 있어 보정치로 보완함\n  - 비고: 여러 개가 있어도 가중치는 1회만 적용됨")
        
    if res["found_wing"]:
        weights += 40
        if res["found_wing_codes"]:
            wing_codes_str = ", ".join([f"{code} ({TASK_DESCRIPTIONS.get(code, 'Unknown')})" for code in res["found_wing_codes"]])
            reasons.append(f"• [Wing Task] {wing_codes_str} 포함 확인 (+40 M/H)\n  - 사유: BOEING STANDARD 공수(예: 6.0 M/H)만으로는 Seal/Gasket 점검·교환 및 Manual Torque 준수 작업까지 반영하기 어려워 보정치로 보완함\n  - 비고: 목록 중 여러 개가 발견되어도 최대 40 M/H만 적용됨")
        else:
            reasons.append("• [Wing Task] 주요 Wing Task Card 포함 확인 (+40 M/H)\n  - 사유: BOEING STANDARD 공수(예: 6.0 M/H)만으로는 Seal/Gasket 점검·교환 및 Manual Torque 준수 작업까지 반영하기 어려워 보정치로 보완함\n  - 비고: 목록 중 여러 개가 발견되어도 최대 40 M/H만 적용됨")
    
    if res["found_27_108"]:
        weights += 50
        reasons.append(f"• [보정 작업 항목] {_format_task_code_with_desc('27-108-00-01', res.get('task_desc'))} 포함 확인 (+50 M/H)\n  - 사유: BOEING STANDARD MANHOUR(6.0)만으로는 실제 작업 반영이 부족하여 보정치로 보완함")
    
    if res["found_47_310"]:
        weights += 10
        reasons.append(f"• [보정 작업 항목] {_format_task_code_with_desc('47-310-00-01', res.get('task_desc'))} 포함 확인 (+10 M/H)\n  - 사유: BOEING STANDARD MANHOUR(0.5)만으로는 실제 작업 반영이 부족하여 보정치로 보완함")

    total = res["mpd_sum"] + res["ad_sum"] + weights
    tat_days = math.ceil(total / 100.0)
    wet_area_min_tat_applied = False
    if res["found_53_380"] and tat_days < 12:
        tat_days = 12
        wet_area_min_tat_applied = True
    
    calc_explanation = (
        f"최종 산출식:\n"
        f"(MPD {res['mpd_sum']:,.1f} + AD/SB/EO {res['ad_sum']:,.1f} + 가중치 {weights})\n"
        f"= {total:,.1f} ÷ 100 = {total/100:.2f}일\n"
        f"➔ 소수점 올림 계산하여 {tat_days}일 예측"
    )
    if wet_area_min_tat_applied:
        calc_explanation += "\n➔ Wet Area 포함으로 최소 12일 보정 적용"
    
    if progress_callback: progress_callback(100)
    
    return {
        "mpd_sum": res["mpd_sum"], "ad_sum": res["ad_sum"], "weights": weights,
        "total": total, "tat_days": tat_days, "reasons": reasons,
        "calculation_logic": calc_explanation
    }

# --- UI (Futuristic Modern Style) ---

def run_ui():
    if not load_workbook:
        messagebox.showerror("오류", "openpyxl 패키지가 필요합니다.")
        return

    # Root Window with DnD Support
    if TkinterDnD:
        class App(ctk.CTk, TkinterDnD.DnDWrapper):
            def __init__(self):
                super().__init__()
                # Try to initialize dnd without calling internal _pytkdndversion directly
                try:
                    self.TkdndVersion = TkinterDnD.TkinterDnD._pytkdndversion(self)
                except AttributeError:
                    # Fallback for different versions of tkinterdnd2
                    try:
                        self.TkdndVersion = self.tk.call('dnd', 'version')
                    except Exception:
                        self.TkdndVersion = None
        root = App()
    else:
        root = ctk.CTk()

    root.title("Work Package 분석기")
    root.geometry("1105x585")
    
    app_icon_path = resource_path("app_icon.ico")
    if os.path.exists(app_icon_path):
        try:
            root.iconbitmap(app_icon_path)
            root.after(200, lambda: root.iconbitmap(app_icon_path))
        except Exception:
            pass

    
    # Main Split Layout
    main_split = ctk.CTkFrame(root, fg_color="transparent")
    main_split.pack(fill=tk.BOTH, expand=True)

    left_frame = ctk.CTkFrame(main_split, fg_color="transparent")
    left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    right_frame = ctk.CTkFrame(main_split, fg_color="#1E1E1E", width=550) 
    right_frame.pack(side=tk.RIGHT, fill=tk.BOTH)
    right_frame.pack_propagate(False)

    terminal_label = ctk.CTkLabel(right_frame, text="프로그램 진행 로그", font=("Inter", 14, "bold"), text_color="#FFFFFF")
    terminal_label.pack(pady=(10, 5))

    terminal_textbox = ctk.CTkTextbox(right_frame, fg_color="#000000", text_color="#00FF00", font=("Consolas", 12))
    terminal_textbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
    terminal_textbox.configure(state="disabled")

    def log_to_terminal(message, is_progress=False):
        timestamp = datetime.datetime.now().strftime("[%H:%M:%S] ")
        terminal_textbox.configure(state="normal")
        
        task_key = message.split("...")[0] if is_progress and "..." in message else None
        current_key = getattr(terminal_textbox, "_progress_task_key", None)
        
        if is_progress and current_key == task_key and getattr(terminal_textbox, "_in_progress_line", False):
            terminal_textbox.delete("end-2c linestart", "end-2c lineend")
            terminal_textbox.insert("end-2c linestart", terminal_textbox._progress_prefix + message)
        else:
            terminal_textbox.insert(tk.END, timestamp + message + "\n")
            if is_progress:
                terminal_textbox._in_progress_line = True
                terminal_textbox._progress_prefix = timestamp
                terminal_textbox._progress_task_key = task_key
            else:
                terminal_textbox._in_progress_line = False
                terminal_textbox._progress_task_key = None
                
        terminal_textbox.see(tk.END)
        terminal_textbox.configure(state="disabled")
        root.update_idletasks()

    wp_var, st_var = tk.StringVar(), tk.StringVar()
    config = _load_config()
    
    # Pre-declare UI variables to avoid UnboundLocalError in nested functions
    btn_run = None
    progress_bar = None
    status_label = None
    progress_container = None
    
    if "wp_path" in config: wp_var.set(config["wp_path"])
    if "status_path" in config: st_var.set(config["status_path"])
    
    # Font System (Clean Windows Malgun Gothic)
    FONT_FAMILY = "맑은 고딕"
    F_TITLE = (FONT_FAMILY, 24, "bold")
    F_SUBTITLE = (FONT_FAMILY, 16, "bold")
    F_LABEL = (FONT_FAMILY, 14, "bold")
    F_BODY = (FONT_FAMILY, 13)
    F_BTN = (FONT_FAMILY, 14, "bold")
    F_SMALL = (FONT_FAMILY, 11)
    F_TAT = (FONT_FAMILY, 34, "bold")

    # Color Constants (Jeju Clean Design)
    BG_COLOR = "#FFFFFF"
    CARD_BG = "#F8F9FA"
    ACCENT = "#FF7F00"
    HOVER = "#E67200"
    TEXT_MAIN = "#1A1A1A"
    TEXT_DIM = "#666666"
    BORDER_COLOR = "#E0E0E0"
    HEADER_TEXT = "#FFFFFF"
    STAT_WEIGHT = "#D97706"

    # --- UI Layout ---
    
    # Global containers for view switching
    input_view = ctk.CTkFrame(root, fg_color=BG_COLOR)
    result_view = ctk.CTkFrame(root, fg_color=BG_COLOR)
    
    def show_view(view):
        input_view.pack_forget()
        result_view.pack_forget()
        view.pack(in_=left_frame, fill=tk.BOTH, expand=True)

    # --- Input View Construction ---
    # Title Section
    title_frame = ctk.CTkFrame(input_view, fg_color=ACCENT, corner_radius=0)
    title_frame.pack(fill=tk.X)
    
    title_label = ctk.CTkLabel(title_frame, text="Work Package 분석기", font=F_TITLE, 
                 text_color=HEADER_TEXT)
    title_label.pack(pady=20)

    # Main Container for inputs
    main_container = ctk.CTkFrame(input_view, fg_color=BG_COLOR)
    main_container.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)
    
    # Input Card
    input_card = ctk.CTkFrame(main_container, fg_color="transparent")
    input_card.pack(fill=tk.BOTH, expand=True)
    
    inner_padding = ctk.CTkFrame(input_card, fg_color="transparent")
    inner_padding.pack(fill=tk.BOTH, expand=True, padx=30, pady=10)

    def create_modern_input(parent, label, var, pick_fn, placeholder):
        label_frame = ctk.CTkFrame(parent, fg_color="transparent")
        label_frame.pack(fill=tk.X, pady=(0, 8))
        
        ctk.CTkLabel(label_frame, text=label, font=F_LABEL, 
                     text_color=TEXT_DIM).pack(side=tk.LEFT)
        
        # Container frame for the entire bar
        entry_frame = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=2, border_width=1, border_color=BORDER_COLOR)
        entry_frame.pack(fill=tk.X, pady=(0, 15))
        
        def on_pick():
            initial_dir = str(config.get("wp_dir" if pick_fn.__name__ == "pick_wp" else "status_dir", ""))
            p = filedialog.askopenfilename(filetypes=[("Excel Files","*.xlsx *.xls")], 
                                         initialdir=initial_dir if os.path.isdir(initial_dir) else None)
            if p: pick_fn(p)
            
        btn_pick = ctk.CTkButton(entry_frame, text="찾아보기", font=F_SMALL,
                                 fg_color=ACCENT, hover_color=HOVER, 
                                 text_color="#FFFFFF",
                                 width=100, height=38, corner_radius=2, command=on_pick)
        btn_pick.pack(side=tk.RIGHT, padx=5, pady=5)

        # entry must have pady to NOT cover the border of entry_frame
        entry = ctk.CTkEntry(entry_frame, textvariable=var, font=F_BODY,
                             fg_color="transparent", border_width=0, height=38,
                             corner_radius=0, placeholder_text=placeholder, text_color=TEXT_MAIN)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(15, 0), pady=5)
    
    # ... (rest of the functions need to use the dynamic variables) ...


    def pick_wp(p):
        wp_var.set(p)
        config["wp_path"] = p
        config["wp_dir"] = os.path.dirname(os.path.normpath(p))
        _save_config(config)
    
    def pick_st(p):
        st_var.set(p)
        config["status_path"] = p
        config["status_dir"] = os.path.dirname(os.path.normpath(p))
        _save_config(config)

    def show_guide():
        g_win = ctk.CTkToplevel(root)
        g_win.title("가중치 계산 기준 가이드")
        if os.path.exists(app_icon_path):
            try:
                g_win.after(100, lambda: g_win.iconbitmap(app_icon_path))
            except Exception:
                pass
        g_win.geometry("700x750")
        g_win.configure(fg_color="#FFFFFF")
        g_win.after(10, g_win.lift)
        
        ctk.CTkLabel(g_win, text="가중치 및 TAT 산출 기준", font=F_SUBTITLE, text_color=TEXT_MAIN).pack(pady=15)
        
        txt_box = ctk.CTkTextbox(g_win, fg_color="#F8F9FA", text_color=TEXT_MAIN, font=F_BODY, corner_radius=0, border_width=1, border_color=BORDER_COLOR)
        txt_box.pack(fill=tk.BOTH, expand=True, padx=25, pady=(0, 25))
        txt_box.insert("0.0", WEIGHT_GUIDE_TEXT)
        txt_box.configure(state="disabled")
        
        # Mouse wheel support for guide
        def _on_mousewheel_guide(event):
            try:
                # 가이드 창이 존재하고, 이벤트가 발생한 위젯이 가이드 창 소속일 때만 스크롤
                if g_win.winfo_exists() and event.widget.winfo_toplevel() == g_win:
                    txt_box.yview_scroll(int(-1 * (event.delta / 120) * 4), "units")
            except:
                pass
        
        # 가이드 창 내부 어디서든 스크롤이 가능하도록 bind_all 사용 (핸들러에서 윈도우 체크 수행)
        g_win.bind_all("<MouseWheel>", _on_mousewheel_guide)
        
        # 가이드 창 종료 시 바인딩 해제
        def _on_close_guide():
            try:
                g_win.unbind_all("<MouseWheel>")
            except:
                pass
            g_win.destroy()
            
        g_win.protocol("WM_DELETE_WINDOW", _on_close_guide)

    # Track last log message to avoid duplicates
    root.last_log_message = ""

    def update_progress(val, message=None, is_progress=False):
        progress_bar.set(val / 100)
        status_text = f"분석 중... {int(val)}%"
        status_label.configure(text=status_text)
        
        if message:
            log_to_terminal(message, is_progress=is_progress)
            root.last_log_message = message
        elif int(val) % 10 == 0:
            msg = status_text
            if msg != root.last_log_message:
                log_to_terminal(msg)
                root.last_log_message = msg
        root.update_idletasks()

    def show_results(res):
        progress_bar.set(1)
        status_label.configure(text="분석 완료")
        btn_run.configure(state="normal", text="분석 시작")
        root.config(cursor="") # 분석 완료 시 커서 복원
        
        if "error" in res:
            messagebox.showerror("오류", res["error"])
            return

        # Clear previous result view content
        for widget in result_view.winfo_children():
            widget.destroy()
            
        # Header
        header = ctk.CTkFrame(result_view, fg_color=ACCENT, corner_radius=0, height=80)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        
        ctk.CTkLabel(header, text="분석 결과 리포트", font=F_SUBTITLE, text_color=HEADER_TEXT).pack(side=tk.LEFT, padx=30)
        
        def back_to_input():
            root.unbind_all("<MouseWheel>")
            root.unbind_all("<Button-4>")
            root.unbind_all("<Button-5>")
            show_view(input_view)
            status_label.configure(text="시스템 준비됨")
            progress_bar.set(0)

        ctk.CTkButton(header, text="다시 분석", font=F_BTN, width=100, height=38, corner_radius=0,
                      fg_color="#333333", hover_color="#000000", 
                      text_color="#FFFFFF", command=back_to_input).pack(side=tk.RIGHT, padx=30)

        ctk.CTkButton(header, text="ⓘ 산출 기준 가이드", font=F_BTN, width=200, height=38, corner_radius=0,
                      fg_color="transparent", border_width=1, border_color=HEADER_TEXT, 
                      text_color=HEADER_TEXT, hover_color=HOVER,
                      command=show_guide).pack(side=tk.RIGHT)

        # Scrollable Content
        scroll = ctk.CTkScrollableFrame(result_view, fg_color="transparent")
        scroll.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Responsive tracking lists
        responsive_labels = []
        summary_cards = []

        # Global mouse wheel support for the result view
        def _on_mousewheel(event):
            if not result_view.winfo_ismapped():
                return
            
            # Use coordinate-based check for better reliability with CustomTkinter
            try:
                x_mouse = root.winfo_pointerx()
                y_mouse = root.winfo_pointery()
                
                left_x = left_frame.winfo_rootx()
                left_y = left_frame.winfo_rooty()
                left_w = left_frame.winfo_width()
                left_h = left_frame.winfo_height()
                
                # Check if mouse is within left_frame bounds
                if not (left_x <= x_mouse <= left_x + left_w and left_y <= y_mouse <= left_y + left_h):
                    return
            except:
                return

            try:
                if scroll.winfo_exists():
                    if event.delta != 0: # Windows/macOS
                        # 스크롤 이동량을 약 3배 이상 대폭 늘려 시원하게 스크롤되도록 개선
                        scroll._parent_canvas.yview_scroll(int(-1 * (event.delta / 2.5)), "units")
                    elif event.num == 4: # Linux scroll up
                        scroll._parent_canvas.yview_scroll(-15, "units")
                    elif event.num == 5: # Linux scroll down
                        scroll._parent_canvas.yview_scroll(15, "units")
            except:
                pass
            
        root.bind_all("<MouseWheel>", _on_mousewheel)
        root.bind_all("<Button-4>", _on_mousewheel)
        root.bind_all("<Button-5>", _on_mousewheel)
        
        # Force scrollbar update after a short delay
        def force_scroll_update():
            try:
                if scroll.winfo_exists():
                    # This triggers internal scrollregion recalculation in CTkScrollableFrame
                    scroll._parent_canvas.configure(scrollregion=scroll._parent_canvas.bbox("all"))
            except: pass
            
        root.after(100, force_scroll_update)
        root.after(500, force_scroll_update) # Second pass for safe measure

        # Summary Cards
        summary_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        summary_frame.pack(fill=tk.X, pady=10)
        summary_frame._is_vertical = False # Track layout state
        
        # Configure equal-width columns (1:1:1 ratio)
        summary_frame.grid_columnconfigure(0, weight=1, uniform="stat_card")
        summary_frame.grid_columnconfigure(1, weight=1, uniform="stat_card")
        summary_frame.grid_columnconfigure(2, weight=1, uniform="stat_card")
        
        def create_stat_card(parent, col_idx, title, value, color=TEXT_MAIN, subtext=""):
            card = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=0, border_width=1, border_color=BORDER_COLOR)
            card.grid(row=0, column=col_idx, sticky="nsew", padx=5)
            # Add minimum height to prevent invisible cards
            card.configure(height=120)
            card.pack_propagate(False)
            
            ctk.CTkLabel(card, text=title, font=F_LABEL, text_color=TEXT_DIM).pack(pady=(15, 2))
            ctk.CTkLabel(card, text=value, font=F_SUBTITLE, text_color=color).pack(pady=(0, 2))
            if subtext:
                ctk.CTkLabel(card, text=subtext, font=F_SMALL, text_color=TEXT_DIM).pack(pady=(0, 10))
            return card

        summary_cards.append(create_stat_card(summary_frame, 0, "MPD 합계", f"{res['mpd_sum']:,.1f}", ACCENT, "Man-Hours"))
        summary_cards.append(create_stat_card(summary_frame, 1, "AD/SB/EO", f"{res['ad_sum']:,.1f}", ACCENT, "Man-Hours"))
        summary_cards.append(create_stat_card(summary_frame, 2, "가중치", f"+{res['weights']}", STAT_WEIGHT, "M/H 추가"))

        # TAT Highlight (Compact)
        tat_card = ctk.CTkFrame(scroll, fg_color=ACCENT, corner_radius=0)
        tat_card.pack(fill=tk.X, pady=(10, 20))
        
        ctk.CTkLabel(tat_card, text="예측 TAT", font=F_LABEL, text_color=HEADER_TEXT).pack(pady=(10, 0))
        ctk.CTkLabel(tat_card, text=f"{res['tat_days']}일", font=F_SUBTITLE, text_color=HEADER_TEXT).pack(pady=(0, 10))

        # Details
        detail_label = ctk.CTkLabel(scroll, text="가중치 적용 내역 및 로직", font=F_SUBTITLE, text_color=TEXT_MAIN)
        detail_label.pack(anchor="w", pady=(10, 15))
        
        def _parse_reason(raw_str):
            lines = [l.strip() for l in raw_str.strip().split('\n') if l.strip()]
            hdr = lines[0] if lines else ""
            m_tag = re.search(r'\[(.*?)\]', hdr)
            tag = m_tag.group(1) if m_tag else "보정 항목"
            m_weight = re.search(r'\(([-+]?\d+\s*M/H|[^)]*미적용[^)]*)\)', hdr)
            weight = m_weight.group(1) if m_weight else ""
            title = hdr
            if m_tag: title = title.split('[' + tag + ']', 1)[-1].strip()
            if m_weight: title = title.replace(m_weight.group(0), '').strip()
            if title.startswith('•'): title = title[1:].strip()
            task_items = []
            if '포함 확인' in title and ',' in title:
                tasks_part = title.replace('포함 확인', '').strip()
                raw_items = [t.strip() for t in re.split(r',\s*(?=\d{2}-)', tasks_part) if t.strip()]
                if len(raw_items) > 1:
                    task_items = raw_items
                    title = f"해당 Task Card 총 {len(raw_items)}건 포함 확인"
            reason_desc = ""
            note_desc = ""
            for sub in lines[1:]:
                if sub.startswith('- 사유:'): reason_desc = sub.replace('- 사유:', '').strip()
                elif sub.startswith('- 비고:'): note_desc = sub.replace('- 비고:', '').strip()
                elif not reason_desc: reason_desc = sub.lstrip('- •').strip()
            is_na = '미적용' in weight or '해당 없음' in title or '정보 없음' in title
            if not weight: weight = '미적용' if is_na else '적용됨'
            return {'tag': tag, 'title': title, 'weight': weight, 'task_items': task_items, 'reason': reason_desc, 'note': note_desc, 'is_na': is_na}

        for reason in res["reasons"]:
            p = _parse_reason(reason)
            is_na = p["is_na"]
            card = ctk.CTkFrame(scroll, fg_color=CARD_BG, corner_radius=2, border_width=1, border_color=BORDER_COLOR)
            card.pack(fill=tk.X, pady=6)
            
            # Header Row (Tag Badge + Title on Left, Weight Badge on Right)
            hdr = ctk.CTkFrame(card, fg_color="transparent")
            hdr.pack(fill=tk.X, padx=16, pady=(12, 6))
            
            tag_bg = "#FFF3E0" if not is_na else "#EEEEEE"
            tag_col = "#E65100" if not is_na else "#757575"
            tag_lbl = ctk.CTkLabel(hdr, text=f" {p['tag']} ", font=(FONT_FAMILY, 12, "bold"),
                                   fg_color=tag_bg, text_color=tag_col, corner_radius=3)
            tag_lbl.pack(side=tk.LEFT)
            
            if p["title"]:
                title_col = TEXT_MAIN if not is_na else TEXT_DIM
                t_lbl = ctk.CTkLabel(hdr, text=f"  {p['title']}", font=(FONT_FAMILY, 13, "bold"),
                                     text_color=title_col)
                t_lbl.pack(side=tk.LEFT)
                
            w_col = ACCENT if not is_na and p["weight"].startswith("+") else TEXT_DIM
            w_lbl = ctk.CTkLabel(hdr, text=p["weight"], font=(FONT_FAMILY, 13, "bold"), text_color=w_col)
            w_lbl.pack(side=tk.RIGHT)
            
            # Task items list (when multiple items exist, e.g. Wing Task)
            if p["task_items"]:
                items_frame = ctk.CTkFrame(card, fg_color="#FFFFFF", corner_radius=2, border_width=1, border_color="#E8E8E8")
                items_frame.pack(fill=tk.X, padx=16, pady=(4, 8))
                for item in p["task_items"]:
                    i_lbl = ctk.CTkLabel(items_frame, text=f"•  {item}", font=(FONT_FAMILY, 12),
                                         text_color="#2A2A2A", anchor="w", justify=tk.LEFT, wraplength=450)
                    i_lbl.pack(fill=tk.X, padx=12, pady=3)
                    responsive_labels.append(i_lbl)
                    
            # Reason & Note section
            if p["reason"] or p["note"]:
                desc_frame = ctk.CTkFrame(card, fg_color="transparent")
                desc_frame.pack(fill=tk.X, padx=16, pady=(2, 12))
                if p["reason"]:
                    r_lbl = ctk.CTkLabel(desc_frame, text=f"사유: {p['reason']}", font=(FONT_FAMILY, 12),
                                         text_color=TEXT_DIM, anchor="w", justify=tk.LEFT, wraplength=450)
                    r_lbl.pack(fill=tk.X, pady=1)
                    responsive_labels.append(r_lbl)
                if p["note"]:
                    n_lbl = ctk.CTkLabel(desc_frame, text=f"비고: {p['note']}", font=(FONT_FAMILY, 12),
                                         text_color="#888888", anchor="w", justify=tk.LEFT, wraplength=450)
                    n_lbl.pack(fill=tk.X, pady=1)
                    responsive_labels.append(n_lbl)

        # Logic Card
        logic_card = ctk.CTkFrame(scroll, fg_color=CARD_BG, corner_radius=0, border_width=1, border_color=ACCENT)
        logic_card.pack(fill=tk.X, pady=(20, 10))
        ctk.CTkLabel(logic_card, text="산출 근거 로직", font=F_LABEL, text_color=ACCENT).pack(pady=(15, 8))
        logic_lbl = ctk.CTkLabel(logic_card, text=res["calculation_logic"], font=F_BODY, text_color=TEXT_MAIN, 
                     justify=tk.CENTER, anchor="center", wraplength=450)
        logic_lbl.pack(padx=20, pady=(0, 20), fill=tk.X, expand=True)
        responsive_labels.append(logic_lbl)

        # Responsive Layout Handler
        def _on_resize(event):
            # event.widget이 scroll 프레임인 경우에만 처리
            if event.widget == scroll:
                new_width = event.width
                # 1. 텍스트 줄바꿈 폭 자동 조절 (스크롤바, 카드 패딩 등을 넉넉히 고려하여 여유 확보)
                wrap_val = max(340, new_width - 80)
                for lbl in responsive_labels:
                    try:
                        if lbl.winfo_exists():
                            lbl.configure(wraplength=wrap_val)
                    except: pass
                
                # 2. 요약 카드 배치 전환 (가로 < 480px 일 때 세로 스택)
                if new_width < 480 and not summary_frame._is_vertical:
                    for i, card in enumerate(summary_cards):
                        card.grid_forget()
                        card.grid(row=i, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
                        card.configure(height=100) # Slightly shorter when stacked
                    summary_frame._is_vertical = True
                elif new_width >= 480 and summary_frame._is_vertical:
                    for i, card in enumerate(summary_cards):
                        card.grid_forget()
                        card.grid(row=0, column=i, columnspan=1, sticky="nsew", padx=5, pady=0)
                        card.configure(height=120)
                    summary_frame._is_vertical = False

        scroll.bind("<Configure>", _on_resize)

        # Switch to result view
        show_view(result_view)
        
        # Force initial resize calculation
        root.after(100, lambda: _on_resize(type('obj', (object,), {'widget': scroll, 'width': scroll.winfo_width()})()))

    def do_analyze_thread():
        pythoncom = None
        com_inited = False
        log_to_terminal("분석 스레드 시작...")
        try:
            try:
                import pythoncom as _pythoncom
                pythoncom = _pythoncom
                pythoncom.CoInitialize()
                com_inited = True
                log_to_terminal("COM 라이브러리 초기화 완료")
            except: 
                log_to_terminal("COM 라이브러리 초기화 실패 (무시됨)")
            
            log_to_terminal(f"Work Package 파일 분석 시작: {os.path.basename(wp_var.get())}")
            res = analyze(wp_var.get(), st_var.get(), update_progress)
            
            if "error" in res:
                log_to_terminal(f"오류 발생: {res['error']}")
            else:
                log_to_terminal("분석 완료. 결과를 표시합니다.")
                log_to_terminal(f"총 Man-Hour: {res['total']:.1f}, 예측 TAT: {res['tat_days']}일")
            
            root.after(0, lambda: show_results(res))
        except Exception as e:
            log_to_terminal(f"예상치 못한 오류 발생: {str(e)}")
            def err(err_e=e):
                status_label.configure(text="분석 실패")
                btn_run.configure(state="normal", text="분석 시작")
                root.config(cursor="") 
                messagebox.showerror("오류", f"예상치 못한 오류: {err_e}")
            root.after(0, err)
        finally:
            if com_inited and pythoncom:
                try: 
                    pythoncom.CoUninitialize()
                    log_to_terminal("COM 라이브러리 해제 완료")
                except: pass

    def start_analysis():
        if not wp_var.get() or not st_var.get():
            messagebox.showwarning("경고", "두 파일을 모두 선택하세요.")
            return
        
        log_to_terminal("--- 분석 프로세스 시작 ---")
        log_to_terminal(f"WP 파일: {wp_var.get()}")
        log_to_terminal(f"상태 파일: {st_var.get()}")
        
        btn_run.configure(state="disabled", text="처리 중...")
        progress_bar.set(0)
        root.config(cursor="wait") 
        threading.Thread(target=do_analyze_thread, daemon=True).start()

    create_modern_input(inner_padding, "Work Package 파일", wp_var, pick_wp, "WP 엑셀 파일을 선택하세요...")
    create_modern_input(inner_padding, "항공기 상태 파일", st_var, pick_st, "Aircraft Status 엑셀 파일을 선택하세요...")
    
    # Progress Section
    progress_container = ctk.CTkFrame(inner_padding, fg_color="transparent")
    progress_container.pack(fill=tk.X, pady=(10, 0))
    
    progress_bar = ctk.CTkProgressBar(progress_container, fg_color=BORDER_COLOR, progress_color=ACCENT, 
                                      height=6, corner_radius=0)
    progress_bar.set(0)
    progress_bar.pack(fill=tk.X, pady=(0, 10))
    
    status_label = ctk.CTkLabel(progress_container, text="시스템 준비됨", font=F_SMALL, text_color=TEXT_DIM)
    status_label.pack()

    # Action Section (Moved here to be below status_label)
    action_frame = ctk.CTkFrame(inner_padding, fg_color="transparent")
    action_frame.pack(fill=tk.X, pady=(15, 0))
    
    btn_run = ctk.CTkButton(action_frame, text="분석 시작", font=F_BTN,
                            fg_color=ACCENT, hover_color=HOVER,
                            text_color="#FFFFFF",
                            height=55, corner_radius=2, command=start_analysis)
    btn_run.pack(fill=tk.X)

    # Info Button
    ctk.CTkButton(main_container, text="ⓘ 가중치 산출 기준 가이드 확인", font=F_SMALL,
                  fg_color="transparent", text_color=TEXT_DIM, hover_color="#F2F2F2",
                  command=show_guide).pack(pady=(10, 0))
    
    # Initial view
    show_view(input_view)
    
    root.mainloop()

if __name__ == "__main__":
    run_ui()
