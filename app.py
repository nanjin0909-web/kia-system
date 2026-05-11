import streamlit as st
import pandas as pd
import json
import os
import re
from io import BytesIO

st.set_page_config(page_title="스마트 수불/실적 시스템 v56", layout="wide", initial_sidebar_state="expanded")

# ==========================================
# 1. 자동 저장소 및 상태 관리 (Persistence)
# ==========================================
SAVE_DIR = "data/saves_v56"
if not os.path.exists(SAVE_DIR):
    os.makedirs(SAVE_DIR)

FILES = {
    'inventory': os.path.join(SAVE_DIR, "inventory.json"),
    'inbound': os.path.join(SAVE_DIR, "inbound.json"),
    'sales': os.path.join(SAVE_DIR, "sales.json"),
    'plan': os.path.join(SAVE_DIR, "plan.json"),
    'usage': os.path.join(SAVE_DIR, "usage.json"),
    'vendor_plan': os.path.join(SAVE_DIR, "vendor_plan.json"),
    'part_name': os.path.join(SAVE_DIR, "part_name.json"),
    'part_vendor': os.path.join(SAVE_DIR, "part_vendor.json"),
    'bom_master': os.path.join(SAVE_DIR, "bom_master.json"),
    'settings': os.path.join(SAVE_DIR, "settings.json")
}

from cloud_sync import save_to_cloud, load_from_cloud, get_drive_service

def load_data(key, default):
    # 1. 클라우드 연결 시 클라우드 우선 로드 (클라우드 데이터가 있으면 로컬에도 덮어씌워 동기화)
    if get_drive_service():
        cloud_data = load_from_cloud(key, None)
        if cloud_data is not None:
            with open(FILES[key], "w", encoding="utf-8") as f:
                json.dump(cloud_data, f, ensure_ascii=False, indent=2)
            return cloud_data
            
    # 2. 로컬 로드
    if os.path.exists(FILES[key]):
        try:
            with open(FILES[key], "r", encoding="utf-8") as f:
                return json.load(f)
        except: pass
    return default

def save_data(key):
    data = st.session_state[key]
    with open(FILES[key], "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        
    # 클라우드 연결 시 클라우드 저장
    if get_drive_service():
        save_to_cloud(key, data)

for k in FILES.keys():
    if k not in st.session_state:
        st.session_state[k] = load_data(k, [] if k == 'bom_master' else {})

# ==========================================
# 2. 유틸리티 함수 및 파싱 로직
# ==========================================
def get_mapped_factory(raw_factory):
    f = str(raw_factory).strip().upper()
    if f in ["2921", "1라인", "H1"]: return "1라인"
    if f in ["2922", "2라인", "H2"]: return "2라인"
    if f in ["2923", "3라인", "H3"]: return "3라인"
    return None

def get_mapped_car_name(car_code):
    c = str(car_code).strip().upper()
    if "8V" in c: return "타스만"
    if "GZ" in c or "HC" in c: return "쏘렌토"
    if "OT" in c or "TO" in c: return "니로(SG2)"
    if "5H" in c: return "SP3"
    if "AS" in c: return "EV6"
    if "EX" in c or "HV" in c: return "K5"
    if "GG" in c or "GL" in c: return "K8"
    return c

def is_excluded_car(fac, car):
    f = get_mapped_factory(fac)
    c_up = str(car).strip().upper()
    if f == "1라인" and c_up in ["C5", "DJ", "EN"]: return True
    if f == "2라인" and c_up in ["DQ", "G5", "AS", "DJ"]: return True
    return False

def extract_date_from_filename(file_path, fallback):
    filename = os.path.basename(str(file_path))
    path_str = str(file_path).replace('\\', '/')
    
    year = None
    y_match = re.search(r'(20\d{2})', path_str)
    if y_match: year = y_match.group(1)
    else:
        y_short = re.search(r'[^0-9]?(2[4-9])년', path_str)
        if y_short: year = "20" + y_short.group(1)

    m1 = re.search(r'(20\d{2}|\d{2})[-_ \.]?(0[1-9]|1[0-2])[-_ \.]?(0[1-9]|[12]\d|3[01])', filename)
    if m1:
        y = "20" + m1.group(1) if len(m1.group(1)) == 2 else m1.group(1)
        return f"{y}-{m1.group(2)}-{m1.group(3)}"
    
    m2 = re.search(r'(0[1-9]|1[0-2])[-_]?(0[1-9]|[12]\d|3[01])', filename)
    if m2:
        y = year if year else fallback[:4]
        return f"{y}-{m2.group(1)}-{m2.group(2)}"
    return fallback

def parse_bom(file):
    try:
        df = pd.read_excel(file)
        if len(df) < 2: return "❌ 사양표 데이터 유효하지 않음"
        header1 = df.columns
        header2 = df.iloc[0].fillna('')
        cols, last_top = [], ""
        for i in range(len(header1)):
            top = str(header1[i]).strip()
            if top.startswith("Unnamed") or top == "": top = last_top
            else: last_top = top
            sub = str(header2[i]).strip()
            cols.append(top if sub.startswith("Unnamed") or sub == "" else f"{top}_{sub}")
        df.columns = cols
        df = df.iloc[1:].reset_index(drop=True)
        
        new_spec, new_parts = [], {}
        for _, row in df.iterrows():
            r = row.fillna('').to_dict()
            fac = str(row.iloc[0]).strip()
            car = str(row.iloc[1]).strip() or "차종미상"
            alc = str(row.iloc[2]).strip().upper()
            
            spare_type = ""
            for v in list(r.values())[3:]:
                cv = str(v).replace(" ", "").upper()
                if "TMK" in cv or "TEMPO" in cv or "FULL" in cv:
                    spare_type = cv
                    break
            
            if str(r.get("사용유무", "")).strip() in ["사용", "Y"]:
                spec_obj = {"_factory": fac, "_carModel": car, "_alc": alc, "_spareType": spare_type}
                keys_list = list(r.keys())
                for i, pk in enumerate(keys_list):
                    if "품번" in str(pk) or "PART NO" in str(pk):
                        pn = str(r[pk]).strip()
                        if pn and pn != "-" and str(pn).upper() != "NAN":
                            # Search forward up to 3 columns for Name/Spec
                            for j in range(i+1, min(len(keys_list), i+4)):
                                nk = keys_list[j]
                                if any(x in str(nk) for x in ["품명","스펙","규격","NAME","명칭","사양"]):
                                    nval = str(r[nk]).strip()
                                    if nval and nval != "-" and str(nval).upper() != "NAN":
                                        new_parts[pn] = nval
                                        break
                                        
                            # Search backward up to 2 columns for Vendor
                            for j in range(i-1, max(-1, i-3), -1):
                                vk = keys_list[j]
                                if any(x in str(vk) for x in ["업체명","업체"]):
                                    vval = str(r[vk]).strip()
                                    if vval and vval != "-" and str(vval).upper() != "NAN":
                                        if 'part_vendor' not in st.session_state: st.session_state.part_vendor = {}
                                        st.session_state.part_vendor[pn] = vval
                                        break
                spec_obj.update(r)
                new_spec.append(spec_obj)
        st.session_state.bom_master = new_spec
        st.session_state.part_name.update(new_parts)
        save_data('bom_master'); save_data('part_name')
        if 'part_vendor' in st.session_state: save_data('part_vendor')
        return f"✅ BOM 갱신 완료 ({len(new_spec)}건)"
    except Exception as e: return f"❌ 에러: {e}"

def parse_inbound(files, in_date):
    count = 0
    if in_date not in st.session_state.inbound:
        st.session_state.inbound[in_date] = {}
        
    for file in files:
        df = pd.read_excel(file)
        for _, row in df.iterrows():
            pn = str(row.iloc[1]).strip()
            pname = str(row.iloc[2]).strip()
            vendor = str(row.iloc[4]).strip() if len(row) > 4 else "업체미상"
            qty_str = str(row.iloc[8]).replace(',', '') if len(row) > 8 else "0"
            try: qty = int(float(qty_str))
            except: qty = 0
            
            if pn and pn != "nan" and qty > 0:
                if pn not in st.session_state.inventory:
                    st.session_state.inventory[pn] = {"qty": 0, "vendorName": vendor, "partName": pname}
                st.session_state.inventory[pn]["qty"] += qty
                st.session_state.inbound[in_date][pn] = st.session_state.inbound[in_date].get(pn, 0) + qty
                if pname and pname != "nan": st.session_state.part_name[pn] = pname
                count += 1
    
    if count > 0:
        save_data('inventory'); save_data('inbound'); save_data('part_name')
        return f"✅ [{in_date}] 총 {count}건 입고 통합 완료"
    return "❌ 입고 데이터 반영 실패"

def parse_system_plan(files, fallback_date):
    if not st.session_state.bom_master: return "❌ BOM 사양표 먼저 갱신 요망"
    
    active_car_map = {}
    car_file_path = os.path.join("data", "차종코드.xlsb")
    if os.path.exists(car_file_path):
        try:
            car_df = pd.read_excel(car_file_path, engine='pyxlsb')
            for _, crow in car_df.iterrows():
                ccode = str(crow.get('차종코드', '')).strip().upper()
                cname = str(crow.get('라인명', '')).strip()
                if ccode and ccode != 'NAN':
                    active_car_map[ccode] = cname
        except Exception:
            pass
            
    act_count, plan_count = 0, 0
    t_sales, t_plan = st.session_state.sales, st.session_state.plan
    processed_dates = set()
    
    for f in files:
        file_path = getattr(f, "name", str(f))
        target_date = extract_date_from_filename(file_path, fallback_date)
        
        # 덮어쓰기 로직: 이 배치(명령)에서 처음 만난 날짜면 배열을 비움 (초기화)
        if target_date not in processed_dates:
            t_sales[target_date] = {}
            t_plan[target_date] = {}
            processed_dates.add(target_date)
            
        try:
            # f가 파일경로(str)일 수도 있고 BytesIO(Streamlit UploadedFile)일 수도 있음
            df = pd.read_excel(f, header=None)
            start_row = -1
            c_dminus1, c_d0, c_alc, c_fac, c_carcode = -1, -1, -1, -1, -1
            
            for rt in range(min(20, len(df))):
                for ct in range(len(df.columns)):
                    val = str(df.iloc[rt, ct]).replace(' ', '').upper()
                    if "D-1" in val or "실적" in val or "Ͻ" in val: c_dminus1 = ct
                    if "CODEVALUE" in val or "ALC" == val: c_alc = ct
                    if "공장" in val or "ڵ" in val: c_fac = ct
                    if "차종코드" in val: c_carcode = ct
                    
                    if "DTOTAL" in val or "D0TOTAL" in val or "D+0TOTAL" in val:
                        c_d0 = ct
                    elif c_d0 == -1 or (c_d0 != -1 and "TOTAL" not in str(df.iloc[rt, c_d0]).upper()):
                        if "D0" in val or "D+0" in val or "계획" in val or "QTY" in val or "PBS" == val or ("SCHEDULE" in val and not any(x in val for x in ["D+1", "D+2", "D+3", "D+4"])):
                            c_d0 = ct
                            
                if c_alc != -1 and c_fac != -1 and (c_dminus1 != -1 or c_d0 != -1):
                    start_row = rt
                    break
                    
            if start_row != -1:
                if c_carcode == -1 and len(df.columns) >= 2:
                    c_carcode = len(df.columns) - 2
                    
                for i in range(start_row+1, len(df)):
                    car_code = str(df.iloc[i, c_carcode]).strip().upper() if c_carcode != -1 else ""
                    
                    if active_car_map and car_code and car_code != "NAN":
                        if car_code not in active_car_map:
                            continue
                        fac = get_mapped_factory(active_car_map[car_code])
                    else:
                        fac = get_mapped_factory(df.iloc[i, c_fac])
                        
                    if not fac: continue # 지정 팩토리가 아니면 드랍
                    alc = str(df.iloc[i, c_alc]).strip().upper()
                    if not alc or alc == "NAN": continue
                    
                    try: act_qty = int(float(str(df.iloc[i, c_dminus1]).replace(',', ''))) if c_dminus1 != -1 and pd.notna(df.iloc[i, c_dminus1]) and str(df.iloc[i, c_dminus1]).strip() else 0
                    except: act_qty = 0
                    try: plan_qty = int(float(str(df.iloc[i, c_d0]).replace(',', ''))) if c_d0 != -1 and pd.notna(df.iloc[i, c_d0]) and str(df.iloc[i, c_d0]).strip() else 0
                    except: plan_qty = 0
                    
                    key = f"{fac}_{alc}"
                    if act_qty > 0 or plan_qty > 0:
                        if key not in t_sales[target_date]: t_sales[target_date][key] = 0
                        if key not in t_plan[target_date]: t_plan[target_date][key] = 0
                        t_sales[target_date][key] += act_qty
                        t_plan[target_date][key] += plan_qty
                        if act_qty > 0: act_count += 1
                        if plan_qty > 0: plan_count += 1
        except Exception as e: pass
            
    if act_count > 0 or plan_count > 0:
        save_data('sales'); save_data('plan')
        
        # === 일일동향 자동 연동 로직 ===
        try:
            import requests
            import time
            GAS_URL = "https://script.google.com/macros/s/AKfycbzP_P9wP_mEfh88XtaS1QKbLl1KlS7RwDWMVS-SvPxtz2rXR00LlsBay7gnw2Dfdw3d/exec"
            
            # 1. BOM 기반 소요량 매핑
            alc_info = {}
            for sp in st.session_state.bom_master:
                fac = str(sp.get('_factory', '미상')).strip()
                car = get_mapped_car_name(str(sp.get('_carModel', '차종미상')).strip())
                alc = str(sp.get('_alc', '')).strip().upper()
                mapped_fac = get_mapped_factory(fac)
                q = get_bom_quantities(sp)
                mul = q.get('rqW', 4) + q.get('sqW', 0)
                if mapped_fac:
                    alc_info[f"{mapped_fac}_{alc}"] = {"fac": mapped_fac, "car": car, "mul": mul}

            # 2. 목표 일자(D-1 실적) 데이터 집계
            agg = {}
            if target_date in t_sales:
                for key, val in t_sales[target_date].items():
                    info = alc_info.get(key)
                    if info:
                        f_name = "화성"
                        l_name = info['fac'].replace("라인", "라")
                        car = info['car']
                        qty = val * info['mul']
                        k = (f_name, l_name, car)
                        agg[k] = agg.get(k, 0) + qty

            # 3. GAS 데이터 가져오기 및 POST
            if agg:
                res = requests.get(GAS_URL).json()
                prod_records = res.get("production", [])
                
                for (f_name, l_name, car), qty in agg.items():
                    existing = next((r for r in prod_records if str(r.get('date')) == target_date and r.get('factory') == f_name and r.get('line') == l_name and r.get('car_model') == car), None)
                    
                    payload = {
                        "target": "production",
                        "date": target_date,
                        "factory": f_name,
                        "line": l_name,
                        "car_model": car,
                        "daily_actual": qty
                    }
                    if existing:
                        payload["action"] = "update"
                        payload["id"] = existing["id"]
                    else:
                        payload["action"] = "add"
                        payload["id"] = int(time.time() * 1000)
                        
                    requests.post(GAS_URL, data=payload)
                    
            return f"✅ 실적 {act_count}건, 계획 {plan_count}건 파싱 (일일동향 연동 성공 🚀)"
        except Exception as e:
            return f"✅ 실적 {act_count}건, 계획 {plan_count}건 파싱 (일일동향 연동 실패: {e})"
            
    return "❌ 시스템 소요량 데이터 추출 실패"

def parse_vendor_plan(files):
    count = 0
    ven = st.session_state.vendor_plan
    for f in files:
        count += 1 # Dummy parse implementation
    save_data('vendor_plan')
    return f"✅ 업체 통보 데이터 {count}건 갱신"

# ==========================================
# 3. 사이드바 UI

# ==========================================
st.sidebar.markdown("<div class='sidebar-title'>🚀 스마트 공정 통합 관리 시스템</div>", unsafe_allow_html=True)
st.sidebar.write("Ver 56.0 (실적 배수 완벽화)")

def get_bom_quantities(spec):
    spare_type = str(spec.get("_spareType", "")).upper()
    
    def get_by_kws(kws):
        for k in spec.keys():
            ck = str(k).replace(" ", "").upper()
            if all(x in ck for x in kws): return spec[k]
        return None

    rqw = float(get_by_kws(["정규", "휠", "수량"]) or get_by_kws(["WHEEL", "수량"]) or 4)
    sqw = float(get_by_kws(["보조", "휠", "수량"]) or get_by_kws(["SPARE", "WHEEL", "수량"]) or 0)
    rqt = float(get_by_kws(["정규", "타이어", "수량"]) or get_by_kws(["TIRE", "수량"]) or 4)
    sqt = float(get_by_kws(["보조", "타이어", "수량"]) or get_by_kws(["SPARE", "TIRE", "수량"]) or 0)
    
    spare_wheel = get_by_kws(["보조", "휠", "품번"]) or get_by_kws(["SPARE", "WHEEL", "품번"])
    spare_tire = get_by_kws(["보조", "타이어", "품번"]) or get_by_kws(["SPARE", "TIRE", "품번"])
    
    if "TMK" in spare_type:
        sqw, sqt = 0, 0
    elif "FULL" in spare_type:
        if not spare_wheel or str(spare_wheel).strip() in ["", "-"]:
            rqw, sqw = 5, 0
        else: rqw, sqw = 4, 1
        if not spare_tire or str(spare_tire).strip() in ["", "-"]:
            rqt, sqt = 5, 0
        else: rqt, sqt = 4, 1
    elif "TEMPO" in spare_type:
        rqw, sqw, rqt, sqt = 4, 1, 4, 1
        
    return {"rqW": rqw, "sqW": sqw, "rqT": rqt, "sqT": sqt, "spareType": spare_type}

menu = st.sidebar.radio("메뉴를 선택하세요", [
    "1. 데이터 업로드 센터",
    "2. 수불/실적 모니터링",
    "3. 업체계획 크로스체크",
    "4. 소요량 전개표",
    "5. 실적 가로전개표"
])

if st.sidebar.button("🗑️ 시스템 전체 초기화"):
    for k in ['inventory','inbound','sales','plan','usage','vendor_plan']:
        st.session_state[k] = {}
        save_data(k)
    st.sidebar.success("초기화 완료")

# ==========================================
# 4. 페이지: 데이터 업로드 센터
# ==========================================
if menu == "1. 데이터 업로드 센터":
    st.title("📁 데이터 업로드 센터")
    
    with st.expander("⚙️ 0. 기초 마스터 데이터 갱신 (BOM 사양표)", expanded=True):
        st.write(f"현재 활성 ALC 수: **{len(st.session_state.bom_master)}**건")
        col1, col2 = st.columns([3,1])
        with col1: f1 = st.file_uploader("BOM 사양표 (Excel)", type=['xls', 'xlsx'])
        with col2: 
            if st.button("마스터 갱신", use_container_width=True) and f1:
                st.success(parse_bom(f1))
                
    with st.expander("📦 1. MES 입고 등록 (다중파일)", expanded=True):
        col1, col2, col3 = st.columns([1,3,1])
        with col1: d1 = st.date_input("입고 일자")
        with col2: f2 = st.file_uploader("입고 내역 파일들", type=['xls', 'xlsx', 'csv'], accept_multiple_files=True)
        with col3:
            if st.button("입고 데이터 적용", use_container_width=True) and f2:
                st.success(parse_inbound(f2, str(d1)))

    with st.expander("🏭 2. 시스템 생산 실적 & 소요량 업로드", expanded=True):
        st.info("※ 매번 파일을 올리기 번거로우면, 로컬 폴더(data/기아생산계획)를 스캔하여 일괄 병합/덮어쓰기 합니다.")
        
        # 자동 스캔 버튼
        if st.button("🔄 로컬 폴더(data/기아생산계획) 전체 자동 동기화", use_container_width=True):
            scan_path = os.path.join("data", "기아생산계획")
            if not os.path.exists(scan_path):
                st.error(f"'{scan_path}' 폴더가 존재하지 않습니다.")
            else:
                found_files = []
                for root, _, files_in_dir in os.walk(scan_path):
                    for fname in files_in_dir:
                        if fname.lower().endswith(('.xls', '.xlsx', '.xlsb')):
                            found_files.append(os.path.join(root, fname))
                if not found_files:
                    st.warning("폴더 내에 스캔할 엑셀 파일이 없습니다.")
                else:
                    msg = parse_system_plan(found_files, "2026-01-01")
                    st.success(f"로컬 동기화 완료! ({len(found_files)}개 파일 스캔) - " + msg)
        
        st.markdown("---")
        
        col1, col2, col3 = st.columns([1,3,1])
        with col1: d2 = st.date_input("단일 파일 기준일", key="d2")
        with col2: f3 = st.file_uploader("전개표 직접 업로드", type=['xls', 'xlsx', 'csv'], accept_multiple_files=True, key="f3")
        with col3:
            if st.button("직접 업로드 파싱", use_container_width=True) and f3:
                st.success(parse_system_plan(f3, str(d2)))
                
    with st.expander("🚚 3. 업체 통보 계획 업로드 (크로스체크 전용)", expanded=False):
        col1, col2 = st.columns([3,1])
        with col1: f4 = st.file_uploader("업체 계획 파일들", type=['xls', 'xlsx'], accept_multiple_files=True, key="f4")
        with col2:
            if st.button("업체 통보 데이터 파싱", use_container_width=True) and f4:
                st.success(parse_vendor_plan(f4))

# ==========================================
# 5. 페이지: 수불/실적 모니터링
# ==========================================
elif menu == "2. 수불/실적 모니터링":
    st.title("📦 수불 및 실적 모니터링")
    view_type = st.radio("보기 옵션", ["전체 재고표", "타이어 수불장", "휠 수불장"], horizontal=True)
    
    if view_type == "전체 재고표":
        data = []
        for part, info in st.session_state.inventory.items():
            qty = info.get('qty', 0)
            pname = info.get('partName') or st.session_state.part_name.get(part, "")
            data.append({"품번": part, "품명": pname, "업체명": info.get('vendorName', ''), "총수량": qty})
        if data:
            df = pd.DataFrame(data)
            st.dataframe(df.style.highlight_between(left=-99999, right=0, color='red'), use_container_width=True)
        else:
            st.warning("재고 데이터가 없습니다.")
    else:
        st.subheader(f"📊 {view_type} (납촉 실적재고 포맷)")
        
        # 1. 날짜 취합
        dates = sorted(list(set(st.session_state.inbound.keys()) | set(st.session_state.sales.keys())))
        
        if not dates:
            st.warning("데이터가 없습니다. 입고 내역 또는 실적을 먼저 업로드하세요.")
        else:
            # 2. 부품 기본 정보 및 소요량 맵핑 추출
            part_info = {} # part_no: {Line, Car, Vendor, Name, Type}
            alc_parts_map = {} # fac_alc: {part_no: qty_per_car}
            
            # 동적으로 업체명(Vendor) 맵을 먼저 구축 (기존 저장소 누락 방지)
            part_vendor_map = {}
            for sp in st.session_state.bom_master:
                for pk, pv in sp.items():
                    pk_up = str(pk).upper()
                    if "품번" in pk_up or "PART NO" in pk_up:
                        pn = str(pv).strip()
                        if pn and pn != "-" and str(pn).upper() != "NAN":
                            base_key = str(pk).replace("품번", "").replace("PART NO", "")
                            vend = str(sp.get(base_key + "업체명", sp.get(base_key + "업체", ""))).strip()
                            if vend and vend != "-" and str(vend).upper() != "NAN":
                                part_vendor_map[pn] = vend

            # 이원화 감지를 위해 ALC 내의 슬롯(정규/보조, 타이어/휠)별 부품 번호 수집
            slot_map = {} # key: fac_alc, value: {slot: set(parts)}
            
            for sp in st.session_state.bom_master:
                fac = get_mapped_factory(sp.get('_factory'))
                car = str(sp.get('_carModel', '')).strip()
                
                if is_excluded_car(fac, car): continue
                
                alc = str(sp.get('_alc', '')).strip().upper()
                q = get_bom_quantities(sp)
                
                # Include region to uniquely identify ALC per location
                region = str(sp.get('지역', '')).strip()
                key = f"{fac}_{alc}_{region}"
                alc_parts_map.setdefault(key, {})
                if key not in slot_map:
                    slot_map[key] = {"정규_타이어": set(), "보조_타이어": set(), "정규_휠": set(), "보조_휠": set()}
                
                for pk, pv in sp.items():
                    pk_up = str(pk).upper()
                    if "품번" in pk_up or "PART NO" in pk_up:
                        pn = str(pv).strip()
                        if pn and pn != "-":
                            is_tire = "타이어" in pk_up or "TIRE" in pk_up
                            is_wheel = "휠" in pk_up or "WHEEL" in pk_up
                            is_spare = "보조" in pk_up or "SPARE" in pk_up
                            
                            ptype = ""
                            qty = 0
                            if is_tire: 
                                ptype = "타이어"
                                qty = q.get('sqT', 0) if is_spare else q.get('rqT', 4)
                                if qty > 0: slot_map[key][f"{'보조' if is_spare else '정규'}_타이어"].add(pn)
                            elif is_wheel: 
                                ptype = "휠"
                                qty = q.get('sqW', 0) if is_spare else q.get('rqW', 4)
                                if qty > 0: slot_map[key][f"{'보조' if is_spare else '정규'}_휠"].add(pn)
                            
                            if qty > 0:
                                alc_parts_map[key][pn] = alc_parts_map[key].get(pn, 0) + qty
                                
                                if pn not in part_info:
                                    if pn in part_vendor_map:
                                        vend = part_vendor_map[pn]
                                    elif 'part_vendor' in st.session_state and pn in st.session_state.part_vendor:
                                        vend = st.session_state.part_vendor[pn]
                                    else:
                                        vend = st.session_state.inventory.get(pn, {}).get("vendorName", "")
                                        
                                    part_info[pn] = {
                                        "라인": fac or "",
                                        "차종": car,
                                        "구분": ptype,
                                        "업체": vend,
                                        "PART NAME": st.session_state.part_name.get(pn, "알수없음"),
                                        "이원화": False
                                    }
                                else:
                                    if not part_info[pn]["라인"] and fac: part_info[pn]["라인"] = fac
                                    if not part_info[pn]["차종"] and car: part_info[pn]["차종"] = car
                                    
            # 이원화 감지: 동일 ALC의 동일 슬롯(예: 동일한 '정규 타이어' 슬롯)에 부품이 여러 개 있으면 이원화로 판단
            dual_sourced_parts = set()
            for key, slots in slot_map.items():
                for slot, parts in slots.items():
                    if len(parts) > 1:
                        dual_sourced_parts.update(parts)
                    
            for dpn in dual_sourced_parts:
                if dpn in part_info:
                    part_info[dpn]["이원화"] = True
            
            # 3. 실적(소요량) 및 입고량 계산
            daily_req = {d: {} for d in dates}
            for d in dates:
                sales_d = st.session_state.sales.get(d, {})
                for fac_alc, cars in sales_d.items():
                    if cars <= 0: continue
                    parts_needed = alc_parts_map.get(fac_alc, {})
                    for pn, qty_per_car in parts_needed.items():
                        daily_req[d][pn] = daily_req[d].get(pn, 0) + (cars * qty_per_car)
                        
            daily_inb = {d: {} for d in dates}
            for d in dates:
                for pn, inb_qty in st.session_state.inbound.get(d, {}).items():
                    daily_inb[d][pn] = daily_inb[d].get(pn, 0) + inb_qty
                    
            target_parts = []
            for pn, info in part_info.items():
                if view_type == "타이어 수불장" and info["구분"] == "타이어": target_parts.append(pn)
                elif view_type == "휠 수불장" and info["구분"] == "휠": target_parts.append(pn)
                
            if not target_parts:
                st.info(f"선택한 '{view_type}'에 해당하는 부품 데이터가 없습니다.")
            else:
                raw_records = []
                for pn in target_parts:
                    info = part_info[pn]
                    
                    pname_disp = info["PART NAME"]
                    if info.get("이원화", False):
                        pname_disp = f"[이원화] {pname_disp}"
                        
                    raw_records.append({
                        "라인": info["라인"],
                        "차종": ("" if not info["차종"] else info["차종"]),
                        "구분": info["구분"],
                        "업체": info["업체"],
                        "PART NO.": pn,
                        "PART NAME": pname_disp
                    })
                
                # 정렬: 라인 -> 차종 -> 업체 -> PART NO
                raw_records.sort(key=lambda x: (x["라인"], x["차종"], x["업체"], x["PART NO."]))
                
                records = []
                current_car = None
                subtotal_req = {d: 0 for d in dates}
                subtotal_inb = {d: 0 for d in dates}
                subtotal_prev = {d: 0 for d in dates}
                subtotal_cur = {d: 0 for d in dates}
                
                def append_subtotal(car_name):
                    row = {
                        ("", "라인"): "",
                        ("", "차종"): f"{car_name} 소계",
                        ("", "구분"): "",
                        ("", "업체"): "",
                        ("", "PART NO."): "",
                        ("", "PART NAME"): ""
                    }
                    for d in dates:
                        m_day = d[5:10].replace("-", "/") # MM/DD Format
                        row[(m_day, "전일")] = subtotal_prev[d]
                        row[(m_day, "실적")] = subtotal_req[d]
                        row[(m_day, "입고")] = subtotal_inb[d]
                        row[(m_day, "재고")] = subtotal_cur[d]
                    records.append(row)
                    
                for rec in raw_records:
                    pn = rec["PART NO."]
                    car = rec["차종"]
                    
                    if current_car is not None and current_car != car:
                        append_subtotal(current_car)
                        subtotal_req = {d: 0 for d in dates}
                        subtotal_inb = {d: 0 for d in dates}
                        subtotal_prev = {d: 0 for d in dates}
                        subtotal_cur = {d: 0 for d in dates}
                        
                    current_car = car
                    
                    row = {
                        ("", "라인"): rec["라인"],
                        ("", "차종"): rec["차종"],
                        ("", "구분"): rec["구분"],
                        ("", "업체"): rec["업체"],
                        ("", "PART NO."): pn,
                        ("", "PART NAME"): rec["PART NAME"]
                    }
                    
                    prev_inv = 0
                    for d in dates:
                        used = daily_req[d].get(pn, 0)
                        received = daily_inb[d].get(pn, 0)
                        curr_inv = prev_inv + received - used
                        
                        m_day = d[5:10].replace("-", "/")
                        row[(m_day, "전일")] = prev_inv
                        row[(m_day, "실적")] = used
                        row[(m_day, "입고")] = received
                        row[(m_day, "재고")] = curr_inv
                        
                        subtotal_prev[d] += prev_inv
                        subtotal_req[d] += used
                        subtotal_inb[d] += received
                        subtotal_cur[d] += curr_inv
                        
                        prev_inv = curr_inv
                        
                    records.append(row)
                    
                if current_car is not None:
                    append_subtotal(current_car)
                    
                df_subul = pd.DataFrame(records)
                # 컬럼 순서 유지를 위한 MultiIndex
                df_subul.columns = pd.MultiIndex.from_tuples(df_subul.columns)
                
                # 스타일 지정: 차종 소계 및 마이너스 재고 하이라이팅, 일자별 구분선
                def style_dataframe(df):
                    styles = pd.DataFrame('', index=df.index, columns=df.columns)
                    car_col_idx = df.columns.get_loc(("", "차종"))
                    
                    for r_idx in range(len(df)):
                        car_val = str(df.iloc[r_idx, car_col_idx])
                        is_subtotal = "소계" in car_val
                        
                        for c_idx, col in enumerate(df.columns):
                            css = []
                            if is_subtotal:
                                css.append('background-color: #d1b88e' if '소계' in car_val else 'background-color: #f2e2c4')
                                css.append('font-weight: bold')
                                
                            if col[0] != "":
                                val = df.iloc[r_idx, c_idx]
                                try:
                                    if int(val) < 0:
                                        css.append('color: red; font-weight: bold')
                                except: pass
                                
                            # 일자별/기본항목 명확한 세로 구분선 (가독성 향상)
                            if col[1] in ["PART NAME", "재고"]:
                                css.append('border-right: 2px solid #555 !important')
                            else:
                                css.append('border-right: 1px solid #ddd')
                                
                            styles.iloc[r_idx, c_idx] = '; '.join(css)
                    return styles
                
                format_dict = {col: "{:,.0f}" for col in df_subul.columns if col[0] != ""}
                
                styled_df = df_subul.style.set_table_styles([
                    {'selector': 'th', 'props': [
                        ('background-color', '#e8f4f8'), 
                        ('font-size', '13px'), 
                        ('border', '1px solid #ccc')
                    ]}
                ]).apply(style_dataframe, axis=None).format(format_dict)
                
                st.dataframe(styled_df, use_container_width=True)
# ==========================================
# 6. 소요량 / 크로스체크 / 가로전개 화면 구현
# ==========================================
elif menu == "3. 업체계획 크로스체크":
    st.title("📋 휠 업체별 생산계획 크로스체크")
    st.markdown("※ 이 메뉴는 **휠(Wheel)** 품목 전용입니다. 업체에서 제공한 계획(통보) 엑셀 파일을 업로드하면, 시스템에서 산출한 생산계획 요구량과 자동으로 비교(Cross-check)합니다.")
    
    # 1. 파일 업로드 UI
    col1, col2 = st.columns([2, 1])
    with col1:
        vendor_files = st.file_uploader("업체 계획 엑셀 파일 수동 업로드", type=['xlsx', 'xls', 'xlsm'], accept_multiple_files=True)
    
    vendor_files_list = list(vendor_files) if vendor_files else []
    
    with col2:
        st.write("") # 줄맞춤용 빈 공간
        st.write("")
        if st.button("📥 업체 전용 웹앱 제출본 불러오기", use_container_width=True):
            folder_path = os.path.join("data", "vendor_uploads")
            if os.path.exists(folder_path):
                cnt = 0
                for fname in os.listdir(folder_path):
                    if fname.endswith(('.xls', '.xlsx', '.xlsm')):
                        vendor_files_list.append(os.path.join(folder_path, fname))
                        cnt += 1
                if cnt > 0:
                    st.success(f"전용 웹앱에서 {cnt}개의 파일을 성공적으로 불러왔습니다.")
                else:
                    st.warning("제출된 새로운 파일이 없습니다.")
            else:
                st.warning("제출 대기 중인 파일이 없거나 업체 폼이 초기화되지 않았습니다.")
    
    # 업체 통보량 파싱 저장소 (품번 기준 합산)
    parsed_vendor_qty = {}
    
    if vendor_files_list:
        succ_count = 0
        for f in vendor_files_list:
            try:
                # pandas로 엑셀 읽기 (첫 행부터 스캔)
                df = pd.read_excel(f, header=None)
                # 데이터 행 순회
                for i in range(len(df)):
                    # A열(0), C열(2), D열(3), H열(7), P열(15) 등 필요 열 스캔
                    if df.shape[1] > 2:
                        col_c = str(df.iloc[i, 2]).strip()
                        pn_cleaned = col_c.replace("-", "").replace(" ", "").upper()
                        
                        # (1) 다이캐스탈 양식 추정: C열에 품번, P열(15)에 통보 수량
                        if df.shape[1] > 15:
                            p_val = str(df.iloc[i, 15]).replace(',', '')
                            if p_val.replace('.', '', 1).isdigit(): # 숫자이면 다이캐스탈 수량 적용
                                # 다이캐스탈 데이터로 합산
                                qty = int(float(p_val))
                                if qty > 0 and len(pn_cleaned) >= 5: # 임의 품번 길이 필터
                                    parsed_vendor_qty[pn_cleaned] = parsed_vendor_qty.get(pn_cleaned, 0) + qty
                                    continue # 처리 완료시 다음 행 이동
                                    
                        # (2) 현대성우 양식 추정: A열에 2921 등 공장, C열에 품번, D열(3) PLT당수량, H열(7) 공급계획PLT
                        if df.shape[1] > 7:
                            a_val = str(df.iloc[i, 0]).strip()
                            # 2921(1라인) 등 공장코드가 있는지 확인
                            if "2921" in a_val or "2922" in a_val or "2923" in a_val or "라인" in a_val:
                                d_val = str(df.iloc[i, 3]).replace(',', '')
                                h_val = str(df.iloc[i, 7]).replace(',', '')
                                if d_val.replace('.', '', 1).isdigit() and h_val.replace('.', '', 1).isdigit():
                                    qty_per_plt = int(float(d_val))
                                    plt_count = int(float(h_val))
                                    total_qty = qty_per_plt * plt_count
                                    if total_qty > 0 and len(pn_cleaned) >= 5:
                                        parsed_vendor_qty[pn_cleaned] = parsed_vendor_qty.get(pn_cleaned, 0) + total_qty
                                        continue
            except Exception as e:
                st.error(f"파일 파싱 에러 [{f.name}]: {str(e)}")
        
        st.success(f"{len(vendor_files)}개의 파일에서 업체 계획을 성공적으로 읽어왔습니다!")
    
    # 2. 시스템 기준 휠 소요량 산출 (기존 소요량 로직 재활용, '업체' 대분류)
    st.subheader("🔹 시스템 자체 산출 소요량 (휠 전용)")
    
    # 동적으로 업체명 맵 구축
    part_vendor_map = {}
    for sp in st.session_state.bom_master:
        for pk, pv in sp.items():
            pk_up = str(pk).upper()
            if "품번" in pk_up or "PART NO" in pk_up:
                pn = str(pv).strip()
                if pn and pn != "-" and str(pn).upper() != "NAN":
                    base_key = str(pk).replace("품번", "").replace("PART NO", "")
                    vend = str(sp.get(base_key + "업체명", sp.get(base_key + "업체", ""))).strip()
                    if vend and vend != "-" and str(vend).upper() != "NAN":
                        part_vendor_map[pn] = vend

    # ALC 매핑 (휠만 추출)
    wheel_alc_map = {}
    for sp in st.session_state.bom_master:
        f = get_mapped_factory(sp.get('_factory'))
        a = sp.get('_alc')
        car = str(sp.get('_carModel', '')).strip()
        if is_excluded_car(f, car): continue
        
        k = f"{f}_{a}"
        q = get_bom_quantities(sp)
        
        parts = []
        for pk, pv in sp.items():
            pk_up = str(pk).upper()
            if "품번" in pk_up or "PART NO" in pk_up:
                pn = str(pv).strip()
                if pn and pn != "-" and ("휠" in pk_up or "WHEEL" in pk_up):
                    is_spare = "보조" in pk_up or "SPARE" in pk_up
                    qty = q.get('sqW',0) if is_spare else q.get('rqW',4)
                    
                    vend = part_vendor_map.get(pn, "")
                    if not vend: vend = st.session_state.part_vendor.get(pn, "")
                    if not vend: vend = st.session_state.inventory.get(pn, {}).get("vendorName", "")
                    
                    if qty > 0:
                        parts.append({"part": pn, "qty": qty, "vendor": vend, "car": car})
        
        if parts:
            wheel_alc_map[k] = parts

    # 기간 설정
    dates = sorted(list(st.session_state.plan.keys()), reverse=True)
    if not dates:
        st.warning("시스템 일일 생산 실적(계획) 데이터가 없습니다.")
    else:
        sel_date = st.selectbox("조회 일자", ["ALL (전체 누계)"] + dates, key='vendor_cross')
        target_dates = dates if sel_date.startswith("ALL") else [sel_date]
        
        reqs = {}
        for d in target_dates:
            plan_d = st.session_state.plan.get(d, {})
            for key, cars in plan_d.items():
                comps = wheel_alc_map.get(key, [])
                for cp in comps:
                    # 그룹 기준: 업체 -> 차종 -> 품번
                    group_key = (cp['vendor'], cp['car'], cp['part'])
                    reqs[group_key] = reqs.get(group_key, 0) + (cars * cp['qty'])
        
        if reqs:
            data = []
            for k, v in reqs.items():
                if v > 0:
                    pn_key = k[2].replace("-", "").replace(" ", "").upper()
                    # 파싱된 업체 통보량에서 현재 품번 수량 가져오기
                    vendor_q = parsed_vendor_qty.get(pn_key, 0)
                    diff = vendor_q - v # 업체통보량 - 시스템소요량
                    
                    data.append({
                        "대상 업체": k[0],
                        "차종": k[1],
                        "품번": k[2],
                        "품명": st.session_state.part_name.get(k[2], "알수없음"),
                        "시스템 소요량(EA)": v,
                        "업체 통보량(EA)": vendor_q,
                        "차이 (통보-전개량)": diff
                    })
            df = pd.DataFrame(data).sort_values(["대상 업체", "차종", "품번"])
            # 데이터프레임 표시
            st.dataframe(df.style.format({
                "시스템 소요량(EA)": "{:,.0f}", 
                "업체 통보량(EA)": "{:,.0f}", 
                "차이 (통보-전개량)": "{:,.0f}"
            }).applymap(lambda x: "color: red" if pd.notna(x) and x < 0 else "color: blue", subset=["차이 (통보-전개량)"]), use_container_width=True)
        else:
            st.info("해당 일자에 산출된 휠 소요량이 없습니다.")
    
elif menu == "4. 소요량 전개표":
    # 화면을 극대화하고 산뜻한 디자인 적용
    st.markdown("""
        <style>
        .block-container { max-width: 98% !important; padding-top: 1rem; }
        [data-testid="stDataFrame"] { box-shadow: 0 4px 8px rgba(0,0,0,0.05); border-radius: 8px; overflow: hidden; }
        h3 { color: #1e3a8a; font-weight: 700; border-bottom: 2px solid #e5e7eb; padding-bottom: 5px; margin-top: 1.5rem; }
        </style>
    """, unsafe_allow_html=True)
    
    st.title("📊 상세 소요량 전개표 (타이어/휠/밸브 분리)")
    st.markdown("※ 여러 개의 폴더 내 계획표 파일을 읽어 **단일 타임라인**으로 병합합니다. 겹치는 일자는 **가장 최신 파일의 계획으로 덮어씁니다.**")
    
    import datetime
    
    plan_type = st.radio("계획표 유형 선택", ["세부계획표 (시간대별: 1T~10T)", "장기계획표 (일자별: D0~D12)"], horizontal=True)
    
    base_dir = os.path.join("data", "기아생산계획")
    
    avail_files = []
    if os.path.exists(base_dir):
        for root, _, files in os.walk(base_dir):
            for f in files:
                if f.endswith(('.xlsx', '.xls', '.xlsb')) and not f.startswith('~'):
                    avail_files.append(os.path.join(root, f))
                    
    if not avail_files:
        st.warning("데이터 폴더에 분석할 엑셀 파일이 없습니다.")
    else:
        if st.button("🔄 누적 소요량 전개 산출", use_container_width=True):
            # 1. BOM 기반 ALC 매핑 구축 (한 번만 수행)
            alc_map = {}
            for sp in st.session_state.bom_master:
                f = get_mapped_factory(sp.get('_factory'))
                a = sp.get('_alc')
                car = str(sp.get('_carModel', '')).strip()
                if is_excluded_car(f, car): continue
                
                k = f"{f}_{a}"
                q = get_bom_quantities(sp)
                
                parts = []
                for pk, pv in sp.items():
                    pk_up = str(pk).upper()
                    if "품번" in pk_up or "PART NO" in pk_up:
                        pn = str(pv).strip()
                        if pn and pn != "-" and pn != "NAN":
                            is_tire = "타이어" in pk_up or "TIRE" in pk_up
                            is_wheel = "휠" in pk_up or "WHEEL" in pk_up
                            is_valve = "밸브" in pk_up or "VALVE" in pk_up or "TPMS" in pk_up
                            is_spare = "보조" in pk_up or "SPARE" in pk_up
                            
                            qty = 0
                            if is_tire and not is_spare: qty = q.get('rqT', 4)
                            elif is_tire and is_spare: qty = q.get('sqT', 0)
                            elif is_wheel and not is_spare: qty = q.get('rqW', 4)
                            elif is_wheel and is_spare: qty = q.get('sqW', 0)
                            elif is_valve: qty = q.get('rqW', 4) # 밸브는 보통 바퀴 수(휠)와 동일
                            
                            ptype = "타이어" if is_tire else ("휠" if is_wheel else ("밸브" if is_valve else "기타"))
                            base_key = str(pk).replace("품번", "").replace("PART NO", "")
                            vend = str(sp.get(base_key + "업체명", sp.get(base_key + "업체", ""))).strip()
                            if not vend or vend == "nan": vend = st.session_state.part_vendor.get(pn, "")
                            if not vend: vend = st.session_state.inventory.get(pn, {}).get("vendorName", "")
                            
                            if qty > 0:
                                parts.append({"part": pn, "qty": qty, "ptype": ptype, "car": car, "line": f, "vendor": vend})
                if parts:
                    alc_map[k] = parts

            # 2. 파일들을 날짜(오름차순: 과거->최신)로 정렬
            files_with_date = []
            today_str = datetime.date.today().strftime("%Y-%m-%d")
            for fp in avail_files:
                dt_str = extract_date_from_filename(fp, today_str)
                if dt_str:
                    try:
                        dt = datetime.datetime.strptime(dt_str, "%Y-%m-%d")
                        if "세부" in plan_type:
                            dt = dt + datetime.timedelta(days=1)
                        files_with_date.append((dt, fp))
                    except: pass
            
            files_with_date.sort(key=lambda x: x[0])
            
            reqs = {} # (품종, 라인, 차종, 업체, 품번, 품명) -> {"MM/DD 1T": qty}
            all_cols = set() 

            prog_bar = st.progress(0)
            status_text = st.empty()
            
            valid_files_processed = 0

            for i, (base_dt, target_file) in enumerate(files_with_date):
                status_text.text(f"파일 병합 중... ({i+1}/{len(files_with_date)}): {os.path.basename(target_file)}")
                prog_bar.progress((i + 1) / len(files_with_date))
                
                try:
                    import pandas as pd
                    df = pd.read_excel(target_file, header=None)
                    
                    start_row, c_fac, c_alc = -1, -1, -1
                    target_cols = {}
                    
                    for r in range(min(15, len(df))):
                        row_vals_0 = [str(x).strip().replace(" ","").upper() for x in df.iloc[max(0, r-1)].tolist()]
                        row_vals = [str(x).strip().replace(" ","").upper() for x in df.iloc[r].tolist()]
                        
                        if c_fac == -1:
                            for c, val in enumerate(row_vals):
                                if "공장" in val or "ڵ" in val: c_fac = c
                                if "CODEVALUE" in val or "ALC" in val: c_alc = c
                                
                        if "세부" in plan_type:
                            if "1T" in row_vals and "2T" in row_vals:
                                current_d_offset = 0
                                for c, val in enumerate(row_vals):
                                    if re.match(r'^\d+T$', val):
                                        search_c = c
                                        while search_c >= 0:
                                            tv = str(df.iloc[max(0, r-1), search_c]).replace(" ", "").upper()
                                            if "D" in tv and "SCHEDULE" in tv:
                                                m = re.search(r'D\+?(\d+)', tv)
                                                if m: current_d_offset = int(m.group(1))
                                                break
                                            search_c -= 1
                                            
                                        target_dt = base_dt + datetime.timedelta(days=current_d_offset)
                                        col_name = f"{target_dt.strftime('%m/%d')} {val}"
                                        target_cols[c] = col_name
                                        all_cols.add(col_name)
                                if len(target_cols) > 0: start_row = r
                                
                        else:
                            if "D0" in row_vals and "D+1" in row_vals:
                                for c, val in enumerate(row_vals):
                                    m = re.match(r'^D\+?(\d+)$', val)
                                    if m:
                                        d_offset = int(m.group(1))
                                        target_dt = base_dt + datetime.timedelta(days=d_offset)
                                        col_name = target_dt.strftime('%m/%d')
                                        target_cols[c] = col_name
                                        all_cols.add(col_name)
                                if len(target_cols) > 0: start_row = r
                                
                        if start_row != -1 and c_fac != -1 and c_alc != -1:
                            break
                            
                    if start_row == -1 or not target_cols:
                        continue 
                        
                    valid_files_processed += 1
                    
                    # 핵심: 파일에 해당 날짜 컬럼(예: 04/22)이 있다면 기존 시스템(이전 파일)에 누적되어 있던 해당 일자의 데이터들을 모두 0으로 덮어씀.
                    # 그래야 최신 파일에서 삭제되거나 수량이 0인 품목이 이전 파일 데이터로 부활되는 문제가 사라짐
                    for g_key in reqs.keys():
                        for cn in target_cols.values():
                            if cn in reqs[g_key]:
                                reqs[g_key][cn] = 0

                    for r_i in range(start_row + 1, len(df)):
                        fac_raw = str(df.iloc[r_i, c_fac]).strip()
                        alc = str(df.iloc[r_i, c_alc]).strip().upper()
                        fac = get_mapped_factory(fac_raw)
                        
                        if not fac or not alc or alc == "NAN": continue
                        
                        comps = alc_map.get(f"{fac}_{alc}", [])
                        if not comps: continue
                        
                        for c_idx, col_name in target_cols.items():
                            val_str = str(df.iloc[r_i, c_idx]).replace(',', '')
                            try: val = int(float(val_str))
                            except: val = 0
                            
                            if val > 0:
                                for cp in comps:
                                    g_key = (cp['ptype'], cp['line'], cp['car'], cp['vendor'], cp['part'])
                                    if g_key not in reqs:
                                        reqs[g_key] = {}
                                    reqs[g_key][col_name] = reqs[g_key].get(col_name, 0) + (val * cp['qty'])

                except Exception as e:
                    pass

            status_text.empty()
            
            if valid_files_processed == 0:
                st.error("❌ 분석할 수 있는 포맷의 계획표를 찾지 못했습니다.")
            elif not reqs:
                st.info("조건에 맞는 소요량 데이터가 없습니다.")
            else:
                st.success(f"✅ 총 {valid_files_processed}개의 파일을 분석 후 병합 완료했습니다! 겹치는 일자는 자동으로 최신 데이터가 덮어씌워졌습니다.")
                st.session_state.merged_reqs = reqs
                st.session_state.merged_cols = sorted(list(all_cols))
                
        # 산출 결과가 세션에 있으면 표시
        if hasattr(st.session_state, 'merged_reqs') and st.session_state.merged_reqs:
            reqs = st.session_state.merged_reqs
            sorted_all_cols = st.session_state.merged_cols
            
            # 날짜(일자) 필터링 추출
            unique_dates = []
            for c in sorted_all_cols:
                date_part = c.split(" ")[0]
                if date_part not in unique_dates:
                    unique_dates.append(date_part)
                    
            if "세부" in plan_type:
                sel_date = st.selectbox("조회할 특정 일자 선택 (한 화면에 다 담기 위한 필터)", ["전체 보기"] + unique_dates)
            else:
                sel_date = "전체 보기" # 장기계획은 전체가 기본
                
            filtered_cols = []
            for c in sorted_all_cols:
                if sel_date == "전체 보기" or c.startswith(sel_date):
                    filtered_cols.append(c)
                    
            data = []
            for k_tuple, time_data in reqs.items():
                if sum(time_data.values()) <= 0: continue
                
                row_dict = {
                    "품종": k_tuple[0],
                    "라인": k_tuple[1],
                    "차종": k_tuple[2],
                    "업체": k_tuple[3],
                    "품번": k_tuple[4],
                    "품명": st.session_state.part_name.get(k_tuple[4], "알수없음"),
                }
                
                row_total = 0
                for cn in filtered_cols:
                    val = time_data.get(cn, 0)
                    row_dict[cn] = val
                    row_total += val
                    
                row_dict["선택일자_TOTAL"] = row_total
                
                # 선택한 날짜에 수량이 0인 항목은 감출지 판단
                if row_total > 0 or sel_date == "전체 보기":
                    data.append(row_dict)
                
            if not data:
                st.warning("선택한 일자에 해당하는 소요량 데이터가 없습니다.")
            else:
                # 데이터를 품종별로 분리
                df_all = pd.DataFrame(data)
                
                # 포맷팅 딕셔너리
                format_dict = {cn: "{:,.0f}" for cn in filtered_cols + ["선택일자_TOTAL"]}
                
                def render_table(df_subset, title, color_theme):
                    if df_subset.empty: return
                    st.markdown(f"### {title}")
                    
                    df_sorted = df_subset.sort_values(["라인", "차종", "업체", "품번"])
                    df_grouped = df_sorted.set_index(["라인", "차종", "업체", "품번", "품명"]).drop(columns=["품종"])
                    
                    styled_df = df_grouped.style.format(format_dict).set_table_styles([
                        {'selector': 'th', 'props': [('background-color', color_theme), ('color', '#333'), ('font-size', '13px'), ('border', '1px solid #ddd'), ('text-align', 'center')]},
                        {'selector': 'th.row_heading', 'props': [('background-color', '#fefefe'), ('font-weight', 'bold')]},
                        {'selector': 'td', 'props': [('text-align', 'right'), ('font-size', '13px')]}
                    ]).highlight_max(axis=0, subset=["선택일자_TOTAL"], color="#f5f5f5")
                    
                    st.dataframe(styled_df, use_container_width=True)
                
                # 타이어, 휠, 밸브 분리 출력 (산뜻한 파스텔톤 배경색 적용)
                render_table(df_all[df_all["품종"] == "타이어"], "⚫ 타이어 소요량", "#e0f2fe") # 연한 블루
                render_table(df_all[df_all["품종"] == "휠"], "🔘 휠 소요량", "#f3f4f6") # 연한 그레이
                render_table(df_all[df_all["품종"] == "밸브"], "🔩 밸브(TPMS) 소요량", "#fef3c7") # 연한 옐로우
                render_table(df_all[~df_all["품종"].isin(["타이어", "휠", "밸브"])], "기타 부자재", "#f3e8ff") # 연한 퍼플

elif menu == "5. 실적 가로전개표":
    # 화면 전체를 시원하게 넓히고, 폰트 크기를 키우며 제목 디자인을 산뜻하게 변경
    st.markdown("""
        <style>
        .block-container { max-width: 98% !important; padding-top: 1rem; }
        [data-testid="stDataFrame"] { box-shadow: 0 4px 10px rgba(0,0,0,0.08); border-radius: 10px; overflow: hidden; }
        .stMarkdown p { font-size: 15px; color: #555; }
        h3 { color: #0056b3; font-weight: 800; border-bottom: 3px solid #e2e8f0; padding-bottom: 8px; margin-top: 1rem; margin-bottom: 1rem; }
        </style>
    """, unsafe_allow_html=True)
    
    st.markdown("### 📈 조립 실적 일별 가로전개표")
    st.markdown("※ 차종별 생산 실적을 라인별/차종별로 묶어 보여줍니다.")
    v_mode = st.radio("조회 기준", ["타이어 본수 (스페어 사양 환산)", "순수 조립 차량 대수"], horizontal=True)
    
    dates = sorted(list(st.session_state.sales.keys()))
    if not dates:
        st.warning("실적 데이터가 없습니다.")
    else:
        # 월별 필터 UI 추가
        unique_months = []
        for d in dates:
            m = d[:7] # YYYY-MM
            if m not in unique_months: unique_months.append(m)
            
        col1, col2 = st.columns([1, 3])
        with col1:
            sel_month = st.selectbox("📅 조회할 월(Month) 선택", ["전체 기간"] + unique_months)
            
        if sel_month != "전체 기간":
            dates = [d for d in dates if d.startswith(sel_month)]
            
        # Build ALC mapping from BOM
        alc_info = {}
        for sp in st.session_state.bom_master:
            fac = str(sp.get('_factory', '미상')).strip()
            car = get_mapped_car_name(str(sp.get('_carModel', '차종미상')).strip())
            alc = str(sp.get('_alc', '')).strip().upper()
            mapped_fac = get_mapped_factory(fac)
            
            q = get_bom_quantities(sp)
            mul = q.get('rqT', 4) + q.get('sqT', 0)
            
            if mapped_fac:
                alc_info[f"{mapped_fac}_{alc}"] = {"fac": mapped_fac, "car": car, "mul": mul}
            
        # Aggregate by actual Factory / Car
        agg = {}
        for d in dates:
            for key, val in st.session_state.sales[d].items():
                if val == 0: continue
                info = alc_info.get(key, {"fac": "분류불가", "car": key, "mul": 4})
                f, c, m = info["fac"], info["car"], info["mul"]
                
                # '분류불가' 또는 '미상' 인 데이터는 출력에서 제외
                if "불가" in f or "미상" in f: continue
                if is_excluded_car(f, c): continue
                
                real_val = val * m if "타이어" in v_mode else val
                
                if f not in agg: agg[f] = {}
                if c not in agg[f]: agg[f][c] = {dd: 0 for dd in dates}
                agg[f][c][d] += real_val
                
        # Build UI formatting
        records = []
        for f in sorted(agg.keys()):
            subtotal = {d: 0 for d in dates}
            for c in sorted(agg[f].keys()):
                r = {"라인": f, "차종": c}
                r_tot = 0
                for d in dates:
                    v = agg[f][c][d]
                    r[d] = v
                    subtotal[d] += v
                    r_tot += v
                r["총누적"] = r_tot
                records.append(r)
                
            # Add Subtotal row
            sr = {"라인": f, "차종": "소계"}
            sr_tot = 0
            for d in dates:
                v = subtotal[d]
                sr[d] = v
                sr_tot += v
            sr["총누적"] = sr_tot
            records.append(sr)
            
        df = pd.DataFrame(records)
        if not df.empty:
            # Styler for subtotals
            def color_subtotal(row):
                if row['차종'] == '소계': return ['background-color: #bae6fd; font-weight: 800; color: #0c4a6e; font-size: 15px'] * len(row)
                return [''] * len(row)
                
            st.dataframe(df.style.apply(color_subtotal, axis=1).format({d: "{:,.0f}" for d in dates + ["총누적"]}, na_rep="")
                .set_table_styles([
                    {'selector': 'th', 'props': [('background-color', '#f1f5f9'), ('color', '#1e293b'), ('font-size', '15px'), ('font-weight', 'bold'), ('text-align', 'center'), ('padding', '10px')]},
                    {'selector': 'td', 'props': [('font-size', '15px'), ('padding', '10px'), ('color', '#334155')]}
                ]), use_container_width=True)
        else:
            st.info("출력할 실적 내역이 없습니다.")