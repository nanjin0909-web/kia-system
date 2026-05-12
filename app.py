import streamlit as st
import pandas as pd
import json
import os
import re
from io import BytesIO

st.set_page_config(page_title="스마트 수불/실적 시스템 v56.2", layout="wide", initial_sidebar_state="expanded")

# ==========================================
# 1. 자동 저장소 및 상태 관리
# ==========================================
SAVE_DIR = "data/saves_v56"
if not os.path.exists(SAVE_DIR):
    os.makedirs(SAVE_DIR, exist_ok=True)

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
    if get_drive_service():
        cloud_data = load_from_cloud(key, None)
        if cloud_data is not None:
            with open(FILES[key], "w", encoding="utf-8") as f:
                json.dump(cloud_data, f, ensure_ascii=False, indent=2)
            return cloud_data
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
    if get_drive_service():
        save_to_cloud(key, data)

# 세션 초기화 및 데이터 동기화
for k in FILES.keys():
    if k not in st.session_state:
        st.session_state[k] = load_data(k, [] if k == 'bom_master' else {})
    else:
        # 이미 세션에 있더라도, 파일이 더 최신일 수 있으므로 (다른 탭 작업 등) 
        # 로컬 파일에서 주기적으로 불러오거나 보존 로직 확인
        pass

# ==========================================
# 2. 유틸리티 함수 및 파싱 로직
# ==========================================
def get_mapped_factory(raw_factory):
    f = str(raw_factory).strip().upper()
    if f in ["2911", "2921", "1라인", "H1"]: return "1라인"
    if f in ["2912", "2922", "2라인", "H2"]: return "2라인"
    if f in ["2913", "2923", "3라인", "H3"]: return "3라인"
    return None

def get_mapped_car_name(car_code):
    c = str(car_code).strip().upper()
    if "8V" in c: return "타스만"
    if "GZ" in c or "HC" in c: return "쏘렌토"
    if "OT" in c or "TO" in c: return "니로"
    if "5H" in c: return "SP3"
    if "AS" in c: return "EV6"
    if "EX" in c or "HV" in c: return "K5"
    if "GG" in c or "GL" in c: return "K8"
    if "DL" in c: return "K5(DL3)"
    if "NQ" in c: return "스포티지"
    return c

def is_excluded_car(fac, car):
    f = get_mapped_factory(fac)
    c_up = str(car).strip().upper()
    if f == "1라인" and c_up in ["C5", "DJ", "EN"]: return True
    if f == "2라인" and c_up in ["DQ", "G5", "AS", "DJ"]: return True
    return False

def extract_date_from_filename(file_path, fallback):
    filename = os.path.basename(str(file_path))
    m1 = re.search(r'(20\d{2})[-_ \.]?(0[1-9]|1[0-2])[-_ \.]?(0[1-9]|[12]\d|3[01])', filename)
    if m1:
        return f"{m1.group(1)}-{m1.group(2)}-{m1.group(3)}"
    m2 = re.search(r'(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])', filename)
    if m2:
        return f"2026-{m2.group(1)}-{m2.group(2)}"
    return fallback

def safe_read_excel(file, **kwargs):
    try:
        if hasattr(file, 'seek'): file.seek(0)
        return pd.read_excel(file, engine='calamine', **kwargs)
    except Exception:
        try:
            if hasattr(file, 'seek'): file.seek(0)
            return pd.read_excel(file, **kwargs)
        except Exception as e:
            raise e

def get_bom_quantities(spec):
    """BOM 스펙에서 수량 정보를 추출 (v56.2 고정형)"""
    spare_type = str(spec.get("_spareType", "")).upper()
    rqw = float(spec.get("정규_수량", 4))
    sqw = float(spec.get("보조_수량", 0))
    # 타이어 수량도 휠과 동일하게 처리 (일반적)
    return {"rqW": rqw, "sqW": sqw, "rqT": rqw, "sqT": sqw, "spareType": spare_type}

def parse_bom(file):
    try:
        df = safe_read_excel(file, header=None)
        new_spec = []
        new_parts = {}
        new_vendors = {}
        
        # 실제 master_bom.xlsx 구조에 맞춘 인덱스 매핑
        for i in range(2, len(df)):
            row = df.iloc[i].fillna('')
            fac = str(row[0]).strip()
            car = str(row[1]).strip()
            alc = str(row[2]).strip().upper()
            if not alc or alc == "NAN" or alc == "": continue
            
            spare_type = str(row[7]).strip().upper()
            
            spec_obj = {
                "_factory": fac,
                "_carModel": car,
                "_alc": alc,
                "_spareType": spare_type,
                "정규_타이어_품번": str(row[13]).strip(),
                "정규_타이어_업체": str(row[11]).strip(),
                "정규_휠_품번": str(row[18]).strip(),
                "정규_휠_업체": str(row[11]).strip(),
                "정규_밸브_품번": str(row[23]).strip(),
                "보조_타이어_품번": str(row[30]).strip(),
                "보조_타이어_업체": str(row[28]).strip(),
                "보조_휠_품번": str(row[33]).strip(),
                "보조_휠_업체": str(row[28]).strip(),
                "보조_밸브_품번": str(row[38]).strip(),
                "정규_수량": row[10] if row[10] != "" else 4,
                "보조_수량": row[27] if row[27] != "" else 0
            }
            
            # 품번 정보 수집 (품명은 일단 비움)
            for p_key, v_key in [("정규_타이어_품번", "정규_타이어_업체"), ("정규_휠_품번", "정규_타이어_업체"), 
                                ("보조_타이어_품번", "보조_타이어_업체"), ("보조_휠_품번", "보조_타이어_업체")]:
                pn = spec_obj[p_key]
                vn = spec_obj[v_key]
                if pn and pn != "-" and pn != "0" and pn != "nan":
                    new_vendors[pn] = vn
                    if pn not in st.session_state.part_name:
                        st.session_state.part_name[pn] = "BOM_PART"
            
            new_spec.append(spec_obj)
            
        st.session_state.bom_master = new_spec
        st.session_state.part_vendor.update(new_vendors)
        save_data('bom_master')
        save_data('part_vendor')
        save_data('part_name')
        return f"✅ BOM 갱신 완료 ({len(new_spec)}건)"
    except Exception as e:
        import traceback
        return f"❌ BOM 에러: {e}\n{traceback.format_exc()}"

def parse_system_plan(files, fallback_date):
    if not st.session_state.bom_master: return "❌ BOM 사양표 먼저 갱신 요망"
    act_count, plan_count = 0, 0
    t_plan = st.session_state.plan
    processed_dates = set()
    
    for f in files:
        file_path = getattr(f, "name", str(f))
        target_date = extract_date_from_filename(file_path, fallback_date)
        if target_date not in processed_dates:
            t_plan[target_date] = {}
            processed_dates.add(target_date)
            
        try:
            df = safe_read_excel(f, header=None)
            c_fac, c_alc, c_total = -1, -1, -1
            start_row = -1
            for r in range(min(15, len(df))):
                row = [str(x).replace(" ", "").upper() for x in df.iloc[r]]
                if "ALC" in row or "CODEVALUE" in row:
                    c_alc = row.index("ALC") if "ALC" in row else row.index("CODEVALUE")
                if "2911" in row or "2921" in row:
                    try: c_fac = row.index("2911")
                    except: c_fac = row.index("2921")
                for i, val in enumerate(row):
                    if "D+0TOTAL" in val or "D0TOTAL" in val or "TOTAL" in val:
                        c_total = i
                        break
                if c_alc != -1 and c_total != -1:
                    start_row = r + 1
                    break
            
            if start_row != -1:
                for i in range(start_row, len(df)):
                    row = df.iloc[i]
                    alc = str(row[c_alc]).strip().upper()
                    if not alc or alc == "NAN" or alc == "": continue
                    fac_raw = row[c_fac] if c_fac != -1 else "2911"
                    fac = get_mapped_factory(fac_raw)
                    if not fac: continue
                    try: qty = int(float(str(row[c_total]).replace(',', '')))
                    except: qty = 0
                    if qty > 0:
                        key = f"{fac}_{alc}"
                        t_plan[target_date][key] = t_plan[target_date].get(key, 0) + qty
                        plan_count += 1
        except Exception as e: pass
            
    if plan_count > 0:
        save_data('plan')
        return f"✅ 계획 {plan_count}건 파싱 완료"
    return "❌ 데이터 추출 실패"

def parse_vendor_plan(files):
    parsed_vendor_qty = {}
    for f in files:
        try:
            df = safe_read_excel(f, header=None)
            fname = f.name if hasattr(f, 'name') else str(f)
            # 핸즈코퍼레이션 양식: B열 품번, Q열 D+0 TOTAL
            if "핸즈" in fname:
                for i in range(4, len(df)):
                    row = df.iloc[i]
                    if pd.isna(row[1]): continue
                    pn = str(row[1]).strip().replace("-", "").upper()
                    try:
                        qty = int(float(str(row[16]).replace(',', '')))
                        if qty > 0:
                            parsed_vendor_qty[pn] = parsed_vendor_qty.get(pn, 0) + qty
                    except: pass
            # 다이캐스탈 양식: C열 품번, P열 수량
            elif "다이" in fname:
                for i in range(len(df)):
                    if df.shape[1] > 15:
                        pn = str(df.iloc[i, 2]).strip().replace("-", "").upper()
                        try:
                            qty = int(float(str(df.iloc[i, 15]).replace(',', '')))
                            if qty > 0:
                                parsed_vendor_qty[pn] = parsed_vendor_qty.get(pn, 0) + qty
                        except: pass
        except: pass
    st.session_state.vendor_plan = parsed_vendor_qty
    save_data('vendor_plan')
    return f"✅ 업체 계획 {len(parsed_vendor_qty)}개 품목 갱신"

# ==========================================
# 3. 사이드바 UI
# ==========================================
st.sidebar.markdown("<div class='sidebar-title'>🚀 스마트 공정 통합 관리 시스템</div>", unsafe_allow_html=True)
st.sidebar.write("Ver 56.2 (BOM 매핑 완벽화)")

menu = st.sidebar.radio("메뉴를 선택하세요", [
    "1. 데이터 업로드 센터",
    "2. 수불/실적 모니터링",
    "3. 업체계획 크로스체크"
])

if menu == "1. 데이터 업로드 센터":
    st.title("📁 데이터 업로드 센터")
    with st.expander("⚙️ 0. 기초 마스터 데이터 갱신 (BOM 사양표)", expanded=True):
        f1 = st.file_uploader("BOM 사양표 (Excel)", type=['xls', 'xlsx'])
        if st.button("마스터 갱신") and f1:
            st.success(parse_bom(f1))
    with st.expander("🏭 1. 시스템 생산 계획 업로드", expanded=True):
        d2 = st.date_input("기준일")
        f3 = st.file_uploader("전개표 업로드", type=['xls', 'xlsx'], accept_multiple_files=True)
        if st.button("계획 파싱") and f3:
            st.success(parse_system_plan(f3, str(d2)))
    with st.expander("🚚 2. 업체 통보 계획 업로드", expanded=True):
        f4 = st.file_uploader("업체 파일", type=['xls', 'xlsx'], accept_multiple_files=True)
        if st.button("업체 데이터 파싱") and f4:
            st.success(parse_vendor_plan(f4))

elif menu == "3. 업체계획 크로스체크":
    st.title("⚖️ 업체계획 크로스체크")
    dates = sorted(list(st.session_state.plan.keys()), reverse=True)
    if not dates:
        st.warning("시스템 계획 데이터가 없습니다.")
    else:
        sel_date = st.selectbox("조회 일자", dates)
        plan_d = st.session_state.plan.get(sel_date, {})
        
        # 소요량 전개
        reqs = {}
        for key, total_qty in plan_d.items():
            # key: fac_alc
            fac, alc = key.split('_')
            # BOM에서 해당 ALC 찾기
            matches = [b for b in st.session_state.bom_master if get_mapped_factory(b['_factory']) == fac and b['_alc'] == alc]
            for m in matches:
                q = get_bom_quantities(m)
                # 휠 소요량만 일단 계산 (정규+보조)
                for p_key, q_val in [("정규_휠_품번", q['rqW']), ("보조_휠_품번", q['sqW'])]:
                    pn = m[p_key]
                    if pn and pn != "-" and pn != "0" and pn != "nan":
                        pn_clean = pn.replace("-", "").upper()
                        reqs[pn_clean] = reqs.get(pn_clean, 0) + (total_qty * q_val)
        
        if reqs:
            data = []
            for pn, req_qty in reqs.items():
                vend_qty = st.session_state.vendor_plan.get(pn, 0)
                data.append({
                    "품번": pn,
                    "업체명": st.session_state.part_vendor.get(pn, "미상"),
                    "시스템 소요량": req_qty,
                    "업체 통보량": vend_qty,
                    "차이": vend_qty - req_qty
                })
            df = pd.DataFrame(data)
            st.dataframe(df.style.map(lambda x: "color: red" if x < 0 else "color: blue", subset=["차이"]), use_container_width=True)
        else:
            st.info("산출된 소요량이 없습니다.")
