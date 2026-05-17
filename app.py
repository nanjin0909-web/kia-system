import streamlit as st
import pandas as pd
import json
import os
import re
from io import BytesIO

st.set_page_config(page_title="스마트 수불/실적 시스템 v56.4", layout="wide", initial_sidebar_state="expanded")

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
    # 클라우드에서 먼저 시도
    cloud_data = load_from_cloud(key, None)
    if cloud_data is not None:
        try:
            with open(FILES[key], "w", encoding="utf-8") as f:
                json.dump(cloud_data, f, ensure_ascii=False, indent=2)
        except: pass
        return cloud_data
    
    if os.path.exists(FILES[key]):
        try:
            with open(FILES[key], "r", encoding="utf-8") as f:
                return json.load(f)
        except: pass
    return default

def save_data(key):
    data = st.session_state[key]
    # 1. 로컬 저장
    try:
        with open(FILES[key], "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        st.error(f"로컬 저장 실패 ({key}): {e}")
    
    # 2. 클라우드 저장
    success = save_to_cloud(key, data)
    return success

# 세션 초기화 및 데이터 동기화
for k in FILES.keys():
    if k not in st.session_state:
        st.session_state[k] = load_data(k, [] if k == 'bom_master' else {})

# ==========================================
# 2. 유틸리티 함수 및 파싱 로직
# ==========================================
def get_mapped_factory(raw_factory):
    f = str(raw_factory).strip().upper()
    if f in ["2911", "2921", "1라인", "H1"]: return "1라인"
    if f in ["2912", "2922", "2라인", "H2"]: return "2라인"
    if f in ["2913", "2923", "3라인", "H3"]: return "3라인"
    return "기타"

def extract_date_from_filename(file_path, fallback):
    filename = os.path.basename(str(file_path))
    m1 = re.search(r'(20\d{2})[-_ \.]?(0[1-9]|1[0-2])[-_ \.]?(0[1-9]|[12]\d|3[01])', filename)
    if m1: return f"{m1.group(1)}-{m1.group(2)}-{m1.group(3)}"
    m2 = re.search(r'(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])', filename)
    if m2: return f"2026-{m2.group(1)}-{m2.group(2)}"
    return fallback

def safe_read_excel(file, **kwargs):
    try:
        if hasattr(file, 'seek'): file.seek(0)
        return pd.read_excel(file, engine='calamine', **kwargs)
    except:
        try:
            if hasattr(file, 'seek'): file.seek(0)
            return pd.read_excel(file, **kwargs)
        except Exception as e: raise e

def get_bom_quantities(spec):
    rqw = float(spec.get("정규_수량", 4) or 4)
    sqw = float(spec.get("보조_수량", 0) or 0)
    return {"rqW": rqw, "sqW": sqw}

def parse_bom(file):
    try:
        df = safe_read_excel(file, header=None)
        new_spec, new_vendors = [], {}
        for i in range(2, len(df)):
            row = df.iloc[i].fillna('')
            alc = str(row[2]).strip().upper()
            if not alc or alc in ("NAN", ""): continue
            spec_obj = {
                "_factory": str(row[0]).strip(),
                "_carModel": str(row[1]).strip(),
                "_alc": alc,
                "정규_타이어_품번": str(row[13]).strip(),
                "정규_휠_품번": str(row[18]).strip(),
                "정규_휠_업체": str(row[11]).strip(),
                "보조_휠_품번": str(row[33]).strip(),
                "보조_휠_업체": str(row[28]).strip(),
                "정규_수량": row[10] if row[10] != "" else 4,
                "보조_수량": row[27] if row[27] != "" else 0
            }
            for p, v in [("정규_휠_품번", "정규_휠_업체"), ("보조_휠_품번", "보조_휠_업체")]:
                pn, vn = spec_obj[p], spec_obj[v]
                if pn and pn not in ("-", "0", "nan"):
                    new_vendors[pn.replace("-","").upper()] = vn
            new_spec.append(spec_obj)
        st.session_state.bom_master = new_spec
        st.session_state.part_vendor.update(new_vendors)
        save_data('bom_master')
        save_data('part_vendor')
        return f"✅ BOM 갱신 완료 ({len(new_spec)}건)"
    except Exception as e: return f"❌ BOM 에러: {e}"

def parse_system_plan(files, fallback_date):
    if not st.session_state.bom_master: return "❌ BOM 사양표 먼저 갱신 요망"
    plan_count = 0
    for f in files:
        target_date = extract_date_from_filename(f.name, fallback_date)
        if target_date not in st.session_state.plan: st.session_state.plan[target_date] = {}
        try:
            df = safe_read_excel(f, header=None)
            c_alc, c_total, c_fac = -1, -1, -1
            for r in range(min(15, len(df))):
                row = [str(x).replace(" ", "").upper() for x in df.iloc[r]]
                if "ALC" in row: c_alc = row.index("ALC")
                if "TOTAL" in row or "D+0TOTAL" in row: c_total = row.index("TOTAL") if "TOTAL" in row else row.index("D+0TOTAL")
                if "2911" in row: c_fac = row.index("2911")
            if c_alc != -1 and c_total != -1:
                for i in range(r+1, len(df)):
                    row = df.iloc[i]
                    alc = str(row[c_alc]).strip().upper()
                    if not alc or alc in ("NAN", ""): continue
                    try: qty = int(float(str(row[c_total]).replace(',', '')))
                    except: qty = 0
                    if qty > 0:
                        fac = get_mapped_factory(row[c_fac] if c_fac != -1 else "2911")
                        key = f"{fac}_{alc}"
                        st.session_state.plan[target_date][key] = st.session_state.plan[target_date].get(key, 0) + qty
                        plan_count += 1
        except: pass
    if plan_count > 0:
        save_data('plan')
        return f"✅ 계획 {plan_count}건 파싱 완료"
    return "❌ 데이터 추출 실패"

def parse_vendor_plan(files):
    all_vendor_plan = {}
    for f in files:
        try:
            df = safe_read_excel(f, header=None)
            fname = f.name.upper()
            if "핸즈" in fname:
                for i in range(4, len(df)):
                    row = df.iloc[i]
                    if pd.isna(row[0]) or str(row[0]).strip() == "": continue
                    pn = str(row[1]).strip().replace("-", "").upper()
                    try:
                        qty = int(float(str(row[16]).replace(',', ''))) # Q열
                        if qty > 0: all_vendor_plan[pn] = all_vendor_plan.get(pn, 0) + qty
                    except: pass
            else:
                # 기본 파서 (B열 품번, P열 수량 가정)
                for i in range(len(df)):
                    pn = str(df.iloc[i, 1]).strip().replace("-", "").upper()
                    try:
                        qty = int(float(str(df.iloc[i, 15]).replace(',', '')))
                        if qty > 0: all_vendor_plan[pn] = all_vendor_plan.get(pn, 0) + qty
                    except: pass
        except: pass
    st.session_state.vendor_plan = all_vendor_plan
    save_data('vendor_plan')
    return f"✅ 업체 계획 {len(all_vendor_plan)}건 갱신"

# ==========================================
# 3. 사이드바 UI 및 메뉴
# ==========================================
st.sidebar.markdown("<div style='font-size:20px; font-weight:bold; color:#1E3A8A;'>🚀 스마트 공정 관리 시스템</div>", unsafe_allow_html=True)
st.sidebar.write(f"Ver 56.4 (클라우드 동기화)")

menu = st.sidebar.radio("메뉴를 선택하세요", [
    "1. 데이터 업로드 센터",
    "2. 수불/실적 모니터링",
    "3. 상세시간 소요량 전개",
    "4. 업체계획 크로스체크"
])

if menu == "1. 데이터 업로드 센터":
    st.title("📁 데이터 업로드 센터")
    st.info("💡 여기서 올린 데이터는 모든 메뉴에서 공유되며 클라우드에 자동 저장됩니다.")
    with st.expander("⚙️ 0. 기초 마스터 데이터 갱신 (BOM 사양표)", expanded=True):
        f1 = st.file_uploader("BOM 사양표 (Excel)", type=['xls', 'xlsx'])
        if st.button("마스터 갱신") and f1: st.success(parse_bom(f1))
    with st.expander("🏭 1. 시스템 생산 계획 업로드", expanded=True):
        d2 = st.date_input("기준일")
        f3 = st.file_uploader("전개표 업로드", type=['xls', 'xlsx'], accept_multiple_files=True)
        if st.button("계획 파싱") and f3: st.success(parse_system_plan(f3, str(d2)))
    with st.expander("🚚 2. 업체 통보 계획 업로드", expanded=True):
        f4 = st.file_uploader("업체 파일", type=['xls', 'xlsx'], accept_multiple_files=True)
        if st.button("업체 데이터 파싱") and f4: st.success(parse_vendor_plan(f4))

elif menu == "2. 수불/실적 모니터링":
    st.title("📊 수불/실적 모니터링")
    st.info("준비 중인 메뉴입니다. (기본 데이터 연동 완료)")

elif menu == "3. 상세시간 소요량 전개":
    st.title("🕒 상세시간 소요량 전개")
    dates = sorted(list(st.session_state.plan.keys()), reverse=True)
    if not dates: st.warning("업로드 센터에서 생산계획을 먼저 올려주세요.")
    else:
        sel_date = st.selectbox("조회 일자", dates)
        plan_d = st.session_state.plan.get(sel_date, {})
        reqs = {}
        for key, total_qty in plan_d.items():
            fac, alc = key.split('_')
            matches = [b for b in st.session_state.bom_master if get_mapped_factory(b['_factory']) == fac and b['_alc'] == alc]
            for m in matches:
                q = get_bom_quantities(m)
                for p_key, q_val in [("정규_휠_품번", q['rqW']), ("보조_휠_품번", q['sqW'])]:
                    pn = m.get(p_key)
                    if pn and pn not in ("-", "0", "nan"):
                        pn_clean = pn.replace("-", "").upper()
                        reqs[pn_clean] = reqs.get(pn_clean, 0) + (total_qty * q_val)
        if reqs:
            st.write(f"### {sel_date} 소요량 집계")
            df = pd.DataFrame([{"품번": k, "업체명": st.session_state.part_vendor.get(k, "미상"), "소요량": v} for k, v in reqs.items()])
            st.dataframe(df, use_container_width=True)
        else: st.info("데이터가 없습니다.")

elif menu == "4. 업체계획 크로스체크":
    st.title("⚖️ 업체계획 크로스체크")
    dates = sorted(list(st.session_state.plan.keys()), reverse=True)
    if not dates: st.warning("시스템 계획 데이터가 없습니다.")
    else:
        sel_date = st.selectbox("조회 일자", dates)
        plan_d = st.session_state.plan.get(sel_date, {})
        reqs = {}
        for key, total_qty in plan_d.items():
            fac, alc = key.split('_')
            matches = [b for b in st.session_state.bom_master if get_mapped_factory(b['_factory']) == fac and b['_alc'] == alc]
            for m in matches:
                q = get_bom_quantities(m)
                for p_key, q_val in [("정규_휠_품번", q['rqW']), ("보조_휠_품번", q['sqW'])]:
                    pn = m.get(p_key)
                    if pn and pn not in ("-", "0", "nan"):
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
            st.dataframe(pd.DataFrame(data).style.map(lambda x: "color: red" if x < 0 else "color: blue", subset=["차이"]), use_container_width=True)
        else: st.info("산출된 소요량이 없습니다.")
