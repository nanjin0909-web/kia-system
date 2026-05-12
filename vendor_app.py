import streamlit as st
import os
import datetime
import json
import pandas as pd
import io

st.set_page_config(page_title="협력사 포털 v56.2", layout="wide")

SAVE_DIR = "data/saves_v56"
if not os.path.exists(SAVE_DIR): os.makedirs(SAVE_DIR)

from cloud_sync import load_from_cloud, get_drive_service

def load_data(key, default):
    path = os.path.join(SAVE_DIR, f"{key}.json")
    if get_drive_service():
        cloud_data = load_from_cloud(key, None)
        if cloud_data is not None:
            return cloud_data
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return default

def get_mapped_factory(raw_factory):
    f = str(raw_factory).strip().upper()
    if f in ["2911", "2921", "1라인", "H1"]: return "1라인"
    if f in ["2912", "2922", "2라인", "H2"]: return "2라인"
    return "기타"

def vend_match(bom_vend, sel_vend):
    if not bom_vend or not sel_vend: return False
    return sel_vend.strip() in bom_vend.strip() or bom_vend.strip() in sel_vend.strip()

st.title("🏭 협력사 공급계획 크로스체크")

vendor_name = st.selectbox("업체 선택", ["다이캐스탈", "현대성우", "핸즈코퍼레이션", "코리아휠"])
uploaded_files = st.file_uploader("공급계획 파일 업로드", type=['xlsx', 'xls'], accept_multiple_files=True)

if st.button("분석 및 크로스체크"):
    if not uploaded_files:
        st.warning("파일을 업로드해주세요.")
    else:
        # 1. 업체 파일 파싱
        parsed_vendor_qty = {}
        for f in uploaded_files:
            df = pd.read_excel(f, header=None)
            if vendor_name == "핸즈코퍼레이션":
                for i in range(4, len(df)):
                    row = df.iloc[i]
                    if pd.isna(row[1]): continue
                    pn = str(row[1]).strip().replace("-", "").upper()
                    try:
                        qty = int(float(str(row[16]).replace(',', '')))
                        if qty > 0:
                            parsed_vendor_qty[pn] = parsed_vendor_qty.get(pn, 0) + qty
                    except: pass
            elif vendor_name == "다이캐스탈":
                for i in range(len(df)):
                    if df.shape[1] > 15:
                        pn = str(df.iloc[i, 2]).strip().replace("-", "").upper()
                        try:
                            qty = int(float(str(df.iloc[i, 15]).replace(',', '')))
                            if qty > 0:
                                parsed_vendor_qty[pn] = parsed_vendor_qty.get(pn, 0) + qty
                        except: pass
        
        # 2. 시스템 데이터 로드 및 비교
        plan = load_data('plan', {})
        bom = load_data('bom_master', [])
        dates = sorted(list(plan.keys()), reverse=True)
        
        if not dates:
            st.error("시스템에 등록된 생산계획이 없습니다.")
        else:
            target_date = dates[0]
            st.subheader(f"📅 기준 일자: {target_date}")
            
            reqs = {}
            plan_d = plan.get(target_date, {})
            for key, total_qty in plan_d.items():
                fac, alc = key.split('_')
                matches = [b for b in bom if get_mapped_factory(b['_factory']) == fac and b['_alc'] == alc]
                for m in matches:
                    # 해당 업체 부품만 필터링
                    for p_key, v_key, q_val in [
                        ("정규_휠_품번", "정규_휠_업체", float(m.get("정규_수량", 4))),
                        ("보조_휠_품번", "보조_휠_업체", float(m.get("보조_수량", 0)))
                    ]:
                        pn = m.get(p_key)
                        vend = m.get(v_key)
                        if pn and pn != "-" and vend_match(vend, vendor_name):
                            pn_clean = pn.replace("-", "").upper()
                            reqs[pn_clean] = reqs.get(pn_clean, 0) + (total_qty * q_val)
            
            if reqs:
                data = []
                for pn, r_qty in reqs.items():
                    v_qty = parsed_vendor_qty.get(pn, 0)
                    data.append({
                        "품번": pn,
                        "기아 소요량": r_qty,
                        "업체 공급량": v_qty,
                        "차이": v_qty - r_qty
                    })
                st.table(pd.DataFrame(data))
            else:
                st.info("해당 업체에 대한 소요량이 없습니다.")
