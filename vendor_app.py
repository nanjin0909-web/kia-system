import streamlit as st
import os
import datetime
import json
import pandas as pd
import io
import zipfile

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

def safe_read_excel(file, **kwargs):
    """
    openpyxl의 custom doc props TypeError를 우회하는 안전한 엑셀 읽기 함수.
    1) calamine 엔진 우선 시도
    2) 실패 시 openpyxl로 시도
    3) openpyxl도 실패 시 docProps/custom.xml을 제거한 후 재시도
    """
    # 1차: calamine 엔진 시도 (가장 빠르고 안전)
    try:
        if hasattr(file, 'seek'): file.seek(0)
        return pd.read_excel(file, engine='calamine', **kwargs)
    except Exception:
        pass

    # 2차: 기본 openpyxl 시도
    try:
        if hasattr(file, 'seek'): file.seek(0)
        return pd.read_excel(file, **kwargs)
    except Exception as e:
        # 3차: custom.xml 제거 후 재시도 (TypeError: StringProperty 등 대응)
        try:
            if hasattr(file, 'seek'): file.seek(0)
            raw = file.read()
            zin = zipfile.ZipFile(io.BytesIO(raw))
            zout_buf = io.BytesIO()
            with zipfile.ZipFile(zout_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename not in ('docProps/custom.xml',):
                        zout.writestr(item, zin.read(item.filename))
            zout_buf.seek(0)
            return pd.read_excel(zout_buf, **kwargs)
        except Exception as e2:
            raise RuntimeError(f"엑셀 읽기 실패: {e} / {e2}")

def get_mapped_factory(raw_factory):
    f = str(raw_factory).strip().upper()
    if f in ["2911", "2921", "1라인", "H1"]: return "1라인"
    if f in ["2912", "2922", "2라인", "H2"]: return "2라인"
    if f in ["2913", "2923", "3라인", "H3"]: return "3라인"
    return "기타"

def vend_match(bom_vend, sel_vend):
    """업체명 유연 매칭"""
    ALIAS = {
        "핸즈": "핸즈코퍼레이션",
        "성우": "현대성우",
        "현대성우캐스팅": "현대성우",
        "현대성우캐스팅(주)충주공장": "현대성우",
        "코리아휠주식회사": "코리아휠",
        "코리아휠 주식회사": "코리아휠",
    }
    if not bom_vend or not sel_vend: return False
    b = ALIAS.get(bom_vend.strip(), bom_vend.strip()).replace(" ", "")
    s = sel_vend.strip().replace(" ", "")
    return b == s or b in s or s in b

st.title("🏭 협력사 공급계획 크로스체크")

vendor_name = st.selectbox("업체 선택", ["다이캐스탈", "현대성우", "핸즈코퍼레이션", "코리아휠"])
uploaded_files = st.file_uploader("공급계획 파일 업로드", type=['xlsx', 'xls', 'xlsm'], accept_multiple_files=True)

col1, col2 = st.columns(2)
with col1:
    btn_preview = st.button("👀 소요량만 미리보기 (파일 없이)", use_container_width=True)
with col2:
    btn_submit = st.button("🚀 제출 및 크로스체크", use_container_width=True)

if btn_preview or btn_submit:
    if btn_submit and not uploaded_files:
        st.warning("파일을 먼저 업로드해주세요.")
        st.stop()

    # 1. 업체 파일 파싱
    parsed_vendor_qty = {}
    if btn_submit and uploaded_files:
        for f in uploaded_files:
            try:
                df = safe_read_excel(f, header=None)
                if vendor_name == "핸즈코퍼레이션":
                    # B열(1): 품번, Q열(16): D+0 TOTAL
                    for i in range(4, len(df)):
                        row = df.iloc[i]
                        if pd.isna(row.iloc[1]) or str(row.iloc[1]).strip() == "": continue
                        pn = str(row.iloc[1]).strip().replace("-", "").upper()
                        if len(pn) < 5: continue
                        try:
                            qty = int(float(str(row.iloc[16]).replace(',', '')))
                            if qty > 0:
                                parsed_vendor_qty[pn] = parsed_vendor_qty.get(pn, 0) + qty
                        except: pass

                elif vendor_name == "다이캐스탈":
                    # C열(2): 품번, P열(15): 수량
                    for i in range(len(df)):
                        if df.shape[1] <= 15: continue
                        pn = str(df.iloc[i, 2]).strip().replace("-", "").upper()
                        if len(pn) < 5: continue
                        try:
                            qty = int(float(str(df.iloc[i, 15]).replace(',', '')))
                            if qty > 0:
                                parsed_vendor_qty[pn] = parsed_vendor_qty.get(pn, 0) + qty
                        except: pass

                elif vendor_name in ["현대성우", "코리아휠"]:
                    # A열(0): 공장코드, C열(2): 품번, D열(3): PLT당수량, H열(7): PLT수
                    for i in range(len(df)):
                        if df.shape[1] <= 7: continue
                        a_val = str(df.iloc[i, 0]).strip()
                        if not any(x in a_val for x in ["2921", "2922", "2923", "라인"]): continue
                        pn = str(df.iloc[i, 2]).strip().replace("-", "").upper()
                        if len(pn) < 5: continue
                        try:
                            qty_per_plt = int(float(str(df.iloc[i, 3]).replace(',', '')))
                            plt_count = int(float(str(df.iloc[i, 7]).replace(',', '')))
                            total = qty_per_plt * plt_count
                            if total > 0:
                                parsed_vendor_qty[pn] = parsed_vendor_qty.get(pn, 0) + total
                        except: pass

            except Exception as e:
                st.error(f"❌ 파일 파싱 오류 [{f.name}]: {e}")

        if parsed_vendor_qty:
            st.success(f"✅ 업체 파일에서 {len(parsed_vendor_qty)}개 품목 파싱 완료")

    # 2. 시스템 데이터 로드
    plan = load_data('plan', {})
    bom = load_data('bom_master', [])
    part_name = load_data('part_name', {})
    dates = sorted(list(plan.keys()), reverse=True)

    if not dates:
        st.error("⚠️ 시스템에 등록된 생산계획이 없습니다. 관리자 앱에서 먼저 생산계획을 업로드해주세요.")
        st.stop()

    # 날짜 선택
    sel_date = st.selectbox("기준 일자 선택", dates)
    st.subheader(f"📅 기준 일자: {sel_date}")

    # 3. BOM 기반 소요량 전개
    reqs = {}
    plan_d = plan.get(sel_date, {})
    for key, total_qty in plan_d.items():
        parts = key.split('_', 1)
        if len(parts) != 2: continue
        fac, alc = parts
        matches = [b for b in bom if get_mapped_factory(b.get('_factory', '')) == fac and b.get('_alc', '') == alc]
        for m in matches:
            for p_key, v_key, q_key in [
                ("정규_휠_품번", "정규_휠_업체", "정규_수량"),
                ("보조_휠_품번", "보조_휠_업체", "보조_수량")
            ]:
                pn = m.get(p_key, "")
                vend = m.get(v_key, "")
                q_val = float(m.get(q_key, 0) or 0)
                if pn and pn not in ("-", "0", "nan", "") and q_val > 0:
                    if vend_match(vend, vendor_name):
                        pn_clean = pn.replace("-", "").upper()
                        reqs[pn_clean] = reqs.get(pn_clean, 0) + (total_qty * q_val)

    if reqs:
        data = []
        for pn, r_qty in sorted(reqs.items()):
            v_qty = parsed_vendor_qty.get(pn, None) if btn_submit else None
            diff = (v_qty - r_qty) if v_qty is not None else None
            data.append({
                "품번": pn,
                "품명": part_name.get(pn, "-"),
                "기아 소요량": int(r_qty),
                "업체 공급량": f"{int(v_qty):,}" if v_qty is not None else "미제출",
                "차이 (공급-소요)": diff
            })

        df_result = pd.DataFrame(data)

        def color_diff(val):
            if val is None or val == "": return ""
            try:
                v = float(val)
                if v < 0: return "color: red; font-weight: bold"
                return "color: blue"
            except: return ""

        st.dataframe(
            df_result.style.map(color_diff, subset=["차이 (공급-소요)"]).format({
                "기아 소요량": "{:,}",
            }),
            use_container_width=True
        )

        # 엑셀 다운로드
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            df_result.to_excel(writer, index=False, sheet_name='크로스체크')
        buf.seek(0)
        st.download_button(
            label="📥 결과 엑셀 다운로드",
            data=buf,
            file_name=f"크로스체크_{vendor_name}_{sel_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info(f"'{vendor_name}' 업체에 해당하는 소요량이 없습니다. BOM 사양표의 업체명을 확인해주세요.")
