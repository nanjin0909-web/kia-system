import streamlit as st
import os
import datetime
import json
import pandas as pd
import shutil
import io

# 페이지 설정
st.set_page_config(page_title="협력사 & 공용 포털", page_icon="🏭", layout="wide")

UPLOAD_DIR = os.path.join("data", "vendor_uploads")
SHARED_DIR = os.path.join("data", "기아생산계획")
SAVE_DIR = "data/saves_v56"

for d in [UPLOAD_DIR, SHARED_DIR, SAVE_DIR]:
    if not os.path.exists(d): os.makedirs(d)

from cloud_sync import load_from_cloud, get_drive_service

def load_data(key, default):
    path = os.path.join(SAVE_DIR, f"{key}.json")
    if get_drive_service():
        cloud_data = load_from_cloud(key, None)
        if cloud_data is not None:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(cloud_data, f, ensure_ascii=False, indent=2)
            return cloud_data
            
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return default

def get_mapped_factory(raw_factory):
    f = str(raw_factory).strip().upper()
    if f in ["2921", "1라인", "H1"]: return "1라인"
    if f in ["2922", "2라인", "H2"]: return "2라인"
    if f in ["2923", "3라인", "H3"]: return "3라인"
    return None

def is_excluded_car(fac, car):
    f = get_mapped_factory(fac)
    c_up = str(car).strip().upper()
    if f == "1라인" and c_up in ["C5", "DJ", "EN"]: return True
    if f == "2라인" and c_up in ["DQ", "G5", "AS", "DJ"]: return True
    return False

def get_bom_quantities(spec):
    spare_type = str(spec.get("_spareType", "")).upper()
    def get_by_kws(kws):
        for k in spec.keys():
            ck = str(k).replace(" ", "").upper()
            if all(x in ck for x in kws): return spec[k]
        return None
    rqw = float(get_by_kws(["정규", "휠", "수량"]) or get_by_kws(["WHEEL", "수량"]) or 4)
    sqw = float(get_by_kws(["보조", "휠", "수량"]) or get_by_kws(["SPARE", "WHEEL", "수량"]) or 0)
    if "TMK" in spare_type: sqw = 0
    elif "FULL" in spare_type: rqw, sqw = (5, 0) if not (get_by_kws(["보조", "휠", "품번"]) or get_by_kws(["SPARE", "WHEEL", "품번"])) else (4, 1)
    elif "TEMPO" in spare_type: rqw, sqw = 4, 1
    return {"rqW": rqw, "sqW": sqw}

# 간단한 보안
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.title("🏭 통합 시스템 포털 로그인")
    pwd = st.text_input("접속 비밀번호를 입력하세요", type="password")
    if st.button("로그인"):
        if pwd == "1234":
            st.session_state.logged_in = True
            st.rerun()
        else:
            st.error("비밀번호가 틀렸습니다.")
    st.stop()

st.title("🏭 협력사 & 공용 시스템 포털")
tab1, tab2 = st.tabs(["🚀 업체 통보 (크로스체크)", "📁 기아생산계획 공유보드 (사내용)"])

with tab1:
    st.subheader("업체 생산계획 업로드 및 결과 확인")
    vendor_name = st.selectbox("어느 업체이신가요?", [
        "선택하세요",
        "다이캐스탈",
        "현대성우",
        "코리아휠",
        "핸즈코퍼레이션",
        "대원알텍",
        "보정입중",
        "기타"
    ])
    
    if vendor_name != "선택하세요":
        uploaded_files = st.file_uploader(f"[{vendor_name}] 생산계획 엑셀 파일을 올려주세요 (미조회 시 생략 가능)", type=['xlsx', 'xls', 'xlsm'], accept_multiple_files=True)
        
        col1, col2 = st.columns(2)
        with col1:
            btn_preview = st.button("👀 파일 없이 소요량만 미리보기", use_container_width=True)
        with col2:
            btn_submit = st.button("🚀 제출 및 크로스체크 결과 보기", use_container_width=True)
            
        if btn_preview or btn_submit:
            parsed_vendor_qty = {}
            succ_count = 0
            
            if btn_submit:
                if not uploaded_files:
                    st.warning("⚠️ 업로드된 파일이 없습니다. (파일 없이 소요량만 보시려면 왼쪽 '미리보기' 버튼을 눌러주세요)")
                    st.stop()
                    
                # 1. 파일 저장 및 파싱
                now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                for f in uploaded_files:
                    try:
                        save_path = os.path.join(UPLOAD_DIR, f"{now_str}_{vendor_name}_{f.name}")
                        with open(save_path, "wb") as out_file: out_file.write(f.getbuffer())
                        succ_count += 1
                        
                        try:
                            df = pd.read_excel(f, header=None)
                        except Exception as ex:
                            if 'StringProperty' in str(ex) or 'openpyxl' in str(ex):
                                import zipfile
                                from io import BytesIO
                                f.seek(0)
                                zin = zipfile.ZipFile(f)
                                zout_buf = BytesIO()
                                zout = zipfile.ZipFile(zout_buf, 'w')
                                for item in zin.infolist():
                                    if 'docProps/custom.xml' not in item.filename:
                                        zout.writestr(item, zin.read(item.filename))
                                zout.close()
                                zout_buf.seek(0)
                                df = pd.read_excel(zout_buf, header=None)
                            else:
                                raise ex
                        for i in range(len(df)):
                            if df.shape[1] > 2:
                                col_c = str(df.iloc[i, 2]).strip()
                                pn_cleaned = col_c.replace("-", "").replace(" ", "").upper()
                                
                                if vendor_name == "다이캐스탈" and df.shape[1] > 15:
                                    p_val = str(df.iloc[i, 15]).replace(',', '')
                                    if p_val.replace('.', '', 1).isdigit():
                                        qty = int(float(p_val))
                                        if qty > 0 and len(pn_cleaned) >= 5:
                                            parsed_vendor_qty[pn_cleaned] = parsed_vendor_qty.get(pn_cleaned, 0) + qty
                                            
                                elif vendor_name == "현대성우" and df.shape[1] > 7:
                                    a_val = str(df.iloc[i, 0]).strip()
                                    if "2921" in a_val or "2922" in a_val or "2923" in a_val or "라인" in a_val:
                                        d_val = str(df.iloc[i, 3]).replace(',', '')
                                        h_val = str(df.iloc[i, 7]).replace(',', '')
                                        if d_val.replace('.', '', 1).isdigit() and h_val.replace('.', '', 1).isdigit():
                                            qty_per_plt = int(float(d_val))
                                            plt_count = int(float(h_val))
                                            if (qty_per_plt * plt_count) > 0 and len(pn_cleaned) >= 5:
                                                parsed_vendor_qty[pn_cleaned] = parsed_vendor_qty.get(pn_cleaned, 0) + (qty_per_plt * plt_count)
                                else:
                                    pass
                    except Exception as e:
                        st.error(f"❌ '{f.name}' 오류: {e}")
                
                if succ_count > 0:
                    st.success(f"🎉 성공적으로 전송 및 분석을 마쳤습니다. (기준: {vendor_name})")
            else:
                st.info("💡 파일 제출 없이 기아 소요량만 먼저 조회합니다.")
                succ_count = 1  # 패스 통과용
                
            if succ_count > 0:
                # 2. 시스템 베이스 로드 및 비교
                plan = load_data('plan', {})
                bom = load_data('bom_master', [])
                inv = load_data('inventory', {})
                part_vendor = load_data('part_vendor', {})
                part_name = load_data('part_name', {})
                
                dates = sorted(list(plan.keys()), reverse=True)
                
                def vend_match(bom_vend, sel_vend):
                    """업체명 유연 매칭: 완전일치 OR 포함관계 OR 별칭 테이블"""
                    # 별칭 정규화 테이블 (BOM 표기 → 표준 업체명)
                    ALIAS = {
                        "핸즈": "핸즈코퍼레이션",
                        "성우": "현대성우",
                        "현대성우캐스팅": "현대성우",
                        "현대성우캐스팅(주)충주공장": "현대성우",
                        "현대성우캐스팅(주) 충주공장": "현대성우",
                        "코리아휠주식회사": "코리아휠",
                        "코리아휠 주식회사": "코리아휠",
                        "센사타": "센싸타",
                        "컨티네탈": "컨티넨탈",
                        "컨티넨탈": "콘티넨탈",
                    }
                    if not bom_vend or not sel_vend: return False
                    b_raw = bom_vend.strip()
                    b = ALIAS.get(b_raw, b_raw).replace(" ", "")
                    s = sel_vend.strip().replace(" ", "")
                    return b == s or b in s or s in b
                if not dates:
                    st.warning("⚠️ 현재 관리자 시스템에 등록된 기아 생산계획(Base Plan) 데이터가 없어 비교할 수 없습니다. (먼저 관리자 시스템을 통해 생산계획을 동기화해야 합니다.)")
                else:
                        target_date = dates[0]
                        reqs = {}
                        
                        part_vendor_map = {}
                        for sp in bom:
                            for pk, pv in sp.items():
                                pk_up = str(pk).upper()
                                if "품번" in pk_up or "PART NO" in pk_up:
                                    pn = str(pv).strip()
                                    base_key = str(pk).replace("품번", "").replace("PART NO", "")
                                    vend = str(sp.get(base_key + "업체명", sp.get(base_key + "업체", ""))).strip()
                                    if vend and vend != "-" and str(vend).upper() != "NAN":
                                        part_vendor_map[pn] = vend
                        
                        for sp in bom:
                            f = get_mapped_factory(sp.get('_factory'))
                            a = sp.get('_alc')
                            car = str(sp.get('_carModel', '')).strip()
                            if is_excluded_car(f, car): continue
                            
                            k = f"{f}_{a}"
                            q = get_bom_quantities(sp)
                            
                            for pk, pv in sp.items():
                                pk_up = str(pk).upper()
                                if "품번" in pk_up or "PART NO" in pk_up:
                                    pn = str(pv).strip()
                                    if pn and pn != "-" and str(pn).upper() != "NAN":
                                        is_spare = "보조" in pk_up or "SPARE" in pk_up
                                        qty = q.get('sqW',0) if is_spare else q.get('rqW',4)
                                        
                                        vend = part_vendor_map.get(pn, "")
                                        if not vend: vend = part_vendor.get(pn, "")
                                        if not vend: vend = inv.get(pn, {}).get("vendorName", "")
                                        
                                        if vend_match(vend, vendor_name) and qty > 0:
                                            plan_qty = plan.get(target_date, {}).get(k, 0)
                                            if plan_qty > 0:
                                                group_key = (car, pn)
                                                reqs[group_key] = reqs.get(group_key, 0) + (plan_qty * qty)

                        
                        if reqs:
                            data = []
                            for k, v in reqs.items():
                                pn_key = k[1].replace("-", "").replace(" ", "").upper()
                                vendor_q = parsed_vendor_qty.get(pn_key, 0) if btn_submit else None
                                diff = vendor_q - v if vendor_q is not None else None
                                
                                data.append({
                                    "차종": k[0],
                                    "품번": k[1],
                                    "품명": part_name.get(k[1], ""),
                                    "기아 필요량": f"{v:,}",
                                    "업체 통보량": f"{vendor_q:,}" if vendor_q is not None else "미제출",
                                    "차이 (통보-필요)": diff
                                })
                                
                            if btn_submit:
                                st.markdown(f"#### 🔍 [ {vendor_name} ] 크로스체크 결과 (시스템: {target_date} 실적 반영분)")
                            else:
                                st.markdown(f"#### 🔍 [ {vendor_name} ] 기아 소요량 사전 조회 (시스템: {target_date} 실적 반영분)")
                                
                            df = pd.DataFrame(data).sort_values(["차종", "품번"])
                            
                            # 차이 컬럼 포매팅 함수
                            def format_diff(val):
                                if val is None: return "-"
                                return f"{val:,}"
                                
                            df["차이 (통보-필요)"] = df["차이 (통보-필요)"].apply(format_diff)
                            
                            def color_diff(val):
                                if isinstance(val, str) and val.startswith("-") and val != "-":
                                    return "color: red; font-weight: bold"
                                return "color: blue"
                                
                            st.dataframe(df.style.map(color_diff, subset=["차이 (통보-필요)"]), use_container_width=True)
                            
                            # 엑셀 다운로드 버튼
                            buffer = io.BytesIO()
                            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                                df.to_excel(writer, index=False, sheet_name='소요량조회')
                            st.download_button(
                                label="📥 엑셀 파일로 다운로드",
                                data=buffer.getvalue(),
                                file_name=f"{vendor_name}_기아소요량_{target_date}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.info("해당 업체의 기아 소요량 데이터가 없거나 매핑되지 않았습니다.")

with tab2:
    st.subheader("📁 기아 생산계획 (Base Plan) 공유보드")
    st.markdown("관리자 부재 시, 다른 담당자가 **기아 생산계획 엑셀 파일(장기/세부)**을 이곳에 직접 업로드하여 서버(메인 관리자 시스템)와 즉각적으로 공유할 수 있습니다.")
    st.info("💡 이곳에 파일을 올리기만 하면, 관리자용 메인 시스템 4번 메뉴에서 [누적 산출] 버튼을 클릭할 때 해당 파일이 자동으로 동기화(병합)됩니다.")
    
    shared_files = st.file_uploader("기아 생산계획 파일 추가", type=['xlsx', 'xls', 'xlsb'], accept_multiple_files=True)
    if st.button("공유보드에 저장 🚀"):
        if shared_files:
            for f in shared_files:
                save_path = os.path.join(SHARED_DIR, f.name)
                with open(save_path, "wb") as out_file:
                    out_file.write(f.getbuffer())
            st.success(f"✅ {len(shared_files)}개의 기아 생산계획 파일이 성공적으로 서버(공유보드)에 등록되었습니다!")
        else:
            st.warning("등록할 파일을 먼저 추가해 주세요.")
            
    st.markdown("---")
    st.markdown("#### 📋 현재 공유보드에 등록된 기아 생산계획 파일 목록")
    existing_files = []
    for root, _, files in os.walk(SHARED_DIR):
        for f in files:
            if f.endswith(('.xlsx', '.xls', '.xlsb')) and not f.startswith('~'):
                rel_path = os.path.relpath(os.path.join(root, f), SHARED_DIR)
                existing_files.append(rel_path)
    
    if existing_files:
        for f in sorted(existing_files):
            st.write(f"📄 `{f}`")
    else:
        st.write("등록된 기아 생산계획 파일이 없습니다.")
