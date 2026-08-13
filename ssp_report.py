import streamlit as st
import pandas as pd
from datetime import timedelta
from pathlib import Path
import io
import re

# 1. 페이지 설정
st.set_page_config(page_title="애드팝콘 통합 분석기", layout="wide")
st.markdown("""
<style>
table { width: 100% !important; border-collapse: collapse; font-size: 13px; margin-bottom: 25px; table-layout: fixed !important; }
th { background-color: #f0f2f6; text-align: center !important; padding: 10px !important; border: 1px solid #dee2e6; font-weight: bold !important; }
td { padding: 8px !important; border: 1px solid #dee2e6; text-align: center !important; }
td:first-child, th:first-child { text-align: left !important; }
.section-title { padding: 10px; background-color: #31333F; color: white; border-radius: 5px; margin: 20px 0 10px 0; font-weight: bold; }
.report-card { background-color: #f8f9fa; border-left: 5px solid #ff4b4b; padding: 15px; margin-bottom: 10px; border-radius: 5px; font-size: 13px; line-height: 1.6; }
.report-card.plus { border-left-color: #28a745; }
.detail-box { padding: 20px; border: 1px solid #e6e9ef; border-radius: 10px; background-color: #fcfcfc; margin-top: 10px; }
</style>
""", unsafe_allow_html=True)

st.title("📊 애드팝콘 통합 분석기 (by. Pole)")

# --- 사이드바 메뉴 ---
menu = st.sidebar.radio("📊 메뉴 선택", ["💰 마진현황", "SSP 리포트", "DSP 리포트"])


# =========================================================================
# 마진현황 메뉴
# =========================================================================
if menu == "💰 마진현황":
    MANUAL_LIST_FILE = Path("manual_list.txt")
    HIDE_LIST_FILE = Path("hide_list.txt")

    def load_list(filepath):
        if filepath.exists():
            with open(filepath, 'r', encoding='utf-8') as f:
                return [line.strip() for line in f.readlines() if line.strip()]
        return []

    def to_numeric_m(val):
        if pd.isna(val): return 0
        try: return int(float(str(val).replace(',', '').strip()))
        except: return 0

    manual_codes = load_list(MANUAL_LIST_FILE)
    hide_keys = load_list(HIDE_LIST_FILE)

    # 제외 목록 표시
    with st.expander(f"📋 수기정산 제외 ({len(manual_codes)}개) / 숨김 업체 ({len(hide_keys)}개)", expanded=False):
        st.markdown("_조회 후 매체명 표시됨. 수정: `manual_list.txt` / `hide_list.txt`_")
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**🚫 수기정산 제외 (ERP코드)**")
            for c in manual_codes: st.text(f"  {c}")
        with col_b:
            st.markdown("**👁️ 숨김 업체 (회사키)**")
            for c in hide_keys: st.text(f"  {c}")

    # 파일 업로드
    st.header("1️⃣ 파일 업로드")
    col1, col2, col3 = st.columns(3)
    with col1:
        file_erp = st.file_uploader("📂 ERP 데이터", type=["xlsx"], key="m_erp")
    with col2:
        file_console1 = st.file_uploader("📂 콘솔 정산 (일반)", type=["xlsx"], key="m_con1")
    with col3:
        file_console2 = st.file_uploader("📂 콘솔 정산 (조기)", type=["xlsx"], key="m_con2")


    if file_erp and (file_console1 or file_console2):
        
            df_erp = pd.read_excel(file_erp, engine='openpyxl')
            df_erp.columns = df_erp.columns.astype(str).str.strip()
            df_erp['매체사코드'] = df_erp['매체사코드'].astype(str).str.strip()
            df_erp['매체비 합계(KRW)'] = df_erp['매체비 합계(KRW)'].apply(to_numeric_m)
            df_erp['매체수수료 합계(KRW)'] = df_erp['매체수수료 합계(KRW)'].apply(to_numeric_m)

            con_frames = []
            if file_console1:
                df_c1 = pd.read_excel(file_console1, engine='openpyxl')
                df_c1.columns = df_c1.columns.astype(str).str.strip()
                con_frames.append(df_c1)
            if file_console2:
                df_c2 = pd.read_excel(file_console2, engine='openpyxl')
                df_c2.columns = df_c2.columns.astype(str).str.strip()
                con_frames.append(df_c2)

            df_con = pd.concat(con_frames, ignore_index=True)
            df_con['ERP 매체코드'] = df_con['ERP 매체코드'].astype(str).str.strip()
            df_con['회사키'] = df_con['회사키'].astype(str).str.strip()
            df_con['순매체비(KRW)'] = df_con['순매체비(KRW)'].apply(to_numeric_m)
            df_con['운영수수료(KRW)'] = df_con['운영수수료(KRW)'].apply(to_numeric_m)
            df_con['세금계산서발행액(KRW)'] = df_con['세금계산서발행액(KRW)'].apply(to_numeric_m)
            df_con['상태'] = df_con['상태'].astype(str).str.strip()

            # 숨김 제외
            df_con = df_con[~df_con['회사키'].isin(hide_keys)]
            # 수기정산 제외
            manual_mask = df_con['ERP 매체코드'].isin(manual_codes)
            df_manual = df_con[manual_mask].copy()
            df_con = df_con[~manual_mask].copy()

            # 이월 분리
            erp_codes_set = set(df_erp['매체사코드'].astype(str).str.strip())
            iwol_raw = pd.concat(con_frames, ignore_index=True)
            iwol_raw.columns = iwol_raw.columns.astype(str).str.strip()
            iwol_raw['ERP 매체코드'] = iwol_raw['ERP 매체코드'].astype(str).str.strip()
            iwol_raw['회사키'] = iwol_raw['회사키'].astype(str).str.strip()
            iwol_raw['상태'] = iwol_raw['상태'].astype(str).str.strip()
            iwol_raw = iwol_raw[iwol_raw['상태'].str.contains('이월', na=False)]
            iwol_raw = iwol_raw[~iwol_raw['회사키'].isin(hide_keys)]
            iwol_raw = iwol_raw[~iwol_raw['ERP 매체코드'].isin(manual_codes)]
            df_iwol = iwol_raw[(~iwol_raw['ERP 매체코드'].isin(erp_codes_set)) | (iwol_raw['ERP 매체코드'].isin(['-','','nan']))].copy()

            # ERP코드 없는 업체
            no_code_mask = (df_con['ERP 매체코드']=='') | (df_con['ERP 매체코드']=='-') | (df_con['ERP 매체코드']=='nan')
            df_no_code = df_con[no_code_mask].copy()
            if '회사키' in df_no_code.columns:
                df_no_code = df_no_code[~df_no_code['회사키'].isin(hide_keys)]
            df_con = df_con[~no_code_mask].copy()

            # 콘솔 합산
            df_con_grouped = df_con.groupby('ERP 매체코드', as_index=False).agg({
                '회사명':'first','회사키':'first',
                '순매체비(KRW)':'sum','운영수수료(KRW)':'sum','세금계산서발행액(KRW)':'sum',
                '상태': lambda x: ', '.join(x.unique())
            })
            df_con_grouped['비고_이월'] = df_con_grouped['상태'].apply(
                lambda x: '⚠️ 이월합산' if '이월' in str(x) else None)
            df_con = df_con_grouped


            # ERP 매칭
            erp_code_counts = df_erp.groupby('매체사코드').size().reset_index(name='ERP건수')
            merged = pd.merge(df_con, erp_code_counts, left_on='ERP 매체코드', right_on='매체사코드', how='left')
            erp_margin = df_erp.groupby('매체사코드', as_index=False).agg({'매체비 합계(KRW)':'sum','매체수수료 합계(KRW)':'sum'})
            merged = pd.merge(merged, erp_margin, left_on='ERP 매체코드', right_on='매체사코드', how='left', suffixes=('','_erp'))

            split_mask = merged['ERP건수'] > 1
            df_split = merged[split_mask].copy()
            df_normal = merged[~split_mask].copy()

            # 탭
            st.markdown("---")
            tab_normal, tab_split, tab_iwol, tab_nocode, tab_erp_only, tab_manual = st.tabs([
                "✅ 정상","⚠️ 분리 필요","🔄 이월","📝 ERP코드 미등록","📦 ERP에만 존재","🚫 수기정산"])

            with tab_normal:
                if not df_normal.empty:
                    df_nv = df_normal[df_normal['매체사코드'].notna()].copy()
                    con_fee = df_con.groupby('ERP 매체코드', as_index=False)['운영수수료(KRW)'].sum()
                    con_fee = con_fee.rename(columns={'운영수수료(KRW)':'콘솔수수료합'})
                    df_nv = pd.merge(df_nv, con_fee, on='ERP 매체코드', how='left')
                    df_nv['콘솔수수료합'] = df_nv['콘솔수수료합'].fillna(0).astype(int)
                    df_nv['ERP_마진'] = df_nv['매체수수료 합계(KRW)'].fillna(0).astype(int)
                    df_nv['차이'] = df_nv['ERP_마진'] - df_nv['콘솔수수료합']
                    df_nv['상태체크'] = df_nv.apply(lambda r: '✅' if r['차이']==0 else f"⚠️ 차이 {r['차이']:,}원", axis=1)
                    if '비고_이월' in df_nv.columns:
                        df_nv.loc[df_nv['비고_이월'].notna(),'상태체크'] = df_nv.loc[df_nv['비고_이월'].notna(),'상태체크'] + ' (이월합산)'

                    st.success(f"✅ {len(df_nv)}건")
                    col_f1, col_f2 = st.columns(2)
                    with col_f1:
                        filt = st.radio("필터:",["전체","✅ 일치만","⚠️ 차이만"], horizontal=True, key="nf")
                    with col_f2:
                        sort = st.radio("정렬:",["마진 높은순","마진 낮은순","차이 큰순"], horizontal=True, key="ns")

                    disp = df_nv[['회사명','ERP 매체코드','회사키','ERP_마진','콘솔수수료합','차이','상태체크']].copy()
                    disp = disp.rename(columns={'ERP_마진':'마진(ERP)','콘솔수수료합':'콘솔 수수료','상태체크':'상태'})
                    if filt=="✅ 일치만": disp = disp[disp['상태']=='✅']
                    elif filt=="⚠️ 차이만": disp = disp[disp['상태']!='✅']
                    if sort=="마진 높은순": disp = disp.sort_values('마진(ERP)', ascending=False)
                    elif sort=="마진 낮은순": disp = disp.sort_values('마진(ERP)')
                    elif sort=="차이 큰순": disp = disp.sort_values('차이', key=abs, ascending=False)

                    download_df = disp.copy()
                    disp['마진(ERP)'] = disp['마진(ERP)'].apply(lambda x: f"{x:,}")
                    disp['콘솔 수수료'] = disp['콘솔 수수료'].apply(lambda x: f"{x:,}")
                    disp['차이'] = disp['차이'].apply(lambda x: f"{x:,}")
                    st.table(disp.reset_index(drop=True))

                    t_m = df_nv['ERP_마진'].sum(); t_c = df_nv['콘솔수수료합'].sum()
                    st.markdown(f"---\n**합계** | 마진(ERP): **{t_m:,}** | 콘솔 수수료: **{t_c:,}** | 차이: **{t_m-t_c:,}**")

                    buf = io.BytesIO()
                    with pd.ExcelWriter(buf, engine='openpyxl') as w:
                        download_df.to_excel(w, index=False)
                    st.download_button("📥 엑셀", buf.getvalue(), "마진_정상.xlsx")

            with tab_split:
                if not df_split.empty:
                    st.warning(f"⚠️ {len(df_split)}건 분리 필요")
                    d = df_split[['회사명','ERP 매체코드','회사키','ERP건수']].copy()
                    d['ERP건수'] = d['ERP건수'].astype(int)
                    d['마진'] = df_split['매체수수료 합계(KRW)'].fillna(0).astype(int).apply(lambda x: f"{x:,}")
                    st.table(d.reset_index(drop=True))
                else: st.info("해당 없음")

            with tab_iwol:
                if not df_iwol.empty:
                    df_iwol['기준월'] = df_iwol['기준월'].astype(str).str.strip()
                    il = df_iwol.sort_values('기준월', ascending=False).drop_duplicates(subset=['회사명'], keep='first')
                    st.info(f"🔄 {len(il)}건 이월")
                    d = il[['회사명','ERP 매체코드','기준월']].copy()
                    d['ERP 매체코드'] = d['ERP 매체코드'].apply(lambda x: '코드 없음' if x in ['-','','nan'] else x)
                    d['콘솔 매체비'] = il['순매체비(KRW)'].apply(to_numeric_m).apply(lambda x: f"{int(x):,}")
                    d['콘솔 수수료'] = il['운영수수료(KRW)'].apply(to_numeric_m).apply(lambda x: f"{int(x):,}")
                    st.table(d.reset_index(drop=True))
                else: st.info("해당 없음")

            with tab_nocode:
                if not df_no_code.empty:
                    df_no_code['순매체비(KRW)'] = df_no_code['순매체비(KRW)'].apply(to_numeric_m)
                    df_no_code['운영수수료(KRW)'] = df_no_code['운영수수료(KRW)'].apply(to_numeric_m)
                    nc = df_no_code.groupby(['회사명','회사키'], as_index=False).agg({'순매체비(KRW)':'sum','운영수수료(KRW)':'sum'})
                    st.warning(f"📝 {len(nc)}건")
                    d = nc[['회사명','회사키']].copy()
                    d['콘솔 매체비'] = nc['순매체비(KRW)'].astype(int).apply(lambda x: f"{x:,}")
                    d['콘솔 수수료'] = nc['운영수수료(KRW)'].astype(int).apply(lambda x: f"{x:,}")
                    st.table(d.reset_index(drop=True))
                else: st.info("해당 없음")

            with tab_erp_only:
                all_con_codes = set(pd.concat(con_frames)['ERP 매체코드'].astype(str).str.strip().unique())
                eo = df_erp[~df_erp['매체사코드'].isin(all_con_codes)].copy()
                eo = eo[~eo['매체사코드'].isin(manual_codes)]
                eo = eo[~eo['매체사코드'].isin(hide_keys)]
                if not eo.empty:
                    eo['매체비 합계(KRW)'] = eo['매체비 합계(KRW)'].apply(to_numeric_m)
                    eo['매체수수료 합계(KRW)'] = eo['매체수수료 합계(KRW)'].apply(to_numeric_m)
                    ea = eo.groupby(['매체사코드','매체사 명'], as_index=False).agg({'매체비 합계(KRW)':'sum','매체수수료 합계(KRW)':'sum'})
                    st.info(f"📦 {len(ea)}건")
                    d = ea.rename(columns={'매체사코드':'ERP코드','매체사 명':'매체명','매체수수료 합계(KRW)':'마진','매체비 합계(KRW)':'매체비'})
                    d['마진'] = d['마진'].astype(int).apply(lambda x: f"{x:,}")
                    d['매체비'] = d['매체비'].astype(int).apply(lambda x: f"{x:,}")
                    st.table(d[['매체명','ERP코드','매체비','마진']].reset_index(drop=True))
                else: st.info("해당 없음")

            with tab_manual:
                if not df_manual.empty:
                    df_manual['순매체비(KRW)'] = df_manual['순매체비(KRW)'].apply(to_numeric_m)
                    df_manual['운영수수료(KRW)'] = df_manual['운영수수료(KRW)'].apply(to_numeric_m)
                    mg = df_manual.groupby(['회사명','ERP 매체코드','회사키'], as_index=False).agg({'순매체비(KRW)':'sum','운영수수료(KRW)':'sum'})
                    st.info(f"🚫 {len(mg)}건")
                    d = mg[['회사명','ERP 매체코드','회사키']].copy()
                    d['콘솔 매체비'] = mg['순매체비(KRW)'].astype(int).apply(lambda x: f"{x:,}")
                    d['콘솔 수수료'] = mg['운영수수료(KRW)'].astype(int).apply(lambda x: f"{x:,}")
                    st.table(d.reset_index(drop=True))
                else: st.info("해당 없음")

            # 제외 목록 상세
            st.markdown("---")
            with st.expander("📋 제외 목록 상세 (매체명 포함)", expanded=False):
                all_con_raw = pd.concat(con_frames, ignore_index=True)
                all_con_raw['ERP 매체코드'] = all_con_raw['ERP 매체코드'].astype(str).str.strip()
                all_con_raw['회사키'] = all_con_raw['회사키'].astype(str).str.strip()
                c2n = dict(zip(all_con_raw['ERP 매체코드'], all_con_raw['회사명']))
                k2n = dict(zip(all_con_raw['회사키'], all_con_raw['회사명']))
                ca, cb = st.columns(2)
                with ca:
                    st.markdown("**🚫 수기정산 제외**")
                    for c in manual_codes: st.text(f"  {c} → {c2n.get(c,'(미매칭)')}")
                with cb:
                    st.markdown("**👁️ 숨김 업체**")
                    for c in hide_keys: st.text(f"  {c} → {k2n.get(c,'(미매칭)')}")


          
# =========================================================================
# SSP / DSP 리포트 메뉴 (기존 코드)
# =========================================================================
else:
    data_type = menu  # "SSP 리포트" or "DSP 리포트"

    def to_excel_with_format(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Report')
            workbook, worksheet = writer.book, writer.sheets['Report']
            money_fmt = workbook.add_format({'num_format': '$#,##0.00'})
            pct_fmt = workbook.add_format({'num_format': '0.0"%"'})
            wrap_fmt = workbook.add_format({'text_wrap': True}); wrap_fmt.set_align('top')
            num_fmt = workbook.add_format({'num_format': '#,##0'})
            for i, col in enumerate(df.columns):
                if '%' in col: worksheet.set_column(i, i, 10, pct_fmt)
                elif '매출' in col or 'eCPM' in col: worksheet.set_column(i, i, 12, money_fmt)
                elif 'AI' in col or '리포트' in col: worksheet.set_column(i, i, 70, wrap_fmt)
                elif '요청' in col or '노출' in col: worksheet.set_column(i, i, 12, num_fmt)
                else: worksheet.set_column(i, i, 15)
        return output.getvalue()

    def clean_df_for_excel(df):
        clean_df = df.copy()
        for col in clean_df.columns:
            clean_df[col] = clean_df[col].astype(str).str.replace(r'<[^>]*>', '', regex=True)
            clean_df[col] = clean_df[col].str.replace('▲','',regex=False).str.replace('▼','',regex=False)
            if any(k in col for k in ['%','매출','eCPM','요청','노출']):
                clean_df[col] = clean_df[col].apply(lambda x: re.sub(r'[^0-9.\-]','',str(x)))
                clean_df[col] = pd.to_numeric(clean_df[col].str.strip(), errors='coerce').fillna(0)
        return clean_df

    def generate_complex_report(media_name, full_df, p_s, p_e, c_s, c_e, c_pla_id, c_pla_name, c_rev, row_data, c_med):
        try:
            rev_a = float(str(row_data['이전 매출']).replace('$','').replace(',',''))
            rev_b = float(str(row_data['현재 매출']).replace('$','').replace(',',''))
            rev_diff = rev_b - rev_a
            if rev_diff == 0: return "변동 없음"
            m_df = full_df[full_df['media_name']==media_name].copy()
            mask_a = (m_df['report_date']>=p_s)&(m_df['report_date']<=p_e)
            mask_b = (m_df['report_date']>=c_s)&(m_df['report_date']<=c_e)
            m_df['pla_full'] = m_df[c_pla_name].astype(str)+" ("+m_df[c_pla_id].astype(str)+")"
            p_a = m_df[mask_a].groupby('pla_full')[c_rev].sum()
            p_b = m_df[mask_b].groupby('pla_full')[c_rev].sum()
            p_diff = (p_b-p_a).fillna(p_b).fillna(-p_a)
            sorted_p = p_diff.sort_values() if rev_diff<0 else p_diff.sort_values(ascending=False)
            top_plas = [f"{p}[기여 {v/rev_diff*100:.1f}%]" for p,v in sorted_p.items() if abs(v/rev_diff*100)>15]
            ssp_a = m_df[mask_a].groupby(c_med)[c_rev].sum()
            ssp_b = m_df[mask_b].groupby(c_med)[c_rev].sum()
            ssp_diff = (ssp_b-ssp_a).fillna(ssp_b).fillna(-ssp_a)
            sorted_s = ssp_diff.sort_values() if rev_diff<0 else ssp_diff.sort_values(ascending=False)
            top_ssps = [s for s,v in sorted_s.items() if abs(v)>abs(rev_diff)*0.1]
            main_type = "미디에이션" if top_ssps else "매체트래픽"
            report = f"【주요원인】 {main_type} ({', '.join(top_ssps[:2])})\n"
            report += f"【상세분석】 주요지면: {', '.join(top_plas[:2]) if top_plas else '전반적 변동'}\n"
            report += f"【수치요약】 매출 ${abs(rev_diff):,.2f} {'하락' if rev_diff<0 else '상승'}."
            return report
        except: return "진단 불가"

    def format_arrow(curr, prev):
        if prev==0: return "<span style='color:red'>▲ New</span>" if curr>0 else "0.0%"
        diff = ((curr-prev)/prev)*100
        color = "red" if diff>0 else "blue"
        arrow = "▲" if diff>0 else "▼"
        return f"<span style='color:{color}'>{arrow} {diff:.1f}%</span>"

    def make_ordered_table(merged_df, name_col, c_req, c_rev, c_imp):
        res = pd.DataFrame()
        res['항목명'] = merged_df[name_col]
        for c, l in zip([c_req,c_imp,c_rev], ['요청','노출','매출']):
            fmt = '{:,.0f}' if l!='매출' else '${:,.2f}'
            res[f'이전 {l}'] = merged_df[f'{c}_A'].map(fmt.format)
            res[f'현재 {l}'] = merged_df[f'{c}_B'].map(fmt.format)
            res[f'{l} %'] = [format_arrow(b,a) for a,b in zip(merged_df[f'{c}_A'], merged_df[f'{c}_B'])]
        ea = (merged_df[c_rev+'_A']/merged_df[c_imp+'_A']*1000).fillna(0)
        eb = (merged_df[c_rev+'_B']/merged_df[c_imp+'_B']*1000).fillna(0)
        res['이전 eCPM'] = ea.map('${:,.2f}'.format); res['현재 eCPM'] = eb.map('${:,.2f}'.format)
        res['eCPM %'] = [format_arrow(b,a) for a,b in zip(ea, eb)]
        return res

    @st.cache_data
    def load_and_merge_data(uploaded_files):
        all_dfs = []
        for f in uploaded_files:
            df = pd.read_csv(f, encoding='utf-8-sig') if f.name.endswith('csv') else pd.concat(pd.read_excel(f, sheet_name=None).values())
            df.columns = df.columns.str.strip()
            df['report_date'] = pd.to_datetime(df['report_date'], format='%Y%m%d', errors='coerce').dt.date
            all_dfs.append(df)
        return pd.concat(all_dfs, ignore_index=True).drop_duplicates()

    if data_type == "SSP 리포트":
        c_med,c_req,c_imp,c_rev,c_app,c_pla_n,c_pla_id = 'thirdparty_name','request_value','impression_value','partner_revenue','media_name','placement_name','placement_id'
    else:
        c_med,c_req,c_imp,c_rev,c_app,c_pla_n,c_pla_id = 'dsp_name','response_value','impression_value','partner_revenue','media_name','placement_name','placement_id'

    uploaded_files = st.file_uploader(f"{data_type} 업로드", type=['xlsx','csv'], accept_multiple_files=True)


    if uploaded_files:
        try:
            df = load_and_merge_data(uploaded_files)
            all_dates = sorted([d for d in df['report_date'].unique() if pd.notnull(d)])
            mode = st.radio("🔍 분석 모드", ["DOD","자유 기간","MOM"], horizontal=True)
            p_s,p_e,c_s,c_e = None,None,None,None
            if mode=="DOD":
                t = st.selectbox("기준일", all_dates[::-1]); c_s=c_e=t; p_s=p_e=t-timedelta(days=1)
            elif mode=="자유 기간":
                c1,c2 = st.columns(2)
                with c1: pr=st.date_input("이전",[all_dates[0],all_dates[0]]); p_s,p_e=pr if len(pr)==2 else (None,None)
                with c2: cr=st.date_input("현재",[all_dates[-1],all_dates[-1]]); c_s,c_e=cr if len(cr)==2 else (None,None)
            elif mode=="MOM":
                df['m']=pd.to_datetime(df['report_date']).dt.to_period('M')
                ms=sorted(df['m'].unique(),reverse=True)
                pm,cm=st.selectbox("이전 월",ms[1:] if len(ms)>1 else ms),st.selectbox("현재 월",ms)
                p_s,p_e,c_s,c_e=pm.start_time.date(),pm.end_time.date(),cm.start_time.date(),cm.end_time.date()

            if p_s and c_s:
                def get_agg(s,e,g):
                    mask=(df['report_date']>=s)&(df['report_date']<=e)
                    return df[mask].groupby(g).agg({c_req:'sum',c_imp:'sum',c_rev:'sum'}).reset_index()

                st.markdown(f"<div class='section-title'>🌐 {data_type} 통합 성과</div>", unsafe_allow_html=True)
                m_p,m_c = get_agg(p_s,p_e,[c_med]), get_agg(c_s,c_e,[c_med])
                m_f = m_c.merge(m_p, on=c_med, how='outer', suffixes=('_B','_A')).fillna(0)
                st.write(make_ordered_table(m_f.sort_values(c_rev+'_B',ascending=False),c_med,c_req,c_rev,c_imp).to_html(escape=False,index=False), unsafe_allow_html=True)

                st.markdown("<div class='section-title'>🏠 매체별 성과 요약</div>", unsafe_allow_html=True)
                a_p,a_c = get_agg(p_s,p_e,[c_app]), get_agg(c_s,c_e,[c_app])
                a_f = a_c.merge(a_p, on=c_app, how='outer', suffixes=('_B','_A')).fillna(0)
                a_f['rev_diff'] = a_f[c_rev+'_B']-a_f[c_rev+'_A']
                base = a_f[a_f[c_req+'_B']>=5000].sort_values(c_rev+'_B',ascending=False)
                search = st.multiselect("🔍 매체명 검색", options=base[c_app].unique())
                f_df = base[base[c_app].isin(search)] if search else base
                pg = st.number_input("페이지",1,max(1,len(f_df)//20+1),1)
                st.write(make_ordered_table(f_df.iloc[(pg-1)*20:pg*20],c_app,c_req,c_rev,c_imp).to_html(escape=False,index=False), unsafe_allow_html=True)

                st.divider()
                st.markdown(f"<div class='section-title'>🎯 매체 상세 분석</div>", unsafe_allow_html=True)
                sel_app = st.selectbox("분석할 매체", base[c_app].unique(), key="app_sel")
                if sel_app:
                    d_p,d_c = get_agg(p_s,p_e,[c_app,c_med]), get_agg(c_s,c_e,[c_app,c_med])
                    d_f = d_c[d_c[c_app]==sel_app].merge(d_p[d_p[c_app]==sel_app], on=[c_app,c_med], how='outer', suffixes=('_B','_A')).fillna(0)
                    st.write(make_ordered_table(d_f.sort_values(c_rev+'_B',ascending=False),c_med,c_req,c_rev,c_imp).to_html(escape=False,index=False), unsafe_allow_html=True)

                if data_type=="SSP 리포트":
                    st.markdown("<div class='section-title'>🏁 지면 분석</div>", unsafe_allow_html=True)
                    pla_app = st.selectbox("지면 분석용 매체", ["선택 안함"]+list(base[c_app].unique()), key="pla_app")
                    if pla_app != "선택 안함":
                        p_l = sorted(df[df[c_app]==pla_app][c_pla_n].unique())
                        sel_p = st.selectbox("지면 선택", p_l, key="sel_p")
                        if sel_p:
                            p_p,p_c = get_agg(p_s,p_e,[c_app,c_pla_n,c_med]), get_agg(c_s,c_e,[c_app,c_pla_n,c_med])
                            fs = p_c[(p_c[c_app]==pla_app)&(p_c[c_pla_n]==sel_p)].merge(p_p[(p_p[c_app]==pla_app)&(p_p[c_pla_n]==sel_p)], on=[c_app,c_pla_n,c_med], how='outer', suffixes=('_B','_A')).fillna(0)
                            tr = pd.DataFrame([{c_med:'🔥 합계',c_req+'_B':fs[c_req+'_B'].sum(),c_req+'_A':fs[c_req+'_A'].sum(),c_imp+'_B':fs[c_imp+'_B'].sum(),c_imp+'_A':fs[c_imp+'_A'].sum(),c_rev+'_B':fs[c_rev+'_B'].sum(),c_rev+'_A':fs[c_rev+'_A'].sum()}])
                            st.write(make_ordered_table(pd.concat([tr,fs.sort_values(c_rev+'_B',ascending=False)],ignore_index=True),c_med,c_req,c_rev,c_imp).to_html(escape=False,index=False), unsafe_allow_html=True)

                st.divider()
                def render_diagnose(target_df, title):
                    st.markdown(f"<div class='section-title'>{title}</div>", unsafe_allow_html=True)
                    table = make_ordered_table(target_df,c_app,c_req,c_rev,c_imp)
                    reports = [generate_complex_report(row['항목명'],df,p_s,p_e,c_s,c_e,c_pla_id,c_pla_n,c_rev,row,c_med) for _,row in table.iterrows()]
                    excel_table = table.copy(); excel_table['AI 복합 진단 리포트'] = reports
                    col1,col2 = st.columns([1,4])
                    with col1: st.download_button(f"📥 {title} 엑셀", to_excel_with_format(clean_df_for_excel(excel_table)), f"{title}.xlsx")
                    with col2:
                        with st.expander(f"📝 {title} 상세 사유"):
                            for idx,r in enumerate(reports):
                                if r!="변동 없음":
                                    style="plus" if "상승" in r else ""
                                    st.markdown(f"<div class='report-card {style}'><b>● {table.iloc[idx]['항목명']}</b><br/>{r.replace(chr(10),'<br/>')}</div>", unsafe_allow_html=True)
                    st.write(table.to_html(escape=False,index=False), unsafe_allow_html=True)

                render_diagnose(base.sort_values('rev_diff').head(30), "📉 매출 하락 Top 30")
                render_diagnose(base.sort_values('rev_diff',ascending=False).head(30), "📈 매출 상승 Top 30")
        except Exception as e: st.error(f"⚠️ 오류: {e}")
