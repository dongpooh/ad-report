import streamlit as st
import pandas as pd
from datetime import timedelta
import io
import re

# 1. 페이지 설정 및 디자인
st.set_page_config(page_title="애드팝콘 통합 분석기", layout="wide")

st.markdown("""
    <style>
    table { width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 25px; }
    th { background-color: #f0f2f6; text-align: center !important; padding: 10px !important; border: 1px solid #dee2e6; }
    td { padding: 8px !important; border: 1px solid #dee2e6; }
    .section-title { padding: 10px; background-color: #31333F; color: white; border-radius: 5px; margin: 20px 0 10px 0; font-weight: bold; }
    .report-card { background-color: #f8f9fa; border-left: 5px solid #ff4b4b; padding: 15px; margin-bottom: 10px; border-radius: 5px; font-size: 13px; line-height: 1.6; }
    .report-card.plus { border-left-color: #28a745; }
    .detail-box { padding: 20px; border: 1px solid #e6e9ef; border-radius: 10px; background-color: #fcfcfc; margin-top: 10px; }
    </style>
    """, unsafe_allow_html=True)

st.title("📊 애드팝콘 통합 분석기 (by. Pole)")

# --- [도구 함수] 엑셀 다운로드 서식 ---
def to_excel_with_format(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Report')
        workbook, worksheet = writer.book, writer.sheets['Report']
        money_fmt = workbook.add_format({'num_format': '$#,##0.00'}) 
        pct_fmt = workbook.add_format({'num_format': '0.0"%"'}) 
        wrap_fmt = workbook.add_format({'text_wrap': True})
        wrap_fmt.set_align('top') 
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
        clean_df[col] = clean_df[col].str.replace('▲', '', regex=False).str.replace('▼', '', regex=False)
        if any(k in col for k in ['%', '매출', 'eCPM', '요청', '노출']):
            clean_df[col] = clean_df[col].apply(lambda x: re.sub(r'[^0-9.\-]', '', str(x)))
            clean_df[col] = pd.to_numeric(clean_df[col].str.strip(), errors='coerce').fillna(0)
    return clean_df

# --- [AI 로직] 복합 원인 생성기 ---
def generate_complex_report(media_name, full_df, p_s, p_e, c_s, c_e, c_pla_id, c_pla_name, c_rev, row_data, c_med):
    try:
        rev_a = float(str(row_data['이전 매출']).replace('$','').replace(',',''))
        rev_b = float(str(row_data['현재 매출']).replace('$','').replace(',',''))
        rev_diff = rev_b - rev_a
        if rev_diff == 0: return "변동 없음"

        m_df = full_df[full_df['media_name'] == media_name].copy()
        mask_a = (m_df['report_date'] >= p_s) & (m_df['report_date'] <= p_e)
        mask_b = (m_df['report_date'] >= c_s) & (m_df['report_date'] <= c_e)
        
        m_df['pla_full'] = m_df[c_pla_name].astype(str) + " (" + m_df[c_pla_id].astype(str) + ")"
        p_a = m_df[mask_a].groupby('pla_full')[c_rev].sum()
        p_b = m_df[mask_b].groupby('pla_full')[c_rev].sum()
        p_diff = (p_b - p_a).fillna(p_b).fillna(-p_a)
        sorted_p_diff = p_diff.sort_values() if rev_diff < 0 else p_diff.sort_values(ascending=False)

        top_plas = []
        for p_info, p_val in sorted_p_diff.items():
            share = (p_val / rev_diff * 100) if rev_diff != 0 else 0
            if abs(share) > 15: top_plas.append(f"{p_info}[기여 {share:.1f}%]")

        ssp_a = m_df[mask_a].groupby(c_med)[c_rev].sum()
        ssp_b = m_df[mask_b].groupby(c_med)[c_rev].sum()
        ssp_diff = (ssp_b - ssp_a).fillna(ssp_b).fillna(-ssp_a)
        sorted_ssp_diff = ssp_diff.sort_values() if rev_diff < 0 else ssp_diff.sort_values(ascending=False)
        top_ssps = [s_name for s_name, s_val in sorted_ssp_diff.items() if abs(s_val) > abs(rev_diff) * 0.1]

        main_type = "미디에이션" if len(top_ssps) > 0 else "매체트래픽"
        report = f"【주요원인】 {main_type} ({', '.join(top_ssps[:2])})\n"
        report += f"【상세분석】 주요지면: {', '.join(top_plas[:2]) if top_plas else '전반적 변동'}\n"
        report += f"【수치요약】 매출 ${abs(rev_diff):,.2f} {'하락' if rev_diff < 0 else '상승'}."
        return report
    except: return "진단 불가"

# --- [도구 함수] 테이블 생성 ---
def format_arrow(curr, prev):
    if prev == 0: return "<span style='color:red'>▲ New</span>" if curr > 0 else "0.0%"
    diff = ((curr - prev) / prev) * 100
    color = "red" if diff > 0 else "blue"
    arrow = "▲" if diff > 0 else "▼"
    return f"<span style='color:{color}'>{arrow} {diff:.1f}%</span>"

def make_ordered_table(merged_df, name_col, c_req, c_rev, c_imp):
    res = pd.DataFrame()
    res['항목명'] = merged_df[name_col]
    for c, l in zip([c_req, c_imp, c_rev], ['요청', '노출', '매출']):
        fmt = '{:,.0f}' if l != '매출' else '${:,.2f}'
        res[f'이전 {l}'] = merged_df[f'{c}_A'].map(fmt.format)
        res[f'현재 {l}'] = merged_df[f'{c}_B'].map(fmt.format)
        res[f'{l} %'] = [format_arrow(b, a) for a, b in zip(merged_df[f'{c}_A'], merged_df[f'{c}_B'])]
    ea = (merged_df[c_rev+'_A'] / merged_df[c_imp+'_A'] * 1000).fillna(0)
    eb = (merged_df[c_rev+'_B'] / merged_df[c_imp+'_B'] * 1000).fillna(0)
    res['이전 eCPM'] = ea.map('${:,.2f}'.format); res['현재 eCPM'] = eb.map('${:,.2f}'.format)
    res['eCPM %'] = [format_arrow(b, a) for a, b in zip(ea, eb)]
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

# --- 메인 실행부 ---
data_type = st.sidebar.radio("📊 데이터 유형", ["SSP 리포트", "DSP 리포트"])
uploaded_files = st.file_uploader(f"{data_type} 업로드", type=['xlsx', 'csv'], accept_multiple_files=True)

if data_type == "SSP 리포트":
    c_med, c_req, c_imp, c_rev, c_app, c_pla_n, c_pla_id = 'thirdparty_name', 'request_value', 'impression_value', 'partner_revenue', 'media_name', 'placement_name', 'placement_id'
else:
    c_med, c_req, c_imp, c_rev, c_app, c_pla_n, c_pla_id = 'dsp_name', 'response_value', 'impression_value', 'partner_revenue', 'media_name', 'placement_name', 'placement_id'

if uploaded_files:
    try:
        df = load_and_merge_data(uploaded_files)
        all_dates = sorted([d for d in df['report_date'].unique() if pd.notnull(d)])
        mode = st.radio("🔍 분석 모드", ["DOD", "자유 기간", "MOM"], horizontal=True)
        p_s, p_e, c_s, c_e = None, None, None, None

        if mode == "DOD":
            t = st.selectbox("기준일", all_dates[::-1]); c_s = c_e = t; p_s = p_e = t - timedelta(days=1)
        elif mode == "자유 기간":
            c1, c2 = st.columns(2)
            with c1: pr = st.date_input("이전", [all_dates[0], all_dates[0]]); p_s, p_e = pr if len(pr)==2 else (None, None)
            with c2: cr = st.date_input("현재", [all_dates[-1], all_dates[-1]]); c_s, c_e = cr if len(cr)==2 else (None, None)
        elif mode == "MOM":
            df['m'] = pd.to_datetime(df['report_date']).dt.to_period('M')
            ms = sorted(df['m'].unique(), reverse=True)
            pm, cm = st.selectbox("이전 월", ms[1:] if len(ms)>1 else ms), st.selectbox("현재 월", ms)
            p_s, p_e, c_s, c_e = pm.start_time.date(), pm.end_time.date(), cm.start_time.date(), cm.end_time.date()

        if p_s and c_s:
            def get_agg(s, e, g):
                mask = (df['report_date'] >= s) & (df['report_date'] <= e)
                return df[mask].groupby(g).agg({c_req:'sum', c_imp:'sum', c_rev:'sum'}).reset_index()

            # 1. 통합 성과
            st.markdown(f"<div class='section-title'>🌐 {data_type} 통합 성과</div>", unsafe_allow_html=True)
            m_p, m_c = get_agg(p_s, p_e, [c_med]), get_agg(c_s, c_e, [c_med])
            m_f = m_c.merge(m_p, on=c_med, how='outer', suffixes=('_B', '_A')).fillna(0)
            st.write(make_ordered_table(m_f.sort_values(c_rev+'_B', ascending=False), c_med, c_req, c_rev, c_imp).to_html(escape=False, index=False), unsafe_allow_html=True)

            # 2. 매체별 요약
            st.markdown("<div class='section-title'>🏠 매체별 성과 요약</div>", unsafe_allow_html=True)
            a_p, a_c = get_agg(p_s, p_e, [c_app]), get_agg(c_s, c_e, [c_app])
            a_f = a_c.merge(a_p, on=c_app, how='outer', suffixes=('_B', '_A')).fillna(0)
            a_f['rev_diff'] = a_f[c_rev+'_B'] - a_f[c_rev+'_A']
            base = a_f[a_f[c_req+'_B'] >= 5000].sort_values(c_rev+'_B', ascending=False)
            search = st.multiselect("🔍 매체명 검색", options=base[c_app].unique())
            f_df = base[base[c_app].isin(search)] if search else base
            pg = st.number_input(f"페이지", 1, max(1, len(f_df)//20+1), 1)
            st.write(make_ordered_table(f_df.iloc[(pg-1)*20:pg*20], c_app, c_req, c_rev, c_imp).to_html(escape=False, index=False), unsafe_allow_html=True)

            # 3. 매체 상세 분석
            st.divider()
            st.markdown(f"<div class='section-title'>🎯 3. 매체 상세 분석 ({data_type}별)</div>", unsafe_allow_html=True)
            sel_app = st.selectbox("분석할 매체 선택", base[c_app].unique(), key="app_detail_sel")
            if sel_app:
                d_p, d_c = get_agg(p_s, p_e, [c_app, c_med]), get_agg(c_s, c_e, [c_app, c_med])
                d_f = d_c[d_c[c_app]==sel_app].merge(d_p[d_p[c_app]==sel_app], on=[c_app, c_med], how='outer', suffixes=('_B', '_A')).fillna(0)
                st.write(make_ordered_table(d_f.sort_values(c_rev+'_B', ascending=False), c_med, c_req, c_rev, c_imp).to_html(escape=False, index=False), unsafe_allow_html=True)

            # [순서변경] 4. 상세 지면 분석 (매체 상세 분석 바로 아래로 이동)
            if data_type == "SSP 리포트":
                st.markdown("<div class='section-title'>🏁 4. 상세 지면(Placement) 분석</div>", unsafe_allow_html=True)
                with st.container():
                    st.markdown("<div class='detail-box'>", unsafe_allow_html=True)
                    # 위에서 선택한 sel_app을 기본값으로 사용하거나 별도 선택
                    pla_app = st.selectbox("지면 분석용 매체 선택", ["선택 안함"] + list(base[c_app].unique()), key="pla_app_final")
                    if pla_app != "선택 안함":
                        p_l = sorted(df[df[c_app] == pla_app][c_pla_n].unique())
                        sel_p = st.selectbox("지면 선택", p_l, key="sel_p_final")
                        if sel_p:
                            p_p, p_c = get_agg(p_s, p_e, [c_app, c_pla_n, c_med]), get_agg(c_s, c_e, [c_app, c_pla_n, c_med])
                            fs = p_c[(p_c[c_app]==pla_app)&(p_c[c_pla_n]==sel_p)].merge(p_p[(p_p[c_app]==pla_app)&(p_p[c_pla_n]==sel_p)], on=[c_app, c_pla_n, c_med], how='outer', suffixes=('_B', '_A')).fillna(0)
                            tr = pd.DataFrame([{c_med: '🔥 지면 합계(Total)', c_req+'_B': fs[c_req+'_B'].sum(), c_req+'_A': fs[c_req+'_A'].sum(), c_imp+'_B': fs[c_imp+'_B'].sum(), c_imp+'_A': fs[c_imp+'_A'].sum(), c_rev+'_B': fs[c_rev+'_B'].sum(), c_rev+'_A': fs[c_rev+'_A'].sum()}])
                            st.write(make_ordered_table(pd.concat([tr, fs.sort_values(c_rev+'_B', ascending=False)], ignore_index=True), c_med, c_req, c_rev, c_imp).to_html(escape=False, index=False), unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)

            # 5. 하락/상승 진단 결과 (가장 아래로)
            st.divider()
            def render_diagnose(target_df, title):
                st.markdown(f"<div class='section-title'>{title}</div>", unsafe_allow_html=True)
                table = make_ordered_table(target_df, c_app, c_req, c_rev, c_imp)
                reports = [generate_complex_report(row['항목명'], df, p_s, p_e, c_s, c_e, c_pla_id, c_pla_n, c_rev, row, c_med) for _, row in table.iterrows()]
                excel_table = table.copy(); excel_table['AI 복합 진단 리포트'] = reports
                col1, col2 = st.columns([1, 4])
                with col1: st.download_button(f"📥 {title} 엑셀", to_excel_with_format(clean_df_for_excel(excel_table)), f"{title}.xlsx")
                with col2:
                    with st.expander(f"📝 {title} 상세 사유 보기 (클릭)"):
                        for idx, r in enumerate(reports):
                            if r != "변동 없음":
                                style = "plus" if "상승" in r else ""
                                st.markdown(f"<div class='report-card {style}'><b>● {table.iloc[idx]['항목명']}</b><br/>{r.replace(chr(10), '<br/>')}</div>", unsafe_allow_html=True)
                st.write(table.to_html(escape=False, index=False), unsafe_allow_html=True)

            render_diagnose(base.sort_values('rev_diff').head(30), "📉 5-1. 매출 하락 Top 30 진단")
            render_diagnose(base.sort_values('rev_diff', ascending=False).head(30), "📈 5-2. 매출 상승 Top 30 진단")

    except Exception as e: st.error(f"⚠️ 오류: {e}")