import streamlit as st
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment
from io import BytesIO
import datetime

st.set_page_config(page_title="Keyword Master Pro", layout="wide")
st.markdown('<h1 style="color: #1E293B; text-align: center;">📈 Keyword Master Pro</h1>', unsafe_allow_html=True)

def analyze(uploaded, cat, start_m):
    files = sorted(uploaded, key=lambda x: x.name)
    y, m = int(start_m[:2]), int(start_m[2:])
    m_names = [f"{(y + (m+i-1)//12):02d}{(m+i-1)%12+1:02d}" for i in range(len(files))]
    kw_map = {}
    for idx, f in enumerate(files):
        df = pd.read_excel(f)
        for _, row in df.iterrows():
            if cat.strip() in str(row['대표 카테고리']):
                kw = str(row['키워드']).strip()
                if kw not in kw_map: kw_map[kw] = [0] * len(files)
                kw_map[kw][idx] = float(row['총 검색수']) if not pd.isna(row['총 검색수']) else 0
    res = {'사계절': [], '시즌': [], '성장': []}
    for kw, counts in kw_map.items():
        avg = sum(counts)/len(counts)
        if avg < 3000: continue
        res['사계절'].append([kw] + counts + [round(avg), "Gold" if avg >= 10000 else "Silver", "정상"])
    return res, m_names

col1, col2 = st.columns([1, 2])
with col1:
    st.subheader("⚙️ 설정")
    cat = st.text_input("카테고리명", value="실버용품")
    s_m = st.text_input("시작월(YYMM)", value="2501")
    uploaded = st.file_uploader("엑셀 파일 전부 선택", accept_multiple_files=True)
    btn = st.button("🚀 분석 시작")

if btn and uploaded:
    with st.spinner("분석 중..."):
        res, m_names = analyze(uploaded, cat, s_m)
        with col2:
            st.success("분석 완료!")
            o = BytesIO()
            with pd.ExcelWriter(o, engine='openpyxl') as w:
                for k, v in res.items(): pd.DataFrame(v, columns=["키워드"]+m_names+["평균","등급","비고"]).to_excel(w, sheet_name=k, index=False)
            st.download_button("📥 결과 다운로드", o.getvalue(), f"{cat}_분석결과.xlsx")