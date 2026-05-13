import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import re
import csv
from datetime import datetime

# ==========================================
# ⚙️ 페이지 및 기본 설정
# ==========================================
st.set_page_config(page_title="멘소래담 마트 통합 수주 자동화", page_icon="🏢", layout="wide")
today_str = datetime.today().strftime("%Y%m%d")

FINAL_COLUMNS = [
    '구분', '수주날짜', '납품일자', '발주코드', '발주처', '배송코드', '배송처', 
    'ME코드', '상품명', '수량', '단가', 'Total Amount'
]

# ==========================================
# 🛠️ 공통 유틸리티 (엑셀 스타일 및 숫자 추출)
# ==========================================
def to_excel_unified(df, sheet_name="통합_수주업로드"):
    numeric_cols = ['수량', '단가', 'Total Amount']
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook, worksheet = writer.book, writer.sheets[sheet_name]
        num_format = workbook.add_format({'num_format': '#,##0'})
        header_format = workbook.add_format({'bold': True, 'bg_color': '#F0F2F6', 'border': 1, 'align': 'center'})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        for col_idx, col_name in enumerate(df.columns):
            if col_name in numeric_cols: worksheet.set_column(col_idx, col_idx, 12, num_format)
            elif col_name in ['상품명', '배송처']: worksheet.set_column(col_idx, col_idx, 30)
            else: worksheet.set_column(col_idx, col_idx, 15)
    return output.getvalue()

def extract_num(val):
    s = str(val).split('(')[0]
    s = re.sub(r'[^\d.]', '', s)
    try: return float(s) if s else 0.0
    except: return 0.0

# ==========================================
# 🔴 [로직 1] TESCO 처리
# ==========================================
def run_tesco_logic(uploaded_file):
    # (Tesco 제품 매핑은 사용자님의 기존 딕셔너리 리스트를 여기에 그대로 넣어주세요)
    FULL_PRODUCT_MAP = {8809020342310: 'ME90521CLA', 8809020342211: 'ME90521CLL', 8809020342419: 'ME90521CLS'} # 예시
    
    if uploaded_file.name.endswith('.csv'):
        content = uploaded_file.getvalue()
        try: text = content.decode('utf-8-sig')
        except: text = content.decode('cp949')
        all_rows = list(csv.reader(io.StringIO(text)))
    else:
        all_rows = pd.read_excel(uploaded_file, header=None).fillna('').astype(str).values.tolist()

    parsed_data = []
    col_map = {}
    for row in all_rows:
        row_strs = [str(x).strip() for x in row]
        if '상품코드' in row_strs and '낱개수량' in row_strs:
            col_map = {k: row_strs.index(v) for k, v in {'상품명':'상품명','상품코드':'상품코드','수량':'낱개수량','단가':'낱개당 단가','금액':'발주금액','납품처':'납품처','납품일자':'납품일자'}.items() if v in row_strs}
            continue
        if not col_map: continue
        try:
            barcode = int(re.sub(r'[^\d]', '', row_strs[col_map['상품코드']]))
            if barcode in FULL_PRODUCT_MAP:
                parsed_data.append({
                    '구분': '0', '수주날짜': today_str, '발주처': 'Tesco', '발주코드': '81020000', '배송코드': '81020000',
                    '납품일자': pd.to_datetime(row_strs[col_map['납품일자']], errors='coerce').strftime('%Y%m%d'),
                    '배송처': row_strs[col_map['납품처']], 'ME코드': FULL_PRODUCT_MAP[barcode], '상품명': row_strs[col_map['상품명']],
                    '수량': extract_num(row_strs[col_map['수량']]), '단가': extract_num(row_strs[col_map['단가']])
                })
        except: pass
    return pd.DataFrame(parsed_data)

# ==========================================
# 🟡 [로직 2] 이마트 처리
# ==========================================
def run_emart_logic(uploaded_file, prod_df):
    try:
        if uploaded_file.name.endswith('.csv'):
            try: raw_df = pd.read_csv(uploaded_file, encoding='utf-8-sig')
            except: raw_df = pd.read_csv(uploaded_file, encoding='cp949')
        else: raw_df = pd.read_excel(uploaded_file)
    except: return pd.DataFrame()

    if '점포코드' not in raw_df.columns: return pd.DataFrame()

    raw_df['점포코드'] = pd.to_numeric(raw_df['점포코드'], errors='coerce').fillna(0).astype(int)
    raw_df['센터코드'] = raw_df['센터코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
    
    emart_map_dict = {
        'E-mart': {'9110': '81010902', '9120': '81010905', '9100': '81010903'},
        'E-mart(TRD)': {'9150': '81033036', '9102': '89011174', '9120': '81011012'},
        'E-mart(노브랜드)': {'9102': '89011175', '9130': '81010904', '9120': '81010968', '9110': '81010969'}
    }
    def process_row(row):
        code, center = row['점포코드'], str(row['센터코드'])
        cust = 'E-mart' if (1000 <= code <= 1999 or code >= 9000) else ('E-mart(TRD)' if 2000 <= code <= 2999 else 'E-mart(노브랜드)')
        return pd.Series([cust, emart_map_dict.get(cust, {}).get(center, center)])

    raw_df[['발주처', '배송코드']] = raw_df.apply(process_row, axis=1)
    raw_df['상품코드'] = raw_df['상품코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
    
    if prod_df is not None:
        merged = pd.merge(raw_df, prod_df, left_on='상품코드', right_on='바코드', how='left')
        merged['ME코드'] = merged['상품코드(기획)'].fillna(merged['상품코드'])
        name_col = '상품명(기획)' if '상품명(기획)' in prod_df.columns else '상품명'
        merged['상품명'] = merged[name_col].fillna(merged.get('상품명', ''))
    else:
        merged = raw_df.copy(); merged['ME코드'] = merged['상품코드']

    merged['구분'], merged['수주날짜'], merged['발주코드'] = '0', today_str, '81010000'
    merged['단가'] = pd.to_numeric(merged.get('발주원가', 0), errors='coerce').fillna(0)
    merged['수량'] = pd.to_numeric(merged.get('수량', 0), errors='coerce').fillna(0)
    date_col = next((c for c in ['센터입하일자', '센터입하일', '점입점일자'] if c in merged.columns), '')
    merged['납품일자'] = merged[date_col].astype(str).str.replace(r'[^0-9]', '', regex=True) if date_col else today_str
    merged['배송처'] = merged['배송코드'] # 실제 이름 매핑 필요시 추가
    return merged

# ==========================================
# 🟢 [로직 3] 롯데마트 처리 (ME코드 매핑 강화)
# ==========================================
def run_lotte_logic(uploaded_file, lotte_prod_df):
    def get_lotte_delivery_info(center_name):
        name = str(center_name)
        if '오산' in name: return '81030907', '롯데 오산상온센터'
        if '김해' in name: return '81030908', '롯데 김해상온센터'
        return '81030000', name

    if uploaded_file.name.endswith('.csv'): 
        try: df_edi = pd.read_csv(uploaded_file, header=None, encoding='utf-8-sig')
        except: uploaded_file.seek(0); df_edi = pd.read_csv(uploaded_file, header=None, encoding='cp949')
    else: df_edi = pd.read_excel(uploaded_file, header=None)

    parsed_list, curr_center, curr_date = [], "", ""
    for _, row in df_edi.dropna(how='all').iterrows():
        r = [str(x).strip() for x in row.tolist()]
        if r[0] == 'ORDERS':
            curr_center = re.sub(r'상온센타|상온센터|센타', '센터', r[5])
            curr_date = re.sub(r'[^0-9]', '', r[7]); continue
        if len(r) > 1 and r[1].startswith('880'):
            qty = int(extract_num(r[6])) * (int(extract_num(r[5])) or 1)
            parsed_list.append({
                '바코드': r[1].replace('.0', ''), '상품명_원본': r[2], '수량': qty, 
                '단가': extract_num(r[7]), '납품일자': curr_date, '원본_배송처': curr_center
            })
    
    df = pd.DataFrame(parsed_list)
    if df.empty: return df
    
    # ⭐ 핵심: 롯데 마스터 파일과 매핑 (바코드를 ME코드로 교체)
    if lotte_prod_df is not None:
        # 데이터 타입을 문자열로 통일하여 매핑 정확도 향상
        df['바코드'] = df['바코드'].astype(str)
        lotte_prod_df['바코드'] = lotte_prod_df['바코드'].astype(str)
        
        df = pd.merge(df, lotte_prod_df, on='바코드', how='left')
        # 매핑된 ME코드가 있으면 쓰고, 없으면 바코드를 그대로 씀 (보통은 매핑되어야 함)
        df['ME코드'] = df['ME코드'].fillna(df['바코드'])
        df['상품명'] = df['마스터_품명'].fillna(df['상품명_원본'])
    else:
        df['ME코드'] = df['바코드']
        df['상품명'] = df['상품명_원본']

    df[['배송코드', '배송처']] = df['원본_배송처'].apply(lambda x: pd.Series(get_lotte_delivery_info(x)))
    df['구분'], df['수주날짜'], df['발주처'], df['발주코드'] = '0', today_str, '롯데마트', '81030000'
    return df

# ==========================================
# 🚀 마스터 파일 로드 (가장 중요)
# ==========================================
with st.sidebar:
    @st.cache_data
    def load_masters():
        emart_m, lotte_m = None, None
        
        # 1. 이마트 마스터 (기존 로직 유지)
        e_files = ["NEW 이마트 서식파일_20260420납품.xlsx", "NEW 이마트 트레이더스(한익스점포확인)_260327납품(평택9여주0대구4).xlsx", "NEW 노브랜드_20260409납품.xlsx"]
        e_list = [pd.read_excel(f).assign(바코드=lambda x: x['바코드'].astype(str).str.replace('.0','')) for f in e_files if os.path.exists(f)]
        if e_list: emart_m = pd.concat(e_list).drop_duplicates(subset=['바코드'])

        # 2. 롯데마트 마스터 (강력한 컬럼 찾기 로직 추가)
        l_file = "2022 롯데마트 서식파일 260417납품.xlsx"
        if os.path.exists(l_file):
            try:
                xls = pd.ExcelFile(l_file)
                df_map = pd.read_excel(xls, sheet_name=0).astype(str)
                df_price = pd.read_excel(xls, sheet_name=1).astype(str)
                
                # '바코드'와 'ME' 혹은 '상품코드'가 포함된 열을 자동으로 찾음
                col_barcode = [c for c in df_map.columns if '바코드' in str(c)][0]
                col_me = [c for c in df_map.columns if 'ME' in str(c) or '기획' in str(c)][0]
                
                # 맵핑 테이블 생성
                mapping_table = df_map[[col_barcode, col_me]].copy()
                mapping_table.columns = ['바코드', 'ME코드']
                mapping_table['바코드'] = mapping_table['바코드'].str.replace('.0', '', regex=False).str.strip()
                
                # 품명 테이블 생성
                col_me_price = [c for c in df_price.columns if 'ME' in str(c) or '상품코드' in str(c)][0]
                col_name = [c for c in df_price.columns if '품명' in str(c) or '상품명' in str(c)][0]
                price_table = df_price[[col_me_price, col_name]].copy()
                price_table.columns = ['ME코드', '마스터_품명']
                
                lotte_m = pd.merge(mapping_table, price_table, on='ME코드', how='left').drop_duplicates(subset=['바코드'])
            except Exception as e:
                st.sidebar.error(f"롯데 마스터 로드 실패: {e}")
        return emart_m, lotte_m

    emart_master, lotte_master = load_masters()
    if emart_master is not None: st.success("✅ 이마트 마스터 로드 완료")
    if lotte_master is not None: st.success("✅ 롯데마트 마스터 로드 완료")

# ==========================================
# 📊 메인 실행 영역
# ==========================================
st.title("📦 통합 마트 수주 자동 변환")
uploaded_files = st.file_uploader("📂 발주서 파일들을 한꺼번에 올리세요", accept_multiple_files=True)

if uploaded_files:
    all_dfs = []
    for f in uploaded_files:
        f.seek(0); sample = str(f.read(2000)); f.seek(0)
        if 'ORDERS' in sample: df = run_lotte_logic(f, lotte_master)
        elif '점포코드' in sample or '센터입하' in sample: df = run_emart_logic(f, emart_master)
        else: df = run_tesco_logic(f)
        
        if not df.empty:
            all_dfs.append(df[df.columns.intersection(FINAL_COLUMNS + ['Total Amount'])])

    if all_dfs:
        merged_df = pd.concat(all_dfs, ignore_index=True).fillna("")
        # 동일 항목 수량 합산
        group_cols = ['구분', '수주날짜', '납품일자', '발주코드', '발주처', '배송코드', '배송처', 'ME코드', '상품명', '단가']
        final_df = merged_df.groupby(group_cols, as_index=False).agg({'수량': 'sum'})
        final_df['Total Amount'] = final_df['수량'] * final_df['단가']
        final_df = final_df[FINAL_COLUMNS]

        st.dataframe(final_df, use_container_width=True)
        st.download_button("📥 통합 결과 엑셀 다운로드", data=to_excel_unified(final_df), file_name=f"통합수주_{today_str}.xlsx", type="primary")
