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
# 🛠️ 공통 유틸리티
# ==========================================
def to_excel_unified(df, sheet_name="통합_수주업로드"):
    numeric_cols = ['수량', '단가', 'Total Amount']
    for col in numeric_cols:
        if col in df.columns:
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
    # (Tesco 제품 매핑은 필요 시 기존 딕셔너리로 교체하세요)
    FULL_PRODUCT_MAP = {8809020342310: 'ME90521CLA', 8809020342211: 'ME90521CLL', 8809020342419: 'ME90521CLS'} 
    
    try:
        if uploaded_file.name.endswith('.csv'):
            content = uploaded_file.getvalue()
            try: text = content.decode('utf-8-sig')
            except: text = content.decode('cp949')
            all_rows = list(csv.reader(io.StringIO(text)))
        else:
            all_rows = pd.read_excel(uploaded_file, header=None).fillna('').astype(str).values.tolist()

        parsed_data = []
        for row in all_rows:
            row_strs = [str(x).strip() for x in row]
            if '상품코드' in row_strs and '낱개수량' in row_strs:
                col_map = {k: row_strs.index(v) for k, v in {'상품명':'상품명','상품코드':'상품코드','수량':'낱개수량','단가':'낱개당 단가','납품처':'납품처','납품일자':'납품일자'}.items() if v in row_strs}
                continue
            if 'col_map' not in locals() or not col_map: continue
            
            try:
                barcode = int(re.sub(r'[^\d]', '', row_strs[col_map['상품코드']]))
                parsed_data.append({
                    '구분': '0', '수주날짜': today_str, '발주처': 'Tesco', '발주코드': '81020000', '배송코드': '81020000',
                    '납품일자': pd.to_datetime(row_strs[col_map['납품일자']], errors='coerce').strftime('%Y%m%d') if '납품일자' in col_map else today_str,
                    '배송처': row_strs[col_map['납품처']] if '납품처' in col_map else '',
                    'ME코드': FULL_PRODUCT_MAP.get(barcode, str(barcode)),
                    '상품명': row_strs[col_map['상품명']] if '상품명' in col_map else '',
                    '수량': extract_num(row_strs[col_map['수량']]) if '수량' in col_map else 0,
                    '단가': extract_num(row_strs[col_map['단가']]) if '단가' in col_map else 0
                })
            except: pass
        return pd.DataFrame(parsed_data)
    except Exception as e:
        return pd.DataFrame()

# ==========================================
# 🟡 [로직 2] 이마트 처리
# ==========================================
def run_emart_logic(uploaded_file, prod_df):
    try:
        if uploaded_file.name.endswith('.csv'):
            try: raw_df = pd.read_csv(uploaded_file, encoding='utf-8-sig')
            except: raw_df = pd.read_csv(uploaded_file, encoding='cp949')
        else: raw_df = pd.read_excel(uploaded_file)

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
        
        if prod_df is not None and not prod_df.empty and '바코드' in prod_df.columns:
            merged = pd.merge(raw_df, prod_df, left_on='상품코드', right_on='바코드', how='left')
            merged['ME코드'] = merged.get('상품코드(기획)', merged['상품코드']).fillna(merged['상품코드'])
            merged['상품명'] = merged.get('상품명(기획)', merged.get('상품명', '')).fillna(merged.get('상품명', ''))
        else:
            merged = raw_df.copy()
            merged['ME코드'] = merged['상품코드']

        merged['구분'], merged['수주날짜'], merged['발주코드'] = '0', today_str, '81010000'
        merged['단가'] = pd.to_numeric(merged.get('발주원가', 0), errors='coerce').fillna(0)
        merged['수량'] = pd.to_numeric(merged.get('수량', 0), errors='coerce').fillna(0)
        
        date_col = next((c for c in ['센터입하일자', '센터입하일', '점입점일자'] if c in merged.columns), '')
        merged['납품일자'] = merged[date_col].astype(str).str.replace(r'[^0-9]', '', regex=True) if date_col else today_str
        merged['배송처'] = merged['배송코드'] 
        return merged
    except Exception as e:
        return pd.DataFrame()

# ==========================================
# 🟢 [로직 3] 롯데마트 처리
# ==========================================
def run_lotte_logic(uploaded_file, lotte_prod_df):
    try:
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
                curr_center = re.sub(r'상온센타|상온센터|센타', '센터', r[5]) if len(r)>5 else ""
                curr_date = re.sub(r'[^0-9]', '', r[7]) if len(r)>7 else today_str
                continue
            if len(r) > 1 and r[1].startswith('880'):
                qty = int(extract_num(r[6] if len(r)>6 else 0)) * (int(extract_num(r[5] if len(r)>5 else 1)) or 1)
                parsed_list.append({
                    '바코드': r[1].replace('.0', ''), '상품명_원본': r[2] if len(r)>2 else '', 
                    '수량': qty, '단가': extract_num(r[7] if len(r)>7 else 0), 
                    '납품일자': curr_date, '원본_배송처': curr_center
                })
        
        df = pd.DataFrame(parsed_list)
        if df.empty: return df
        
        # 마스터 파일 매핑 로직 (안전성 강화)
        if lotte_prod_df is not None and not lotte_prod_df.empty and '바코드' in lotte_prod_df.columns:
            df['바코드'] = df['바코드'].astype(str)
            lotte_prod_df['바코드'] = lotte_prod_df['바코드'].astype(str)
            df = pd.merge(df, lotte_prod_df, on='바코드', how='left')
            df['ME코드'] = df.get('ME코드', df['바코드']).fillna(df['바코드'])
            df['상품명'] = df.get('마스터_품명', df['상품명_원본']).fillna(df['상품명_원본'])
        else:
            df['ME코드'] = df['바코드']
            df['상품명'] = df['상품명_원본']

        df[['배송코드', '배송처']] = df['원본_배송처'].apply(lambda x: pd.Series(get_lotte_delivery_info(x)))
        df['구분'], df['수주날짜'], df['발주처'], df['발주코드'] = '0', today_str, '롯데마트', '81030000'
        return df
    except Exception as e:
        return pd.DataFrame()

# ==========================================
# 🚀 마스터 파일 로드 (초강력 방어 로직 적용)
# ==========================================
with st.sidebar:
    @st.cache_data
    def load_masters():
        emart_m, lotte_m = None, None
        
        # 1. 이마트 마스터 파일 안전하게 로드
        e_files = ["NEW 이마트 서식파일_20260420납품.xlsx", "NEW 이마트 트레이더스(한익스점포확인)_260327납품(평택9여주0대구4).xlsx", "NEW 노브랜드_20260409납품.xlsx"]
        e_list = []
        for f in e_files:
            if os.path.exists(f):
                try:
                    xls = pd.ExcelFile(f)
                    target_sheet = xls.sheet_names[0]
                    for s in xls.sheet_names:
                        if any(k in s for k in ['제품', '상품', '단가']):
                            target_sheet = s; break
                            
                    d = pd.read_excel(xls, sheet_name=target_sheet)
                    d.columns = d.columns.astype(str).str.strip()
                    
                    if '바코드' in d.columns:
                        d['바코드'] = d['바코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
                        e_list.append(d)
                    elif '상품코드' in d.columns:
                        d['바코드'] = d['상품코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
                        e_list.append(d)
                except Exception:
                    continue # 파일 구조가 달라도 절대 에러 내지 않고 다음 파일로 넘어감
                    
        if e_list:
            emart_m = pd.concat(e_list, ignore_index=True)
            if '바코드' in emart_m.columns:
                emart_m = emart_m.drop_duplicates(subset=['바코드'])

        # 2. 롯데마트 마스터 파일 안전하게 로드
        l_file = "2022 롯데마트 서식파일 260417납품.xlsx"
        if os.path.exists(l_file):
            try:
                xls = pd.ExcelFile(l_file)
                df_map = pd.read_excel(xls, sheet_name=0).astype(str)
                df_price = pd.read_excel(xls, sheet_name=1).astype(str)
                
                barcode_cols = [c for c in df_map.columns if '바코드' in str(c)]
                me_cols = [c for c in df_map.columns if 'ME' in str(c) or '기획' in str(c)]
                price_me_cols = [c for c in df_price.columns if 'ME' in str(c) or '상품코드' in str(c)]
                price_name_cols = [c for c in df_price.columns if '품명' in str(c) or '상품명' in str(c)]
                
                if barcode_cols and me_cols and price_me_cols and price_name_cols:
                    mapping_table = df_map[[barcode_cols[0], me_cols[0]]].copy()
                    mapping_table.columns = ['바코드', 'ME코드']
                    mapping_table['바코드'] = mapping_table['바코드'].str.replace('.0', '', regex=False).str.strip()
                    
                    price_table = df_price[[price_me_cols[0], price_name_cols[0]]].copy()
                    price_table.columns = ['ME코드', '마스터_품명']
                    
                    lotte_m = pd.merge(mapping_table, price_table, on='ME코드', how='left')
                    if '바코드' in lotte_m.columns:
                        lotte_m = lotte_m.drop_duplicates(subset=['바코드'])
            except Exception:
                pass # 롯데마트 파일이 이상해도 절대 멈추지 않음

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
    with st.spinner("🔄 데이터를 변환 중입니다..."):
        for f in uploaded_files:
            try:
                f.seek(0); sample = str(f.read(2000)); f.seek(0)
                if 'ORDERS' in sample: df = run_lotte_logic(f, lotte_master)
                elif '점포코드' in sample or '센터입하' in sample: df = run_emart_logic(f, emart_master)
                else: df = run_tesco_logic(f)
                
                if df is not None and not df.empty:
                    # 열 누락 방지 안전장치
                    for col in FINAL_COLUMNS + ['Total Amount']:
                        if col not in df.columns: df[col] = ""
                    all_dfs.append(df)
            except Exception as e:
                st.warning(f"⚠️ {f.name} 파일 처리에 실패하여 건너뜁니다.")

    if all_dfs:
        merged_df = pd.concat(all_dfs, ignore_index=True).fillna("")
        
        # 필요한 열이 다 있는지 최종 검사
        for col in FINAL_COLUMNS:
            if col not in merged_df.columns: merged_df[col] = ""
            
        # 동일 항목 수량 합산
        group_cols = ['구분', '수주날짜', '납품일자', '발주코드', '발주처', '배송코드', '배송처', 'ME코드', '상품명', '단가']
        merged_df['수량'] = pd.to_numeric(merged_df['수량'], errors='coerce').fillna(0)
        merged_df['단가'] = pd.to_numeric(merged_df['단가'], errors='coerce').fillna(0)
        
        final_df = merged_df.groupby(group_cols, as_index=False, dropna=False).agg({'수량': 'sum'})
        final_df['Total Amount'] = final_df['수량'] * final_df['단가']
        final_df = final_df[FINAL_COLUMNS]

        st.dataframe(final_df, use_container_width=True)
        st.download_button("📥 통합 결과 엑셀 다운로드", data=to_excel_unified(final_df), file_name=f"통합수주_{today_str}.xlsx", type="primary")
