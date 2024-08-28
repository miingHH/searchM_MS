__version__ = "0.1"
app_name = "3M_Report"

import streamlit as st
import pandas as pd
import os

st.set_page_config(layout='centered', page_title=f'{app_name} {__version__}')
ss = st.session_state
if 'debug' not in ss: ss['debug'] = {}

header1 = st.empty()  # for errors / messages
header2 = st.empty()  # for errors / messages
header3 = st.empty()  # for errors / messages

def ui_spacer(n=2, line=False, next_n=0):
    for _ in range(n):
        st.write('')
    if line:
        st.tabs([' '])
    for _ in range(next_n):
        st.write('')

def ui_info():
    st.markdown(f"""
    # {app_name}
    version {__version__}
    """)

def format_number(value):
    if pd.isna(value) or value == '':
        return ''
    return '{:,.0f}'.format(float(value)).replace(',', '.')

predefined_row_order = [
    "P&E_스카치디스펜서", "P&E_스카치", "P&E_포스트잇", "P&E_플래그", "P&E_가위", "P&E_순간접착제",
    "P&E_이젤패드", "SBR_P&E_스카치", "SBR_P&E_스카치디스펜서", "SBR_P&E_포스트잇", "P&E_스카치_파워링크",
    "P&E_포스트잇_파워링크", "CWB_후투로", "CWB_넥스케어&패치", "CWB_넥스케어_물방쿠/하이드로", "CWB_마스크",
    "3M_CWB_이어플러그", "3M_CWB_문닫힘방지", "3M_CWB_콘센트안전커버", "SBR_CWB_후투로", "SBR_CWB_넥스케어",
    "CWB_후투로_파워링크", "CWB_넥스케어_파워링크", "FUTURO_2405"
]
alternative_predefined_row_order = [
    "CWB_안전장갑(®)", "CWB_유아안전용품(®)", "CWB_이어플러그(®)", "CWB_블레미쉬(®)", "CWB_후투로(®)", "CWB_후투로_컴포트핏(®)",
    "CWB_후투로_컴포트핏(®) 매출 최적화", "CWB_치실", "CWB_밴드&물집방지(®)", "CWB_안전장갑(D)", "CWB_이어플러그(D)", 
    "PnE_스카치_테이프(®)", "PnE_스카치_집중권장", "PnE_스카치_쿠팡전용", "PnE_가위(®)", "PnE_스카치_원터치디스펜서(®)", 
    "PnE_포스트잇(®)", "PnE_포스트잇_집중권장쿠팡", "PnE_포스트잇_플래그(®)", "PnE_주방가위(D)", "PnE_접착제(D)", 
    "PnE_사무_테이프(D)", "PnE_사무_포스트잇(D)", "PnE_AI스마트광고(®)"
]

def process_money_files(uploaded_files, save_directory):
    combined_missing_campaign_name_df = pd.DataFrame()

    for uploaded_file in uploaded_files:
        try:
            # Read the uploaded Excel file
            df = pd.read_excel(uploaded_file, header=None)

            # Identify which row has the column names
            if '캠페인 이름' in df.iloc[0].values:
                header_row = 0
            elif '캠페인 이름' in df.iloc[1].values:
                header_row = 1
            else:
                # Append the whole DataFrame to combined_missing_campaign_name_df
                combined_missing_campaign_name_df = pd.concat([combined_missing_campaign_name_df, df], ignore_index=True)
                continue

            # Set the header row
            df.columns = df.iloc[header_row]
            df = df.drop([header_row])

            # Reset the index
            df = df.reset_index(drop=True)

            # Identify columns B, D, and E by name
            column_b = '캠페인 이름'
            column_d = '상태'
            column_e = '하루예산'

            # Select only columns B, D, and E
            df = df[[column_b, column_d, column_e]]

            # Remove '원' and convert '하루예산' column to numeric, handling errors
            df[column_e] = df[column_e].str.replace('원', '').str.replace(',', '')
            df[column_e] = pd.to_numeric(df[column_e], errors='coerce')

            # Create a DataFrame with predefined row order and empty E column for first sheet
            new_df = pd.DataFrame({column_b: predefined_row_order, column_e: ['']*len(predefined_row_order)})

            # Fill new_df with the values from df if they exist and if D column is '노출가능'
            for value in predefined_row_order:
                if value in df[column_b].values:
                    row = df[df[column_b] == value].iloc[0]
                    if row[column_d] == '노출가능':
                        new_df.loc[new_df[column_b] == value, column_e] = row[column_e]

            # Format the '하루예산' column
            new_df[column_e] = new_df[column_e].apply(format_number)

        except Exception as e:
            st.error(f"Error processing file {uploaded_file.name}: {e}")

    if not combined_missing_campaign_name_df.empty:
        combined_missing_campaign_name_df.columns = combined_missing_campaign_name_df.iloc[0]  # Set the first row as header
        combined_missing_campaign_name_df = combined_missing_campaign_name_df.drop([0]).reset_index(drop=True)  # Drop the header row
        
        # Rearrange columns according to the specified order
        arranged_df = pd.DataFrame({
            '캠페인 이름': combined_missing_campaign_name_df.iloc[:, 0],
            '하루예산': combined_missing_campaign_name_df.iloc[:, 2],
            '광고On/Off': combined_missing_campaign_name_df.iloc[:, 1],
            '목표수익률': combined_missing_campaign_name_df.iloc[:, 4],
            'Remark': combined_missing_campaign_name_df.iloc[:, 3]
        })

    else:
        arranged_df = pd.DataFrame(columns=['캠페인 이름', '하루예산', '광고On/Off', '목표수익률', 'Remark'])

    # Create a DataFrame for the second sheet based on alternative_predefined_row_order
    filtered_missing_df = pd.DataFrame({'캠페인 이름': alternative_predefined_row_order, '하루예산': ['']*len(alternative_predefined_row_order)})

    for value in alternative_predefined_row_order:
        if not combined_missing_campaign_name_df.empty and value in combined_missing_campaign_name_df.iloc[:, 0].values:
            filtered_missing_df.loc[filtered_missing_df['캠페인 이름'] == value, '하루예산'] = combined_missing_campaign_name_df[combined_missing_campaign_name_df.iloc[:, 0] == value].iloc[0, 2]
            filtered_missing_df.loc[filtered_missing_df['캠페인 이름'] == value, '광고On/Off'] = combined_missing_campaign_name_df[combined_missing_campaign_name_df.iloc[:, 0] == value].iloc[0, 1]
            filtered_missing_df.loc[filtered_missing_df['캠페인 이름'] == value, '목표수익률'] = combined_missing_campaign_name_df[combined_missing_campaign_name_df.iloc[:, 0] == value].iloc[0, 4]
            filtered_missing_df.loc[filtered_missing_df['캠페인 이름'] == value, 'Remark'] = combined_missing_campaign_name_df[combined_missing_campaign_name_df.iloc[:, 0] == value].iloc[0, 3]

    # Define the save path
    save_path = os.path.join(save_directory, "3M_네이버 광고비 광고상품 관리 현황.xlsx")
    with pd.ExcelWriter(save_path) as writer:
        new_df.to_excel(writer, sheet_name='Sheet1', index=False)
        filtered_missing_df.to_excel(writer, sheet_name='Sheet2', index=False)

    return save_path
    
def process_weekly_files(uploaded_files, save_directory):
    combined_df1 = pd.DataFrame(columns=[
        '캠페인', '광고그룹', '소재', '노출수', '클릭수', '총비용(VAT포함,원)', '전환수', '전환매출액(원)'
    ])
    combined_df2 = pd.DataFrame(columns=[
        '캠페인', '광고그룹', '키워드', '노출수', '클릭수', '총비용(VAT포함,원)', '전환수', '전환매출액(원)'
    ])
    combined_df3 = pd.DataFrame(columns=[
        '캠페인', '광고집행 옵션ID', '노출수', '클릭수', '광고비', '총 주문수(14일)', '총 전환매출액(14일)'
    ])
    
    for uploaded_file in uploaded_files:
        try:
            # Read the uploaded Excel file and detect header row
            df = pd.read_excel(uploaded_file, header=None)
            
            # Check both the first and second rows for headers
            if '캠페인유형' in df.iloc[0].values:
                header_row = 0
            elif '캠페인유형' in df.iloc[1].values:
                header_row = 1
            else:
                header_row = 0 if any(df.iloc[0].str.contains('캠페인명|광고집행 옵션ID', na=False)) else 1
            
            df.columns = df.iloc[header_row]
            df = df.drop([header_row])
            
            if '캠페인유형' in df.columns:
                # Filter rows based on '캠페인유형'
                df = df[df['캠페인유형'].isin(['파워링크', '쇼핑검색'])]

                if '소재' in df.columns:
                    # Aggregate rows by '캠페인', '광고그룹', '소재' for combined_df1
                    agg_df1 = df.groupby(['캠페인', '광고그룹', '소재'], as_index=False).agg({
                        '노출수': 'sum',
                        '클릭수': 'sum',
                        '총비용(VAT포함,원)': 'sum',
                        '전환수': 'sum',
                        '전환매출액(원)': 'sum'
                    })
                    combined_df1 = pd.concat([combined_df1, agg_df1], ignore_index=True)
                elif '검색어' in df.columns:
                    # Aggregate rows by '캠페인', '광고그룹', '검색어' for combined_df2
                    agg_df2 = df.groupby(['캠페인', '광고그룹', '검색어'], as_index=False).agg({
                        '노출수': 'sum',
                        '클릭수': 'sum',
                        '총비용(VAT포함,원)': 'sum',
                        '전환수': 'sum',
                        '전환매출액(원)': 'sum'
                    }).rename(columns={'검색어': '키워드'})
                    combined_df2 = pd.concat([combined_df2, agg_df2], ignore_index=True)
            else:
                # For files without '캠페인유형' column
                if '캠페인명' in df.columns and '광고집행 옵션ID' in df.columns:
                    # Filter rows based on '캠페인명' containing 'CWB' or 'PnE'
                    df = df[df['캠페인명'].str.contains('CWB|PnE', na=False)]

                    # Aggregate rows by '캠페인명', '광고집행 옵션ID' for combined_df3
                    agg_df3 = df.groupby(['캠페인명', '광고집행 옵션ID'], as_index=False).agg({
                        '노출수': 'sum',
                        '클릭수': 'sum',
                        '광고비': 'sum',
                        '총 주문수(14일)': 'sum',
                        '총 전환매출액(14일)': 'sum'
                    }).rename(columns={'캠페인명': '캠페인'})
                    combined_df3 = pd.concat([combined_df3, agg_df3], ignore_index=True)
                else:
                    st.error(f"File {uploaded_file.name} does not contain required columns.")
        except Exception as e:
            st.error(f"Error processing file {uploaded_file.name}: {e}")
    
    # Save the combined dataframes to an Excel file with three sheets
    save_path = os.path.join(save_directory, "3M_주간 변환 파일.xlsx")
    with pd.ExcelWriter(save_path) as writer:
        combined_df1.to_excel(writer, sheet_name='Sheet1', index=False)
        combined_df2.to_excel(writer, sheet_name='Sheet2', index=False)
        combined_df3.to_excel(writer, sheet_name='Sheet3', index=False)
    
    return save_path

def ui_excel_file():
    st.write('##### 저장 경로 설정')
    save_directory = st.text_input("Save Directory for 3M", value=os.path.expanduser("~"))  # Default는 Home
    st.write("---")

    st.write('### 3M 주간 변환 파일 처리')
    campaign_files = st.file_uploader("raw 파일을 전부 업로드하세요", key='주간버튼', type=["xlsx"], accept_multiple_files=True)
    
    if st.button('파일 생성', key='주간변환'):
        if campaign_files:
            save_path = process_weekly_files(campaign_files, save_directory)
            st.success(f"결과 파일 saved as: {save_path}")
            st.dataframe(pd.read_excel(save_path, sheet_name='Sheet1'))
        else:
            st.warning("raw 파일을 업로드하세요.")
    st.write("---")

    st.write('### 3M 광고비/광고상품관리 처리')
    uploaded_files = st.file_uploader("raw 파일을 전부 업로드하세요", key='광고비버튼', type=["xlsx"], accept_multiple_files=True)
    
    if st.button('파일 생성', key='광고비/상품관리'):
        if uploaded_files:
            save_path = process_money_files(uploaded_files, save_directory)
            st.success(f"결과 파일 saved as: {save_path}")
            st.dataframe(pd.read_excel(save_path, sheet_name='Sheet1'))
        else:
            st.warning("raw 파일을 업로드하세요.")
    st.write("---")

def ui_output():
    output = ss.get('output', '')
    st.markdown(output)

def b_clear():
    if st.button('clear output'):
        ss['output'] = ''

with st.sidebar:
    ui_info()
    ui_spacer(2)

ui_excel_file()
ui_output()