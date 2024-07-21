import streamlit as st
import pandas as pd
import msoffcrypto
from io import BytesIO
from datetime import datetime

def read_excel_file(uploaded_file, password=None, header_row=0):
    try:
        if password:
            # 비밀번호가 있는 경우
            decrypted = BytesIO()
            file = msoffcrypto.OfficeFile(uploaded_file)
            file.load_key(password=password)
            file.decrypt(decrypted)
            df = pd.read_excel(decrypted, engine='openpyxl', header=header_row)
        else:
            # 비밀번호가 없는 경우
            df = pd.read_excel(uploaded_file, engine='openpyxl', header=header_row)
        return df
    except Exception as e:
        st.error(f"파일을 읽는 중 오류 발생: {e}")
        return None

def main():
    st.title("스마트스토어 to CJ배송")

    today = datetime.today().strftime('%Y%m%d')

    st.header("주문내역 올리기")
    uploaded_file1 = st.file_uploader(" ", type=["xlsx", "xls"], key="upload1")

    if uploaded_file1 is not None:
        df1 = read_excel_file(uploaded_file1, password='1111', header_row=1)
        if df1 is not None:
            df_transformed1 = pd.DataFrame()
            df_transformed1["받는분성명"] = df1["수취인명"].apply(lambda x: x + '-' if len(x) == 1 else x)
            df_transformed1["받는분전화번호"] = df1["수취인연락처1"]
            df_transformed1["받는분기타연락처"] = df1["수취인연락처2"]
            df_transformed1["받는분주소(전체, 분할)"] = df1["통합배송지"]
            df_transformed1["품목명"] = "기타"
            df_transformed1["내품명"] = "기타"
            df_transformed1["배송메세지1"] = df1["배송메세지"]

            df_transformed1 = df_transformed1.drop_duplicates(subset=["받는분주소(전체, 분할)"])

            st.write(df_transformed1)

            output1 = BytesIO()
            with pd.ExcelWriter(output1, engine='xlsxwriter') as writer:
                df_transformed1.to_excel(writer, index=False)
            processed_data1 = output1.getvalue()

            st.download_button(
                label="cj택배용 엑셀 다운로드",
                data=processed_data1,
                file_name=f"{today}_택배.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
