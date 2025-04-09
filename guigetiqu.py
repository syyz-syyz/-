import streamlit as st
import pandas as pd
import os


def extract_data(df, selected_column):
    unit_list = "包倍笔袋刀个罐盒斤块排瓶条箱桶支"
    h_values = []
    i_values = []

    for cell_value in df[selected_column]:
        if pd.isna(cell_value):
            h_values.append(cell_value)
            i_values.append("")
            continue
        cell_value = str(cell_value)
        h_value = ""
        i_value = ""
        has_number = False
        has_non_number = False
        has_english = False
        has_chinese = False
        number_appeared = False

        for char in cell_value:
            if char.isdigit():
                has_number = True
                number_appeared = True
            else:
                has_non_number = True
            if char.isdigit() or (char in unit_list and number_appeared) or (char in "-_*."):
                i_value += char
                if char in unit_list:
                    has_chinese = True
            elif ((65 <= ord(char) <= 90) or (97 <= ord(char) <= 122)) and number_appeared:
                i_value += char
                has_english = True
            else:
                h_value += char
                if char in unit_list:
                    number_appeared = False

        if has_number and has_non_number and (has_english or has_chinese):
            h_values.append(h_value)
            i_values.append(i_value)
        else:
            h_values.append(cell_value)
            i_values.append("")

    df['H'] = h_values
    df['I'] = i_values
    return df


def main():
    st.title("Excel 列拆分工具")

    # 上传 Excel 文件
    uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx", "xls"])

    if uploaded_file is not None:
        # 读取 Excel 文件
        excel_file = pd.ExcelFile(uploaded_file)

        # 获取所有表名
        sheet_names = excel_file.sheet_names

        # 选择工作表
        selected_sheet = st.selectbox("选择工作表", sheet_names)

        # 获取所选工作表中的数据
        df = excel_file.parse(selected_sheet)

        # 获取所有列名
        column_names = df.columns.tolist()

        # 选择要拆分的列
        selected_column = st.selectbox("选择要拆分的列", column_names)

        if st.button("拆分列"):
            # 调用函数进行拆分
            result_df = extract_data(df.copy(), selected_column)

            # 筛选出选择列、H列和I列
            final_df = result_df[[selected_column, 'H', 'I']]

            # 取前 10 行
            final_df = final_df.head(10)

            # 保存为 Excel 文件
            output_file = "output.xlsx"
            final_df.to_excel(output_file, index=False)

            # 提供下载链接
            if os.path.exists(output_file):
                with open(output_file, "rb") as file:
                    st.download_button(
                        label="下载拆分后的文件",
                        data=file,
                        file_name="output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )


if __name__ == "__main__":
    main()
    
