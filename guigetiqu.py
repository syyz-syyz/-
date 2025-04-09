import streamlit as st
import pandas as pd
import os


def extract_data(df, selected_column, fixed_phrases):
    unit_list = "包倍笔袋刀个罐盒斤块排瓶条箱桶支克升"
    product_name_values = []
    product_specification_values = []

    for cell_value in df[selected_column]:
        if pd.isna(cell_value):
            product_name_values.append(cell_value)
            product_specification_values.append("")
            continue
        cell_value = str(cell_value)
        product_name = ""
        product_specification = ""
        has_number = False
        has_non_number = False
        has_english = False
        has_chinese = False
        number_appeared = False
        index = 0

        while index < len(cell_value):
            found_fixed_phrase = False
            for phrase in fixed_phrases:
                if cell_value[index:].startswith(phrase):
                    product_name += phrase
                    index += len(phrase)
                    found_fixed_phrase = True
                    break
            if found_fixed_phrase:
                continue

            char = cell_value[index]
            if char.isdigit():
                has_number = True
                number_appeared = True
            else:
                has_non_number = True
            if char.isdigit() or (char in unit_list and number_appeared) or (char in "-_*."):
                product_specification += char
                if char in unit_list:
                    has_chinese = True
            elif ((65 <= ord(char) <= 90) or (97 <= ord(char) <= 122)) and number_appeared:
                product_specification += char
                has_english = True
            else:
                product_name += char
                if char in unit_list:
                    number_appeared = False
            index += 1

        if has_number and has_non_number and (has_english or has_chinese):
            product_name_values.append(product_name)
            product_specification_values.append(product_specification)
        else:
            product_name_values.append(cell_value)
            product_specification_values.append("")

    df['产品名称'] = product_name_values
    df['产品规格'] = product_specification_values
    return df


def main():
    st.title("Excel 列拆分工具")

    # 上传 Excel 文件
    uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx", "xls"])

    # 默认固定词组
    default_fixed_phrases = ["0添加", "0度", "99%"]

    # 让用户输入固定词组
    fixed_phrases_input = st.text_input("输入要保留的固定词组，用逗号分隔（例如：0添加,0度,99%）")
    user_fixed_phrases = [phrase.strip() for phrase in fixed_phrases_input.split(',') if phrase.strip()]
    fixed_phrases = default_fixed_phrases + user_fixed_phrases

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
            result_df = extract_data(df.copy(), selected_column, fixed_phrases)

            # 筛选出选择列、产品名称列和产品规格列
            final_df = result_df[[selected_column, '产品名称', '产品规格']]

            # 取前 10 行用于预览
            preview_df = final_df.head(30)

            # 在网页上显示前 10 行预览
            st.write("拆分后数据的前 30 行预览：")
            st.dataframe(preview_df)

            # 保存完整数据为 Excel 文件
            output_file = "output.xlsx"
            final_df.to_excel(output_file, index=False)

            # 提供下载链接
            if os.path.exists(output_file):
                with open(output_file, "rb") as file:
                    st.download_button(
                        label="下载拆分后的完整文件",
                        data=file,
                        file_name="output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )


if __name__ == "__main__":
    main()
    
