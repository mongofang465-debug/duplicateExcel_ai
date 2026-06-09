import pandas as pd
import streamlit as st
from io import BytesIO

st.title("Excel 去重 & 规则筛选 MVP (用户自定义)")

# 1️⃣ 上传文件
uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])

if uploaded_file:
    # 读取所有 sheet
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
     # st.write("检测到 Sheets：", sheet_names)

    # 用户选择数据 sheet
    data_sheet = st.selectbox("选择【数据Sheet】", sheet_names)
    df = pd.read_excel(xls, sheet_name=data_sheet)

    st.write("数据预览：")
    st.dataframe(df.head())

    # -------------------------
    # 2️⃣ 用户选择判重列
    # -------------------------
    st.subheader("Step 1: 选择判重列（相同值归为一组）")
    group_cols = st.multiselect("判重列", options=df.columns.tolist())

    if group_cols:
        # -------------------------
        # 3️⃣ 用户选择规则列 + 对应规则
        # -------------------------
        st.subheader("Step 2: 选择规则列 & 规则类型")
        rule_cols = st.multiselect("规则列", options=[c for c in df.columns if c not in group_cols])

        rule_defs = []
        if rule_cols:
            st.write("为每个规则列选择规则类型：")
            for col in rule_cols:
                rule_type = st.selectbox(
                    f"{col} 的规则类型",
                    options=["max", "min", "latest", "earliest", "asc", "desc", "value=指定值"],
                    key=col
                )
                rule_defs.append(rule_type)

            # 如果有 value=指定值，需要用户输入具体值
            rule_params = []
            for col, rule in zip(rule_cols, rule_defs):
                if rule.startswith("value="):
                    val = st.text_input(f"{col} 指定值", key=col+"_val")
                    rule_params.append(val)
                else:
                    rule_params.append(None)

            # -------------------------
            # 4️⃣ 核心规则处理函数
            # -------------------------
            def apply_rule(df, group_cols, rule_cols, rule_defs, rule_params):
                temp_df = df.copy()
                sort_cols = []
                ascending_list = []

                for col, rule, param in zip(rule_cols, rule_defs, rule_params):
                    sort_col = "_sort_" + col

                    # 规则解析
                    if rule.startswith("value=") or param is not None:
                        val = param if param is not None else rule.split("=")[1]
                        temp_df[sort_col] = temp_df[col].apply(lambda x: 0 if str(x) == str(val) else 1)
                        ascending = True

                    elif rule == "max":
                        temp_df[sort_col] = temp_df[col]
                        ascending = False
                    elif rule == "min":
                        temp_df[sort_col] = temp_df[col]
                        ascending = True
                    elif rule == "latest":
                        temp_df[sort_col] = pd.to_datetime(temp_df[col], errors='coerce')
                        ascending = False
                    elif rule == "earliest":
                        temp_df[sort_col] = pd.to_datetime(temp_df[col], errors='coerce')
                        ascending = True
                    elif rule == "asc":
                        temp_df[sort_col] = temp_df[col]
                        ascending = True
                    elif rule == "desc":
                        temp_df[sort_col] = temp_df[col]
                        ascending = False
                    else:
                        temp_df[sort_col] = temp_df[col]
                        ascending = False

                    sort_cols.append(sort_col)
                    ascending_list.append(ascending)

                # 排序 + 分组取第一条
                temp_df = temp_df.sort_values(by=sort_cols, ascending=ascending_list)
                df_result = temp_df.groupby(group_cols, as_index=False).first()

                # 删除临时排序列
                df_result = df_result[[c for c in df_result.columns if not c.startswith("_sort_")]]
                return df_result

            # -------------------------
            # 5️⃣ 执行处理
            # -------------------------
            if st.button("执行去重规则"):
                result_df = apply_rule(df, group_cols, rule_cols, rule_defs, rule_params)
                st.success("处理完成！预览前5行：")
                st.dataframe(result_df.head())

                # 下载
                output = BytesIO()
                result_df.to_excel(output, index=False)
                output.seek(0)

                st.download_button(
                    label="下载处理后的 Excel",
                    data=output,
                    file_name="output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )