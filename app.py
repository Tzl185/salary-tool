import streamlit as st
import zipfile
import os
import tempfile
import pandas as pd
from openpyxl import load_workbook

def process_file_a(folder_path):
    all_data = []
    positive_values = {}
    wage_cols_global = []

    for filename in os.listdir(folder_path):
        if filename.endswith(('.xls', '.xlsx')) and not filename.startswith('~$'):
            filepath = os.path.join(folder_path, filename)
            try:
                df = pd.read_excel(filepath, header=3)
                budget_unit_col = df.columns[1]
                wage_cols = df.columns[16:30]
                wage_cols_global = wage_cols  # 记住工资列

                df_filtered = df[[budget_unit_col] + list(wage_cols)]
                df_filtered[wage_cols] = df_filtered[wage_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
                df_grouped = df_filtered.groupby(budget_unit_col).sum()
                all_data.append(df_grouped)
            except Exception as e:
                st.warning(f"处理文件 {filename} 出错: {e}")

    if all_data:
        df_all = pd.concat(all_data)
        df_final = df_all.groupby(df_all.index).sum()

        for budget_unit, row in df_final.iterrows():
            for wage_type in wage_cols_global:
                value = row[wage_type]
                if value > 0:
                    if "绩效工资" in wage_type:
                        wage_type = wage_type.replace("绩效工资", "基础性绩效")
                    key = (str(budget_unit).strip(), str(wage_type).strip())
                    positive_values[key] = value
        return df_final, positive_values
    else:
        return None, None

def update_file_b(file_b_path, positive_values):
    wb = load_workbook(file_b_path)
    sheet = wb.active
    j_col_index = 10
    match_count = 0

    for row_idx in range(2, sheet.max_row + 1):
        unit_info = str(sheet.cell(row=row_idx, column=1).value or "").replace("-", "").replace(" ", "")
        project = str(sheet.cell(row=row_idx, column=2).value or "")
        matched = False

        for (budget_unit, wage_type), value in positive_values.items():
            budget_unit_cleaned = budget_unit.replace(" ", "")
            if (budget_unit_cleaned in unit_info or unit_info in budget_unit_cleaned) and wage_type in project:
                sheet.cell(row=row_idx, column=j_col_index).value = value
                match_count += 1
                matched = True
                break
    temp_output = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(temp_output.name)
    return temp_output.name

# --- Streamlit UI ---
st.title("工资调整自动汇总与填充工具")

zip_file = st.file_uploader("请上传包含Excel文件的.zip压缩包", type="zip")
file_b = st.file_uploader("请上传‘项目细化导入模板-正数.xlsx’(即文件B)", type="xlsx")

if zip_file and file_b:
    with tempfile.TemporaryDirectory() as tmpdir:
        # 解压zip文件
        zip_path = os.path.join(tmpdir, "uploaded.zip")
        with open(zip_path, "wb") as f:
            f.write(zip_file.read())

        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(tmpdir)

        # 读取压缩包内容
        df_a, positive_values = process_file_a(tmpdir)
        if df_a is None:
            st.error("处理失败：没有识别到有效数据")
        else:
            st.success("已成功汇总工资数据")

            # 保存汇总表
            summary_path = os.path.join(tmpdir, "汇总结果.xlsx")
            df_a.to_excel(summary_path)
            with open(summary_path, "rb") as f:
                st.download_button("下载汇总结果.xlsx", f, file_name="汇总结果.xlsx")

            # 更新文件B
            b_path = os.path.join(tmpdir, "file_b.xlsx")
            with open(b_path, "wb") as f:
                f.write(file_b.read())
            updated_b_path = update_file_b(b_path, positive_values)

            with open(updated_b_path, "rb") as f:
                st.download_button("下载更新后的文件B", f, file_name="更新后的文件B.xlsx")
