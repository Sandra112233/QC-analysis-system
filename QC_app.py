# QC_app.py
# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import io
from datetime import datetime
from database import init_db, save_record, load_all_records, delete_records

# ==================== 页面设置 ====================
st.set_page_config(page_title="QC数据智能分析系统", layout="wide")
st.title("QC数据智能分析系统")

# 初始化数据库
init_db()

# ==================== 侧边栏：基本信息录入 ====================
st.sidebar.header("📋 基本信息")

product_name = st.sidebar.text_input("品名", value="")
batch_no = st.sidebar.text_input("批号", value="")
spec = st.sidebar.text_input("规格", value="")
inspector = st.sidebar.text_input("检验人", value="")
inspection_date = st.sidebar.date_input("检验日期", value=datetime.now().date())

# ==================== 主区域：上传与模板生成 ====================
tab1, tab2 = st.tabs(["📤 上传数据 & 生成模板一", "📂 历史记录"])

with tab1:
    st.subheader("上传仪器原始数据")
    uploaded_file = st.file_uploader("选择仪器导出的 .xls 文件", type=["xls", "xlsx"])

    if uploaded_file is not None:
        df_raw = pd.read_excel(uploaded_file)

        st.subheader("📊 原始数据预览")
        st.dataframe(df_raw, use_container_width=True)

        # 找出数据中所有通道
        available_targets = df_raw["Target Name"].dropna().unique().tolist()
        default_targets = ["CY5", "FAM", "Texas Red", "VIC"]
        channels = [t for t in default_targets if t in available_targets]

        # 提取样本列表
        samples = df_raw["Sample Name"].dropna().unique().tolist()

        # 构建模板一数据
        template_data = []
        for sample in samples:
            if not sample or pd.isna(sample):
                continue
            row_data = {"编号": sample}
            sample_rows = df_raw[df_raw["Sample Name"] == sample]

            for ch in channels:
                ch_row = sample_rows[sample_rows["Target Name"] == ch]
                if len(ch_row) > 0:
                    try:
                        ct_val = ch_row["Cт"].values[0]
                    except KeyError:
                        ct_val = ch_row["Ct"].values[0]
                    try:
                        ct_val = round(float(ct_val), 2)
                    except (ValueError, TypeError):
                        ct_val = "Undetermined"
                    row_data[f"{ch}通道Ct值"] = ct_val
                else:
                    row_data[f"{ch}通道Ct值"] = "Undetermined"

            row_data["检测结果"] = ""
            row_data["结果判读"] = ""
            template_data.append(row_data)

        df_template = pd.DataFrame(template_data)

        st.subheader("📋 模板一预览")
        st.dataframe(df_template, use_container_width=True)

        # 下载模板一
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_template.to_excel(writer, sheet_name="原始记录附页", index=False)
        output.seek(0)

        st.download_button(
            label="📥 下载模板一 (Excel)",
            data=output,
            file_name=f"模板一_{batch_no}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # 保存到数据库
        if st.button("💾 保存到历史记录"):
            record = {
                "upload_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "product_name": product_name,
                "batch_no": batch_no,
                "spec": spec,
                "inspector": inspector,
                "inspection_date": str(inspection_date),
                "data": {
                    "raw_data": df_raw.to_dict(orient="records"),
                    "template_data": template_data,
                    "channels": channels
                }
            }
            save_record(record)
            st.success("✅ 已保存到历史记录！")
            st.rerun()

# ==================== 历史记录 ====================
with tab2:
    st.subheader("📂 历史记录")
    records = load_all_records()

    if len(records) == 0:
        st.info("暂无历史记录")
    else:
        selected_ids = []
        for rec in records:
            col1, col2, col3, col4 = st.columns([3, 2, 2, 1])
            with col1:
                st.write(f"**{rec['product_name']}**")
            with col2:
                st.write(f"批号: {rec['batch_no']}")
            with col3:
                st.write(f"上传: {rec['upload_time']}")
            with col4:
                selected = st.checkbox("选择", key=f"sel_{rec['id']}")
                if selected:
                    selected_ids.append(rec['id'])

        st.divider()

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("📥 生成模板二（合并选中记录）", disabled=len(selected_ids) < 1):
                st.info("模板二功能待开发")
        with col_btn2:
            if st.button("🗑 删除选中记录", disabled=len(selected_ids) < 1):
                delete_records(selected_ids)
                st.success("已删除")
                st.rerun()