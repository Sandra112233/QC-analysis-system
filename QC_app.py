# QC_app.py
# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import io
import csv
import os
from datetime import datetime
from database import init_db, save_record, load_all_records, delete_records

# ==================== 页面设置 ====================
st.set_page_config(page_title="QC数据智能分析系统", layout="wide")
st.title("QC数据智能分析系统")

init_db()

# ==================== 加载判读规则 ====================
@st.cache_data
def load_rules():
    rules = []
    with open("rules.csv", "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rules.append(row)
    return rules

rules = load_rules()

# ==================== 解析编号前缀函数 ====================
def parse_range(prefix_str):
    """解析 P1-P10 这样的范围，返回列表"""
    result = []
    parts = prefix_str.split(",")
    for part in parts:
        part = part.strip()
        if "-" in part:
            # 找数字部分
            import re
            match = re.match(r"([A-Za-z]+)(\d+)-([A-Za-z]+)?(\d+)", part)
            if match:
                prefix1 = match.group(1)
                start = int(match.group(2))
                prefix2 = match.group(3) if match.group(3) else prefix1
                end = int(match.group(4))
                for i in range(start, end + 1):
                    result.append(f"{prefix1}{i}")
        else:
            result.append(part)
    return result

def match_rule(sample_name, rules):
    """根据样本编号匹配规则"""
    for rule in rules:
        prefixes = parse_range(rule["编号前缀"])
        if sample_name in prefixes:
            return rule
    return None

# ==================== 判读函数 ====================
def check_channel(value, rule_str):
    """检查单个通道是否满足规则"""
    if not rule_str or rule_str.strip() == "" or rule_str.strip() == "无要求":
        return True  # 无要求则通过

    if value == "Undetermined" or value is None or (isinstance(value, float) and np.isnan(value)):
        if "Undetermined" in rule_str:
            return True
        else:
            return False

    try:
        ct_val = float(value)
    except (ValueError, TypeError):
        return False

    if "Undetermined" in rule_str:
        return True
    if "≤" in rule_str:
        import re
        nums = re.findall(r"[\d.]+", rule_str)
        if nums:
            threshold = float(nums[0])
            return ct_val <= threshold
    if "≥" in rule_str:
        import re
        nums = re.findall(r"[\d.]+", rule_str)
        if nums:
            threshold = float(nums[0])
            return ct_val >= threshold

    return False

def do_judge(row_data, channels, rule):
    """执行判读，返回 检测结果、结果判读、判读规则文字"""
    all_pass = True
    for ch in channels:
        value = row_data.get(f"{ch}通道Ct值", "Undetermined")
        rule_str = rule.get(f"{ch}规则", "")
        if not check_channel(value, rule_str):
            all_pass = False
            break

    expected = rule.get("预期结果", "")
    if all_pass:
        result = expected
    else:
        result = "不符合"

    verdict = "符合规定" if result == expected else "不符合规定"
    rule_text = rule.get("结果判读规则", "")

    return result, verdict, rule_text

# ==================== 侧边栏 ====================
st.sidebar.header("📋 基本信息")
product_name = st.sidebar.text_input("品名", value="")
batch_no = st.sidebar.text_input("批号", value="")
spec = st.sidebar.text_input("规格", value="")
inspector = st.sidebar.text_input("检验人", value="")
inspection_date = st.sidebar.date_input("检验日期", value=datetime.now().date())

# ==================== 主区域 ====================
tab1, tab2 = st.tabs(["📤 上传数据 & 生成模板一", "📂 历史记录"])

with tab1:
    st.subheader("上传仪器原始数据")
    uploaded_file = st.file_uploader("选择仪器导出的 .xls 文件", type=["xls", "xlsx"])

    if uploaded_file is not None:
        # 第一步：读取原始数据
        try:
            df_raw = pd.read_excel(uploaded_file, header=6)
            if "Well" not in str(df_raw.columns) and "Sample Name" not in str(df_raw.columns):
                raise ValueError("列名不对")
        except:
            df_full = pd.read_excel(uploaded_file, header=None)
            header_row = None
            for i in range(len(df_full)):
                row_vals = df_full.iloc[i].astype(str).tolist()
                if "Well" in row_vals and "Sample Name" in row_vals:
                    header_row = i
                    break
            if header_row is None:
                st.error("找不到表头行，请确认文件格式。")
                st.stop()
            df_raw = pd.read_excel(uploaded_file, header=header_row)

        df_raw.columns = [str(c).strip() for c in df_raw.columns]
        df_raw = df_raw.dropna(how="all")

        # 找Ct值列
        ct_col = None
        for col in df_raw.columns:
            col_clean = col.replace(" ", "")
            if col_clean in ["Ct", "Cт", "CT"]:
                ct_col = col
                break
        if ct_col is None:
            for col in df_raw.columns:
                if "Ct" in str(col).replace(" ", "") or "Cт" in str(col).replace(" ", ""):
                    ct_col = col
                    break
        if ct_col is None:
            st.error(f"找不到Ct值列，当前列名：{list(df_raw.columns)}")
            st.stop()

        st.subheader("📊 原始数据预览")
        st.dataframe(df_raw, use_container_width=True)

        # 第二步：提取通道和样本
        available_targets = df_raw["Target Name"].dropna().unique().tolist()
        known_targets = ["CY5", "FAM", "Texas Red", "VIC"]
        channels = [t for t in known_targets if t in available_targets]

        samples = df_raw["Sample Name"].dropna().unique().tolist()
        samples = [s for s in samples if str(s).strip() != ""]

        # 第三步：构建模板一
        template_data = []
        current_category = ""
        current_quality = ""

        for sample in samples:
            rule = match_rule(str(sample), rules)
            if rule is None:
                category = ""
                quality = ""
            else:
                category = rule["参考品大类"]
                quality = rule["质量标准"]

            # 合并单元格逻辑：同一大类只显示一次
            if category != current_category:
                display_category = category
                display_quality = quality
                current_category = category
                current_quality = quality
            else:
                display_category = ""
                display_quality = ""

            row_data = {
                "参考品大类": display_category,
                "编号": sample,
                "质量标准": display_quality,
            }

            sample_rows = df_raw[df_raw["Sample Name"] == sample]

            for ch in channels:
                ch_row = sample_rows[sample_rows["Target Name"] == ch]
                if len(ch_row) > 0:
                    ct_val = ch_row[ct_col].values[0]
                    try:
                        ct_val = round(float(ct_val), 2)
                    except (ValueError, TypeError):
                        ct_val = str(ct_val)
                    row_data[f"{ch}通道Ct值"] = ct_val
                else:
                    row_data[f"{ch}通道Ct值"] = "Undetermined"

            # 判读
            if rule:
                result, verdict, rule_text = do_judge(row_data, channels, rule)
                row_data["检测结果"] = result
                row_data["结果判读"] = verdict
                row_data["结果判读规则"] = rule_text
            else:
                row_data["检测结果"] = ""
                row_data["结果判读"] = ""
                row_data["结果判读规则"] = ""

            template_data.append(row_data)

        # 第四步：处理R1/R2统计行
        for prefix, label in [("R1", "重复性参考品R1"), ("R2", "重复性参考品R2")]:
            r_rows = [r for r in template_data if str(r["编号"]).startswith(prefix) and r["编号"] == prefix]
            if len(r_rows) == 0:
                r_samples = [r for r in template_data if str(r["编号"]).startswith(prefix)]
                if len(r_samples) >= 2:
                    for ch in channels:
                        values = []
                        for r in r_samples:
                            val = r.get(f"{ch}通道Ct值", "Undetermined")
                            try:
                                values.append(float(val))
                            except (ValueError, TypeError):
                                pass
                        if len(values) > 0:
                            avg = round(np.mean(values), 2)
                            std = round(np.std(values, ddof=1), 4) if len(values) > 1 else 0
                            cv = round(std / avg * 100, 2) if avg != 0 else 0
                        else:
                            avg = "/"
                            std = "/"
                            cv = "/"

                    template_data.append({
                        "参考品大类": "",
                        "编号": "平均值",
                        "质量标准": "/",
                        **{f"{ch}通道Ct值": avg for ch in channels},
                        "检测结果": "/",
                        "结果判读": "/",
                        "结果判读规则": ""
                    })
                    template_data.append({
                        "参考品大类": "",
                        "编号": "标准偏差",
                        "质量标准": "/",
                        **{f"{ch}通道Ct值": std for ch in channels},
                        "检测结果": "/",
                        "结果判读": "/",
                        "结果判读规则": ""
                    })
                    cv_val = cv if isinstance(cv, str) else f"{cv}%"
                    cv_verdict = "符合规定" if (isinstance(cv, (int, float)) and cv <= 5) else ""
                    template_data.append({
                        "参考品大类": "",
                        "编号": "变异系数（CV值）",
                        "质量标准": "/",
                        **{f"{ch}通道Ct值": cv_val for ch in channels},
                        "检测结果": "/",
                        "结果判读": cv_verdict,
                        "结果判读规则": '数值小于等于"5"'
                    })

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
