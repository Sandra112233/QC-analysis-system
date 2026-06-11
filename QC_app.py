# QC_app.py
# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import os
from datetime import datetime
from copy import copy
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from database import init_db, save_record, load_all_records, delete_records

# ==================== 项目规则配置 ====================
PROJECT_CONFIGS = {
    "新冠甲乙流": {
        "channels": ["CY5", "FAM", "Texas Red", "VIC"],
        "channel_labels": {
            "CY5": "CY5通道Ct值\n（内标）",
            "FAM": "FAM通道Ct值\n（甲流）",
            "Texas Red": "Texas Red通道Ct值（乙流）",
            "VIC": "VIC通道Ct值\n（新冠）"
        },
        "pathogens": [
            {"name": "甲型流感病毒", "channel": "FAM", "threshold": 38},
            {"name": "乙型流感病毒", "channel": "Texas Red", "threshold": 38},
            {"name": "新冠病毒", "channel": "VIC", "threshold": 38},
        ],
        "use_prefix": True,
        "reference_categories": {
            "N": "阴性参考品",
            "P": "阳性参考品",
            "S": "最低检出限参考品",
            "R1": "重复性参考品R1",
            "R2": "重复性参考品R2",
            "R3": "重复性参考品R3",
            "YANG": "阳性质控品",
            "YIN": "阴性质控品"
        }
    },
    "通用": {
        "channels": ["CY5", "FAM", "Texas Red", "VIC"],
        "channel_labels": {
            "CY5": "CY5通道Ct值\n（内标）",
            "FAM": "FAM通道Ct值",
            "Texas Red": "Texas Red通道Ct值",
            "VIC": "VIC通道Ct值"
        },
        "pathogens": [
            {"name": "", "channel": "FAM", "threshold": 38},
            {"name": "", "channel": "Texas Red", "threshold": 38},
            {"name": "", "channel": "VIC", "threshold": 38},
        ],
        "use_prefix": False,
        "reference_categories": {
            "N": "阴性参考品",
            "P": "阳性参考品",
            "S": "最低检出限参考品",
            "R1": "重复性参考品R1",
            "R2": "重复性参考品R2",
            "R3": "重复性参考品R3",
            "YANG": "阳性质控品",
            "YIN": "阴性质控品"
        }
    }
}

# ==================== 判读规则（每个参考品类型的通道规则） ====================
JUDGE_RULES = {
    "N": {"CY5": "≤38", "FAM": "Undetermined或≥42", "TexasRed": "Undetermined或≥42", "VIC": "Undetermined或≥42", "expected": "阴性"},
    "P1-P10": {"CY5": "≤38", "FAM": "≤38", "TexasRed": "Undetermined或≥42", "VIC": "Undetermined或≥42", "expected": "阳性"},
    "P11-P14": {"CY5": "≤38", "FAM": "Undetermined或≥42", "TexasRed": "≤38", "VIC": "Undetermined或≥42", "expected": "阳性"},
    "P15-P20": {"CY5": "≤38", "FAM": "Undetermined或≥42", "TexasRed": "Undetermined或≥42", "VIC": "≤38", "expected": "阳性"},
    "S1-S5": {"CY5": "≤38", "FAM": "≤38", "TexasRed": "Undetermined或≥42", "VIC": "Undetermined或≥42", "expected": "阳性"},
    "S6-S7": {"CY5": "≤38", "FAM": "Undetermined或≥42", "TexasRed": "≤38", "VIC": "Undetermined或≥42", "expected": "阳性"},
    "S8": {"CY5": "≤38", "FAM": "Undetermined或≥42", "TexasRed": "Undetermined或≥42", "VIC": "≤38", "expected": "阳性"},
    "R1": {"CY5": "无要求", "FAM": "≤38", "TexasRed": "≤38", "VIC": "≤38", "expected": "阳性"},
    "R2": {"CY5": "无要求", "FAM": "≤38", "TexasRed": "≤38", "VIC": "≤38", "expected": "阳性"},
    "R3": {"CY5": "≤38", "FAM": "Undetermined或≥42", "TexasRed": "Undetermined或≥42", "VIC": "Undetermined或≥42", "expected": "阴性"},
    "YANG": {"CY5": "无要求", "FAM": "≤38", "TexasRed": "≤38", "VIC": "≤38", "expected": "阳性"},
    "YIN": {"CY5": "≤38", "FAM": "Undetermined", "TexasRed": "Undetermined", "VIC": "Undetermined", "expected": "阴性"},
}

# 质量标准文字
QUALITY_STANDARDS = {
    "N": "均为阴性",
    "P1-P10": "甲型流感病毒阳性",
    "P11-P14": "乙型流感病毒阳性",
    "P15-P20": "2019-nCoV新型冠状病毒阳性",
    "S1-S5": "甲型流感病毒阳性",
    "S6-S7": "乙型流感病毒阳性",
    "S8": "2019-nCoV新型冠状病毒阳性",
    "R1": "检测重复性参考品R1，重复检测10次，R1检测结果应均为阳性，且各重复性参考品检测结果Ct值的变异系数CV值均≤5%（内标通道无需进行统计）。",
    "R2": "检测重复性参考品R2，重复检测10次，R2检测结果应均为阳性，且各重复性参考品检测结果Ct值的变异系数CV值均≤5%（内标通道无需进行统计）。",
    "R3": "检测重复性参考品R3，重复检测10次，R3检测结果应为阴性。",
    "YANG": "FAM、Texas Red、VIC检测通道均存在明显扩增曲线，且Ct值≤32，CY5通道有或无扩增曲线",
    "YIN": "为阴性，CY5通道存在明显扩增曲线，且Ct值≤38，其他通道无扩增曲线。",
}

# ==================== 页面设置 ====================
st.set_page_config(page_title="QC数据智能分析系统", layout="wide")
st.title("QC数据智能分析系统")

init_db()

# ==================== 项目选择 ====================
project_name = st.selectbox("选择项目", list(PROJECT_CONFIGS.keys()))
config = PROJECT_CONFIGS[project_name]

# ==================== 函数 ====================
def parse_range(prefix_str):
    result = []
    parts = prefix_str.split(",")
    for part in parts:
        part = part.strip()
        if "-" in part:
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

def match_judge_rule(sample_name):
    for key, rule in JUDGE_RULES.items():
        prefixes = parse_range(key)
        if sample_name in prefixes:
            return rule
    return None

def check_channel(value, rule_str):
    if not rule_str or rule_str.strip() == "" or rule_str.strip() == "无要求":
        return True
    if value == "Undetermined" or value is None or (isinstance(value, float) and np.isnan(value)):
        return "Undetermined" in rule_str
    try:
        ct_val = float(value)
    except (ValueError, TypeError):
        return False
    if "Undetermined" in rule_str:
        return True
    nums = re.findall(r"[\d.]+", rule_str)
    if not nums:
        return False
    threshold = float(nums[0])
    if "≤" in rule_str:
        return ct_val <= threshold
    if "≥" in rule_str:
        return ct_val >= threshold
    return False

def do_judge(row_data, channels, judge_rule):
    # 第一步：检查无效
    cy5_val = row_data.get("CY5通道Ct值", "Undetermined")
    if cy5_val == "Undetermined" or (isinstance(cy5_val, (int, float)) and cy5_val > 38):
        return "无效", "不符合规定", "CY5通道Ct值为Undetermined或Ct>38，结果无效。"

    # 第二步：按病原优先级判读
    pathogens = config["pathogens"]
    for pathogen in pathogens:
        ch = pathogen["channel"]
        if ch not in channels:
            continue
        ch_val = row_data.get(f"{ch}通道Ct值", "Undetermined")
        try:
            ch_ct = float(ch_val)
            if ch_ct <= pathogen["threshold"]:
                if config["use_prefix"]:
                    result_name = f"{pathogen['name']}阳性"
                else:
                    result_name = "阳性"
                # 生成规则文字
                rule_parts = []
                for c in channels:
                    label = config["channel_labels"].get(c, c)
                    r_str = judge_rule.get(c, "") if judge_rule else ""
                    if r_str and r_str != "无要求":
                        rule_parts.append(f'"{label.replace(chr(10), "")}"为"{r_str}"')
                rule_text = ";".join(rule_parts)
                return result_name, "符合规定", rule_text
        except (ValueError, TypeError):
            pass

    # 第三步：判阴性
    # 所有病毒通道都是 Undetermined 或 ≥42
    all_negative = True
    for pathogen in pathogens:
        ch = pathogen["channel"]
        if ch not in channels:
            continue
        ch_val = row_data.get(f"{ch}通道Ct值", "Undetermined")
        if ch_val == "Undetermined":
            continue
        try:
            ch_ct = float(ch_val)
            if ch_ct < 42:
                all_negative = False
                break
        except (ValueError, TypeError):
            pass

    if all_negative:
        rule_parts = []
        for c in channels:
            label = config["channel_labels"].get(c, c)
            r_str = judge_rule.get(c, "") if judge_rule else ""
            if r_str and r_str != "无要求":
                rule_parts.append(f'"{label.replace(chr(10), "")}"为"{r_str}"')
        rule_text = ";".join(rule_parts)
        return "阴性", "符合规定", rule_text

    return "不符合", "不符合规定", ""

def get_category(sample_name):
    s = str(sample_name)
    for prefix, cat in config["reference_categories"].items():
        if s.startswith(prefix):
            # R1/R2/R3 特殊处理
            if prefix in ["R1", "R2", "R3"]:
                if s == prefix or re.match(rf"^{prefix}\d", s):
                    continue
                return cat
            return cat
    # 更细致的匹配
    if s.startswith("R1"):
        return config["reference_categories"].get("R1", "重复性参考品R1")
    if s.startswith("R2"):
        return config["reference_categories"].get("R2", "重复性参考品R2")
    if s.startswith("R3"):
        return config["reference_categories"].get("R3", "重复性参考品R3")
    return ""

def get_quality(sample_name):
    s = str(sample_name)
    for key, quality in QUALITY_STANDARDS.items():
        prefixes = parse_range(key)
        if s in prefixes:
            return quality
    return ""

# ==================== 侧边栏 ====================
st.sidebar.header("📋 基本信息")
product_name = st.sidebar.text_input("品名", value="")
batch_no = st.sidebar.text_input("批号", value="")
spec = st.sidebar.text_input("规格", value="")
inspector = st.sidebar.text_input("检验人", value="")
inspection_date = st.sidebar.date_input("检验日期", value=datetime.now().date())
ref_batch = st.sidebar.text_input("企业参考品批号", value="")

# ==================== 主区域 ====================
tab1, tab2 = st.tabs(["📤 上传数据 & 生成模板一", "📂 历史记录"])

with tab1:
    st.subheader("上传仪器原始数据")
    uploaded_file = st.file_uploader("选择仪器导出的 .xls 文件", type=["xls", "xlsx"])

    if uploaded_file is not None:
        # 读取数据
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

        # 找Ct列
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

        # 提取通道
        available_targets = df_raw["Target Name"].dropna().unique().tolist()
        channels = [t for t in config["channels"] if t in available_targets]

        # 提取样本并排序
        samples_raw = df_raw["Sample Name"].dropna().unique().tolist()
        samples_raw = [s for s in samples_raw if str(s).strip() != ""]

        category_order = {"N": 1, "P": 2, "S": 3, "R": 4, "YANG": 5, "YIN": 6}
        def sort_key(sample):
            s = str(sample)
            for prefix, order in category_order.items():
                if s.startswith(prefix):
                    nums = re.findall(r"\d+", s)
                    num = int(nums[0]) if nums else 0
                    return (order, num)
            return (99, 0)

        samples = sorted(samples_raw, key=sort_key)

        # 构建模板一
        template_data = []
        current_category = ""

        for sample in samples:
            category = get_category(str(sample))
            quality = get_quality(str(sample))
            judge_rule = match_judge_rule(str(sample))

            display_category = category if category != current_category else ""
            display_quality = quality if category != current_category else ""
            if category != current_category:
                current_category = category

            row_data = {
                "参考品": display_category,
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
            result, verdict, rule_text = do_judge(row_data, channels, judge_rule)
            row_data["检测结果"] = result
            row_data["结果判读"] = verdict
            row_data["结果判读规则"] = rule_text

            template_data.append(row_data)

        # R1/R2/R3 统计行
        for prefix in ["R1", "R2", "R3"]:
            r_rows = [r for r in template_data if str(r["编号"]) == prefix]
            if len(r_rows) >= 2:
                avg_row = {"参考品": "", "编号": "平均值", "质量标准": "/"}
                std_row = {"参考品": "", "编号": "标准偏差", "质量标准": "/"}
                cv_row = {"参考品": "", "编号": "变异系数（CV值）", "质量标准": "/"}
                cv_values = {}
                for ch in channels:
                    vals = []
                    for r in r_rows:
                        v = r.get(f"{ch}通道Ct值", "Undetermined")
                        try:
                            vals.append(float(v))
                        except (ValueError, TypeError):
                            pass
                    if len(vals) > 0:
                        a = round(np.mean(vals), 2)
                        s = round(np.std(vals, ddof=1), 4) if len(vals) > 1 else 0
                        c = round(s / a * 100, 2) if a != 0 else 0
                    else:
                        a, s, c = "/", "/", "/"
                    avg_row[f"{ch}通道Ct值"] = a
                    std_row[f"{ch}通道Ct值"] = s
                    cv_row[f"{ch}通道Ct值"] = f"{c}%" if not isinstance(c, str) else c
                    cv_values[ch] = c
                avg_row["检测结果"] = "/"
                avg_row["结果判读"] = "/"
                avg_row["结果判读规则"] = ""
                std_row["检测结果"] = "/"
                std_row["结果判读"] = "/"
                std_row["结果判读规则"] = ""
                cv_row["检测结果"] = "/"
                if prefix == "R3":
                    cv_row["结果判读"] = ""
                    cv_row["结果判读规则"] = ""
                else:
                    cv_ok = all(isinstance(cv_values.get(ch), (int, float)) and cv_values.get(ch) <= 5 for ch in channels)
                    cv_row["结果判读"] = "符合规定" if cv_ok else ""
                    cv_row["结果判读规则"] = '数值小于等于"5"'
                template_data.append(avg_row)
                template_data.append(std_row)
                template_data.append(cv_row) 

        # 构建最终列顺序
        final_columns = ["参考品", "编号", "质量标准"]
        for ch in channels:
            final_columns.append(config["channel_labels"].get(ch, f"{ch}通道Ct值"))
        final_columns += ["检测结果", "结果判读", "结果判读规则"]

        df_template = pd.DataFrame(template_data)
        # 重命名列
        rename_map = {}
        for ch in channels:
            rename_map[f"{ch}通道Ct值"] = config["channel_labels"].get(ch, f"{ch}通道Ct值")
        df_template = df_template.rename(columns=rename_map)
        # 确保列顺序
        existing_cols = [c for c in final_columns if c in df_template.columns]
        df_template = df_template[existing_cols]

        st.subheader("📋 模板一预览")
        st.dataframe(df_template, use_container_width=True)

        # 下载 Excel（含合并单元格）
        output = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "原始记录附页"

        # 写入公共字段标题行
        public_fields = ["日期", "企业参考品批号", "成品批号", "规格"]
        for i, field in enumerate(public_fields):
            ws.cell(row=1, column=i+1, value=field)

        # 空行
        ws.cell(row=2, column=1, value="")

        # 写入数据表头（从第3行开始）
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        header_font = Font(bold=True, size=10)
        header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

        for j, col_name in enumerate(existing_cols):
            cell = ws.cell(row=3, column=j+1, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')

        # 写入数据（从第4行开始）
        data_font = Font(size=10)
        for i, row_data in enumerate(template_data):
            row_num = 4 + i
            for j, col_name in enumerate(existing_cols):
                # 反查原始key
                orig_key = col_name
                for ch in channels:
                    if config["channel_labels"].get(ch) == col_name:
                        orig_key = f"{ch}通道Ct值"
                        break
                value = row_data.get(orig_key, row_data.get(col_name, ""))
                cell = ws.cell(row=row_num, column=j+1, value=value)
                cell.font = data_font
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')

        # 合并参考品列
        merge_ranges = []
        start_row = 4
        prev_cat = None
        for i, row_data in enumerate(template_data):
            cat = row_data.get("参考品", "")
            if cat != "" and cat is not None:
                if prev_cat is not None and start_row < 4 + i - 1:
                    merge_ranges.append((start_row, 4 + i - 1))
                start_row = 4 + i
                prev_cat = cat
        if prev_cat is not None and start_row < 4 + len(template_data) - 1:
            merge_ranges.append((start_row, 4 + len(template_data) - 1))

        for start, end in merge_ranges:
            if end > start:
                ws.merge_cells(start_row=start, start_column=1, end_row=end, end_column=1)

        # 调整列宽
        for j, col_name in enumerate(existing_cols):
            ws.column_dimensions[get_column_letter(j+1)].width = max(15, len(str(col_name))*2)

        wb.save(output)
        output.seek(0)

        st.download_button(
            label="📥 下载模板一 (Excel)",
            data=output,
            file_name=f"模板一_{batch_no}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if st.button("💾 保存到历史记录"):
            record = {
                "upload_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "product_name": product_name,
                "batch_no": batch_no,
                "spec": spec,
                "inspector": inspector,
                "inspection_date": str(inspection_date),
                "data": {"template_data": template_data, "channels": channels, "project": project_name}
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
