import re
import os
import shutil
import traceback
from copy import copy
from datetime import datetime

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
from openpyxl.styles import Font

# ==================== 配置 ====================
# 规则已写死，source_field 使用源表实际字段名（不再依赖列字母，列顺序变动不影响）
# action 类型: "直接取数" | "求和" | "多个日期取最早~最晚" | "多个日期取最晚"
# source_field 为 __COMPOUND_xxx__ 时走复合处理逻辑

HARDCODED_RULES = {
    "重特大工程": [
        {"field": "合同号",              "source_table": "主表",  "source_field": "合同编号",           "action": "直接取数", "note": "直接取数"},
        {"field": "合同名称",            "source_table": "主表",  "source_field": "合同名称",           "action": "直接取数", "note": "直接取数"},
        {"field": "厂家",               "source_table": "主表",  "source_field": "供应商名称",         "action": "直接取数", "note": "直接取数"},
        {"field": "数量",               "source_table": "次表1", "source_field": "采购数量",           "action": "求和",     "note": "一份合同可能对应多条物资名字，因此要对该合同下的所有数量进行求和"},
        {"field": "单位",               "source_table": "次表1", "source_field": "计量单位",           "action": "直接取数", "note": "一般多条物资明细的单位均一致，直接取数"},
        {"field": "合同最终结算金额（元）", "source_table": "主表",  "source_field": "合同金额（元）",     "action": "直接取数", "note": "直接取数"},
        {"field": "付款比例",            "source_table": "主表",  "source_field": "合同支付比例",       "action": "直接取数", "note": "直接取数"},
        {"field": "合同签订日期",         "source_table": "主表",  "source_field": "合同签约完成时间",   "action": "直接取数", "note": "直接取数"},
        {"field": "投标偏差（是/否）",     "source_table": "主表",  "source_field": "备注",               "action": "直接取数", "note": "若提出了正偏差，则填正偏差；若提出了负偏差，则填负偏差；如果未提出偏差则填否"},
        {"field": "合同要求到货时间",      "source_table": "主表",  "source_field": "合同交货日期",       "action": "直接取数", "note": "直接取数"},
        {"field": "现场最终要求交货日期(系统)", "source_table": "次表1", "source_field": "合同最终交货日期", "action": "直接取数", "note": "如果明细中的日期为1个，则直接取数；如果有多个且完全重复，则取该重复日期；如果有多个且不完全重复，则更新为最早时间~最晚时间"},
        {"field": "到货比例（%）",  "source_table": "主表",  "source_field": "到货比例",           "action": "直接取数", "note": "直接取数"},
        {"field": "已到货金额（元）",       "source_table": "主表",  "source_field": "已到货金额（元）",   "action": "直接取数", "note": "直接取数"},
        {"field": "实际到货日期",         "source_table": "次表3", "source_field": "实际到货日期",     "action": "直接取数", "note": "如果次表3中只有一条数据则直接取数；如果有多个且完全重复则取该重复日期；如果有多个且不完全重复则取最早时间~最晚时间"},
        {"field": "全部到货验收时间",       "source_table": "主表",  "source_field": "__COMPOUND_ACCEPTANCE__",    "action": "直接取数", "note": "同一合同编号下，主表AA列到货比例为100%后，取次表3该合同所有明细的最晚验收日期；不为100%则/"},
        {"field": "已付金额（元）",         "source_table": "主表",  "source_field": "已支付金额（元）",   "action": "直接取数", "note": "直接取数"},
        {"field": "累计付款比例(百分比)",    "source_table": "主表",  "source_field": "付款比例",           "action": "直接取数", "note": "直接取数"},
        {"field": "预付款支付时间",         "source_table": "次表2", "source_field": "__COMPOUND_PREPAY__",         "action": "直接取数", "note": "筛选次表2付款类型=预付款，取实际付款日期；如果没有数据则更新为/"},
        {"field": "冻结款支付时间",         "source_table": "次表2", "source_field": "__COMPOUND_DELIVERY_PAY__",     "action": "直接取数", "note": "筛选次表2付款类型=设计冻结款，取实际付款日期；如果没有数据则更新为/"},
        {"field": "结清款到期时间",         "source_table": "次表2", "source_field": "结清款到期日期",     "action": "直接取数", "note": "直接取数"},
        {"field": "质保到期时间",          "source_table": "次表2", "source_field": "__COMPOUND_WARRANTY__",        "action": "直接取数", "note": "次表2质保期开始计算之日 + 主表备注中'质保期为XX个月'的月数 = 质保到期时间；次表2为空则/"},
        {"field": "数量",               "source_table": "次表4", "source_field": "__COMPOUND_RETRIEVAL__",       "action": "直接取数", "note": "次表4筛选领料类型=退库/出库，取数量列汇总求和；为空则/"},
        {"field": "单位",               "source_table": "次表4", "source_field": "__COMPOUND_RETRIEVAL_UNIT__",  "action": "直接取数", "note": "次表4筛选领料类型=退库/出库，取计量单位列第一个值；为空则/"},
    ],
    "改扩建工程": [
        {"field": "合同号",              "source_table": "主表",  "source_field": "合同编号",           "action": "直接取数", "note": "直接取数"},
        {"field": "合同名称",            "source_table": "主表",  "source_field": "合同名称",           "action": "直接取数", "note": "直接取数"},
        {"field": "厂家",               "source_table": "主表",  "source_field": "供应商名称",         "action": "直接取数", "note": "直接取数"},
        {"field": "数量",               "source_table": "次表1", "source_field": "采购数量",           "action": "求和",     "note": "一份合同可能对应多条物资名字，因此要对该合同下的所有数量进行求和"},
        {"field": "单位",               "source_table": "次表1", "source_field": "计量单位",           "action": "直接取数", "note": "一般多条物资明细的单位均一致，直接取数"},
        {"field": "合同最终结算金额（元）", "source_table": "主表",  "source_field": "合同金额（元）",     "action": "直接取数", "note": "直接取数"},
        {"field": "付款比例",            "source_table": "主表",  "source_field": "合同支付比例",       "action": "直接取数", "note": "直接取数"},
        {"field": "合同签订日期",         "source_table": "主表",  "source_field": "合同签约完成时间",   "action": "直接取数", "note": "直接取数"},
        {"field": "投标偏差（是/否）",     "source_table": "主表",  "source_field": "备注",               "action": "直接取数", "note": "若提出了正偏差，则填正偏差；若提出了负偏差，则填负偏差；如果未提出偏差则填否"},
        {"field": "合同要求到货时间",      "source_table": "主表",  "source_field": "合同交货日期",       "action": "直接取数", "note": "直接取数"},
        {"field": "现场最终要求交货日期(系统)", "source_table": "次表1", "source_field": "合同最终交货日期", "action": "直接取数", "note": "如果明细中的日期为1个，则直接取数；如果有多个且完全重复，则取该重复日期；如果有多个且不完全重复，则更新为最早时间~最晚时间"},
        {"field": "实际到货日期",         "source_table": "次表3", "source_field": "实际到货日期",     "action": "直接取数", "note": "如果次表3中只有一条数据则直接取数；如果有多个且完全重复则取该重复日期；如果有多个且不完全重复则取最早时间~最晚时间"},
        {"field": "全部到货验收时间",       "source_table": "主表",  "source_field": "__COMPOUND_ACCEPTANCE__",    "action": "直接取数", "note": "同一合同编号下，主表AA列到货比例为100%后，取次表3该合同所有明细的最晚验收日期；不为100%则/"},
        {"field": "出入库时间",           "source_table": "次表4", "source_field": "__COMPOUND_WAREHOUSE__",      "action": "直接取数", "note": "次表4筛选领料类型=领料，用合同号取J列日期；一个日期直接取，多个取最早~最晚，空则/"},
        {"field": "已付金额（元）",         "source_table": "主表",  "source_field": "已支付金额（元）",   "action": "直接取数", "note": "直接取数"},
        {"field": "累计付款比例(百分比)",    "source_table": "主表",  "source_field": "付款比例",           "action": "直接取数", "note": "直接取数"},
        {"field": "结清款到期时间",         "source_table": "次表2", "source_field": "结清款到期日期",     "action": "直接取数", "note": "直接取数"},
        {"field": "质保到期时间",          "source_table": "次表2", "source_field": "__COMPOUND_WARRANTY__",        "action": "直接取数", "note": "次表2质保期开始计算之日 + 主表备注中'质保期为XX个月'的月数 = 质保到期时间；次表2为空则/"},
        {"field": "数量",               "source_table": "次表4", "source_field": "__COMPOUND_RETRIEVAL__",       "action": "直接取数", "note": "次表4筛选领料类型=退库/出库，取数量列汇总求和；为空则/"},
        {"field": "单位",               "source_table": "次表4", "source_field": "__COMPOUND_RETRIEVAL_UNIT__",  "action": "直接取数", "note": "次表4筛选领料类型=退库/出库，取计量单位列第一个值；为空则/"},
    ],
    "技改项目": [
        {"field": "合同号",              "source_table": "主表",  "source_field": "合同编号",           "action": "直接取数", "note": "直接取数"},
        {"field": "合同名称",            "source_table": "主表",  "source_field": "合同名称",           "action": "直接取数", "note": "直接取数"},
        {"field": "厂家",               "source_table": "主表",  "source_field": "供应商名称",         "action": "直接取数", "note": "直接取数"},
        {"field": "数量",               "source_table": "次表1", "source_field": "采购数量",           "action": "求和",     "note": "一份合同可能对应多条物资名字，因此要对该合同下的所有数量进行求和"},
        {"field": "单位",               "source_table": "次表1", "source_field": "计量单位",           "action": "直接取数", "note": "一般多条物资明细的单位均一致，直接取数"},
        {"field": "合同最终结算金额（元）", "source_table": "主表",  "source_field": "合同金额（元）",     "action": "直接取数", "note": "直接取数"},
        {"field": "付款比例",            "source_table": "主表",  "source_field": "合同支付比例",       "action": "直接取数", "note": "直接取数"},
        {"field": "合同签订日期",         "source_table": "主表",  "source_field": "合同签约完成时间",   "action": "直接取数", "note": "直接取数"},
        {"field": "投标偏差（是/否）",     "source_table": "主表",  "source_field": "备注",               "action": "直接取数", "note": "若提出了正偏差，则填正偏差；若提出了负偏差，则填负偏差；如果未提出偏差则填否"},
        {"field": "合同要求到货时间",      "source_table": "主表",  "source_field": "合同交货日期",       "action": "直接取数", "note": "直接取数"},
        {"field": "现场最终要求交货日期(系统)", "source_table": "次表1", "source_field": "合同最终交货日期", "action": "直接取数", "note": "如果明细中的日期为1个，则直接取数；如果有多个且完全重复，则取该重复日期；如果有多个且不完全重复，则更新为最早时间~最晚时间"},
        {"field": "实际到货日期",         "source_table": "次表3", "source_field": "实际到货日期",     "action": "直接取数", "note": "如果次表3中只有一条数据则直接取数；如果有多个且完全重复则取该重复日期；如果有多个且不完全重复则取最早时间~最晚时间"},
        {"field": "全部到货验收时间",       "source_table": "主表",  "source_field": "__COMPOUND_ACCEPTANCE__",    "action": "直接取数", "note": "同一合同编号下，主表AA列到货比例为100%后，取次表3该合同所有明细的最晚验收日期；不为100%则/"},
        {"field": "出入库时间",           "source_table": "次表4", "source_field": "__COMPOUND_WAREHOUSE__",      "action": "直接取数", "note": "次表4筛选领料类型=领料，用合同号取J列日期；一个日期直接取，多个取最早~最晚，空则/"},
        {"field": "已付金额（元）",         "source_table": "主表",  "source_field": "已支付金额（元）",   "action": "直接取数", "note": "直接取数"},
        {"field": "累计付款比例(百分比)",    "source_table": "主表",  "source_field": "付款比例",           "action": "直接取数", "note": "直接取数"},
        {"field": "结清款到期时间",         "source_table": "次表2", "source_field": "结清款到期日期",     "action": "直接取数", "note": "直接取数"},
        {"field": "质保到期时间",          "source_table": "次表2", "source_field": "__COMPOUND_WARRANTY__",        "action": "直接取数", "note": "次表2质保期开始计算之日 + 主表备注中'质保期为XX个月'的月数 = 质保到期时间；次表2为空则/"},
        {"field": "数量",               "source_table": "次表4", "source_field": "__COMPOUND_RETRIEVAL__",       "action": "直接取数", "note": "次表4筛选领料类型=退库/出库，取数量列汇总求和；为空则/"},
        {"field": "单位",               "source_table": "次表4", "source_field": "__COMPOUND_RETRIEVAL_UNIT__",  "action": "直接取数", "note": "次表4筛选领料类型=退库/出库，取计量单位列第一个值；为空则/"},
    ],
}

# 源表读取配置: {别名: (header_row, skip_rows)}
SOURCE_TABLE_CONFIG = {
    "主表": {"header_row": 2, "skip_rows": [0]},
    "次表1": {"header_row": 2, "skip_rows": [0]},
    "次表2": {"header_row": 2, "skip_rows": [0]},
    "次表3": {"header_row": 1, "skip_rows": None},
    "次表4": {"header_row": 2, "skip_rows": [0]},
}

# 合同基本信息输出字段（需要更新到模板的字段）
CONTRACT_INFO_FIELDS = [
    "合同号", "合同名称", "厂家", "数量", "单位",
    "合同签订金额", "合同最终结算金额", "付款比例",
    "合同签订日期", "供应商联系人及电话", "合同要求到货时间",
    "物资名称", "累计付款比例(百分比)", "已付金额（元）",
    "投标偏差（是/否）",
    "预付款支付时间", "冻结款支付时间", "结清款到期时间", "质保到期时间",
    "逆向物资数量", "逆向物资单位",
    "现场最终要求交货日期(系统)",
    "全部到货验收时间", "到货验收时间", "出入库时间",
    "到货比例（%）", "已到货金额（元）",
    "实际到货日期",
]

# 模板中重复列名的处理：第二次出现的"数量"→"逆向物资数量"，第二次出现的"单位"→"逆向物资单位"
DUPLICATE_FIELD_ALIASES = {
    "数量": "逆向物资数量",
    "单位": "逆向物资单位",
}

FIELD_ALIASES = {
    '合同签订金额': ['合同签订金额', '合同金额', '合同金额(元)', '合同签订金额(元)'],
    '合同最终结算金额': ['合同最终结算金额', '合同最终结算金额(元)', '合同最终结算金额（元）'],
    '预付款支付时间': ['预付款支付时间'],
    '冻结款支付时间': ['冻结款支付时间'],
    '现场最终要求交货日期(系统)': ['现场最终要求交货日期(系统)', '现场最终要求交货日期（系统）'],
    '到货比例（%）': ['到货比例（%）', '到货比例'],
    '全部到货验收时间': ['全部到货验收时间', '到货验收时间'],
    '实际到货日期': ['实际到货日期'],
}

SEGMENT_KEYWORDS = ['云南段', '广西段', '广东段']


# ==================== 工具函数 ====================
def safe_date_parse(date_str):
    if pd.isna(date_str) or str(date_str).strip() == "":
        return ""
    date_str = str(date_str).strip()
    m = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', date_str)
    if m:
        return f"{m.group(1)}-{m.group(2).zfill(2)}-{m.group(3).zfill(2)}"
    m = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', date_str)
    if m:
        return f"{m.group(1)}-{m.group(2).zfill(2)}-{m.group(3).zfill(2)}"
    return date_str


def ratio_parse(ratio_str):
    if pd.isna(ratio_str):
        return ""
    ratio_str = str(ratio_str).strip()
    try:
        if ":" in ratio_str:
            parts = [p.strip() for p in ratio_str.split(":") if p.strip()]
            if parts:
                total = sum(float(x) for x in parts)
                return "/".join([f"{float(x)/total*100:.1f}%" for x in parts])
        if "%" in ratio_str:
            return ratio_str
        ratio = float(ratio_str)
        if 0 <= ratio <= 1:
            return f"{ratio*100:.1f}%"
        return f"{ratio}%"
    except:
        return ratio_str


def _safe_float(val):
    """安全地将值转为float，处理逗号、空格等格式"""
    if val is None or val == '':
        return 0.0
    try:
        s = str(val).strip().replace(',', '').replace('，', '').replace(' ', '').replace('　', '')
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def percent_to_decimal(pct_str):
    if not pct_str or pd.isna(pct_str):
        return 0.0
    pct_str = str(pct_str).strip().replace('%', '')
    try:
        return float(pct_str) / 100.0
    except:
        return 0.0


def values_differ(old_val, new_val):
    if old_val is None:
        old_val = ""
    if new_val is None:
        new_val = ""
    try:
        return float(old_val) != float(new_val)
    except (ValueError, TypeError):
        pass

    def to_str(v):
        if isinstance(v, datetime):
            return v.strftime("%Y-%m-%d")
        return str(v).strip()

    return to_str(old_val) != to_str(new_val)


def copy_row_style(ws, from_row, to_row, max_col=None):
    if max_col is None:
        max_col = ws.max_column
    for col in range(1, max_col + 1):
        src = ws.cell(row=from_row, column=col)
        dst = ws.cell(row=to_row, column=col)
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format


def get_first_column_as_sequence(ws, data_start, max_row):
    header_cell = ws.cell(row=data_start - 1, column=1)
    if header_cell.value:
        clean_col = re.sub(r'[()（）\s]', '', str(header_cell.value)).lower()
        for field in CONTRACT_INFO_FIELDS:
            if re.sub(r'[()（）\s]', '', field).lower() == clean_col:
                return None
    if max_row >= data_start:
        cell = ws.cell(row=max_row, column=1)
        try:
            return int(cell.value) if cell.value is not None else 0
        except:
            return None
    return 0


def clean_project_name(name):
    return re.sub(r'[（(].*?[）)]', '', str(name)).strip()


def clean_filename(name):
    name = re.sub(r'[\\/*?:"<>|]', '_', name)
    name = name.strip('. ')
    if not name:
        name = "未命名"
    return name


def find_all_summary_files(summary_folder, project_name):
    """在文件夹中查找所有匹配项目名称的xlsx模板文件"""
    clean_name = clean_project_name(project_name)
    result = []
    for fname in os.listdir(summary_folder):
        if not fname.endswith('.xlsx') or fname.startswith('~$'):
            continue
        clean_fname = clean_project_name(fname)
        if clean_name in clean_fname:
            result.append(os.path.join(summary_folder, fname))
    return result


def _has_segment_sheets(filepath):
    try:
        wb = openpyxl.load_workbook(filepath)
        for sn in wb.sheetnames:
            for seg in SEGMENT_KEYWORDS:
                if seg in sn:
                    wb.close()
                    return True
        wb.close()
    except Exception:
        pass
    return False


def classify_templates(file_paths):
    line_path = None
    device_path = None
    for fp in file_paths:
        if _has_segment_sheets(fp):
            line_path = fp
        else:
            device_path = fp
    return line_path, device_path


def extract_segment(name):
    name_str = str(name)
    for seg in SEGMENT_KEYWORDS:
        if seg in name_str:
            return seg
    return None


def extract_template_tag(filepath):
    fname = os.path.basename(filepath)
    m = re.search(r'\[.*?\]', fname)
    return m.group(0) if m else ''


def split_data_by_sub_project(main_proj, sub_proj):
    """按单项工程名称是否包含线路工程拆分主表和次表数据"""
    main_sub_col = None
    for col in main_proj.columns:
        if '单项工程名称' in str(col):
            main_sub_col = col
            break
    sub_sub_col = None
    for col in sub_proj.columns:
        if '单项工程名称' in str(col):
            sub_sub_col = col
            break
    if main_sub_col is None or sub_sub_col is None:
        return (main_proj, sub_proj), (None, None)
    main_line = main_proj[main_proj[main_sub_col].astype(str).str.contains('线路工程')].copy()
    main_device = main_proj[~main_proj[main_sub_col].astype(str).str.contains('线路工程')].copy()
    line_contracts = set(
        main_line['合同编号'].astype(str).str.strip()) if '合同编号' in main_line.columns else set()
    device_contracts = set(
        main_device['合同编号'].astype(str).str.strip()) if '合同编号' in main_device.columns else set()
    sub_line = pd.DataFrame()
    sub_device = pd.DataFrame()
    if '合同编号' in sub_proj.columns:
        sub_proj_copy = sub_proj.copy()
        sub_proj_copy['_contract'] = sub_proj_copy['合同编号'].astype(str).str.strip()
        sub_line = sub_proj_copy[sub_proj_copy['_contract'].isin(line_contracts)].drop(columns=['_contract'])
        sub_device = sub_proj_copy[sub_proj_copy['_contract'].isin(device_contracts)].drop(columns=['_contract'])
    return (main_line, sub_line), (main_device, sub_device)


# ==================== 源表读取 ====================
def _auto_detect_header_row(filepath, sheet_name):
    """自动检测某个sheet的表头行号（从第1行开始找，包含'合同编号'或'合同号'的行为表头）"""
    # 先无表头读取前20行来探测
    try:
        df_preview = pd.read_excel(filepath, sheet_name=sheet_name, header=None, nrows=20)
    except:
        return 0  # 读不了就用第1行
    for row_idx in range(min(20, len(df_preview))):
        row_vals = [str(v).strip() for v in df_preview.iloc[row_idx].values if pd.notna(v)]
        row_text = ' '.join(row_vals)
        if '合同编号' in row_text or '合同号' in row_text:
            return row_idx  # 返回0-based行号（作为header参数）
    return 0  # 没找到就用第1行


def read_source_table_all_sheets(filepath):
    """读取源表Excel所有sheet，自动识别表头并拼接，返回 DataFrame"""
    xl = pd.ExcelFile(filepath)
    all_dfs = []
    for sn in xl.sheet_names:
        header_row = _auto_detect_header_row(filepath, sn)
        try:
            df = pd.read_excel(filepath, sheet_name=sn, header=header_row)
        except:
            # 如果指定header_row失败，尝试默认读取
            df = pd.read_excel(filepath, sheet_name=sn)
        # 跳过空sheet
        if df.empty:
            continue
        df.columns = df.columns.astype(str).str.strip()
        df.columns = df.columns.str.replace('​', '')
        df.columns = df.columns.str.replace('（', '(').str.replace('）', ')')
        all_dfs.append(df)
    if not all_dfs:
        return pd.DataFrame()
    combined = pd.concat(all_dfs, ignore_index=True)
    return combined


# ==================== 字段名匹配 ====================
def find_field_in_df(df, field_names):
    """在DataFrame的列中查找匹配的字段名。field_names按优先级排列。"""
    for name in field_names:
        if name in df.columns:
            return name
        # 模糊匹配（去括号、去空格）
        clean_target = re.sub(r'[()（）\s]', '', name).lower()
        for col in df.columns:
            if re.sub(r'[()（）\s]', '', str(col)).lower() == clean_target:
                return col
    return None


# ==================== 构建规则（显式字段名，无需列字母转换） ====================
def build_source_field_mapping(rules):
    """处理HARDCODED_RULES中的重复字段名，source_field已在规则中显式指定。
    返回: rules增强版，增加 _original_field 字段，重复字段加后缀
    """
    for rule in rules:
        rule['_original_field'] = rule['field']

    # 处理重复字段名: 统计出现次数，重复的加后缀(数量_次表1, 数量_次表4)
    field_counts = {}
    for rule in rules:
        fn = rule['_original_field']
        field_counts[fn] = field_counts.get(fn, 0) + 1

    field_occurrence = {}
    for rule in rules:
        fn = rule['_original_field']
        if field_counts.get(fn, 1) > 1:
            occ = field_occurrence.get(fn, 0)
            field_occurrence[fn] = occ + 1
            if occ > 0:
                rule['field'] = f"{fn}_{rule['source_table']}"

    # source_field 已在规则中显式指定，直接使用
    for rule in rules:
        rule['skip_source'] = False

    return rules


# ==================== 数据提取引擎 ====================
def extract_contract_data(main_df, sub_dfs, rules):
    """根据规则从各源表提取合同数据。
    返回: DataFrame，列为规则中的输出字段名，行为每个合同
    """
    # 确保合同编号列存在
    if '合同编号' not in main_df.columns:
        for col in main_df.columns:
            if '合同编号' in str(col):
                main_df = main_df.rename(columns={col: '合同编号'})
                break

    main_df['合同编号'] = main_df['合同编号'].astype(str).str.strip()
    contracts = main_df['合同编号'].unique()

    result_rows = []
    for contract in contracts:
        row = {}
        row['合同号'] = contract

        # 按规则提取每个字段
        for rule in rules:
            field = rule['field']
            src_table = rule['source_table']
            src_field = rule.get('source_field')
            action = rule['action']
            note = rule['note']

            if rule.get('skip_source'):
                continue

            if src_field is None:
                continue

            # 检查是否为复合字段
            if str(src_field).startswith('__COMPOUND_'):
                value = _apply_compound_rule(contract, main_df, sub_dfs, src_field)
                row[field] = value
                continue

            # 获取源表数据
            if src_table == '主表':
                src_df = main_df
            else:
                src_df = sub_dfs.get(src_table)

            if src_df is None:
                continue

            # 确保源表有合同编号列
            contract_col = find_field_in_df(src_df, ['合同编号', '合同号'])
            if contract_col is None:
                continue
            src_df[contract_col] = src_df[contract_col].astype(str).str.strip()
            src_match = src_df[src_df[contract_col] == contract]

            # 确保源字段存在（显式字段名，直接传入查找）
            actual_field = find_field_in_df(src_df, [src_field])
            if actual_field is None:
                continue

            value = _apply_extract_rule(src_match, actual_field, action, note)
            row[field] = value

        result_rows.append(row)

    result = pd.DataFrame(result_rows)
    return result


def _apply_extract_rule(src_match, field, action, note):
    """对匹配的源数据应用提取规则"""
    if src_match.empty:
        return ""

    values = src_match[field].dropna()

    if values.empty:
        return ""

    # 直接取数（需先检查日期范围/最晚逻辑，否则直接取第一个值）
    if action == '直接取数':
        # 检查是否需要日期范围处理（最早~最晚）
        if '最早' in note and '最晚' in note:
            try:
                dates = pd.to_datetime(values, errors='coerce').dropna()
                if len(dates) == 0:
                    return ""
                elif len(dates) == 1:
                    return dates.iloc[0].strftime('%Y-%m-%d')
                elif dates.nunique() == 1:
                    # 多个日期但全部相同，只取这个日期，不展示范围
                    return dates.iloc[0].strftime('%Y-%m-%d')
                else:
                    # 多个不完全相同的日期，展示最早~最晚
                    return f"{dates.min().strftime('%Y-%m-%d')}~{dates.max().strftime('%Y-%m-%d')}"
            except:
                pass
        # 多个日期取最晚
        if '最晚' in note:
            try:
                dates = pd.to_datetime(values, errors='coerce').dropna()
                if len(dates) > 0:
                    return dates.max().strftime('%Y-%m-%d')
            except:
                pass
        return values.iloc[0]

    # 筛选去重求和 (用于数量字段)
    if '求和' in action or '汇总求和' in note:
        try:
            return values.astype(float).sum()
        except:
            return values.iloc[0]

    # 多个日期取最早~最晚范围
    if '最早时间' in note and '最晚时间' in note:
        try:
            dates = pd.to_datetime(values, errors='coerce').dropna()
            if len(dates) == 0:
                return ""
            elif len(dates) == 1:
                return dates.iloc[0].strftime('%Y-%m-%d')
            elif dates.nunique() == 1:
                # 多个日期但全部相同，只取这个日期，不展示范围
                return dates.iloc[0].strftime('%Y-%m-%d')
            else:
                # 多个不完全相同的日期，展示最早~最晚
                return f"{dates.min().strftime('%Y-%m-%d')}~{dates.max().strftime('%Y-%m-%d')}"
        except:
            pass
        return values.iloc[0]

    # 多个日期取最早
    if '最早完成日期' in note or '最早' in note:
        try:
            dates = pd.to_datetime(values, errors='coerce').dropna()
            if len(dates) > 0:
                return dates.min().strftime('%Y-%m-%d')
        except:
            pass
        return values.iloc[0]

    # 默认返回第一个值
    return values.iloc[0]


def _apply_compound_rule(contract, main_df, sub_dfs, compound_type):
    """处理复合字段，所有字段名均使用源表实际字段名"""
    if compound_type == '__COMPOUND_PREPAY__':
        # 次表2筛选"预付款"，取实际付款日期
        sub2 = sub_dfs.get('次表2')
        if sub2 is None:
            return ''
        contract_col = find_field_in_df(sub2, ['合同编号', '合同号'])
        type_col = find_field_in_df(sub2, ['付款类型'])
        date_col = find_field_in_df(sub2, ['实际付款日期'])
        if not all([contract_col, type_col, date_col]):
            return ''
        match = sub2[(sub2[contract_col].astype(str).str.strip() == str(contract).strip()) &
                     (sub2[type_col].astype(str).str.contains('预付款'))]
        if match.empty:
            return '/'
        dates = pd.to_datetime(match[date_col], errors='coerce').dropna()
        return dates.iloc[0].strftime('%Y-%m-%d') if not dates.empty else '/'

    elif compound_type == '__COMPOUND_DELIVERY_PAY__':
        # 次表2筛选"设计冻结款"，取实际付款日期
        sub2 = sub_dfs.get('次表2')
        if sub2 is None:
            return ''
        contract_col = find_field_in_df(sub2, ['合同编号', '合同号'])
        type_col = find_field_in_df(sub2, ['付款类型'])
        date_col = find_field_in_df(sub2, ['实际付款日期'])
        if not all([contract_col, type_col, date_col]):
            return ''
        match = sub2[(sub2[contract_col].astype(str).str.strip() == str(contract).strip()) &
                     (sub2[type_col].astype(str).str.contains('设计冻结款'))]
        if match.empty:
            return '/'
        dates = pd.to_datetime(match[date_col], errors='coerce').dropna()
        return dates.iloc[0].strftime('%Y-%m-%d') if not dates.empty else '/'

    elif compound_type == '__COMPOUND_WARRANTY__':
        # 次表2 BC列(质保期开始计算之日) + 主表 BM列(备注中"质保期为XX个月")
        # = 质保到期时间(日期类型)
        sub2 = sub_dfs.get('次表2')
        main_df_local = main_df
        if sub2 is None:
            return '/'

        # 1. 获取质保期开始计算之日
        contract_col = find_field_in_df(sub2, ['合同编号', '合同号'])
        bc_col = find_field_in_df(sub2, ['质保期开始计算之日'])
        if not all([contract_col, bc_col]):
            return '/'
        match = sub2[sub2[contract_col].astype(str).str.strip() == str(contract).strip()]
        if match.empty:
            return '/'
        dates = pd.to_datetime(match[bc_col], errors='coerce').dropna()
        if dates.empty:
            return '/'
        start_date = dates.iloc[0]  # 取第一个日期

        # 2. 从主表备注解析质保期月数（如"质保期为120个月" → 120）
        months = None
        main_contract_col = find_field_in_df(main_df_local, ['合同编号', '合同号'])
        if main_contract_col:
            mm = main_df_local[main_df_local[main_contract_col].astype(str).str.strip() == str(contract).strip()]
            if not mm.empty:
                bm_col = find_field_in_df(main_df_local, ['备注'])
                if bm_col:
                    remark = str(mm[bm_col].iloc[0]) if pd.notna(mm[bm_col].iloc[0]) else ''
                    m = re.search(r'质保期[为是]?\s*(\d+)\s*个?月', remark)
                    if m:
                        months = int(m.group(1))

        if months is None:
            # 解析不到质保期月数，仅返回开始日期
            return start_date.strftime('%Y-%m-%d')

        # 3. 计算质保到期时间 = 开始日期 + 质保期月数
        end_date = start_date + pd.DateOffset(months=months)
        return end_date.strftime('%Y-%m-%d')

    elif compound_type == '__COMPOUND_MATERIALS__':
        # 从次表1聚合物资名称
        sub1 = sub_dfs.get('次表1')
        if sub1 is None:
            return ''
        contract_col = find_field_in_df(sub1, ['合同编号', '合同号'])
        mat_col = find_field_in_df(sub1, ['物资名称'])
        if not all([contract_col, mat_col]):
            return ''
        match = sub1[sub1[contract_col].astype(str).str.strip() == str(contract).strip()]
        if match.empty:
            return ''
        materials = match[mat_col].dropna().astype(str).str.strip().tolist()
        return '，'.join(dict.fromkeys(materials))  # 去重保序

    elif compound_type == '__COMPOUND_WAREHOUSE__':
        # 改扩建工程：次表4 B列筛选"领料"，用合同号取J列日期
        # 一个日期直接取，多个日期取最早~最晚，空则/
        sub4 = sub_dfs.get('次表4')
        if sub4 is None:
            return ''
        contract_col = find_field_in_df(sub4, ['合同编号', '合同号'])
        type_col = find_field_in_df(sub4, ['领料类型'])
        time_col = find_field_in_df(sub4, ['领料时间'])
        if not all([contract_col, type_col, time_col]):
            return ''
        match = sub4[(sub4[contract_col].astype(str).str.strip() == str(contract).strip()) &
                     (sub4[type_col].astype(str).str.contains('领料'))]
        if match.empty:
            return '/'
        dates = pd.to_datetime(match[time_col], errors='coerce').dropna()
        if dates.empty:
            return '/'
        if len(dates) == 1 or dates.nunique() == 1:
            return dates.iloc[0].strftime('%Y-%m-%d')
        else:
            return f"{dates.min().strftime('%Y-%m-%d')}~{dates.max().strftime('%Y-%m-%d')}"

    elif compound_type == '__COMPOUND_ACCEPTANCE__':
        # 到货验收时间：先判断主表到货比例是否100%
        # 100%则取次表3该合同所有明细的最晚验收日期，否则返回/
        arrival_col = find_field_in_df(main_df, ['到货比例'])
        if not arrival_col:
            return '/'
        main_contract_col = find_field_in_df(main_df, ['合同编号', '合同号'])
        if not main_contract_col:
            return '/'
        mm = main_df[main_df[main_contract_col].astype(str).str.strip() == str(contract).strip()]
        if mm.empty:
            return '/'
        arrival_rate = str(mm[arrival_col].iloc[0]).strip()
        # 判断到货比例是否100%
        try:
            rate_val = float(arrival_rate.replace('%', ''))
            if rate_val < 100:
                return '/'
        except:
            return '/'
        # 到货比例100%，取次表3最晚验收日期
        sub3 = sub_dfs.get('次表3')
        if sub3 is None:
            return '/'
        contract_col3 = find_field_in_df(sub3, ['合同编号', '合同号'])
        date_col = find_field_in_df(sub3, ['验收日期'])
        if not all([contract_col3, date_col]):
            return '/'
        match = sub3[sub3[contract_col3].astype(str).str.strip() == str(contract).strip()]
        if match.empty:
            return '/'
        dates = pd.to_datetime(match[date_col], errors='coerce').dropna()
        if dates.empty:
            return '/'
        return dates.max().strftime('%Y-%m-%d')

    elif compound_type == '__COMPOUND_RETRIEVAL__':
        # 次表4筛选"退库"或"出库"，取数量列汇总求和
        sub4 = sub_dfs.get('次表4')
        if sub4 is None:
            return ''
        contract_col = find_field_in_df(sub4, ['合同编号', '合同号'])
        type_col = find_field_in_df(sub4, ['领料类型'])
        qty_col = find_field_in_df(sub4, ['数量'])
        if not all([contract_col, type_col, qty_col]):
            return ''
        match = sub4[(sub4[contract_col].astype(str).str.strip() == str(contract).strip()) &
                     (sub4[type_col].astype(str).str.contains('退库|出库'))]
        if match.empty:
            return '/'
        try:
            return match[qty_col].astype(float).sum()
        except:
            return match[qty_col].iloc[0] if not match.empty else '/'

    elif compound_type == '__COMPOUND_RETRIEVAL_UNIT__':
        # 次表4筛选"退库"或"出库"，取计量单位列第一个值
        sub4 = sub_dfs.get('次表4')
        if sub4 is None:
            return ''
        contract_col = find_field_in_df(sub4, ['合同编号', '合同号'])
        type_col = find_field_in_df(sub4, ['领料类型'])
        unit_col = find_field_in_df(sub4, ['计量单位'])
        if not all([contract_col, type_col, unit_col]):
            return ''
        match = sub4[(sub4[contract_col].astype(str).str.strip() == str(contract).strip()) &
                     (sub4[type_col].astype(str).str.contains('退库|出库'))]
        if match.empty:
            return '/'
        return str(match[unit_col].iloc[0]) if not match.empty else '/'

    return ''


# ==================== 结果构建 ====================
def build_result(main_df, sub_dfs, rules, category):
    """构建结果DataFrame，匹配CONTRACT_INFO_FIELDS格式"""

    raw = extract_contract_data(main_df, sub_dfs, rules)

    # 字段映射: 规则输出字段 → 结果DataFrame列名
    field_map = {
        '合同号': '合同号',
        '合同名称': '合同名称',
        '厂家': '厂家',
        '数量': '数量',
        '数量_次表1': '数量',
        '数量_次表4': '逆向物资数量',
        '单位': '单位',
        '单位_次表1': '单位',
        '单位_次表4': '逆向物资单位',
        '合同签订金额': '合同签订金额',
        '合同签订金额（元）': '合同签订金额',
        '合同最终结算金额': '合同最终结算金额',
        '合同最终结算金额（元）': '合同最终结算金额',
        '付款比例': '付款比例',
        '合同签订日期': '合同签订日期',
        '投标偏差（是/否）': '投标偏差（是/否）',
        '合同要求到货时间': '合同要求到货时间',
        '已付金额（元）': '已付金额（元）',
        '累计付款比例(百分比)': '累计付款比例(百分比)',
        '累计付款比例': '累计付款比例(百分比)',
        '物资名称': '物资名称',
        '结清款到期时间': '结清款到期时间',
        '质保到期时间': '质保到期时间',
        '预付款支付时间': '预付款支付时间',
        '冻结款支付时间': '冻结款支付时间',
        '现场最终要求交货日期(系统)': '现场最终要求交货日期(系统)',
        '到货比例（%）': '到货比例（%）',
        '到货比例': '到货比例（%）',
        '已到货金额（元）': '已到货金额（元）',
        '全部到货验收时间': '全部到货验收时间',
        '实际到货日期': '实际到货日期',
    }

    result = pd.DataFrame()
    result['合同号'] = raw.get('合同号', pd.Series(''))

    # 合同名称
    result['合同名称'] = raw.get('合同名称', pd.Series(''))

    # 厂家: 优先用规则字段, 否则从主表取
    if '厂家' in raw.columns:
        result['厂家'] = raw['厂家']
    else:
        supplier_col = find_field_in_df(main_df, ['供应商名称'])
        if supplier_col:
            main_df['_contract'] = main_df['合同编号'].astype(str).str.strip()
            supplier_map = main_df.set_index('_contract')[supplier_col].to_dict()
            result['厂家'] = result['合同号'].map(supplier_map).fillna('')
        else:
            result['厂家'] = ''

    # 数量: 从规则输出取(次表1的采购数量汇总)
    qty_field = '数量' if '数量' in raw.columns else '数量_次表1'
    result['数量'] = raw.get(qty_field, pd.Series(''))

    # 单位: 从规则输出取(次表1的计量单位)
    unit_field = '单位' if '单位' in raw.columns else '单位_次表1'
    result['单位'] = raw.get(unit_field, pd.Series(''))

    # 合同签订金额: 从主表取
    amt_col = find_field_in_df(main_df, ['合同金额(元)', '合同金额（元）', '合同金额'])
    if amt_col:
        main_df['_contract'] = main_df['合同编号'].astype(str).str.strip()
        amt_map = main_df.set_index('_contract')[amt_col].to_dict()
        result['合同签订金额'] = result['合同号'].map(amt_map).fillna('')
    else:
        result['合同签订金额'] = ''

    # 合同最终结算金额: 直接从主表合同金额（元）取数覆盖
    amt_col_final = find_field_in_df(main_df, ['合同金额(元)', '合同金额（元）', '合同金额'])
    if amt_col_final:
        main_df['_contract'] = main_df['合同编号'].astype(str).str.strip()
        amt_map_final = main_df.set_index('_contract')[amt_col_final].to_dict()
        result['合同最终结算金额'] = result['合同号'].map(amt_map_final).fillna('')
    else:
        result['合同最终结算金额'] = ''

    # 付款比例
    result['付款比例'] = raw.get('付款比例', pd.Series(''))

    # 合同签订日期（空值写"合同签订中"）
    signing_date = raw.get('合同签订日期', pd.Series(''))
    result['合同签订日期'] = signing_date.apply(
        lambda x: safe_date_parse(x) if pd.notna(x) and str(x).strip() != '' else '合同签订中'
    )

    # 供应商联系人及电话: 从主表取
    contact_col = find_field_in_df(main_df, ['供应商联系人'])
    phone_col = find_field_in_df(main_df, ['供应商联系人电话'])
    if contact_col and phone_col:
        main_df['_contract'] = main_df['合同编号'].astype(str).str.strip()
        contact_map = main_df.set_index('_contract')[contact_col].fillna('').astype(str).to_dict()
        phone_map = main_df.set_index('_contract')[phone_col].fillna('').astype(str).to_dict()
        result['供应商联系人及电话'] = result['合同号'].apply(
            lambda c: str(contact_map.get(c, '')) + str(phone_map.get(c, '')))
    else:
        result['供应商联系人及电话'] = ''

    # 合同要求到货时间
    delivery = raw.get('合同要求到货时间', pd.Series(''))
    result['合同要求到货时间'] = delivery.apply(safe_date_parse)

    # 物资名称: 从次表1聚合
    if '物资名称' in raw.columns:
        result['物资名称'] = raw['物资名称']
    else:
        sub1 = sub_dfs.get('次表1')
        if sub1 is not None:
            c_col = find_field_in_df(sub1, ['合同编号', '合同号'])
            m_col = find_field_in_df(sub1, ['物资名称'])
            if c_col and m_col:
                sub1_copy = sub1.copy()
                sub1_copy['_c'] = sub1_copy[c_col].astype(str).str.strip()
                mat_map = {}
                for c, grp in sub1_copy.groupby('_c'):
                    mats = grp[m_col].dropna().astype(str).str.strip().tolist()
                    mat_map[c] = '，'.join(dict.fromkeys(mats))
                result['物资名称'] = result['合同号'].map(mat_map).fillna('')
            else:
                result['物资名称'] = ''
        else:
            result['物资名称'] = ''

    # 已付金额
    result['已付金额（元）'] = raw.get('已付金额（元）', pd.Series(''))

    # 累计付款比例
    accum_rate = raw.get('累计付款比例(百分比)',
                         raw.get('累计付款比例', pd.Series('')))
    result['累计付款比例(百分比)'] = accum_rate.apply(
        lambda x: ratio_parse(x) if pd.notna(x) and x != '' else '')

    # 投标偏差：根据文本内容判断正偏差/负偏差/否
    def _parse_bid_deviation(val):
        if pd.isna(val) or str(val).strip() == '':
            return '否'
        s = str(val)
        if '正偏差' in s:
            return '正偏差'
        elif '负偏差' in s:
            return '负偏差'
        else:
            return '否'

    raw_deviation = raw.get('投标偏差（是/否）', pd.Series(''))
    result['投标偏差（是/否）'] = raw_deviation.apply(_parse_bid_deviation)

    # 预付款支付时间：从次表2取预付款的实际付款日期
    result['预付款支付时间'] = raw.get('预付款支付时间', pd.Series(''))

    # 冻结款支付时间：从次表2取设计冻结款的实际付款日期
    result['冻结款支付时间'] = raw.get('冻结款支付时间', pd.Series(''))

    # 结清款到期时间：从次表2结清款到期日期直接取数，为空则/
    due_date = raw.get('结清款到期时间', pd.Series(''))
    result['结清款到期时间'] = due_date.apply(lambda x: '/' if pd.isna(x) or str(x).strip() == '' else x)

    # 质保到期时间：从次表2质保期开始计算之日 + 主表备注复合取数
    result['质保到期时间'] = raw.get('质保到期时间', pd.Series(''))

    # 现场最终要求交货日期(系统)：从次表1合同最终交货日期取
    result['现场最终要求交货日期(系统)'] = raw.get('现场最终要求交货日期(系统)', pd.Series(''))

    # 全部到货验收时间 (重特大/改扩建/技改项目)：从复合规则取
    result['全部到货验收时间'] = raw.get('全部到货验收时间', pd.Series(''))
    result['到货验收时间'] = raw.get('到货验收时间', pd.Series(''))  # 兼容旧模板

    # 出入库时间 (改扩建/技改项目)：从次表4复合规则取
    result['出入库时间'] = raw.get('出入库时间', pd.Series(''))

    # 到货比例（%）/ 已到货金额（从主表直接取数，source_field不变）
    result['到货比例（%）'] = raw.get('到货比例（%）', pd.Series(''))
    result['已到货金额（元）'] = raw.get('已到货金额（元）', pd.Series(''))

    # 实际到货日期（从次表3 AE列取，支持最早~最晚，空则/）
    actual_arrival = raw.get('实际到货日期', pd.Series(''))
    result['实际到货日期'] = actual_arrival.apply(lambda x: '/' if pd.isna(x) or str(x).strip() == '' else x)

    # 逆向物资数量：从次表4筛选退库/出库，取数量列汇总
    result['逆向物资数量'] = raw.get('数量_次表4', pd.Series(''))

    # 逆向物资单位：从次表4筛选退库/出库，取计量单位列第一个值
    result['逆向物资单位'] = raw.get('单位_次表4', pd.Series(''))

    # 排序
    try:
        result['_sort'] = pd.to_datetime(result['合同签订日期'], errors='coerce')
        result = result.sort_values('_sort', na_position='last').drop(columns=['_sort'])
    except:
        result = result.sort_values('合同签订日期')
    result = result.reset_index(drop=True)

    return result


# ==================== 模板更新 ====================
def _update_template(result, summary_path, project_name):
    """更新单个汇总表模板"""
    if not os.path.exists(summary_path):
        return False, f"汇总表文件不存在: {summary_path}"
    wb = openpyxl.load_workbook(summary_path)
    segment_sheets = {}
    for sn in wb.sheetnames:
        for seg in SEGMENT_KEYWORDS:
            if seg in sn:
                segment_sheets[seg] = sn
                break
    if len(segment_sheets) >= 2:
        return _process_multi_sheet(wb, result, segment_sheets, summary_path, project_name)
    else:
        return _process_single_sheet(wb, result, summary_path, project_name)


def _process_single_sheet(wb, result, summary_path, project_name):
    ws = wb.active
    header_row = None
    for row in ws.iter_rows(min_row=1, max_row=10):
        for cell in row:
            if cell.value and "合同号" in str(cell.value):
                header_row = cell.row
                break
        if header_row:
            break
    if header_row is None:
        header_row = 4

    col_map = {}
    field_occ = {}  # 跟踪每个字段在表头中出现的次数，用于处理重复列名
    for cell in ws[header_row]:
        if cell.value:
            col_name = str(cell.value).strip()
            clean_col = re.sub(r'[()（）\s]', '', col_name).lower()
            for field in CONTRACT_INFO_FIELDS:
                aliases = FIELD_ALIASES.get(field, [field])
                for alias in aliases:
                    if re.sub(r'[()（）\s]', '', alias).lower() == clean_col:
                        occ = field_occ.get(field, 0)
                        field_occ[field] = occ + 1
                        if occ == 0:
                            col_map[field] = cell.column
                        elif field in DUPLICATE_FIELD_ALIASES:
                            col_map[DUPLICATE_FIELD_ALIASES[field]] = cell.column
                        break
    if '合同号' not in col_map:
        headers = [str(ws.cell(row=header_row, column=c).value) for c in range(1, min(ws.max_column + 1, 20))]
        return False, (f"汇总表中未找到合同号列\n  文件: {os.path.basename(summary_path)}\n"
                       f"  表头行({header_row}行)前20列: {headers}")

    data_start = header_row + 1
    contract_row_map = {}
    for row in ws.iter_rows(min_row=data_start, max_row=ws.max_row,
                            min_col=col_map['合同号'], max_col=col_map['合同号']):
        cell = row[0]
        if cell.value is not None:
            contract_row_map[str(cell.value).strip()] = cell.row

    blue = "0000FF"
    green = "008000"
    yellow = "FFA500"
    updated_count = 0
    changed_contracts = []
    new_rows_data = []

    for _, r_row in result.iterrows():
        contract = str(r_row['合同号']).strip()
        if contract in contract_row_map:
            row_idx = contract_row_map[contract]
            row_changed = False
            for field in CONTRACT_INFO_FIELDS:
                if field == '已付金额（元）':
                    continue
                if field == '供应商联系人及电话':
                    continue
                if field == '合同签订金额':
                    continue  # 已有行不覆盖合同签订金额，仅新增行填入
                if field in col_map and field in r_row:
                    col_idx = col_map[field]
                    old_cell = ws.cell(row=row_idx, column=col_idx)
                    if values_differ(old_cell.value, r_row[field]):
                        old_cell.value = r_row[field]
                        new_font = copy(old_cell.font) if old_cell.font else Font()
                        new_font.color = blue
                        old_cell.font = new_font
                        row_changed = True
            if '已付金额（元）' in col_map:
                amt_col = col_map.get('合同签订金额')
                rat_col = col_map.get('累计付款比例(百分比)')
                if amt_col and rat_col:
                    amt_val = ws.cell(row=row_idx, column=amt_col).value
                    rat_val = ws.cell(row=row_idx, column=rat_col).value
                    try:
                        amt = _safe_float(amt_val)
                        rat = percent_to_decimal(rat_val)
                        paid = amt * rat
                        paid_cell = ws.cell(row=row_idx, column=col_map['已付金额（元）'])
                        if values_differ(paid_cell.value, paid):
                            paid_cell.value = paid
                            paid_cell.number_format = '#,##0.00'
                            new_font = copy(paid_cell.font) if paid_cell.font else Font()
                            new_font.color = blue
                            paid_cell.font = new_font
                            row_changed = True
                    except:
                        pass
            if row_changed:
                updated_count += 1
                changed_contracts.append(contract)
        else:
            new_rows_data.append(r_row)

    if new_rows_data:
        template_row = ws.max_row if ws.max_row >= data_start else data_start
        seq_index = get_first_column_as_sequence(ws, data_start, ws.max_row)
        seq = seq_index if seq_index is not None else 0
        for r_row in new_rows_data:
            # 跳过空合同号，不插入
            contract = str(r_row.get('合同号', '')).strip()
            if not contract:
                continue
            ins_row = ws.max_row + 1
            copy_row_style(ws, template_row, ins_row)
            if seq_index is not None:
                seq += 1
                seq_cell = ws.cell(row=ins_row, column=1)
                seq_cell.value = seq
                # 序号列继承模板行第1列字体，改绿色
                src_font0 = ws.cell(row=template_row, column=1).font
                new_font0 = copy(src_font0) if src_font0 else Font()
                new_font0.color = green
                seq_cell.font = new_font0
            for field in CONTRACT_INFO_FIELDS:
                if field == '已付金额（元）':
                    continue
                # 合同签订金额在新增行中填入（与已有行跳过逻辑相反）
                if field in col_map and field in r_row:
                    cell = ws.cell(row=ins_row, column=col_map[field])
                    cell.value = r_row[field]
                    # 继承模板行对应列的字体，仅改颜色为绿色
                    src_cell = ws.cell(row=template_row, column=col_map[field])
                    new_font = copy(src_cell.font) if src_cell.font else Font()
                    new_font.color = green
                    cell.font = new_font
            if '已付金额（元）' in col_map:
                amt_col = col_map.get('合同签订金额')
                rat_col = col_map.get('累计付款比例(百分比)')
                if amt_col and rat_col:
                    amt = _safe_float(r_row['合同签订金额'])
                    rat = percent_to_decimal(r_row['累计付款比例(百分比)'])
                    paid = amt * rat
                    paid_cell = ws.cell(row=ins_row, column=col_map['已付金额（元）'])
                    paid_cell.value = paid
                    paid_cell.number_format = '#,##0.00'
                    # 继承模板行字体，改绿色
                    src_font2 = ws.cell(row=template_row, column=col_map['已付金额（元）']).font
                    new_font2 = copy(src_font2) if src_font2 else Font()
                    new_font2.color = green
                    paid_cell.font = new_font2
            template_row = ins_row

    result_contracts = set(result['合同号'].astype(str).str.strip())
    yellow_font = Font(color=yellow)
    for row in ws.iter_rows(min_row=data_start, max_row=ws.max_row,
                            min_col=col_map['合同号'], max_col=col_map['合同号']):
        cell = row[0]
        if cell.value is not None:
            c = str(cell.value).strip()
            if c and c not in result_contracts:
                for col_idx in range(1, ws.max_column + 1):
                    target = ws.cell(row=cell.row, column=col_idx)
                    if target.value is not None:
                        new_font = copy(target.font) if target.font else Font()
                        new_font.color = yellow
                        target.font = new_font

    wb.save(summary_path)

    changed_info = "、".join(changed_contracts) if changed_contracts else "无"
    return True, (f"项目 {project_name} 更新完成\n"
                  f"更新合同: {updated_count} 个\n"
                  f"新增合同: {len(new_rows_data)} 个\n"
                  f"变更合同号: {changed_info}")


def _process_multi_sheet(wb, result, segment_sheets, summary_path, project_name):
    blue = "0000FF"
    green = "008000"
    yellow = "FFA500"

    sheet_info = {}
    for seg, sn in segment_sheets.items():
        ws = wb[sn]
        header_row = None
        for row in ws.iter_rows(min_row=1, max_row=10):
            for cell in row:
                if cell.value and "合同号" in str(cell.value):
                    header_row = cell.row
                    break
            if header_row:
                break
        if header_row is None:
            header_row = 4

        col_map = {}
        field_occ = {}
        for cell in ws[header_row]:
            if cell.value:
                col_name = str(cell.value).strip()
                clean_col = re.sub(r'[()（）\s]', '', col_name).lower()
                for field in CONTRACT_INFO_FIELDS:
                    aliases = FIELD_ALIASES.get(field, [field])
                    for alias in aliases:
                        if re.sub(r'[()（）\s]', '', alias).lower() == clean_col:
                            occ = field_occ.get(field, 0)
                            field_occ[field] = occ + 1
                            if occ == 0:
                                col_map[field] = cell.column
                            elif field in DUPLICATE_FIELD_ALIASES:
                                col_map[DUPLICATE_FIELD_ALIASES[field]] = cell.column
                            break

        data_start = header_row + 1
        contract_row_map = {}
        if '合同号' in col_map:
            for row in ws.iter_rows(min_row=data_start, max_row=ws.max_row,
                                    min_col=col_map['合同号'], max_col=col_map['合同号']):
                cell = row[0]
                if cell.value is not None:
                    contract_row_map[str(cell.value).strip()] = cell.row

        sheet_info[seg] = {
            'ws': ws, 'col_map': col_map, 'contract_row_map': contract_row_map,
            'header_row': header_row, 'data_start': data_start
        }

    global_contract_map = {}
    for seg, info in sheet_info.items():
        for contract in info['contract_row_map']:
            if contract not in global_contract_map:
                global_contract_map[contract] = seg

    def determine_segment(contract_name):
        name = str(contract_name)
        for seg in SEGMENT_KEYWORDS:
            if seg in name:
                return seg
        return None

    updated_count = 0
    new_count = 0
    changed_contracts = []
    new_contracts_data = {}

    for _, r_row in result.iterrows():
        contract = str(r_row['合同号']).strip()
        if contract in global_contract_map:
            seg = global_contract_map[contract]
            info = sheet_info[seg]
            ws = info['ws']
            col_map = info['col_map']
            row_idx = info['contract_row_map'][contract]
            row_changed = False
            for field in CONTRACT_INFO_FIELDS:
                if field == '已付金额（元）':
                    continue
                if field == '供应商联系人及电话':
                    continue
                if field == '合同签订金额':
                    continue  # 已有行不覆盖合同签订金额，仅新增行填入
                if field in col_map and field in r_row:
                    col_idx = col_map[field]
                    old_cell = ws.cell(row=row_idx, column=col_idx)
                    if values_differ(old_cell.value, r_row[field]):
                        old_cell.value = r_row[field]
                        new_font = copy(old_cell.font) if old_cell.font else Font()
                        new_font.color = blue
                        old_cell.font = new_font
                        row_changed = True
            if '已付金额（元）' in col_map:
                amt_col = col_map.get('合同签订金额')
                rat_col = col_map.get('累计付款比例(百分比)')
                if amt_col and rat_col:
                    amt_val = ws.cell(row=row_idx, column=amt_col).value
                    rat_val = ws.cell(row=row_idx, column=rat_col).value
                    try:
                        amt = _safe_float(amt_val)
                        rat = percent_to_decimal(rat_val)
                        paid = amt * rat
                        paid_cell = ws.cell(row=row_idx, column=col_map['已付金额（元）'])
                        if values_differ(paid_cell.value, paid):
                            paid_cell.value = paid
                            paid_cell.number_format = '#,##0.00'
                            new_font = copy(paid_cell.font) if paid_cell.font else Font()
                            new_font.color = blue
                            paid_cell.font = new_font
                            row_changed = True
                    except:
                        pass
            if row_changed:
                updated_count += 1
                changed_contracts.append(contract)
        else:
            target_seg = determine_segment(r_row['合同名称'])
            if target_seg and target_seg in sheet_info:
                if target_seg not in new_contracts_data:
                    new_contracts_data[target_seg] = []
                new_contracts_data[target_seg].append(r_row)
                new_count += 1

    for seg, rows in new_contracts_data.items():
        info = sheet_info[seg]
        ws = info['ws']
        col_map = info['col_map']
        data_start = info['data_start']
        if '合同号' not in col_map:
            continue
        template_row = ws.max_row if ws.max_row >= data_start else data_start
        seq_index = get_first_column_as_sequence(ws, data_start, ws.max_row)
        seq = seq_index if seq_index is not None else 0
        for r_row in rows:
            ins_row = ws.max_row + 1
            copy_row_style(ws, template_row, ins_row)
            if seq_index is not None:
                seq += 1
                seq_cell = ws.cell(row=ins_row, column=1)
                seq_cell.value = seq
                # 序号列继承模板行第1列字体，改绿色
                src_font0 = ws.cell(row=template_row, column=1).font
                new_font0 = copy(src_font0) if src_font0 else Font()
                new_font0.color = green
                seq_cell.font = new_font0
            for field in CONTRACT_INFO_FIELDS:
                if field == '已付金额（元）':
                    continue
                # 合同签订金额在新增行中填入（与已有行跳过逻辑相反）
                if field in col_map and field in r_row:
                    cell = ws.cell(row=ins_row, column=col_map[field])
                    cell.value = r_row[field]
                    # 继承模板行对应列的字体，仅改颜色为绿色
                    src_cell = ws.cell(row=template_row, column=col_map[field])
                    new_font = copy(src_cell.font) if src_cell.font else Font()
                    new_font.color = green
                    cell.font = new_font
            if '已付金额（元）' in col_map:
                amt_col = col_map.get('合同签订金额')
                rat_col = col_map.get('累计付款比例(百分比)')
                if amt_col and rat_col:
                    amt = _safe_float(r_row['合同签订金额'])
                    rat = percent_to_decimal(r_row['累计付款比例(百分比)'])
                    paid = amt * rat
                    paid_cell = ws.cell(row=ins_row, column=col_map['已付金额（元）'])
                    paid_cell.value = paid
                    paid_cell.number_format = '#,##0.00'
                    # 继承模板行字体，改绿色
                    src_font2 = ws.cell(row=template_row, column=col_map['已付金额（元）']).font
                    new_font2 = copy(src_font2) if src_font2 else Font()
                    new_font2.color = green
                    paid_cell.font = new_font2
            template_row = ins_row

    result_contracts = set(result['合同号'].astype(str).str.strip())
    yellow_font = Font(color=yellow)
    for seg, info in sheet_info.items():
        ws = info['ws']
        col_map = info['col_map']
        data_start = info['data_start']
        if '合同号' not in col_map:
            continue
        for row in ws.iter_rows(min_row=data_start, max_row=ws.max_row,
                                min_col=col_map['合同号'], max_col=col_map['合同号']):
            cell = row[0]
            if cell.value is not None:
                c = str(cell.value).strip()
                if c and c not in result_contracts:
                    for col_idx in range(1, ws.max_column + 1):
                        target = ws.cell(row=cell.row, column=col_idx)
                        if target.value is not None:
                            new_font = copy(target.font) if target.font else Font()
                            new_font.color = yellow
                            target.font = new_font

    wb.save(summary_path)

    changed_info = "、".join(changed_contracts) if changed_contracts else "无"
    return True, (f"项目 {project_name} 更新完成\n"
                  f"更新合同: {updated_count} 个\n"
                  f"新增合同: {new_count} 个\n"
                  f"变更合同号: {changed_info}")


# ==================== 单项目处理 ====================
def process_project(main_df, sub_dfs, project_name, category, rules,
                    summary_path, device_template_path=None):
    """处理单个项目"""
    try:
        # 1. 筛选主表
        main_name_col = find_field_in_df(main_df,
                                         ['项目名称', '工程名称', '工程项目名称'])
        if not main_name_col:
            return False, (f"主表中未找到项目名称/工程名称列\n"
                           f"文件中的列名: {list(main_df.columns)[:20]}...")

        proj_str = str(project_name).strip()
        clean_proj = clean_project_name(proj_str)
        main_mask = main_df[main_name_col].astype(str).str.strip().apply(
            lambda x: x == proj_str or x == clean_proj)
        main_proj = main_df[main_mask].copy()

        # 湖南-贵州电力灵活互济工程：剔除合同名称中含"迁改工程"的合同
        clean_proj_for_filter = clean_project_name(proj_str)
        if '湖南-贵州电力灵活互济工程' in clean_proj_for_filter:
            contract_name_col = find_field_in_df(main_proj, ['合同名称'])
            if contract_name_col:
                main_proj = main_proj[~main_proj[contract_name_col].astype(str).str.contains('迁改工程')].copy()

        if main_proj.empty:
            return False, (f"项目 \"{project_name}\" 在主表中无匹配数据\n"
                           f"主表项目名列: {main_name_col}\n"
                           f"搜索的项目名: {proj_str} / {clean_proj}\n"
                           f"主表中存在的项目名: {list(main_df[main_name_col].dropna().astype(str).str.strip().unique())[:10]}")

        # 2. 筛选各次表
        sub_proj_dfs = {}
        for table_name, sub_df in sub_dfs.items():
            sub_name_col = find_field_in_df(sub_df,
                                            ['项目名称', '工程名称', '工程项目名称'])
            if sub_name_col:
                sub_mask = sub_df[sub_name_col].astype(str).str.strip().apply(
                    lambda x: x == proj_str or x == clean_proj)
                sub_proj_dfs[table_name] = sub_df[sub_mask].copy()
            else:
                # 次表3没有项目名称列，通过合同编号关联
                contract_col = find_field_in_df(sub_df, ['合同编号', '合同号'])
                if contract_col:
                    main_contracts = set(main_proj['合同编号'].astype(str).str.strip())
                    sub_df_copy = sub_df.copy()
                    sub_df_copy[contract_col] = sub_df_copy[contract_col].astype(str).str.strip()
                    sub_proj_dfs[table_name] = sub_df_copy[
                        sub_df_copy[contract_col].isin(main_contracts)
                    ]
                else:
                    sub_proj_dfs[table_name] = sub_df.copy()

        # 3. 标准化主表列名
        main_proj['合同编号'] = main_proj['合同编号'].astype(str).str.strip()

        # 4. 判断是否多模板处理
        if device_template_path and os.path.exists(device_template_path):
            sub1_proj = sub_proj_dfs.get('次表1', pd.DataFrame())
            (main_line, sub_line), (main_device, sub_device) = \
                split_data_by_sub_project(main_proj, sub1_proj)

            msgs = []
            all_ok = True

            if not main_line.empty:
                line_sub_dfs = {}
                for tn, sdf in sub_proj_dfs.items():
                    if tn == '次表1':
                        line_sub_dfs[tn] = sub_line
                    else:
                        contract_col = find_field_in_df(sdf, ['合同编号', '合同号'])
                        if contract_col:
                            line_contracts = set(
                                main_line['合同编号'].astype(str).str.strip())
                            sdf_copy = sdf.copy()
                            sdf_copy[contract_col] = sdf_copy[contract_col].astype(str).str.strip()
                            line_sub_dfs[tn] = sdf_copy[
                                sdf_copy[contract_col].isin(line_contracts)
                            ]
                        else:
                            line_sub_dfs[tn] = sdf
                result = build_result(main_line, line_sub_dfs, rules, category)
                if not result.empty:
                    ok, msg = _update_template(result, summary_path, project_name)
                    all_ok = all_ok and ok
                    msgs.append(f"[线路材料] {msg}")
                else:
                    msgs.append("[线路材料] 无有效合同数据")
            else:
                msgs.append("[线路材料] 无线路工程数据，跳过")

            if not main_device.empty:
                device_sub_dfs = {}
                for tn, sdf in sub_proj_dfs.items():
                    if tn == '次表1':
                        device_sub_dfs[tn] = sub_device
                    else:
                        contract_col = find_field_in_df(sdf, ['合同编号', '合同号'])
                        if contract_col:
                            device_contracts = set(
                                main_device['合同编号'].astype(str).str.strip())
                            sdf_copy = sdf.copy()
                            sdf_copy[contract_col] = sdf_copy[contract_col].astype(str).str.strip()
                            device_sub_dfs[tn] = sdf_copy[
                                sdf_copy[contract_col].isin(device_contracts)
                            ]
                        else:
                            device_sub_dfs[tn] = sdf
                result = build_result(main_device, device_sub_dfs, rules, category)
                if not result.empty:
                    ok, msg = _update_template(result, device_template_path, project_name)
                    all_ok = all_ok and ok
                    msgs.append(f"[设备] {msg}")
                else:
                    msgs.append("[设备] 无有效合同数据")
            else:
                msgs.append("[设备] 无设备数据，跳过")

            return all_ok, "\n".join(msgs)
        else:
            result = build_result(main_proj, sub_proj_dfs, rules, category)
            if result.empty:
                return False, f"项目 {project_name} 无有效合同数据"
            return _update_template(result, summary_path, project_name)

    except Exception as e:
        tb_info = traceback.format_exc()
        return False, (f"处理项目 \"{project_name}\" 时出错:\n"
                       f"错误: {str(e)}\n"
                       f"详情:\n{tb_info}")


# ==================== 工程列表读取 ====================
def read_engineering_list(filepath):
    """读取工程列表，返回 [(项目名称, 工程类别), ...]"""
    df = pd.read_excel(filepath)
    df.columns = df.columns.astype(str).str.strip()

    name_col = None
    for col in df.columns:
        clean = re.sub(r'[()（）\s]', '', col).lower()
        if '工程项目名称' in clean or '项目名称' in clean:
            name_col = col
            break

    if not name_col:
        raise ValueError("工程列表文件中未找到'工程项目名称'或'项目名称'列")

    category_col = None
    for col in df.columns:
        clean = re.sub(r'[()（）\s]', '', col).lower()
        if '工程类别' in clean or '类别' in clean:
            category_col = col
            break

    projects = []
    seen = set()
    for _, row in df.iterrows():
        name = str(row[name_col]).strip() if pd.notna(row[name_col]) else ''
        if not name:
            continue
        category = str(row[category_col]).strip() if category_col and pd.notna(
            row.get(category_col, '')) else '重特大工程'
        if name not in seen:
            seen.add(name)
            projects.append((name, category))

    return projects


# ==================== GUI ====================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("工程物资一本账更新工具")
        self.root.geometry("700x900")
        self.root.configure(bg="#f0f0f0")

        self.eng_var = tk.StringVar()
        self.main_var = tk.StringVar()
        self.sub1_var = tk.StringVar()
        self.sub2_var = tk.StringVar()
        self.sub3_var = tk.StringVar()
        self.sub4_var = tk.StringVar()
        self.folder_var = tk.StringVar()
        self.time_var = tk.StringVar(value=datetime.now().strftime("%Y%m%d"))

        self.project_data = []

        tk.Label(root, text="工程物资一本账更新工具", font=("微软雅黑", 16, "bold"),
                 bg="#f0f0f0", fg="#2c3e50").pack(pady=15)

        frame = tk.LabelFrame(root, text="文件选择", bg="#f0f0f0", padx=10, pady=10)
        frame.pack(fill="x", padx=20, pady=5)

        self._file_row(frame, "工程列表:", self.eng_var, self.load_engineering)
        self._file_row(frame, "主表(采购合同主单):", self.main_var)
        self._file_row(frame, "次表1(采购合同明细):", self.sub1_var)
        self._file_row(frame, "次表2(合同支付综合):", self.sub2_var)
        self._file_row(frame, "次表3(到货验收):", self.sub3_var)
        self._file_row(frame, "次表4(综合领料):", self.sub4_var)
        self._folder_row(frame, "汇总表文件夹:", self.folder_var)

        time_frame = tk.Frame(frame, bg="#f0f0f0")
        time_frame.pack(fill="x", pady=5)
        tk.Label(time_frame, text="自定义后缀:", bg="#f0f0f0", width=18, anchor="w").pack(side="left")
        tk.Entry(time_frame, textvariable=self.time_var, width=20).pack(side="left", padx=5)
        tk.Label(time_frame, text="（用于输出文件名）",
                 bg="#f0f0f0", fg="#555555", font=("微软雅黑", 9)).pack(side="left", padx=5)

        proj_frame = tk.LabelFrame(root, text="选择工程项目（多选）", bg="#f0f0f0", padx=10, pady=10)
        proj_frame.pack(fill="both", expand=True, padx=20, pady=5)
        list_frame = tk.Frame(proj_frame)
        list_frame.pack(fill="both", expand=True)
        scroll = tk.Scrollbar(list_frame)
        scroll.pack(side="right", fill="y")
        self.listbox = tk.Listbox(list_frame, selectmode=tk.MULTIPLE, yscrollcommand=scroll.set,
                                  font=("微软雅黑", 10), width=50, height=8)
        self.listbox.pack(side="left", fill="both", expand=True)
        scroll.config(command=self.listbox.yview)
        btnf = tk.Frame(proj_frame, bg="#f0f0f0")
        btnf.pack(pady=5)
        tk.Button(btnf, text="全选", command=lambda: self.listbox.select_set(0, tk.END)).pack(
            side="left", padx=5)
        tk.Button(btnf, text="取消全选",
                  command=lambda: self.listbox.selection_clear(0, tk.END)).pack(side="left", padx=5)

        self.run_btn = tk.Button(root, text="开始更新汇总表", command=self.run,
                                 bg="#4CAF50", fg="white", font=("微软雅黑", 12, "bold"),
                                 width=20, height=2)
        self.run_btn.pack(pady=15)

        self.status = tk.StringVar()
        self.status.set("请加载工程列表并选择项目")
        tk.Label(root, textvariable=self.status, bg="#f0f0f0", fg="#555555").pack(pady=5)

    def _file_row(self, parent, label, var, cmd=None):
        f = tk.Frame(parent, bg="#f0f0f0")
        f.pack(fill="x", pady=3)
        tk.Label(f, text=label, bg="#f0f0f0", width=22, anchor="w").pack(side="left")
        e = tk.Entry(f, textvariable=var, width=40)
        e.pack(side="left", padx=5)
        tk.Button(f, text="浏览...",
                  command=cmd if cmd else lambda v=var: self.browse_file(v)).pack(side="left")

    def _folder_row(self, parent, label, var):
        f = tk.Frame(parent, bg="#f0f0f0")
        f.pack(fill="x", pady=3)
        tk.Label(f, text=label, bg="#f0f0f0", width=22, anchor="w").pack(side="left")
        e = tk.Entry(f, textvariable=var, width=40)
        e.pack(side="left", padx=5)
        tk.Button(f, text="浏览...", command=lambda: self.browse_folder(var)).pack(side="left")

    def browse_file(self, var):
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx;*.xls")])
        if path:
            var.set(path)

    def browse_folder(self, var):
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    def load_engineering(self):
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx;*.xls")])
        if not path:
            return
        self.eng_var.set(path)
        try:
            self.project_data = read_engineering_list(path)
            self.listbox.delete(0, tk.END)
            for name, category in self.project_data:
                self.listbox.insert(tk.END, f"[{category}] {name}")
            self.status.set(f"已加载 {len(self.project_data)} 个项目")
        except Exception as e:
            messagebox.showerror("错误", f"加载工程列表失败:\n{str(e)}")

    def run(self):
        main_path = self.main_var.get()
        sub1_path = self.sub1_var.get()
        sub2_path = self.sub2_var.get()
        sub3_path = self.sub3_var.get()
        sub4_path = self.sub4_var.get()
        eng_path = self.eng_var.get()
        folder = self.folder_var.get()
        time_suffix = self.time_var.get().strip()

        missing = []
        if not main_path: missing.append("主表(采购合同主单)")
        if not sub1_path: missing.append("次表1(采购合同明细)")
        if not sub2_path: missing.append("次表2(合同支付综合)")
        if not sub3_path: missing.append("次表3(到货验收)")
        if not sub4_path: missing.append("次表4(综合领料)")
        if not eng_path: missing.append("工程列表")
        if not folder: missing.append("汇总表文件夹")
        if missing:
            messagebox.showerror("错误", f"请选择以下文件/文件夹:\n  " + "\n  ".join(missing))
            return
        if not time_suffix:
            messagebox.showerror("错误", "请填写自定义后缀（例如 20260521）")
            return

        selected = self.listbox.curselection()
        if not selected:
            messagebox.showerror("错误", "请至少选择一个项目")
            return
        projects = [self.project_data[i] for i in selected]

        clean_time = clean_filename(time_suffix)
        if not clean_time:
            clean_time = datetime.now().strftime("%Y%m%d")

        output_dir = os.path.join(folder, "更新结果")
        os.makedirs(output_dir, exist_ok=True)

        self.run_btn.config(state=tk.DISABLED, text="处理中...")
        self.root.update()

        try:
            self.status.set("读取源表数据...")
            self.root.update()

            # 读取所有源表（遍历所有sheet，自动识别表头）
            main_df = read_source_table_all_sheets(main_path)
            sub1_df = read_source_table_all_sheets(sub1_path)
            sub2_df = read_source_table_all_sheets(sub2_path)
            sub3_df = read_source_table_all_sheets(sub3_path)
            sub4_df = read_source_table_all_sheets(sub4_path)

            sub_dfs = {
                '次表1': sub1_df,
                '次表2': sub2_df,
                '次表3': sub3_df,
                '次表4': sub4_df,
            }

            # 标准化主表列名
            for col in main_df.columns:
                clean = re.sub(r'[()（）\s]', '', str(col)).lower()
                if clean == '合同编号' or clean == '合同号':
                    main_df = main_df.rename(columns={col: '合同编号'})
                    break
            else:
                if '合同编号' not in main_df.columns:
                    messagebox.showerror("错误",
                        f"主表中未找到合同编号/合同号列\n文件: {os.path.basename(main_path)}\n现有列: {list(main_df.columns)[:15]}...")
                    return

            # 构建规则（source_field 已在规则中显式指定，无需列字母转换）
            rules_cache = {}
            for category, raw_rules in HARDCODED_RULES.items():
                rules_cache[category] = build_source_field_mapping(raw_rules)

            success = []
            fail = []
            for proj_name, proj_category in projects:
                rules = rules_cache.get(proj_category, rules_cache.get('重特大工程', []))
                if not rules:
                    fail.append(f"{proj_name}: 未找到{proj_category}的规则定义")
                    continue

                template_files = find_all_summary_files(folder, proj_name)
                if not template_files:
                    clean_name = clean_project_name(proj_name)
                    fail.append(f"{proj_name}: 在汇总表文件夹中未找到匹配的模板文件\n  搜索文件夹: {folder}\n  搜索关键词: {clean_name}\n  提示: 模板文件名需包含项目名称(去除括号内容)")
                    continue

                clean_proj = clean_filename(proj_name)

                if len(template_files) >= 2:
                    line_path, device_path = classify_templates(template_files)
                    if not line_path or not device_path:
                        fnames = "\n    ".join(template_files)
                        fail.append(f"{proj_name}: 检测到{len(template_files)}个模板文件，但无法区分为线路材料模板和设备模板\n  匹配到的文件:\n    {fnames}\n  提示: 线路材料模板应包含{SEGMENT_KEYWORDS}之一的sheet名")
                        continue

                    line_tag = extract_template_tag(line_path)
                    device_tag = extract_template_tag(device_path)

                    line_new = os.path.join(output_dir,
                                            f"{clean_proj}{line_tag}物资管理一本账{clean_time}.xlsx")
                    device_new = os.path.join(output_dir,
                                              f"{clean_proj}{device_tag}物资管理一本账{clean_time}.xlsx")

                    shutil.copyfile(line_path, line_new)
                    shutil.copyfile(device_path, device_new)

                    ok, msg = process_project(
                        main_df.copy(), {k: v.copy() for k, v in sub_dfs.items()},
                        proj_name, proj_category, rules,
                        line_new, device_new
                    )
                    if ok:
                        success.append(msg)
                    else:
                        for p in [line_new, device_new]:
                            if os.path.exists(p):
                                os.remove(p)
                        fail.append(f"{proj_name}: {msg}")
                else:
                    original_path = template_files[0]
                    new_filename = f"{clean_proj}物资管理一本账{clean_time}.xlsx"
                    new_path = os.path.join(output_dir, new_filename)
                    shutil.copyfile(original_path, new_path)

                    ok, msg = process_project(
                        main_df.copy(), {k: v.copy() for k, v in sub_dfs.items()},
                        proj_name, proj_category, rules,
                        new_path
                    )
                    if ok:
                        success.append(msg)
                    else:
                        if os.path.exists(new_path):
                            os.remove(new_path)
                        fail.append(f"{proj_name}: {msg}")

            result_text = f"成功 {len(success)} 个项目"
            if fail:
                result_text += "\n\n失败项目:\n" + "\n".join(fail)
            messagebox.showinfo("处理结果", result_text)
            self.status.set("处理完成")
        except Exception as e:
            messagebox.showerror("错误", f"运行失败:\n{str(e)}")
        finally:
            self.run_btn.config(state=tk.NORMAL, text="开始更新汇总表")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
