import re
import sys
import os
import io
import json
import queue
import time
import shutil
import tempfile
import threading
import traceback
from copy import copy
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
from openpyxl.styles import Font
import requests
import urllib3

import matplotlib
matplotlib.use("Agg")  # 无界面后端，打包exe必备
from matplotlib import pyplot as plt
from matplotlib import font_manager
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==================== v8 新增: AI 报告生成 ====================
# 智能体平台(知识服务)调用,链路同"业辅差旅费"已验证:
#   generateAppKey -> v2/analysis(每文件单独上传) -> getFileState轮询 -> v6/chat -> 结构化JSON
#   docx/matplotlib 生成 Word 报告。原 v7 逻辑未改动。
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
    """安全地将值转为float，处理逗号、空格等格式；解析失败时提取文本中第一个数字
    （如 "78498.34（已完成合同部分解除流程，原合同金额90655.2）" → 78498.34）"""
    if val is None or val == '':
        return 0.0
    try:
        s = str(val).strip().replace(',', '').replace('，', '').replace(' ', '').replace('　', '')
        return float(s)
    except (ValueError, TypeError):
        # 混合文本中提取第一个数字
        m = re.search(r'\d+(?:\.\d+)?', str(val))
        if m:
            try:
                return float(m.group())
            except (ValueError, TypeError):
                pass
        return 0.0


def percent_to_decimal(pct_str):
    if not pct_str or pd.isna(pct_str):
        return 0.0
    pct_str = str(pct_str).strip().replace('%', '')
    try:
        return float(pct_str) / 100.0
    except:
        m = re.search(r'\d+(?:\.\d+)?', str(pct_str))
        if m:
            try:
                return float(m.group()) / 100.0
            except (ValueError, TypeError):
                pass
        return 0.0


def _phone_to_str(val):
    """电话/联系人值转字符串：float整数去".0"后缀，空值/NaN转空串"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ''
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    s = str(val).strip()
    if s in ('nan', 'None', ''):
        return ''
    return s


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

            # v6: 次表1统计数量/单位/日期时，筛除物资名称为"备品备件"和"专用工具"的行
            if src_table == '次表1' and rule.get('_original_field') in ('数量', '单位', '现场最终要求交货日期(系统)'):
                mat_col = find_field_in_df(src_match, ['物资名称'])
                if mat_col:
                    src_match = src_match[~src_match[mat_col].astype(str).str.strip().isin(
                        ['备品备件', '专用工具'])]

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
        # 单位字段：取同一合同编号下出现次数最多的单位（众数），明细单位不一致时避免取值漂移
        if '计量单位' in str(field):
            try:
                counts = values.astype(str).value_counts()
                return counts.index[0]
            except Exception:
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
        contact_map = main_df.set_index('_contract')[contact_col].apply(_phone_to_str).to_dict()
        phone_map = main_df.set_index('_contract')[phone_col].apply(_phone_to_str).to_dict()
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
def _update_template(result, summary_path, project_name, excluded=None):
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
        return _process_multi_sheet(wb, result, segment_sheets, summary_path, project_name,
                                    excluded=excluded)
    else:
        return _process_single_sheet(wb, result, summary_path, project_name,
                                     excluded=excluded)


def _process_single_sheet(wb, result, summary_path, project_name, excluded=None):
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

    # v7: 删除"零星采购生成合同"的已有行（物理删除，从下往上避免索引漂移）
    if excluded:
        rows_to_delete = sorted(
            (r for c, r in contract_row_map.items() if c in excluded), reverse=True)
        for r in rows_to_delete:
            ws.delete_rows(r)
        # 重建行映射
        contract_row_map = {}
        for row in ws.iter_rows(min_row=data_start, max_row=ws.max_row,
                                min_col=col_map['合同号'], max_col=col_map['合同号']):
            cell = row[0]
            if cell.value is not None:
                contract_row_map[str(cell.value).strip()] = cell.row
        # 序号列若为连续数字则重排
        seq_vals = []
        for row in ws.iter_rows(min_row=data_start, max_row=ws.max_row,
                                min_col=1, max_col=1):
            v = row[0].value
            if isinstance(v, (int, float)):
                seq_vals.append(v)
        if seq_vals and len(seq_vals) >= 2 and all(
                seq_vals[i] == seq_vals[0] + i for i in range(len(seq_vals))):
            n = int(seq_vals[0])
            for row in ws.iter_rows(min_row=data_start, max_row=ws.max_row,
                                    min_col=1, max_col=1):
                if isinstance(row[0].value, (int, float)):
                    row[0].value = n
                    n += 1

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


def _process_multi_sheet(wb, result, segment_sheets, summary_path, project_name, excluded=None):
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

    # v7: 删除"零星采购生成合同"的已有行（物理删除，从下往上避免索引漂移）
    if excluded:
        for seg, info in sheet_info.items():
            ws = info['ws']
            col_map = info['col_map']
            data_start = info['data_start']
            if '合同号' not in col_map:
                continue
            rows_to_delete = sorted(
                (r for c, r in info['contract_row_map'].items() if c in excluded),
                reverse=True)
            for r in rows_to_delete:
                ws.delete_rows(r)
            # 重建该 sheet 的行映射
            info['contract_row_map'] = {}
            for row in ws.iter_rows(min_row=data_start, max_row=ws.max_row,
                                    min_col=col_map['合同号'], max_col=col_map['合同号']):
                cell = row[0]
                if cell.value is not None:
                    info['contract_row_map'][str(cell.value).strip()] = cell.row
            # 序号列若为连续数字则重排
            seq_vals = []
            for row in ws.iter_rows(min_row=data_start, max_row=ws.max_row,
                                    min_col=1, max_col=1):
                v = row[0].value
                if isinstance(v, (int, float)):
                    seq_vals.append(v)
            if seq_vals and len(seq_vals) >= 2 and all(
                    seq_vals[i] == seq_vals[0] + i for i in range(len(seq_vals))):
                n = int(seq_vals[0])
                for row in ws.iter_rows(min_row=data_start, max_row=ws.max_row,
                                        min_col=1, max_col=1):
                    if isinstance(row[0].value, (int, float)):
                        row[0].value = n
                        n += 1

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

        # v7: 排除"零星采购生成合同"——新增时跳过，已存在于模板中的一并删除
        excluded_contracts = set()
        cat_col = find_field_in_df(main_proj, ['合同类别', '合同类型'])
        if cat_col is not None:
            proj_copy = main_proj.copy()
            # pandas 3.x 的 astype(str) 不会把 NaN 转成字符串，先 fillna 保证集合元素为字符串
            proj_copy['合同编号'] = proj_copy['合同编号'].fillna('').astype(str).str.strip()
            excluded_contracts = set(
                proj_copy.loc[proj_copy[cat_col].astype(str).str.strip() == '零星采购生成合同',
                              '合同编号'].tolist())
        if excluded_contracts:
            main_proj['合同编号'] = main_proj['合同编号'].fillna('').astype(str).str.strip()
            main_proj = main_proj[~main_proj['合同编号'].isin(excluded_contracts)].copy()

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
                    ok, msg = _update_template(result, summary_path, project_name, excluded_contracts)
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
                    ok, msg = _update_template(result, device_template_path, project_name, excluded_contracts)
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
            return _update_template(result, summary_path, project_name, excluded_contracts)

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


# ============================================================
# v8 新增: AI 报告生成 —— 智能体(知识服务)调用 + Word 渲染
# ============================================================

# 智能体平台配置(硬编码,不显示在界面上,最终用户不可见)
# ★ agent_id = 报告智能体的ID(平台"智能体ID"),2026-08-31 已填入
AI_DEFAULT = {
    'base_url': 'https://10.10.65.104:5030/knowledgeService',
    'app_id': '1000292200023',
    'app_secret': 'ffc74446b705407e9afd4e8aaa746bba',
    'agent_id': '696693185',  # 报告智能体ID(2026-08-31 用户提供)
}

# ★提示词已配置在智能体平台(以平台配置为准),此常量仅存档/供平台同步参考,不再随 chat 请求发送
# 业务关注点: ①临近现场要求交货日期→生产及发运进度 ②临近投产时间→物资是否全部到货 (自动提醒不做)
AI_REPORT_PROMPT = """你是一名电网工程物资供应管理数据分析师。请基于上传的Excel文件的实际内容，生成一份《项目物资管理一本账分析报告》。

【当前日期】消息末尾提供"当前日期:YYYY-MM-DD"。所有"剩余天数""临近""逾期"判断必须以此日期为基准，禁止使用任何训练知识中的日期。

【输入文件构成】
- 每个文件 = 一个工程项目的"物资管理一本账"表（工具更新后的成品文件），一个或多个。
- 文件第1行（A1）为标题，可能含投产/投运时间，例如："××工程物资管理一本账信息表（投产时间：2026-11-30）"或"（投运时间：2026年12月25日）"。仅当A1括号内明确给出日期时才认定存在投产时间，否则该文件视为"文件未提供投产时间"；该时间适用于本文件全部明细行。
- 明细行自第4行开始（第2行为分组行、第3行为字段名行），行级字段包括：序号、合同号、合同名称（工程名称）、厂家、数量、单位、合同签订金额、合同要求到货时间、交货日期相关列、生产情况、发运情况、实际到货日期、全部到货验收时间、已付金额、累计付款比例 等。
- **交货日期相关列的列名在不同模板下有差异**（例如改扩建工程模板为"现场最终要求交货日期（实际）"，重特大工程模板为"现场实际要求交货日期"；系统列为"现场最终要求交货日期(系统)"），必须按关键词识别：列名同时含"交货日期"和"实际"的是"实际交货日期"列；含"交货日期"和"系统"的是"系统交货日期"列。
- 文件名与归属以消息末尾"已上传文件"标注为准。

【统计口径（必须严格执行）】
1. 明细行 = 有"合同号"的行；表头重复行、无合同号的汇总/备注行不计入。
2. 交货日期取值优先级：实际交货日期列（列名同时含"交货日期"和"实际"，如"现场最终要求交货日期（实际）"或"现场实际要求交货日期"）→ 系统交货日期列（列名含"交货日期"和"系统"）→ 合同要求到货时间 → 均无则视为"无交货日期"。
3. 生产情况归一为4类（按文字包含判断，原文保留在"其他"或说明中）：
   - 已完成生产：包含"已完成生产""已生产完毕""已生产"
   - 生产中：包含"生产中""开始生产""设计联络会""图纸""排产""未完成"等未完工表述
   - 未开始：无生产信息或明确未开始
   - 其他：无法归类的
4. 发运情况归一为4类：已到货 / 已发运（在途）/ 未发运（含空白）/ 其他（保留原文）。
5. 比例统计一律用"物资条目数（明细行数）"口径；合同金额可另列统计。各行数量单位不同（套/台/吨/米/千米等），禁止直接相加。
6. 剩余天数 = 交货日期 − 当前日期（按天）。分档：已逾期（剩余<0）/ 30天内到期（0≤剩余≤30）/ 31-60天 / 60天以上 / 无交货日期。
7. 文件里没有的数据必须写"文件未提供"，不得假设、推算或编造；计算必须可复核（每类统计都要能在明细表中找到对应条目）。
8. 金额列若为混合文本（例如"78498.34（已完成合同部分解除流程，原合同金额90655.2）"），取其中第一个数字作为该行金额，原始文本要在说明中标明。
9. 数量、单位、监造/抽检数量等列为"/""—"或空时，视为文件未提供；明细表与正文中对应数量、单位一律输出"未提供"。**严禁输出任何"数字+单位"组合**（例如"1台""1套""1米"），无论合同标题、设备名称、常识再像，都不允许猜测数量或单位。

【报告内容要求（围绕业务人员核心关注点）】
1. 总体概况：工程项目数、物资明细条目数、合同总金额、整体到货率（已到货条目数÷明细条目数）、付款进度概览。
2. 生产进度分析：生产情况4类比例（条目数）及主要问题，重点说明未完工条目的生产状态原文。
3. 发运与到货进度分析：发运情况4类比例（条目数）、已发运未到货条目数、实际到货日期晚于交货日期（逾期到货）的条目数。
4. ★交货日期临近风险（核心）：仅统计尚未到货的物资（发运情况≠已到货，含已发运/在途/未发运），逐档列出"已逾期 / 30天内到期 / 31-60天"的物资明细表（列：序号、合同号、厂家、数量、单位、交货日期、剩余天数、生产情况、发运情况、实际到货日期），每档一段文字说明主要风险与所属工程。已到货的条目不纳入本清单；对已到货但交货日期已过的条目，在"发运与到货进度分析"中统计"逾期到货"（实际到货日期晚于交货日期）条数与清单。
5. ★投产前到货情况（核心）：按文件A1的投产/投运时间（未提供则写"文件未提供投产时间"）：投产时间在报告日期之后的，列出"投产时间前仍未到货（发运情况≠已到货）"的物资明细；投产时间已过的，写明"投产时间已于××到期，仍有 ×× 条物资未到货"。
6. 存在问题与建议：问题须能从数据直接得出，避免空泛套话；不要写"系统自动提醒"之类无法由本报告实现的内容。

【JSON输出结构】只输出一个JSON对象（思考过程如需请放<think>标签；markdown代码块、多余说明文字都不要出现）：
{
  "title": "报告标题（含工程/报告范围说明）",
  "summary": ["摘要要点1", "摘要要点2"],
  "sections": [
    {
      "heading": "章节标题",
      "paragraphs": ["段落文字1", "段落文字2"],
      "table": {"caption": "表标题", "columns": ["列1", "列2"], "rows": [["值1", "值2"], ["值1", "值2"]]}
    }
  ],
  "charts": [
    {
      "caption": "图表标题",
      "type": "pie",
      "categories": ["类别1", "类别2"],
      "series": [{"name": "系列名", "values": [100, 200]}]
    }
  ],
  "conclusion": ["结论1", "结论2"]
}
图表规则：
- type：bar柱状图 / pie饼图 / line折线图；柱状/折线支持多个series，pie只用第一个series。
- 至少3张：①生产情况分布饼图（已完成生产/生产中/未开始/其他，values=条目数）；②发运情况分布饼图（已到货/已发运/未发运/其他，values=条目数）；③交货日期到期分布柱状图（仅统计未到货条目，与第4章口径一致；categories=[已逾期,30天内,31-60天,60天以上,无交货日期]，series=[{"name":"条目数","values":[...]}]）。
- 图表数据点必须从文件实际内容汇总得出，与正文、表格口径一致。
- table和charts可为空数组，但title/summary/conclusion必须有内容。"""


class AIReportError(RuntimeError):
    """AI报告环节的异常"""
    pass


_AI_LOG_PATH = None


def _ai_log(msg):
    """AI报告调试日志——2026-09-01 已完成乱码定位(v3接口全部正常),日志不再生成"""
    pass


def get_app_key(base_url, app_id, app_secret):
    """获取接口凭证appKey(header认证用)"""
    _ai_log(f"get_app_key: url={base_url} appId={app_id}")
    try:
        r = requests.post(f"{base_url}/extSecret/generateAppKey",
                          json={"appId": app_id, "appSecret": app_secret},
                          verify=False, timeout=30)
    except requests.exceptions.RequestException as e:
        _ai_log(f"get_app_key FAIL: {e}")
        raise AIReportError(f"无法连接智能体平台: {e}")
    try:
        j = json.loads(r.content.decode('utf-8'))
    except Exception:
        _ai_log(f"get_app_key 非JSON(响应头Content-Type={r.headers.get('Content-Type')!r}): {r.content[:200]!r}")
        raise AIReportError(f"平台返回非JSON: {r.content.decode('utf-8', 'ignore')[:200]}")
    obj = j.get('resultObject') or {}
    key = obj.get('appKey') if isinstance(obj, dict) else None
    if not key:
        _ai_log(f"get_app_key 无appKey: {json.dumps(j, ensure_ascii=False)[:200]}")
        raise AIReportError(f"获取appKey失败: {json.dumps(j, ensure_ascii=False)[:200]}")
    _ai_log(f"get_app_key OK: {key[:12]}...")
    return key


def upload_file_to_agent(base_url, headers, agent_id, file_path):
    """上传单个文件到智能体平台,返回上传结果对象(fileId/fileState)"""
    _ai_log(f"upload start: file={file_path!r} agentId={agent_id}")
    try:
        with open(file_path, 'rb') as f:
            r = requests.post(f"{base_url}/extChatApi/v2/analysis",
                              headers=headers,
                              data={'agentId': str(agent_id), 'extractFile': '0'},
                              files={'files': (os.path.basename(file_path), f)},
                              verify=False, timeout=300)
    except requests.exceptions.RequestException as e:
        _ai_log(f"upload FAIL: {e}")
        raise AIReportError(f"上传文件失败: {e}")
    try:
        j = json.loads(r.content.decode('utf-8'))
    except Exception:
        _ai_log(f"upload 非JSON: {r.content[:200]!r}")
        raise AIReportError(f"上传返回非JSON: {r.content.decode('utf-8', 'ignore')[:200]}")
    _ai_log(f"upload resp: {json.dumps(j, ensure_ascii=False)[:300]}")
    objs = j.get('resultObject') or []
    if not objs or not objs[0].get('fileId'):
        _ai_log(f"upload 无fileId: {json.dumps(j, ensure_ascii=False)[:300]}")
        raise AIReportError(f"上传未返回fileId: {json.dumps(j, ensure_ascii=False)[:200]}")
    _ai_log(f"upload OK fileId={objs[0].get('fileId')}")
    return objs[0]


def wait_file_ready(base_url, headers, agent_id, file_id, timeout_s=180):
    """轮询文件解析状态: 3=完成, -1=失败, 其他等待; 超时抛异常"""
    deadline = time.time() + timeout_s
    last = '0'
    while time.time() < deadline:
        try:
            r = requests.post(f"{base_url}/extChatApi/getFileState",
                              headers={**headers, 'Content-Type': 'application/json'},
                              json={"agentId": str(agent_id), "fileIds": [int(file_id)]},
                              verify=False, timeout=30)
            j = json.loads(r.content.decode('utf-8'))
        except Exception as e:
            raise AIReportError(f"查询文件状态失败: {e}")
        objs = j.get('resultObject') or []
        st = str(objs[0].get('fileState')) if objs else ''
        if st != last:
            _ai_log(f"fileState={st} fileId={file_id}")
        last = st
        if st == '3':
            _ai_log(f"fileState OK fileId={file_id}")
            return
        if st == '-1':
            _ai_log(f"fileState FAIL(-1) fileId={file_id}")
            raise AIReportError("文件解析失败(平台无法识别该Excel,可能格式不支持)")
        time.sleep(3)
    _ai_log(f"fileState TIMEOUT fileId={file_id} last={last}")
    raise AIReportError(f"文件解析超时({timeout_s}秒,最后状态:{last})")


def chat_with_agent(base_url, headers, agent_id, file_id, prompt, file_name):
    """调用v3/chat,带文件fileId,非流式,返回智能体文本
    ★2026-09-01 改v3:平台v6非流式有BUG(响应中文全被替换成'?'),
      麒麟机四交叉实测 v6+stream:false 乱码 / v6+stream:true 正常 / v1+非流式 正常 /
      v3+非流式+全参数 正常(2026-09-01实测),v3支持files传参且请求体结构与v6一致"""
    # 提示词已配置在智能体平台(见顶部 AI_REPORT_PROMPT 存档),此处只发当前日期+文件引用
    body = {
        "top_p": 0.9, "frequency_penalty": 0.5, "max_tokens": 20000,
        "presence_penalty": 0.5, "temperature": 0.7,
        "messages": [{
            "role": "user",
            "content": f"[当前日期: {datetime.now().strftime('%Y-%m-%d')}]\n[已上传文件: {file_name}]",
            "files": [int(file_id)],
        }],
        "agentId": str(agent_id), "stream": False,
    }
    _ai_log(f"chat send: {json.dumps(body, ensure_ascii=False)[:300]}")
    try:
        r = requests.post(f"{base_url}/extChatApi/v3/chat",
                          headers={**headers, 'Content-Type': 'application/json'},
                          json=body, verify=False, timeout=600)
    except requests.exceptions.RequestException as e:
        _ai_log(f"chat FAIL: {e}")
        raise AIReportError(f"调用智能体失败: {e}")
    _ai_log(f"chat http={r.status_code} Content-Type={r.headers.get('Content-Type')!r} "
            f"Content-Encoding={r.headers.get('Content-Encoding')!r} bytes={len(r.content)}")
    if r.status_code != 200:
        raise AIReportError(f"chat HTTP {r.status_code}: {r.content.decode('utf-8', 'ignore')[:200]}")
    try:
        j = json.loads(r.content.decode('utf-8'))
    except Exception:
        _ai_log(f"chat 非JSON, 前300字节: {r.content[:300]!r}")
        raise AIReportError(f"智能体返回非JSON: {r.content.decode('utf-8', 'ignore')[:200]}")
    text = ''
    # v6 格式: choices[0].message.content
    if j.get('choices') and isinstance(j['choices'], list) and j['choices'][0] \
            and j['choices'][0].get('message', {}).get('content'):
        text = j['choices'][0]['message']['content']
    # 兼容旧 v1 格式: choices.response[]
    elif j.get('choices', {}).get('response'):
        for item in reversed(j['choices']['response']):
            c = item.get('content')
            if c and c.strip() and c.strip() != '{}' and not c.strip().startswith('<think>'):
                text = c
                break
    if not text:
        _ai_log(f"chat 返回为空: {json.dumps(j, ensure_ascii=False)[:300]}")
        raise AIReportError(f"智能体返回为空: {json.dumps(j, ensure_ascii=False)[:200]}")
    _ai_log(f"chat OK len={len(text)} '?'数={text.count('?')} 前200: {text[:200]!r}")
    return text


def parse_report_json(text):
    """剥离<think>思考块和markdown代码块,解析JSON(含截取{}容错)"""
    ti = text.find('<think>')
    te = text.rfind('</think>')
    if ti >= 0:
        if te > ti:
            text = (text[:ti] + text[te + 8:]).strip()
        else:
            text = text[:ti].strip()
    text = text.replace('```json', '').replace('```JSON', '').replace('```', '').strip()
    try:
        return json.loads(text)
    except Exception:
        pass
    lb, rb = text.find('{'), text.rfind('}')
    if lb >= 0 and rb > lb:
        try:
            return json.loads(text[lb:rb + 1])
        except Exception:
            pass
    # 兜底: 返回夹带思考文字/多段内容时,扫描每个'{'取第一个合法JSON块
    decoder_obj = json.JSONDecoder()
    for i in range(len(text)):
        if text[i] == '{':
            try:
                obj, _ = decoder_obj.raw_decode(text[i:])
                return obj
            except Exception:
                continue
    raise AIReportError("智能体返回内容不是合法JSON，无法生成报告")


# ---------- 图表 ----------

_CJK_FONT_CANDIDATES = ["Microsoft YaHei", "SimHei", "WenQuanYi Zen Hei",
                        "WenQuanYi Micro Hei", "Noto Sans CJK SC", "Source Han Sans CN"]


def pick_cjk_font():
    """选取本机可用的中文字体(Windows优先微软雅黑, 麒麟优先文泉驿)"""
    installed = {f.name for f in font_manager.fontManager.ttflist}
    for name in _CJK_FONT_CANDIDATES:
        if name in installed:
            return name
    return None


def _to_float(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(',', '').replace('%', '').strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def generate_chart(spec, out_png, cjk_font):
    """按智能体返回的图表规格生成PNG: bar/pie/line"""
    typ = (spec.get('type') or 'bar').lower()
    caption = spec.get('caption', '')
    cats = [str(c) for c in (spec.get('categories') or [])]
    series = spec.get('series') or []
    if not series or not series[0].get('values'):
        raise ValueError("图表没有数据")
    if cjk_font:
        plt.rcParams['font.sans-serif'] = [cjk_font]
        plt.rcParams['axes.unicode_minus'] = False

    fig, ax = plt.subplots(figsize=(8, 4.2), dpi=150)
    if typ == 'pie':
        s = series[0]
        values = [_to_float(v) for v in s.get('values') or []]
        ax.pie(values, labels=cats[:len(values)], autopct='%1.1f%%', startangle=90,
               counterclock=False)
        ax.axis('equal')
    elif typ == 'line':
        x = range(len(cats))
        for s in series:
            values = [_to_float(v) for v in s.get('values') or []]
            ax.plot(list(x)[:len(values)], values, marker='o', label=str(s.get('name', '')))
        ax.set_xticks(list(x))
        ax.set_xticklabels(cats, rotation=30, ha='right')
        ax.legend()
        ax.grid(axis='y', alpha=0.3)
    else:  # bar
        n = len(cats)
        width = 0.8 / max(len(series), 1)
        x = list(range(n))
        for i, s in enumerate(series):
            values = [_to_float(v) for v in s.get('values') or []]
            ax.bar([xi + i * width for xi in x][:len(values)], values, width,
                   label=str(s.get('name', '')))
        if len(series) > 1:
            off = width * (len(series) - 1) / 2
            labels_pos = [xi + off for xi in x]
        else:
            labels_pos = x
        ax.set_xticks(labels_pos)
        ax.set_xticklabels(cats, rotation=30, ha='right')
        ax.legend()
        ax.grid(axis='y', alpha=0.3)
    if caption:
        ax.set_title(caption, fontsize=11)
    fig.tight_layout()
    fig.savefig(out_png, bbox_inches='tight')
    plt.close(fig)


# ---------- Word 渲染 ----------

def _set_cjk_style_fonts(doc, cjk_font):
    """正文+标题样式设置中文字体"""
    name = cjk_font or 'Microsoft YaHei'
    normal = doc.styles['Normal']
    normal.font.name = name
    normal.font.size = Pt(11)
    normal.element.rPr.rFonts.set(qn('w:eastAsia'), name)
    for lvl in range(0, 4):
        st = doc.styles[f'Heading {lvl}'] if lvl else doc.styles['Title']
        st.font.name = name
        st.element.rPr.rFonts.set(qn('w:eastAsia'), name)


def _add_center_caption(doc, text, size=9, italic=False, after_blank=None):
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = cap.add_run(text)
    run.font.size = Pt(size)
    run.font.italic = italic
    return cap


def render_word(report, docx_path, chart_tmp_dir):
    """把智能体返回的结构化JSON渲染成Word报告(章节/表格/图表/结论)"""
    if not isinstance(report, dict):
        raise AIReportError("报告JSON不是对象结构")
    cjk_font = pick_cjk_font()
    doc = Document()
    _set_cjk_style_fonts(doc, cjk_font)

    title = str(report.get('title') or os.path.splitext(os.path.basename(docx_path))[0])
    h = doc.add_heading(title, level=0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

    summary = report.get('summary') or []
    if summary:
        doc.add_heading('摘要', level=1)
        for p in summary:
            doc.add_paragraph(str(p))

    for sec in report.get('sections') or []:
        if not isinstance(sec, dict):
            continue
        doc.add_heading(str(sec.get('heading') or ''), level=1)
        for p in sec.get('paragraphs') or []:
            doc.add_paragraph(str(p))
        table = sec.get('table')
        if table and table.get('columns'):
            cols = [str(c) for c in table['columns']]
            rows = table.get('rows') or []
            t = doc.add_table(rows=1 + len(rows), cols=len(cols), style='Table Grid')
            t.alignment = WD_TABLE_ALIGNMENT.CENTER
            for j, c in enumerate(cols):
                cell = t.cell(0, j)
                cell.text = ''
                run = cell.paragraphs[0].add_run(c)
                run.bold = True
            for i, row in enumerate(rows):
                for j in range(len(cols)):
                    t.cell(i + 1, j).text = str(row[j] if j < len(row) else '')
            if table.get('caption'):
                _add_center_caption(doc, f"表：{table['caption']}")

    charts = report.get('charts') or []
    for idx, spec in enumerate(charts, 1):
        if not isinstance(spec, dict):
            continue
        png = os.path.join(chart_tmp_dir, f'chart_{idx}.png')
        try:
            generate_chart(spec, png, cjk_font)
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.add_run().add_picture(png, width=Inches(6.0))
            _add_center_caption(doc, f"图{idx}：{spec.get('caption', '')}")
        except Exception as e:
            doc.add_paragraph(f"(图表生成失败: {e})")

    conclusion = report.get('conclusion') or []
    if conclusion:
        doc.add_heading('结论', level=1)
        for p in conclusion:
            doc.add_paragraph(str(p))

    gen = doc.add_paragraph()
    gen.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = gen.add_run(f'报告生成时间：{datetime.now().strftime("%Y-%m-%d")}')
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    doc.save(docx_path)


def process_one_file(file_path, base_url, app_id, app_secret, agent_id, prompt, on_status=None):
    """单个汇总表文件的全流程: 上传->等解析->chat->渲染Word。在子线程中调用。"""
    def status(msg):
        if on_status:
            on_status(msg)

    base = os.path.basename(file_path)
    status(f"连接平台... {base}")
    key = get_app_key(base_url, app_id, app_secret)
    headers = {'appId': app_id, 'appKey': key}

    status(f"上传文件... {base}")
    up = upload_file_to_agent(base_url, headers, agent_id, file_path)
    status(f"{base}: 平台解析中...")
    wait_file_ready(base_url, headers, agent_id, up['fileId'])

    status(f"{base}: 智能体分析中...")
    text = chat_with_agent(base_url, headers, agent_id, up['fileId'], prompt, base)
    report = parse_report_json(text)
    _ai_log(f"parse OK title={str(report.get('title', ''))[:60]!r} keys={list(report.keys())}")
    status(f"{base}: 生成Word报告...")

    out_dir = os.path.join(os.path.dirname(os.path.abspath(file_path)), "AI报告")
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(base)[0]
    docx_path = os.path.join(out_dir, f"{stem}_AI分析报告_{datetime.now().strftime('%Y%m%d')}.docx")
    tmp_dir = tempfile.mkdtemp(prefix="ai_charts_")
    try:
        render_word(report, docx_path, tmp_dir)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    _ai_log(f"render done: {docx_path}")
    return True, f"{base} 报告已生成", docx_path


# ==================== GUI ====================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("电网工程物资一本账智能管理工具")
        self.root.configure(bg="#f0f0f0")

        self.eng_var = tk.StringVar()
        self.main_var = tk.StringVar()
        self.sub1_var = tk.StringVar()
        self.sub2_var = tk.StringVar()
        self.sub3_var = tk.StringVar()
        self.sub4_var = tk.StringVar()
        self.folder_var = tk.StringVar()
        self.time_var = tk.StringVar(value=datetime.now().strftime("%Y%m%d"))

        # v8: AI报告生成(智能体配置硬编码,不暴露给用户)
        self.report_files = []
        self.report_status_var = tk.StringVar(value="AI报告: 选择汇总表文件后点击生成")
        self._report_queue = queue.Queue()
        self._report_results = []
        self._report_total = 0
        self._report_done = 0
        self._report_dirs = []   # v9: 各份报告实际输出目录(打开输出目录用)

        self.project_data = []

        tk.Label(root, text="电网工程物资一本账智能管理工具", font=("微软雅黑", 14, "bold"),
                 bg="#f0f0f0", fg="#2c3e50").pack(pady=(8, 4))

        frame = tk.LabelFrame(root, text="文件选择", bg="#f0f0f0", padx=10, pady=6)
        frame.pack(fill="x", padx=20, pady=3)

        self._file_row(frame, "工程列表:", self.eng_var, self.load_engineering)
        self._file_row(frame, "主表(采购合同主单):", self.main_var)
        self._file_row(frame, "次表1(采购合同明细):", self.sub1_var)
        self._file_row(frame, "次表2(合同支付综合):", self.sub2_var)
        self._file_row(frame, "次表3(到货验收):", self.sub3_var)
        self._file_row(frame, "次表4(综合领料):", self.sub4_var)
        self._folder_row(frame, "汇总表文件夹:", self.folder_var)

        time_frame = tk.Frame(frame, bg="#f0f0f0")
        time_frame.pack(fill="x", pady=3)
        tk.Label(time_frame, text="自定义后缀:", bg="#f0f0f0", width=18, anchor="w").pack(side="left")
        tk.Entry(time_frame, textvariable=self.time_var).pack(side="left", fill="x",
                                                              expand=True, padx=5)
        tk.Label(time_frame, text="（用于输出文件名）",
                 bg="#f0f0f0", fg="#555555", font=("微软雅黑", 9)).pack(side="left", padx=5)

        proj_frame = tk.LabelFrame(root, text="选择工程项目（多选）", bg="#f0f0f0", padx=10, pady=6)
        proj_frame.pack(fill="both", expand=True, padx=20, pady=3)
        list_frame = tk.Frame(proj_frame)
        list_frame.pack(fill="both", expand=True)
        scroll = tk.Scrollbar(list_frame)
        scroll.pack(side="right", fill="y")
        self.listbox = tk.Listbox(list_frame, selectmode=tk.MULTIPLE, yscrollcommand=scroll.set,
                                  font=("微软雅黑", 10), width=60, height=8)
        self.listbox.pack(side="left", fill="both", expand=True)
        scroll.config(command=self.listbox.yview)
        btnf = tk.Frame(proj_frame, bg="#f0f0f0")
        btnf.pack(pady=3)
        tk.Button(btnf, text="全选", command=lambda: self.listbox.select_set(0, tk.END)).pack(
            side="left", padx=5)
        tk.Button(btnf, text="取消全选",
                  command=lambda: self.listbox.selection_clear(0, tk.END)).pack(side="left", padx=5)

        self.run_btn = tk.Button(root, text="开始更新汇总表", command=self.run,
                                 bg="#4CAF50", fg="white", font=("微软雅黑", 12, "bold"),
                                 width=20, height=2)
        self.run_btn.pack(pady=(6, 8))

        # ---------- v8: AI报告生成区(智能体配置硬编码,界面上不显示) ----------
        ai_frame = tk.LabelFrame(root, text="AI 报告生成（智能体分析汇总表 → Word 报告）",
                                 bg="#f0f0f0", padx=10, pady=6)
        ai_frame.pack(fill="x", padx=20, pady=3)

        filef = tk.Frame(ai_frame, bg="#f0f0f0")
        filef.pack(fill="x", pady=2)
        tk.Button(filef, text="选择汇总表文件...", command=self.pick_report_files).pack(
            side="left", padx=5)
        tk.Label(filef, text="(可多选已生成的汇总表xlsx)",
                 bg="#f0f0f0", fg="#555555", font=("微软雅黑", 9)).pack(side="left")
        tk.Button(filef, text="清空", command=self.clear_report_files).pack(
            side="right", padx=5)
        rlist_frame = tk.Frame(ai_frame)
        rlist_frame.pack(fill="x", pady=2)
        rscroll = tk.Scrollbar(rlist_frame)
        rscroll.pack(side="right", fill="y")
        self.report_listbox = tk.Listbox(rlist_frame, height=3,
                                         yscrollcommand=rscroll.set, font=("微软雅黑", 10))
        self.report_listbox.pack(side="left", fill="both", expand=True)
        rscroll.config(command=self.report_listbox.yview)

        btn_rows = tk.Frame(root, bg="#f0f0f0")
        btn_rows.pack(pady=(4, 3))
        self.report_btn = tk.Button(btn_rows, text="AI生成报告", command=self.run_report,
                                    bg="#2196F3", fg="white", font=("微软雅黑", 11, "bold"),
                                    width=20)
        self.report_btn.pack(side="left", padx=5)
        tk.Button(btn_rows, text="打开输出目录", command=self.open_report_dir,
                  font=("微软雅黑", 10), width=14).pack(side="left", padx=5)

        self.status = tk.StringVar()
        self.status.set("请加载工程列表并选择项目")
        tk.Label(root, textvariable=self.status, bg="#f0f0f0", fg="#555555").pack(pady=1)
        tk.Label(root, textvariable=self.report_status_var, bg="#f0f0f0",
                 fg="#2196F3", font=("微软雅黑", 9)).pack(pady=1)

        # 窗口自适应:高度按内容需求,不超过屏幕可用高度(任务栏+标题栏留余),保证全界面可见
        self.root.update_idletasks()
        req_h = self.root.winfo_reqheight()
        avail_h = self.root.winfo_screenheight() - 120
        win_h = min(req_h, avail_h)
        self.root.geometry(f"900x{win_h}")
        self.root.minsize(760, win_h)

        self.root.after(200, self._report_poll)

    def _file_row(self, parent, label, var, cmd=None):
        f = tk.Frame(parent, bg="#f0f0f0")
        f.pack(fill="x", pady=2)
        tk.Label(f, text=label, bg="#f0f0f0", width=22, anchor="w").pack(side="left")
        e = tk.Entry(f, textvariable=var)  # 无固定宽度,随窗口拉伸
        e.pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(f, text="浏览...",
                  command=cmd if cmd else lambda v=var: self.browse_file(v)).pack(side="left")

    def _folder_row(self, parent, label, var):
        f = tk.Frame(parent, bg="#f0f0f0")
        f.pack(fill="x", pady=2)
        tk.Label(f, text=label, bg="#f0f0f0", width=22, anchor="w").pack(side="left")
        e = tk.Entry(f, textvariable=var)
        e.pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(f, text="浏览...", command=lambda: self.browse_folder(var)).pack(side="left")

    def browse_file(self, var):
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx;*.xls")])
        if path:
            var.set(path)

    def browse_folder(self, var):
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    # ---------- v8: AI报告生成 ----------
    def pick_report_files(self):
        paths = filedialog.askopenfilenames(
            title="选择已生成的汇总表文件(可多选)",
            filetypes=[("Excel文件", "*.xlsx;*.xls")])
        if paths:
            self.report_files = list(paths)
            self.report_listbox.delete(0, tk.END)
            for p in paths:
                self.report_listbox.insert(tk.END, os.path.basename(p))
            self.report_status_var.set(
                f"AI报告: 已选择 {len(paths)} 个文件")

    def clear_report_files(self):
        self.report_files = []
        self.report_listbox.delete(0, tk.END)
        self.report_status_var.set("AI报告: 选择汇总表文件后点击生成")

    def run_report(self):
        files = list(self.report_files)
        if not files:
            messagebox.showerror("错误", "请先选择汇总表文件")
            return
        base_url = AI_DEFAULT['base_url']
        app_id = AI_DEFAULT['app_id']
        app_secret = AI_DEFAULT['app_secret']
        agent_id = AI_DEFAULT['agent_id']
        if not agent_id:
            messagebox.showerror("错误", "当前版本尚未配置智能体ID，请联系管理员")
            return

        self._report_results = []
        self._report_total = len(files)
        self._report_done = 0
        self.report_btn.config(state=tk.DISABLED, text="生成中...")
        self.report_status_var.set(f"AI报告: 开始生成 {len(files)} 份报告...")

        def worker():
            def on_status(msg):
                self._report_queue.put(('status', msg))

            results = []
            with ThreadPoolExecutor(max_workers=min(5, len(files))) as ex:
                futs = {ex.submit(process_one_file, f, base_url, app_id,
                                  app_secret, agent_id, AI_REPORT_PROMPT,
                                  on_status): f for f in files}
                for fut in as_completed(futs):
                    f = futs[fut]
                    try:
                        ok, msg, docx_path = fut.result()
                        results.append((f, ok, msg, docx_path))
                    except Exception as e:
                        results.append((f, False, f"失败: {e}", None))
                    self._report_queue.put(('done', None))
            self._report_queue.put(('finish', results))

        threading.Thread(target=worker, daemon=True).start()

    def _report_poll(self):
        """主线程轮询后台队列,更新进度/弹结果(避免跨线程操作tkinter)"""
        try:
            while True:
                kind, payload = self._report_queue.get_nowait()
                if kind == 'status':
                    self.report_status_var.set(f"AI报告: {payload}")
                elif kind == 'done':
                    self._report_done += 1
                    self.report_status_var.set(
                        f"AI报告: 进度 {self._report_done}/{self._report_total}")
                elif kind == 'finish':
                    self.report_btn.config(state=tk.NORMAL, text="AI生成报告")
                    self._show_report_results(payload)
        except queue.Empty:
            pass
        self.root.after(200, self._report_poll)

    def _show_report_results(self, results):
        # v9: 记录本批成功报告的输出目录("打开输出目录"按钮用,对齐智能办公应用)
        self._report_dirs = [os.path.dirname(p) for _, ok, _, p in results if ok and p]
        ok_msgs, fail_msgs = [], []
        for f, ok, msg, path in results:
            if ok:
                ok_msgs.append(f"✔ {os.path.basename(f)}\n    → {path}")
            else:
                fail_msgs.append(f"✘ {os.path.basename(f)}\n    {msg}")
        text = f"成功 {len(ok_msgs)} 份 / 共 {len(results)}"
        if ok_msgs:
            text += "\n\n已生成:\n" + "\n".join(ok_msgs)
        if fail_msgs:
            text += "\n\n失败: \n" + "\n".join(fail_msgs)
        self.report_status_var.set(f"AI报告: 完成 {len(ok_msgs)} 份")
        if fail_msgs:
            messagebox.showwarning("AI报告生成完成", text)
        else:
            messagebox.showinfo("AI报告生成完成", text)

    def open_report_dir(self):
        """v9: 打开最近一次AI报告输出目录(各文件同目录的"AI报告"子目录),对齐智能办公应用"""
        dirs = [d for d in self._report_dirs if os.path.isdir(d)]
        if not dirs:
            messagebox.showinfo("提示", "还没有AI报告产物，请先生成报告")
            return
        d = dirs[-1]
        self.report_status_var.set(f"AI报告: 正在打开输出目录 {d}")
        try:
            import subprocess
            if sys.platform == "win32":
                # ★explorer优先:中文路径下 os.startfile 在部分环境(受限桌面/策略)静默无响应
                try:
                    subprocess.Popen(["explorer", d])
                    return
                except Exception:
                    pass
                os.startfile(d)
            else:
                subprocess.Popen(["xdg-open", d])
        except Exception as e:
            messagebox.showerror("打开输出目录失败", str(e))

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
    # DPI感知:否则Windows按缩放放大窗口导致超屏/内容被裁(用户反馈界面看不全的根因)
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass
    root = tk.Tk()
    app = App(root)
    root.mainloop()
