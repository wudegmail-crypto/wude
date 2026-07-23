import re
import os
import sys
import shutil
import traceback
from copy import copy
from datetime import datetime

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
from openpyxl.styles import Font

# ==================== 核心配置 ====================
UPDATE_FIELDS = [
    "合同号", "合同名称", "厂家", "数量", "单位", "合同签订金额",
    "付款比例", "合同签订日期", "供应商联系人及电话", "合同要求到货时间",
    "物资名称", "合同最终结算金额", "累计付款比例(百分比)", "已付金额（元）"
]

# 汇总表可能使用的别名
FIELD_ALIASES = {
    '合同签订金额': ['合同签订金额', '合同金额', '合同金额(元)', '合同签订金额(元)'],
}

# ==================== 工具函数 ====================
def safe_date_parse(date_str, fmt="%Y-%m-%d"):
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

def merge_materials(group):
    materials = []
    for _, row in group.iterrows():
        name = str(row.get('物资名称', '')).strip()
        qty = str(row.get('采购数量', '')).strip()
        if name and qty:
            materials.append(f"{name}({qty})")
    return "，".join(materials) if materials else ""

def aggregate_qty_unit(group):
    if group.empty or '计量单位' not in group.columns or '采购数量' not in group.columns:
        return 0, ""
    unit_sums = group.groupby('计量单位')['采购数量'].sum()
    if unit_sums.empty:
        return 0, ""
    best_unit = unit_sums.idxmax()
    best_qty = unit_sums.max()
    return best_qty, best_unit

def get_quantity_display(group):
    if group.empty or '物资名称' not in group.columns or '采购数量' not in group.columns:
        return ""
    material_groups = group.groupby('物资名称')['采购数量'].sum().reset_index()
    if material_groups.empty:
        return ""
    if len(material_groups) == 1:
        return str(material_groups.iloc[0]['采购数量'])
    lines = []
    for _, mat_row in material_groups.iterrows():
        name = str(mat_row['物资名称']).strip()
        qty = mat_row['采购数量']
        if not name:
            continue
        if ',' in name:
            name = name.split(',')[0].strip()
        elif '，' in name:
            name = name.split('，')[0].strip()
        if len(name) > 20:
            name = name[:20] + '…'
        lines.append(f"{name}（{qty}）")
    return "\n".join(lines)

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

def percent_to_decimal(pct_str):
    """将百分比字符串转换成小数，例如'20%' -> 0.2"""
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
        old_num = float(old_val)
        new_num = float(new_val)
        return old_num != new_num
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
    for col in range(1, max_col+1):
        src = ws.cell(row=from_row, column=col)
        dst = ws.cell(row=to_row, column=col)
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format

def get_first_column_as_sequence(ws, data_start, max_row):
    header_cell = ws.cell(row=data_start-1, column=1)
    if header_cell.value:
        clean_col = re.sub(r'[()（）\s]', '', str(header_cell.value)).lower()
        for field in UPDATE_FIELDS:
            if re.sub(r'[()（）\s]', '', field).lower() == clean_col:
                return None
    if max_row >= data_start:
        cell = ws.cell(row=max_row, column=1)
        try:
            return int(cell.value) if cell.value is not None else 0
        except:
            return None
    return 0

def is_file_writable(filepath):
    try:
        with open(filepath, 'a'):
            pass
        return True
    except (IOError, PermissionError):
        return False

# ==================== 主处理函数 ====================
def process_files(main_file, sub_file, summary_file):
    try:
        # ----- 1. 读取数据 -----
        main_df = pd.read_excel(main_file, skiprows=1)
        sub_df = pd.read_excel(sub_file, skiprows=1)

        # ----- 2. 列名清洗 -----
        for df in (main_df, sub_df):
            df.columns = df.columns.str.strip()
            df.columns = df.columns.str.replace('（', '(').str.replace('）', ')')
            df.columns = df.columns.str.replace('\u200b', '')

        # ----- 3. 主表列映射 -----
        required_main = {
            '合同编号': ['合同编号'],
            '合同名称': ['合同名称'],
            '供应商名称': ['供应商名称'],
            '合同金额(元)': ['合同金额(元)', '合同金额', '金额(元)'],
            '合同签约完成时间': ['合同签约完成时间', '签约完成时间', '合同签订日期'],
            '合同支付比例': ['合同支付比例', '支付比例'],  # 用于“付款比例”字段
            '付款比例': ['付款比例'],  # 用于“累计付款比例(百分比)”字段
            '供应商联系人': ['供应商联系人', '联系人'],
            '供应商联系人电话': ['供应商联系人电话', '联系人电话'],
        }
        col_map = {}
        for std, aliases in required_main.items():
            if std in main_df.columns:
                col_map[std] = std
                continue
            for alias in aliases:
                if alias in main_df.columns:
                    col_map[std] = alias
                    break
            else:
                clean_std = re.sub(r'[()（）\s]', '', std).lower()
                for act in main_df.columns:
                    if re.sub(r'[()（）\s]', '', act).lower() == clean_std:
                        col_map[std] = act
                        break
                else:
                    return False, f"主表缺少必要列: {std}\n\n主表实际列名:\n" + "\n".join(main_df.columns)
        for std, act in col_map.items():
            if act != std:
                main_df.rename(columns={act: std}, inplace=True)

        # ----- 4. 次表必要列检查 -----
        sub_required = ['合同编号', '物资名称', '采购数量', '计量单位', '价税合计(元)', '合同要求交货日期']
        for col in sub_required:
            if col not in sub_df.columns:
                matched = False
                clean_need = re.sub(r'[()（）\s]', '', col).lower()
                for act in sub_df.columns:
                    if re.sub(r'[()（）\s]', '', act).lower() == clean_need:
                        sub_df.rename(columns={act: col}, inplace=True)
                        matched = True
                        break
                if not matched:
                    return False, f"次表缺少必要列: {col}"

        main_df['合同编号'] = main_df['合同编号'].astype(str).str.strip()
        sub_df['合同编号'] = sub_df['合同编号'].astype(str).str.strip()

        # ----- 5. 次表聚合 -----
        sub_grouped = sub_df.groupby('合同编号')
        sub_agg_list = []
        for contract, group in sub_grouped:
            qty_total, unit = aggregate_qty_unit(group)
            materials = merge_materials(group)
            qty_display = get_quantity_display(group)
            total_price = group['价税合计(元)'].sum() if '价税合计(元)' in group.columns else 0
            delivery_dates = group['合同要求交货日期'].dropna()
            delivery_date = delivery_dates.iloc[0] if not delivery_dates.empty else ""
            sub_agg_list.append({
                '合同编号': contract,
                '物资名称': materials,
                '数量': qty_display,
                '单位': unit,
                '合同最终结算金额': total_price,
                '合同要求交货日期': delivery_date
            })
        sub_processed = pd.DataFrame(sub_agg_list)

        # ----- 6. 合并主次表 -----
        merged = pd.merge(main_df, sub_processed, on='合同编号', how='left')

        # ----- 7. 生成更新结果（不含已付金额（元），后续动态计算） -----
        result = pd.DataFrame()
        result['合同号'] = merged['合同编号']
        result['合同名称'] = merged['合同名称']
        result['厂家'] = merged['供应商名称']
        result['数量'] = merged['数量']
        result['单位'] = merged['单位']
        result['合同签订金额'] = merged['合同金额(元)']
        result['合同签订日期'] = merged['合同签约完成时间']
        contact = merged['供应商联系人'].fillna('').astype(str)
        phone = merged['供应商联系人电话'].fillna('').astype(str)
        result['供应商联系人及电话'] = contact + phone
        result['合同要求到货时间'] = merged['合同要求交货日期'].apply(safe_date_parse)
        # 付款比例 → 来自主表“合同支付比例”
        if '合同支付比例' in merged.columns:
            result['付款比例'] = merged['合同支付比例']
        else:
            result['付款比例'] = ""
        result['物资名称'] = merged['物资名称']
        result['合同最终结算金额'] = merged['合同最终结算金额']
        # 累计付款比例(百分比) → 来自主表“付款比例”
        if '付款比例' in merged.columns:
            result['累计付款比例(百分比)'] = merged['付款比例']
        else:
            result['累计付款比例(百分比)'] = ""
        # 已付金额（元）暂不填写，将在更新写入时动态计算

        # ----- 8. 备份 -----
        backup_file = os.path.splitext(summary_file)[0] + "_备份.xlsx"
        shutil.copyfile(summary_file, backup_file)

        # ----- 9. 更新汇总表 -----
        wb = openpyxl.load_workbook(summary_file)
        ws = wb.active

        # 寻找列名行
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

        # 构建字段 -> 列号映射（支持别名）
        col_map = {}
        for cell in ws[header_row]:
            if cell.value:
                col_name = str(cell.value).strip()
                clean_col = re.sub(r'[()（）\s]', '', col_name).lower()
                for field in UPDATE_FIELDS:
                    if field in col_map:
                        continue
                    possible = FIELD_ALIASES.get(field, [field])
                    for alias in possible:
                        if re.sub(r'[()（）\s]', '', alias).lower() == clean_col:
                            col_map[field] = cell.column
                            break
        if '合同号' not in col_map:
            return False, "汇总表中未找到“合同号”列，请检查文件第4行是否为列名。"

        data_start = header_row + 1
        # 建立现有合同 -> 行号映射
        contract_row_map = {}
        for row in ws.iter_rows(min_row=data_start, max_row=ws.max_row,
                                min_col=col_map['合同号'], max_col=col_map['合同号']):
            cell = row[0]
            if cell.value is not None:
                contract_row_map[str(cell.value).strip()] = cell.row

        blue_font_color = "0000FF"
        green_font_color = "008000"

        updated_count = 0
        changed_contracts = []
        new_rows_data = []

        for _, r_row in result.iterrows():
            contract_num = str(r_row['合同号']).strip()
            if contract_num in contract_row_map:
                # ---------- 现有合同 ----------
                row_idx = contract_row_map[contract_num]
                row_changed = False
                for field in UPDATE_FIELDS:
                    if field == '已付金额（元）':
                        continue  # 后面统一处理
                    if field == '供应商联系人及电话':
                        # 已有行不更新该字段，跳过
                        continue
                    if field in col_map and field in r_row:
                        col_idx = col_map[field]
                        old_cell = ws.cell(row=row_idx, column=col_idx)
                        if values_differ(old_cell.value, r_row[field]):
                            old_cell.value = r_row[field]
                            new_font = copy(old_cell.font) if old_cell.font else Font()
                            new_font.color = blue_font_color
                            old_cell.font = new_font
                            row_changed = True
                # 处理已付金额（元）：计算并写入
                if '已付金额（元）' in col_map:
                    # 获取最新的合同签订金额和累计付款比例(百分比)
                    amount_col = col_map.get('合同签订金额')
                    ratio_col = col_map.get('累计付款比例(百分比)')
                    if amount_col and ratio_col:
                        # 优先取本行已更新的值（如果有更新过则使用r_row中的新值，否则用旧值）
                        amount = r_row['合同签订金额'] if (amount_col in col_map and
                                values_differ(ws.cell(row=row_idx, column=amount_col).value,
                                              r_row['合同签订金额'])) else ws.cell(row=row_idx, column=amount_col).value
                        ratio = r_row['累计付款比例(百分比)'] if (ratio_col in col_map and
                                values_differ(ws.cell(row=row_idx, column=ratio_col).value,
                                              r_row['累计付款比例(百分比)'])) else ws.cell(row=row_idx, column=ratio_col).value

                        # 如果是由于其他字段变化导致本次未更新金额/比例，但合同本身金额/比例不变，直接使用旧值
                        # 更简单的方式：直接读取当前单元格的值（因为write是立即生效的）
                        amount_val = ws.cell(row=row_idx, column=amount_col).value
                        ratio_val = ws.cell(row=row_idx, column=ratio_col).value
                        try:
                            amt = float(amount_val) if amount_val not in [None, ""] else 0
                            rat = percent_to_decimal(ratio_val)
                            paid = amt * rat
                            paid_cell = ws.cell(row=row_idx, column=col_map['已付金额（元）'])
                            # 如果已付金额（元）发生变化，则更新并标蓝
                            if values_differ(paid_cell.value, paid):
                                paid_cell.value = paid
                                paid_cell.number_format = '#,##0.00'  # 设置数值格式，保留两位小数
                                new_font = copy(paid_cell.font) if paid_cell.font else Font()
                                new_font.color = blue_font_color
                                paid_cell.font = new_font
                                row_changed = True
                        except:
                            pass

                if row_changed:
                    updated_count += 1
                    changed_contracts.append(contract_num)
            else:
                # ---------- 新增合同 ----------
                new_rows_data.append(r_row)

        # ----- 10. 新增行（带格式继承） -----
        if new_rows_data:
            template_row = ws.max_row if ws.max_row >= data_start else data_start
            seq_index = get_first_column_as_sequence(ws, data_start, ws.max_row)
            start_seq = seq_index if seq_index is not None else None

            for r_row in new_rows_data:
                insert_row = ws.max_row + 1
                copy_row_style(ws, template_row, insert_row)

                # 序号列
                if start_seq is not None:
                    start_seq += 1
                    cell = ws.cell(row=insert_row, column=1)
                    cell.value = start_seq
                    current_font = copy(cell.font) if cell.font else Font()
                    current_font.color = green_font_color
                    cell.font = current_font

                # 写入业务字段（供应商联系人及电话正常写入）
                for field in UPDATE_FIELDS:
                    if field == '已付金额（元）':
                        continue
                    if field in col_map and field in r_row:
                        cell = ws.cell(row=insert_row, column=col_map[field])
                        cell.value = r_row[field]
                        current_font = copy(cell.font) if cell.font else Font()
                        current_font.color = green_font_color
                        cell.font = current_font

                # 新增行的已付金额（元）计算
                if '已付金额（元）' in col_map:
                    amount_col = col_map.get('合同签订金额')
                    ratio_col = col_map.get('累计付款比例(百分比)')
                    if amount_col and ratio_col:
                        amt = float(r_row['合同签订金额']) if r_row['合同签订金额'] not in [None, ""] else 0
                        rat = percent_to_decimal(r_row['累计付款比例(百分比)'])
                        paid = amt * rat
                        paid_cell = ws.cell(row=insert_row, column=col_map['已付金额（元）'])
                        paid_cell.value = paid
                        paid_cell.number_format = '#,##0.00'  # 设置数值格式，保留两位小数
                        current_font = copy(paid_cell.font) if paid_cell.font else Font()
                        current_font.color = green_font_color
                        paid_cell.font = current_font

                template_row = insert_row

        wb.save(summary_file)

        new_count = len(new_rows_data)
        changed_info = "、".join(changed_contracts) if changed_contracts else "无"
        return True, f"处理成功！\n更新合同数: {updated_count}\n新增合同数: {new_count}\n变更合同号: {changed_info}\n备份文件: {backup_file}"

    except Exception as e:
        error_info = traceback.format_exc()
        return False, f"处理失败:\n{str(e)}\n\n详细错误:\n{error_info}"

# ==================== GUI 界面（同之前，略作调整） ====================
class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("合同数据合并工具")
        self.root.geometry("600x450")
        self.root.configure(bg="#f0f0f0")
        title_font = ("微软雅黑", 16, "bold")

        self.file_entries = {
            'main': tk.StringVar(),
            'sub': tk.StringVar(),
            'summary': tk.StringVar()
        }

        tk.Label(root, text="合同数据合并工具", font=title_font, bg="#f0f0f0", fg="#2c3e50").pack(pady=20)
        tk.Label(root, text="请选择文件（注意：第一行不是列名，真正的列名在第二行）",
                 bg="#f0f0f0", fg="#555555").pack(pady=(0, 10))

        self._create_file_row("主表文件:", "main")
        self._create_file_row("次表文件:", "sub")
        self._create_file_row("汇总表文件:", "summary")

        clear_btn = tk.Button(root, text="清空所选文件", command=self.clear_files,
                              bg="#f0ad4e", fg="white", font=("微软雅黑", 9))
        clear_btn.pack(pady=(5, 0))

        self.process_btn = tk.Button(root, text="更新汇总表", command=self.process,
                                     bg="#4CAF50", fg="white", font=("微软雅黑", 12, "bold"),
                                     width=20, height=2)
        self.process_btn.pack(pady=20)

        self.status_var = tk.StringVar()
        self.status_var.set("准备就绪，请选择文件")
        tk.Label(root, textvariable=self.status_var, bg="#f0f0f0", fg="#555555").pack(pady=10)

        tk.Label(root, text="合同数据合并工具",
                 font=("微软雅黑", 8), bg="#f0f0f0", fg="#777777").pack(side="bottom", pady=5)

    def _create_file_row(self, label_text, file_key):
        frame = tk.Frame(self.root, bg="#f0f0f0")
        frame.pack(fill="x", padx=40, pady=5)
        tk.Label(frame, text=label_text, bg="#f0f0f0", width=10, anchor="w").pack(side="left")
        entry = tk.Entry(frame, width=40, textvariable=self.file_entries[file_key])
        entry.pack(side="left", padx=5)
        tk.Button(frame, text="浏览...",
                  command=lambda: self.select_file(self.file_entries[file_key])).pack(side="left")

    def select_file(self, entry_var):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx;*.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            entry_var.set(file_path)

    def clear_files(self):
        for key in self.file_entries:
            self.file_entries[key].set("")
        self.status_var.set("已清空所选文件，请重新选择")

    def process(self):
        main_file = self.file_entries['main'].get()
        sub_file = self.file_entries['sub'].get()
        summary_file = self.file_entries['summary'].get()

        if not all([main_file, sub_file, summary_file]):
            messagebox.showerror("错误", "请选择所有文件")
            return
        for path, name in zip([main_file, sub_file, summary_file], ["主表", "次表", "汇总表"]):
            if not os.path.exists(path):
                messagebox.showerror("错误", f"{name}文件不存在:\n{path}")
                return

        if not is_file_writable(summary_file):
            messagebox.showerror("文件被占用", "汇总表文件可能正在被Excel或其他程序打开，请关闭后重试。")
            return

        self.process_btn.config(state=tk.DISABLED, text="处理中...")
        self.status_var.set("正在处理数据，请稍候...")
        self.root.update()

        try:
            success, message = process_files(main_file, sub_file, summary_file)
            if success:
                messagebox.showinfo("完成", message)
                self.status_var.set("处理完成！")
            else:
                error_win = tk.Toplevel(self.root)
                error_win.title("错误详情")
                error_win.geometry("600x400")
                tk.Label(error_win, text="处理过程中发生错误:", font=("微软雅黑", 10, "bold")).pack(pady=10)
                text_frame = tk.Frame(error_win)
                text_frame.pack(fill="both", expand=True, padx=10, pady=5)
                scrollbar = tk.Scrollbar(text_frame)
                scrollbar.pack(side="right", fill="y")
                error_text = tk.Text(text_frame, wrap="word", yscrollcommand=scrollbar.set)
                error_text.pack(fill="both", expand=True)
                error_text.insert("1.0", message)
                error_text.config(state="disabled")
                scrollbar.config(command=error_text.yview)
                tk.Button(error_win, text="关闭", command=error_win.destroy).pack(pady=10)
                self.status_var.set("处理失败 - 查看错误详情")
        except Exception as e:
            messagebox.showerror("错误", f"未知错误:\n{str(e)}")
            self.status_var.set("处理失败")
        finally:
            self.process_btn.config(state=tk.NORMAL, text="更新汇总表")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()