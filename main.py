# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl
import pyodbc
import traceback
import datetime
from collections import defaultdict
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from tkcalendar import DateEntry  # 需要安装: pip install tkcalendar

# ============== 用户配置区 (自动适配驱动) ==============

def get_best_sql_driver():
    """
    自动检测当前电脑安装了哪个 SQL Server 驱动。
    优先级：ODBC 18 > 17 > 13 > Native Client > 系统自带 SQL Server
    """
    try:
        installed_drivers = [d for d in pyodbc.drivers()]
    except Exception:
        return "SQL Server" # 如果获取失败，返回保底驱动
    
    # 驱动优先级列表 (越靠前越好)
    driver_preference = [
        "ODBC Driver 18 for SQL Server",   # 最新版
        "ODBC Driver 17 for SQL Server",   # 主流版
        "ODBC Driver 13 for SQL Server",   # 旧版
        "SQL Server Native Client 11.0",   # SQL 2012 时代
        "SQL Server Native Client 10.0",   # SQL 2008 时代
        "SQL Server"                       # Windows XP/7/10/11 自带通用驱动 (保底)
    ]

    for drv in driver_preference:
        if drv in installed_drivers:
            return drv
    
    # 如果一个都没找到，返回默认值尝试
    return "SQL Server"

# 动态获取当前电脑的最佳驱动
CURRENT_DRIVER = get_best_sql_driver()

# 构造连接字符串
DB_CONN_STRING = (
    f"DRIVER={{{CURRENT_DRIVER}}};"
    "SERVER=192.168.0.117;"
    "DATABASE=FQD;"
    "UID=zhitan;"
    "PWD=Zt@forcome;"
    "TrustServerCertificate=yes;"
)

print(f"系统启动: 检测到并使用数据库驱动 -> {CURRENT_DRIVER}")

# 截图中的关键配置
ROW_IDX_HEADER_MAIN = 2  # 主表头所在行
ROW_IDX_HEADER_DATE = 3  # 日期表头所在行
ROW_IDX_DATA_START = 4   # 数据起始行

COL_NAME_WORKSHOP = "车间"
COL_NAME_WO_TYPE = "单别"
COL_NAME_WO_NO = "工单单号"

# 设定要保留的列索引 (1-based, 对应 Excel 的 A=1, B=2...)
KEEP_COL_INDICES = [2, 3, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14, 15, 16, 20]


# ============== 应用程序类 ==============

class DailyPlanAvailabilityApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"每日排程齐套分析工具 v5.3 (含净需求列) - 驱动: {CURRENT_DRIVER}")
        self.root.geometry("1100x700") #稍微加宽窗口

        # 样式定义
        self.red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        self.green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
        self.header_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
        self.thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                  top=Side(style='thin'), bottom=Side(style='thin'))

        # 变量绑定
        self.file_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.selected_workshop = tk.StringVar()

        # 日期范围控制
        self.is_date_range = tk.BooleanVar(value=False)

        # 缓存数据
        self.date_column_map = {}
        self.col_map_main = {}
        self.header_names_map = {}

        self._create_widgets()

    def _create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. 文件选择
        file_frame = ttk.LabelFrame(main_frame, text="1. 数据源", padding="5")
        file_frame.pack(fill=tk.X, pady=5)

        ttk.Entry(file_frame, textvariable=self.file_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="浏览Excel...", command=self._select_file).pack(side=tk.LEFT, padx=5)

        ttk.Label(file_frame, text="   工作表:").pack(side=tk.LEFT)
        self.sheet_combo = ttk.Combobox(file_frame, textvariable=self.sheet_name, state="disabled", width=15)
        self.sheet_combo.pack(side=tk.LEFT, padx=5)
        self.sheet_combo.bind("<<ComboboxSelected>>", self._on_sheet_selected)

        # 2. 筛选设置
        filter_frame = ttk.LabelFrame(main_frame, text="2. 计划筛选", padding="10")
        filter_frame.pack(fill=tk.X, pady=5)

        # --- 日期选择区域 ---
        date_frame = ttk.Frame(filter_frame)
        date_frame.pack(side=tk.LEFT, fill=tk.X)

        ttk.Checkbutton(date_frame, text="选择日期范围", variable=self.is_date_range, command=self._toggle_date_mode).pack(
            side=tk.LEFT, padx=(0, 10))

        ttk.Label(date_frame, text="开始日期:").pack(side=tk.LEFT)
        self.date_start = DateEntry(date_frame, width=12, background='darkblue', foreground='white', borderwidth=2,
                                    date_pattern='yyyy/mm/dd')
        self.date_start.pack(side=tk.LEFT, padx=5)

        self.lbl_end = ttk.Label(date_frame, text="结束日期:")
        self.date_end = DateEntry(date_frame, width=12, background='darkblue', foreground='white', borderwidth=2,
                                  date_pattern='yyyy/mm/dd')

        self._toggle_date_mode()

        ttk.Label(filter_frame, text="选择车间:").pack(side=tk.LEFT, padx=(30, 5))
        self.workshop_combo = ttk.Combobox(filter_frame, textvariable=self.selected_workshop, state="disabled",
                                           width=20)
        self.workshop_combo.pack(side=tk.LEFT, padx=5)

        # 3. 操作区
        action_frame = ttk.LabelFrame(main_frame, text="3. 执行", padding="10")
        action_frame.pack(fill=tk.X, pady=10)

        btn = ttk.Button(action_frame, text="生成缺料分析文件 (另存为)", command=self._run_analysis_batch)
        btn.pack(fill=tk.X, padx=100)

        # 4. 日志
        self.log_text = tk.Text(main_frame, height=15, state="disabled", font=("Consolas", 9), bg="#F0F0F0")
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self._log(f"程序已启动。当前使用数据库驱动: {CURRENT_DRIVER}")
        if CURRENT_DRIVER == "SQL Server":
            self._log("警告: 未检测到ODBC Driver 17/18，正在使用系统自带老版本驱动，可能会影响性能。")

    def _toggle_date_mode(self):
        if self.is_date_range.get():
            self.lbl_end.pack(side=tk.LEFT, padx=(10, 0))
            self.date_end.pack(side=tk.LEFT, padx=5)
        else:
            self.lbl_end.pack_forget()
            self.date_end.pack_forget()

    def _log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        self.root.update_idletasks()

    def _select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            self.file_path.set(path)
            try:
                wb = openpyxl.load_workbook(path, read_only=True)
                self.sheet_combo['values'] = wb.sheetnames
                if wb.sheetnames:
                    self.sheet_combo.current(0)
                    self._on_sheet_selected(None)
                self.sheet_combo.config(state="readonly")
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件: {e}")

    def _on_sheet_selected(self, event):
        file_path = self.file_path.get()
        sheet_name = self.sheet_name.get()
        if not file_path or not sheet_name: return

        self._log("正在扫描Excel结构...")
        try:
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            ws = wb[sheet_name]

            self.col_map_main = {}
            self.header_names_map = {}

            for idx, cell in enumerate(ws[ROW_IDX_HEADER_MAIN], start=1):
                val = str(cell.value).strip() if cell.value else ""
                if val:
                    self.col_map_main[val] = idx
                if idx in KEEP_COL_INDICES:
                    self.header_names_map[idx] = val

            required_cols = [COL_NAME_WORKSHOP, COL_NAME_WO_TYPE, COL_NAME_WO_NO]
            missing = [c for c in required_cols if c not in self.col_map_main]
            if missing:
                messagebox.showwarning("警告", f"未找到关键列: {missing}")
                return

            self.date_column_map = {}
            for cell in ws[ROW_IDX_HEADER_DATE]:
                val = cell.value
                date_obj = self._parse_excel_date(val)
                if date_obj:
                    self.date_column_map[date_obj] = cell.column

            date_keys = sorted(list(self.date_column_map.keys()))
            if not date_keys:
                self._log("警告: 在第3行未找到任何日期格式的表头！")
            else:
                self._log(f"Excel中检测到排程日期范围: {date_keys[0]} 至 {date_keys[-1]}")

            col_ws_idx = self.col_map_main[COL_NAME_WORKSHOP]
            workshops = set()
            for row in ws.iter_rows(min_row=ROW_IDX_DATA_START, min_col=col_ws_idx, max_col=col_ws_idx,
                                    values_only=True):
                if row[0]: workshops.add(str(row[0]).strip())

            self.workshop_combo['values'] = ["全部车间"] + sorted(list(workshops))
            self.workshop_combo.current(0)
            self.workshop_combo.config(state="readonly")

        except Exception as e:
            traceback.print_exc()
            self._log(f"扫描失败: {e}")

    def _parse_excel_date(self, val):
        if val is None: return None
        try:
            dt = None
            if isinstance(val, datetime.datetime):
                dt = val.date()
            elif isinstance(val, datetime.date):
                dt = val
            elif isinstance(val, (int, float)):
                dt = (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(val))).date()
            elif isinstance(val, str):
                parts = val.strip().split('/')
                if len(parts) == 3:
                    dt = datetime.datetime.strptime(val.strip(), "%Y/%m/%d").date()
            return dt
        except:
            return None

    def _get_target_dates(self):
        start_date = self.date_start.get_date()

        if self.is_date_range.get():
            end_date = self.date_end.get_date()
            if end_date < start_date:
                messagebox.showerror("日期错误", "结束日期不能早于开始日期！")
                return []
        else:
            end_date = start_date

        target_dates = []
        curr = start_date
        while curr <= end_date:
            target_dates.append(curr)
            curr += datetime.timedelta(days=1)

        return target_dates

    def _run_analysis_batch(self):
        target_dates = self._get_target_dates()
        if not target_dates: return

        file_path = self.file_path.get()
        sheet_name = self.sheet_name.get()
        target_workshop = self.selected_workshop.get()

        valid_dates = []
        for d in target_dates:
            if d in self.date_column_map:
                valid_dates.append(d)
            else:
                self._log(f"跳过: Excel中未找到日期列 {d}")

        if not valid_dates:
            messagebox.showwarning("无有效日期", "所选日期在Excel表头中均未找到对应列。")
            return

        if len(valid_dates) == 1:
            date_str = valid_dates[0].strftime("%Y-%m-%d")
            default_name = f"{date_str}缺料分析.xlsx"
        else:
            start_s = valid_dates[0].strftime("%Y-%m-%d")
            end_s = valid_dates[-1].strftime("%Y-%m-%d")
            default_name = f"{start_s}至{end_s}缺料分析.xlsx"

        save_path = filedialog.asksaveasfilename(
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not save_path: return

        try:
            self._log("=" * 50)
            self._log(f"开始批量分析... 共 {len(valid_dates)} 天")

            new_wb = openpyxl.Workbook()
            if "Sheet" in new_wb.sheetnames:
                del new_wb["Sheet"]

            for d in valid_dates:
                sheet_title = d.strftime("%Y-%m-%d")
                self._log(f"正在处理: {sheet_title}")

                col_idx = self.date_column_map[d]
                plans_data = self._extract_data_for_date(file_path, sheet_name, col_idx, target_workshop)

                if not plans_data:
                    self._log(f"  -> {sheet_title} 无排产数据")
                    new_ws = new_wb.create_sheet(title=sheet_title)
                    self._write_headers(new_ws)
                    continue

                wo_keys = list(set(p['wo_key'] for p in plans_data))
                self._log(f"  -> 查询 {len(wo_keys)} 个工单的BOM...")
                wo_details = self._fetch_erp_data(wo_keys)

                all_parts = set()
                for w in wo_details.values():
                    for b in w['bom']: all_parts.add(b['part'])

                self._log(f"  -> 查询 {len(all_parts)} 种物料的库存...")
                inventory = self._fetch_inventory(list(all_parts))

                results = self._simulate(plans_data, wo_details, inventory)

                new_ws = new_wb.create_sheet(title=sheet_title)
                self._write_new_sheet(new_ws, plans_data, results)
                self._log(f"  -> {sheet_title} 处理完成")

            new_wb.save(save_path)
            messagebox.showinfo("完成", f"文件已保存至:\n{save_path}")
            self._log("全部完成。")

        except Exception as e:
            traceback.print_exc()
            self._log(f"错误: {e}")
            messagebox.showerror("运行错误", str(e))

    def _extract_data_for_date(self, file_path, sheet_name, date_col_idx, filter_ws):
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        ws = wb[sheet_name]

        c_ws = self.col_map_main[COL_NAME_WORKSHOP]
        c_type = self.col_map_main[COL_NAME_WO_TYPE]
        c_no = self.col_map_main[COL_NAME_WO_NO]

        extracted_rows = []

        for row in ws.iter_rows(min_row=ROW_IDX_DATA_START):
            try:
                if date_col_idx > len(row): continue
                daily_qty = row[date_col_idx - 1].value

                if isinstance(daily_qty, (int, float)) and daily_qty > 0:
                    curr_ws = row[c_ws - 1].value
                    curr_ws = str(curr_ws).strip() if curr_ws else "未分类"

                    if filter_ws != "全部车间" and curr_ws != filter_ws:
                        continue

                    row_data = {}
                    for target_col_idx in KEEP_COL_INDICES:
                        if target_col_idx <= len(row):
                            val = row[target_col_idx - 1].value
                            row_data[target_col_idx] = val
                        else:
                            row_data[target_col_idx] = None

                    wo_type = row[c_type - 1].value
                    wo_no = row[c_no - 1].value

                    if wo_type and wo_no:
                        extracted_rows.append({
                            'base_data': row_data,
                            'wo_key': (str(wo_type).strip(), str(wo_no).strip()),
                            'daily_qty': float(daily_qty),
                            'workshop': curr_ws
                        })
            except IndexError:
                continue
        return extracted_rows

    def _fetch_erp_data(self, wo_keys):
        if not wo_keys: return {}
        conditions = []
        for t, n in wo_keys:
            conditions.append(f"(TA.TA001='{t}' AND TA.TA002='{n}')")

        if not conditions: return {}

        batch_size = 200
        all_data = defaultdict(lambda: {'total': 0, 'bom': []})

        for i in range(0, len(conditions), batch_size):
            batch_conds = conditions[i:i + batch_size]
            where_sql = " OR ".join(batch_conds)

            # 查询增加 MB.MB004 (单位)
            sql = f"""
                SELECT 
                    RTRIM(TA.TA001) as ta001, RTRIM(TA.TA002) as ta002, 
                    TA.TA015 as wo_total_qty,
                    RTRIM(TB.TB003) as part_no, 
                    ISNULL(RTRIM(MB.MB002),'') as part_name,
                    ISNULL(RTRIM(MB.MB004),'') as unit, 
                    TB.TB004 as req_qty, TB.TB005 as iss_qty
                FROM MOCTA TA
                INNER JOIN MOCTB TB ON TA.TA001 = TB.TB001 AND TA.TA002 = TB.TB002
                LEFT JOIN INVMB MB ON TB.TB003 = MB.MB001
                WHERE {where_sql}
            """
            try:
                with pyodbc.connect(DB_CONN_STRING) as conn:
                    df = pd.read_sql(sql, conn)
                    for _, row in df.iterrows():
                        k = (row['ta001'], row['ta002'])
                        all_data[k]['total'] = float(row['wo_total_qty'])
                        all_data[k]['bom'].append({
                            'part': row['part_no'],
                            'name': row['part_name'],
                            'unit': row['unit'],
                            'req': float(row['req_qty']),
                            'iss': float(row['iss_qty'])
                        })
            except Exception as e:
                self._log(f"数据库查询批次失败: {e}")

        return all_data

    def _fetch_inventory(self, parts):
        if not parts: return {}
        unique_parts = list(set(parts))
        inventory = {}
        batch_size = 500

        for i in range(0, len(unique_parts), batch_size):
            batch_parts = unique_parts[i:i + batch_size]
            p_str = ",".join(f"'{p}'" for p in batch_parts)
            sql = f"SELECT RTRIM(MC001) as p, SUM(MC007) as q FROM INVMC WHERE MC001 IN ({p_str}) GROUP BY MC001"
            try:
                with pyodbc.connect(DB_CONN_STRING) as conn:
                    df = pd.read_sql(sql, conn)
                inventory.update(pd.Series(df.q.values, index=df.p).to_dict())
            except:
                pass
        return inventory

    def _simulate(self, plans_data, wo_data, inventory):
        """
        修正逻辑：
        1. 保持原有的最小齐套率计算逻辑（min(stock/net_demand)）。
        2. 新增 max_net_demand_sets (工单净需求)，用于Excel展示分母，方便核对数据。
        """
        running_inv = inventory.copy()
        calculated_results = []

        for p in plans_data:
            wo_key = p['wo_key']
            daily_qty = p['daily_qty']

            info = wo_data.get(wo_key)

            res_item = {
                'rate': 0.0,
                'achievable': 0,
                'net_demand_sets': 0, # 新增：净需求套数 (分母)
                'shortage_str': "",
                'is_short': False
            }

            if not info or not info['bom']:
                res_item['shortage_str'] = "无ERP信息"
                res_item['is_short'] = True
                calculated_results.append(res_item)
                continue

            wo_total_qty = info['total']

            min_kitting_rate = 1.0 
            min_possible_sets = 9999999
            
            # 用于记录该工单在ERP层面最大的“净需求套数”
            # 即：这单实际还需要多少套料？
            max_net_demand_sets = 0.0 
            
            shortage_details_list = []
            to_deduct = {}
            is_fully_kitted = True

            has_requirement = False 

            for bom in info['bom']:
                part = bom['part']
                remain_issue = max(0, bom['req'] - bom['iss'])
                unit_usage = bom['req'] / wo_total_qty if wo_total_qty > 0 else 0
                theo_demand = daily_qty * unit_usage
                
                # 净需求
                net_demand = min(remain_issue, theo_demand)
                
                # 反推这一个物料对应的"套数需求"
                # 用于计算 max_net_demand_sets (Excel显示用)
                if unit_usage > 0:
                    sets_demand = net_demand / unit_usage
                    if sets_demand > max_net_demand_sets:
                        max_net_demand_sets = sets_demand

                if net_demand <= 0.0001: continue
                
                has_requirement = True
                to_deduct[part] = net_demand
                current_stock = running_inv.get(part, 0)

                # 1. 计算齐套率 (保持原有逻辑)
                item_rate = current_stock / net_demand if net_demand > 0 else 1.0
                if item_rate > 1.0: item_rate = 1.0
                min_kitting_rate = min(min_kitting_rate, item_rate)

                # 2. 计算可产数量
                can_make = int(current_stock // unit_usage) if unit_usage > 0 else 999999
                min_possible_sets = min(min_possible_sets, can_make)

                # 3. 记录缺料
                if current_stock < net_demand - 0.0001:
                    is_fully_kitted = False
                    short_qty = net_demand - current_stock
                    unit_str = bom['unit']
                    shortage_details_list.append(f"{part} {bom['name']}(缺{short_qty:g}{unit_str})")

            # 修正可产数量
            actual_possible_sets = min(int(daily_qty), min_possible_sets)
            
            # 如果没有净需求（不需要领料）
            if not has_requirement:
                actual_possible_sets = int(daily_qty)
                is_fully_kitted = True
                min_kitting_rate = 1.0
                max_net_demand_sets = 0 # 没有欠料

            # 如果 max_net_demand_sets 为0 (可能是浮点误差或完全没需求)，但有排产，
            # 为了数据展示，可以视为全齐套
            if max_net_demand_sets < 0.0001 and has_requirement:
                 max_net_demand_sets = daily_qty # 兜底逻辑

            res_item['rate'] = min_kitting_rate
            res_item['achievable'] = actual_possible_sets
            res_item['net_demand_sets'] = int(max_net_demand_sets) # 存入结果
            res_item['is_short'] = not is_fully_kitted
            
            if shortage_details_list:
                res_item['shortage_str'] = "\n".join(shortage_details_list)

            calculated_results.append(res_item)

            if is_fully_kitted:
                for part, qty in to_deduct.items():
                    running_inv[part] -= qty

        return calculated_results

    def _write_headers(self, ws):
        current_col = 1
        for idx in KEEP_COL_INDICES:
            cell = ws.cell(row=1, column=current_col)
            cell.value = self.header_names_map.get(idx, "")
            cell.fill = self.header_fill
            cell.border = self.thin_border
            current_col += 1

        # 修改表头：插入 "净需求"
        new_headers = ["齐套率", "可产数量", "净需求", "缺料信息"]
        for h in new_headers:
            cell = ws.cell(row=1, column=current_col)
            cell.value = h
            cell.font = Font(bold=True)
            cell.fill = self.header_fill
            cell.border = self.thin_border
            current_col += 1

    def _write_new_sheet(self, ws, plans_data, results):
        self._write_headers(ws)

        font_normal = Font(name="微软雅黑", size=9)
        align_wrap = Alignment(vertical="center", wrap_text=True)
        align_center = Alignment(vertical="center", horizontal="center", wrap_text=False)

        for i, (plan, res) in enumerate(zip(plans_data, results)):
            row_idx = i + 2

            # 1. 基础列
            current_col = 1
            for col_idx in KEEP_COL_INDICES:
                cell = ws.cell(row=row_idx, column=current_col)
                val = plan['base_data'].get(col_idx)
                cell.value = val
                cell.font = font_normal
                cell.border = self.thin_border
                cell.alignment = Alignment(vertical="center")
                current_col += 1

            # 2. 计算结果
            # 齐套率
            c_rate = ws.cell(row=row_idx, column=current_col)
            c_rate.value = res['rate']
            c_rate.number_format = '0%'
            c_rate.border = self.thin_border
            c_rate.alignment = align_center

            # 可产数量
            c_qty = ws.cell(row=row_idx, column=current_col + 1)
            c_qty.value = res['achievable']
            c_qty.border = self.thin_border
            c_qty.alignment = align_center
            
            # 净需求 (新增列)
            c_net = ws.cell(row=row_idx, column=current_col + 2)
            c_net.value = res['net_demand_sets']
            c_net.border = self.thin_border
            c_net.alignment = align_center

            # 缺料信息
            c_info = ws.cell(row=row_idx, column=current_col + 3)
            c_info.value = res['shortage_str']
            c_info.alignment = align_wrap 
            c_info.border = self.thin_border
            c_info.font = font_normal

            # 3. 颜色标记
            fill = self.red_fill if res['is_short'] else self.green_fill
            c_rate.fill = fill
            c_qty.fill = fill
            c_net.fill = fill
            c_info.fill = fill

        # 设置列宽
        ws.column_dimensions['A'].width = 15
        # 缺料信息在最后一列 (index + 4)
        ws.column_dimensions[openpyxl.utils.get_column_letter(len(KEEP_COL_INDICES) + 4)].width = 60


if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = DailyPlanAvailabilityApp(root)
        root.mainloop()
    except Exception as e:
        import tkinter.messagebox
        tkinter.messagebox.showerror("启动失败", str(e))
