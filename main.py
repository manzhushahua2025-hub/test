# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl
import pyodbc
import traceback
import datetime
import copy
import math
from collections import defaultdict
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from tkcalendar import DateEntry


# ============== 用户配置区 ==============
def get_best_sql_driver():
    try:
        installed_drivers = [d for d in pyodbc.drivers()]
    except Exception:
        return "SQL Server"

    driver_preference = [
        "ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server",
        "ODBC Driver 13 for SQL Server", "SQL Server Native Client 11.0",
        "SQL Server"
    ]
    for drv in driver_preference:
        if drv in installed_drivers: return drv
    return "SQL Server"


CURRENT_DRIVER = get_best_sql_driver()
DB_CONN_STRING = (
    f"DRIVER={{{CURRENT_DRIVER}}};SERVER=192.168.0.117;DATABASE=FQD;"
    "UID=zhitan;PWD=Zt@forcome;TrustServerCertificate=yes;"
)

# 基础数据起始行
ROW_IDX_DATA_START = 4

# 关键列名配置
COL_NAME_WORKSHOP = "车间"
COL_NAME_WO_TYPE = "单别"
COL_NAME_WO_NO = "工单单号"

# 仅用于识别列索引，不影响读取内容
FULL_COL_RANGE = range(1, 21)
REMOVE_COLS = [1, 9, 17, 18]
KEEP_COL_INDICES = [c for c in FULL_COL_RANGE if c not in REMOVE_COLS]


# ============== 应用程序类 ==============
class DailyPlanAvailabilityApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"每日排程齐套分析工具 v11.1 (逻辑锁定+A列标注版) - {CURRENT_DRIVER}")
        self.root.geometry("1150x750")

        # 颜色定义 (保持不变)
        self.red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")  # 缺料
        self.green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")  # 齐套
        self.yellow_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")  # 完结/超出
        self.gray_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")  # 已领完

        self.file_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.selected_workshop = tk.StringVar()
        self.is_date_range = tk.BooleanVar(value=False)
        self.date_column_map = {}
        self.col_map_main = {}

        self._create_widgets()

    def _create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        file_frame = ttk.LabelFrame(main_frame, text="1. 数据源 (读取原文件 -> 另存副本)", padding="5")
        file_frame.pack(fill=tk.X, pady=5)
        ttk.Entry(file_frame, textvariable=self.file_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="浏览Excel...", command=self._select_file).pack(side=tk.LEFT, padx=5)
        ttk.Label(file_frame, text="   工作表:").pack(side=tk.LEFT)
        self.sheet_combo = ttk.Combobox(file_frame, textvariable=self.sheet_name, state="disabled", width=15)
        self.sheet_combo.pack(side=tk.LEFT, padx=5)
        self.sheet_combo.bind("<<ComboboxSelected>>", self._on_sheet_selected)

        filter_frame = ttk.LabelFrame(main_frame, text="2. 计划筛选 (分析结果将写入副本 A 列)", padding="10")
        filter_frame.pack(fill=tk.X, pady=5)
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

        action_frame = ttk.LabelFrame(main_frame, text="3. 执行", padding="10")
        action_frame.pack(fill=tk.X, pady=10)
        ttk.Button(action_frame, text="执行分析并另存文件", command=self._run_analysis_annotate).pack(fill=tk.X, padx=100)

        self.log_text = tk.Text(main_frame, height=15, state="disabled", font=("Consolas", 9), bg="#F0F0F0")
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=5)

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
                # 使用普通加载以保留格式（后续另存用）
                wb = openpyxl.load_workbook(path, read_only=True)
                self.sheet_combo['values'] = wb.sheetnames
                if wb.sheetnames:
                    self.sheet_combo.current(0)
                    self._on_sheet_selected(None)
                self.sheet_combo.config(state="readonly")
                wb.close()
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件: {e}")

    # --- 逻辑锁定：保持 v10.4 的混合扫描表头逻辑 ---
    def _on_sheet_selected(self, event):
        file_path = self.file_path.get()
        sheet_name = self.sheet_name.get()
        if not file_path or not sheet_name: return
        try:
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            ws = wb[sheet_name]

            self.col_map_main = {}

            # 1. 混合扫描表头 (Row 3 优先，然后 Row 2)
            scan_rows = [3, 2]
            for r in scan_rows:
                for idx, cell in enumerate(ws[r], start=1):
                    val = str(cell.value).strip() if cell.value else ""
                    if val:
                        if val not in self.col_map_main:
                            self.col_map_main[val] = idx

            # 2. 日期列扫描 (严格扫描第3行)
            self.date_column_map = {}
            for cell in ws[3]:
                val = cell.value
                dt = self._parse_excel_date(val)
                if dt: self.date_column_map[dt] = cell.column

            # 检查关键列
            if not self.col_map_main.get(COL_NAME_WO_NO):
                self._log("警告: 未能在第2或3行找到'工单单号'列。")

            col_ws_idx = self.col_map_main.get(COL_NAME_WORKSHOP)
            workshops = set()
            if col_ws_idx:
                for row in ws.iter_rows(min_row=ROW_IDX_DATA_START, min_col=col_ws_idx, max_col=col_ws_idx,
                                        values_only=True):
                    if row[0]: workshops.add(str(row[0]).strip())

            self.workshop_combo['values'] = ["全部车间"] + sorted(list(workshops))
            self.workshop_combo.current(0)
            self.workshop_combo.config(state="readonly")

            self._log(f"分析完成: 找到 {len(self.date_column_map)} 个日期列。")
            wb.close()
        except Exception as e:
            traceback.print_exc()
            self._log(f"扫描失败: {e}")

    def _parse_excel_date(self, val):
        if val is None: return None
        try:
            if isinstance(val, datetime.datetime): return val.date()
            if isinstance(val, datetime.date): return val
            if isinstance(val, (int, float)):
                return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(val))).date()
            if isinstance(val, str):
                parts = val.strip().split('/')
                if len(parts) >= 2:
                    try:
                        return datetime.datetime.strptime(val.strip(), "%Y/%m/%d").date()
                    except:
                        today = datetime.date.today()
                        dt = datetime.datetime.strptime(val.strip(), "%m/%d")
                        return dt.replace(year=today.year).date()
            return None
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

        dates = []
        curr = start_date
        while curr <= end_date:
            dates.append(curr)
            curr += datetime.timedelta(days=1)
        return dates

    def _fetch_erp_data(self, keys):
        if not keys: return {}
        conditions = [f"(TA.TA001='{t}' AND TA.TA002='{n}')" for t, n in keys]
        data = defaultdict(lambda: {'total': 0, 'bom': []})
        batch_size = 200
        for i in range(0, len(conditions), batch_size):
            batch = conditions[i:i + batch_size]
            sql = f"""
                SELECT RTRIM(TA.TA001) t, RTRIM(TA.TA002) n, TA.TA015 total,
                       RTRIM(TB.TB003) p, ISNULL(RTRIM(MB.MB002),'') name, 
                       ISNULL(RTRIM(MB.MB004),'') unit, TB.TB004 req, TB.TB005 iss
                FROM MOCTA TA
                INNER JOIN MOCTB TB ON TA.TA001=TB.TB001 AND TA.TA002=TB.TB002
                LEFT JOIN INVMB MB ON TB.TB003=MB.MB001
                WHERE {" OR ".join(batch)}
            """
            try:
                with pyodbc.connect(DB_CONN_STRING) as conn:
                    df = pd.read_sql(sql, conn)
                    for _, r in df.iterrows():
                        data[(r['t'], r['n'])]['total'] = float(r['total'])
                        data[(r['t'], r['n'])]['bom'].append({
                            'part': r['p'], 'name': r['name'], 'unit': r['unit'],
                            'req': float(r['req']), 'iss': float(r['iss'])
                        })
            except:
                pass
        return data

    def _fetch_inventory(self, parts):
        if not parts: return {}
        inv = {}
        parts = list(set(parts))
        batch_size = 500
        for i in range(0, len(parts), batch_size):
            p_str = ",".join(f"'{p}'" for p in parts[i:i + batch_size])
            sql = f"SELECT RTRIM(MC001) p, SUM(MC007) q FROM INVMC WHERE MC001 IN ({p_str}) GROUP BY MC001"
            try:
                with pyodbc.connect(DB_CONN_STRING) as conn:
                    df = pd.read_sql(sql, conn)
                inv.update(pd.Series(df.q.values, index=df.p).to_dict())
            except:
                pass
        return inv

    # --- v11.0 核心：直接修改原文件副本 A 列 ---
    def _run_analysis_annotate(self):
        target_dates = self._get_target_dates()
        if not target_dates: return

        # 1. 另存为...
        file_path = self.file_path.get()
        default_name = f"齐套分析_{target_dates[0].strftime('%Y%m%d')}.xlsx"
        save_path = filedialog.asksaveasfilename(initialfile=default_name, defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")])
        if not save_path: return

        try:
            self._log("正在加载完整文件 (这可能需要几秒钟)...")
            # 必须使用 load_workbook 并不带 read_only 以便修改和保存格式
            wb_edit = openpyxl.load_workbook(file_path)
            ws_edit = wb_edit[self.sheet_name.get()]

            valid_dates = [d for d in target_dates if d in self.date_column_map]
            if not valid_dates:
                messagebox.showerror("错误", "所选日期在Excel第3行中均未找到对应列")
                return

            # 2. 预扫描涉及的工单 (所有选中日期)
            all_wo_keys = set()
            c_type = self.col_map_main.get(COL_NAME_WO_TYPE, 5)  # 默认5
            c_no = self.col_map_main.get(COL_NAME_WO_NO, 6)  # 默认6
            c_ws = self.col_map_main.get(COL_NAME_WORKSHOP)
            target_ws = self.selected_workshop.get()

            self._log("正在扫描涉及工单...")

            # 我们需要先知道哪些行在哪些日期有产量，建立 Row -> Date -> Qty 映射
            row_plan_map = defaultdict(list)  # {row_idx: [(date, qty), ...]}

            for d in valid_dates:
                col_idx = self.date_column_map[d]
                for row in ws_edit.iter_rows(min_row=ROW_IDX_DATA_START):
                    try:
                        qty = row[col_idx - 1].value
                        if isinstance(qty, (int, float)) and qty > 0:
                            # 过滤车间
                            if c_ws:
                                curr_ws = str(row[c_ws - 1].value).strip() if row[c_ws - 1].value else "未分类"
                                if target_ws != "全部车间" and curr_ws != target_ws: continue

                            wt = row[c_type - 1].value
                            wn = row[c_no - 1].value
                            if wt and wn:
                                key = (str(wt).strip(), str(wn).strip())
                                all_wo_keys.add(key)
                                # 逻辑锁定：强制整数输入
                                int_qty = int(round(float(qty)))
                                row_plan_map[row[0].row].append((d, int_qty, key))
                    except:
                        continue

            if not all_wo_keys:
                messagebox.showinfo("无数据", "所选范围内无排产计划")
                return

            self._log(f"找到 {len(all_wo_keys)} 个工单，正在查询ERP数据...")
            static_wo_data = self._fetch_erp_data(list(all_wo_keys))

            all_parts = set()
            for w in static_wo_data.values():
                for b in w['bom']: all_parts.add(b['part'])

            static_inventory = self._fetch_inventory(list(all_parts))

            # 3. 滚动推演逻辑 (v10.7 整数闭环内核)
            running_inv = copy.deepcopy(static_inventory)
            running_wo_issued = defaultdict(float)
            for k, v in static_wo_data.items():
                for b in v['bom']:
                    running_wo_issued[(k[0], k[1], b['part'])] = b['iss']

            # row_idx -> 最终要写入 A 列的文本 list
            row_annotations = defaultdict(list)
            # row_idx -> 最终颜色的优先级 (Red > Yellow > Gray > Green)
            row_color_priority = defaultdict(int)

            for d in sorted(valid_dates):
                self._log(f"正在推演日期: {d}")
                # 找出当天有产量的所有行
                daily_active_rows = []
                for ridx, plans in row_plan_map.items():
                    for (p_date, p_qty, p_key) in plans:
                        if p_date == d:
                            daily_active_rows.append((ridx, p_qty, p_key))

                # 对当天数据进行计算
                for ridx, plan_qty, key in daily_active_rows:
                    info = static_wo_data.get(key)
                    if not info: continue

                    # --- 逻辑锁定：v10.7 严谨整数计算 ---
                    # 1. 确定剩余能力上限
                    max_possible = 999999
                    for b in info['bom']:
                        u = b['req'] / info['total'] if info['total'] > 0 else 0
                        if u > 0:
                            rem = max(0, b['req'] - running_wo_issued[(key[0], key[1], b['part'])])
                            max_possible = min(max_possible, int(rem // u))

                    # 2. 整数闭环
                    net_demand = min(plan_qty, max_possible)
                    excess = plan_qty - net_demand
                    if excess < 0: excess = 0

                    # 3. 齐套与缺料
                    min_rate = 1.0
                    can_do = 999999
                    short_msgs = []

                    to_deduct = {}

                    for b in info['bom']:
                        u = b['req'] / info['total'] if info['total'] > 0 else 0
                        if u > 0:
                            stock = max(0, running_inv.get(b['part'], 0))
                            need = net_demand * u

                            if need > 0:
                                rate = stock / need
                                if rate < min_rate: min_rate = min(1.0, rate)

                            can_do = min(can_do, int(stock // u))

                            if stock < need - 0.001:
                                diff = need - stock
                                short_msgs.append(f"{b['name']}({b['part']})缺{diff:g}{b['unit']}")

                            to_deduct[b['part']] = plan_qty * u

                    achievable = min(net_demand, can_do)

                    # 4. 状态判断
                    status_code = 1  # Green

                    if net_demand == 0 and excess > 0:
                        status_code = 2  # Gray
                    elif min_rate < 0.999:
                        status_code = 4  # Red
                    elif excess > 0:
                        status_code = 3  # Yellow

                    # 5. 更新库存 (Rolling)
                    for p_part, p_q in to_deduct.items():
                        if p_part in running_inv: running_inv[p_part] -= p_q
                        running_wo_issued[(key[0], key[1], p_part)] += p_q

                    # 6. 生成 A 列文本 (符合您的格式要求)
                    # 齐套率为XX；可产数量为XX个；工单净需求量为XX个；超出工单的数量为XX个；此工单的缺料信息：...
                    msg_body = (f"齐套率为{min_rate:.0%}; "
                                f"可产数量为{achievable}个; "
                                f"工单净需求量为{net_demand}个; "
                                f"超出工单的数量为{excess}个; "
                                f"此工单的缺料信息：{','.join(short_msgs) if short_msgs else '无'}")

                    # 如果多天推演，加上日期前缀
                    if len(target_dates) > 1:
                        msg = f"[{d.strftime('%m-%d')}] {msg_body}"
                    else:
                        msg = msg_body  # 单天则不加日期前缀，保持简洁

                    row_annotations[ridx].append(msg)
                    if status_code > row_color_priority[ridx]:
                        row_color_priority[ridx] = status_code

            # 4. 写入 Excel
            self._log("正在写入 A 列并标注颜色...")
            font_style = Font(name="微软雅黑", size=9)
            align_style = Alignment(vertical="center", wrap_text=True)

            for ridx, msgs in row_annotations.items():
                cell = ws_edit.cell(row=ridx, column=1)  # A列

                # 合并信息
                full_text = "\n".join(msgs)
                cell.value = full_text
                cell.font = font_style
                cell.alignment = align_style

                # 上色
                prio = row_color_priority[ridx]
                if prio == 4:
                    cell.fill = self.red_fill
                elif prio == 3:
                    cell.fill = self.yellow_fill
                elif prio == 2:
                    cell.fill = self.gray_fill
                elif prio == 1:
                    cell.fill = self.green_fill

            # 调整 A 列宽度
            ws_edit.column_dimensions['A'].width = 60

            wb_edit.save(save_path)
            messagebox.showinfo("成功", f"文件已处理完毕！\n结果保存在：{save_path}")
            self._log("全部完成。")

        except Exception as e:
            traceback.print_exc()
            self._log(f"错误: {str(e)}")
            messagebox.showerror("运行错误", str(e))


if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = DailyPlanAvailabilityApp(root)
        root.mainloop()
    except Exception as e:
        import tkinter.messagebox

        tkinter.messagebox.showerror("启动失败", str(e))
