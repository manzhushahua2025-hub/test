# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl
import pyodbc
import traceback
import datetime
import copy
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
# 数据库连接 (只读权限)
DB_CONN_STRING = (
    f"DRIVER={{{CURRENT_DRIVER}}};SERVER=192.168.0.117;DATABASE=FQD;"
    "UID=zhitan;PWD=Zt@forcome;TrustServerCertificate=yes;"
)

# --- 关键修正 1: 根据截图，表头位于第3行 ---
ROW_IDX_HEADER_MAIN = 3  
ROW_IDX_HEADER_DATE = 3  
ROW_IDX_DATA_START = 4   

COL_NAME_WORKSHOP = "车间"
COL_NAME_WO_TYPE = "单别"
COL_NAME_WO_NO = "工单单号"

# 15 代表 O列 (原始表中的“排产数量”列位置)
KEEP_COL_INDICES = [2, 3, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14, 15, 16, 20]

# ============== 应用程序类 ==============
class DailyPlanAvailabilityApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"每日排程齐套分析工具 v9.0 (表头与排产量终极修正版) - {CURRENT_DRIVER}")
        self.root.geometry("1150x750")

        # 颜色定义
        self.red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")     # 红色：缺料
        self.green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")   # 绿色：齐套
        self.yellow_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")  # 黄色：排产超出
        self.gray_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")    # 灰色：已领完/完结
        
        self.header_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
        self.thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                  top=Side(style='thin'), bottom=Side(style='thin'))

        self.file_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.selected_workshop = tk.StringVar()
        self.is_date_range = tk.BooleanVar(value=False)
        self.date_column_map = {}
        self.col_map_main = {}
        self.header_names_map = {}
        self._create_widgets()

    def _create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        file_frame = ttk.LabelFrame(main_frame, text="1. 数据源 (程序只读，不修改原文件)", padding="5")
        file_frame.pack(fill=tk.X, pady=5)
        ttk.Entry(file_frame, textvariable=self.file_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="浏览Excel...", command=self._select_file).pack(side=tk.LEFT, padx=5)
        ttk.Label(file_frame, text="   工作表:").pack(side=tk.LEFT)
        self.sheet_combo = ttk.Combobox(file_frame, textvariable=self.sheet_name, state="disabled", width=15)
        self.sheet_combo.pack(side=tk.LEFT, padx=5)
        self.sheet_combo.bind("<<ComboboxSelected>>", self._on_sheet_selected)

        filter_frame = ttk.LabelFrame(main_frame, text="2. 计划筛选 (按日期顺序强行推演库存)", padding="10")
        filter_frame.pack(fill=tk.X, pady=5)
        date_frame = ttk.Frame(filter_frame)
        date_frame.pack(side=tk.LEFT, fill=tk.X)
        ttk.Checkbutton(date_frame, text="选择日期范围", variable=self.is_date_range, command=self._toggle_date_mode).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(date_frame, text="开始日期:").pack(side=tk.LEFT)
        self.date_start = DateEntry(date_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy/mm/dd')
        self.date_start.pack(side=tk.LEFT, padx=5)
        self.lbl_end = ttk.Label(date_frame, text="结束日期:")
        self.date_end = DateEntry(date_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy/mm/dd')
        self._toggle_date_mode()
        ttk.Label(filter_frame, text="选择车间:").pack(side=tk.LEFT, padx=(30, 5))
        self.workshop_combo = ttk.Combobox(filter_frame, textvariable=self.selected_workshop, state="disabled", width=20)
        self.workshop_combo.pack(side=tk.LEFT, padx=5)

        action_frame = ttk.LabelFrame(main_frame, text="3. 执行 (结果将存为新文件)", padding="10")
        action_frame.pack(fill=tk.X, pady=10)
        ttk.Button(action_frame, text="生成多天推演缺料分析", command=self._run_analysis_batch).pack(fill=tk.X, padx=100)

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
                # read_only=True 模式下，openpyxl 会读取所有行，包括被筛选隐藏的行
                wb = openpyxl.load_workbook(path, read_only=True)
                self.sheet_combo['values'] = wb.sheetnames
                if wb.sheetnames:
                    self.sheet_combo.current(0)
                    self._on_sheet_selected(None)
                self.sheet_combo.config(state="readonly")
                wb.close()
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件: {e}")

    def _on_sheet_selected(self, event):
        file_path = self.file_path.get()
        sheet_name = self.sheet_name.get()
        if not file_path or not sheet_name: return
        try:
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            ws = wb[sheet_name]
            self.col_map_main = {}
            self.header_names_map = {}
            # 修正：从第3行读取表头
            for idx, cell in enumerate(ws[ROW_IDX_HEADER_MAIN], start=1):
                val = str(cell.value).strip() if cell.value else ""
                if val: self.col_map_main[val] = idx
                if idx in KEEP_COL_INDICES: self.header_names_map[idx] = val
            
            self.date_column_map = {}
            for cell in ws[ROW_IDX_HEADER_DATE]:
                val = cell.value
                dt = self._parse_excel_date(val)
                if dt: self.date_column_map[dt] = cell.column
            
            col_ws_idx = self.col_map_main.get(COL_NAME_WORKSHOP)
            workshops = set()
            if col_ws_idx:
                for row in ws.iter_rows(min_row=ROW_IDX_DATA_START, min_col=col_ws_idx, max_col=col_ws_idx, values_only=True):
                    if row[0]: workshops.add(str(row[0]).strip())
            
            self.workshop_combo['values'] = ["全部车间"] + sorted(list(workshops))
            self.workshop_combo.current(0)
            self.workshop_combo.config(state="readonly")
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
                if len(parts) == 3: return datetime.datetime.strptime(val.strip(), "%Y/%m/%d").date()
            return None
        except: return None

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

    def _run_analysis_batch(self):
        target_dates = self._get_target_dates()
        if not target_dates: return
        
        file_path = self.file_path.get()
        sheet_name = self.sheet_name.get()
        target_workshop = self.selected_workshop.get()
        
        valid_dates = [d for d in target_dates if d in self.date_column_map]
        if not valid_dates:
            messagebox.showwarning("无有效日期", "所选日期在Excel表头中均未找到对应列。")
            return

        date_str = valid_dates[0].strftime("%Y-%m-%d")
        default_name = f"{date_str}至{valid_dates[-1].strftime('%Y-%m-%d')}推演分析.xlsx"
        
        save_path = filedialog.asksaveasfilename(initialfile=default_name, defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not save_path: return

        try:
            self._log("="*50)
            self._log(f"开始跨天强行推演... 共 {len(valid_dates)} 天")
            
            new_wb = openpyxl.Workbook()
            if "Sheet" in new_wb.sheetnames: del new_wb["Sheet"]

            self._log("正在预加载所有日期的排产数据...")
            # openpyxl 的 read_only 模式会自动忽略筛选，读取所有行数据
            all_plans_by_date = {} 
            all_wo_keys = set()
            
            for d in valid_dates:
                col_idx = self.date_column_map[d]
                plans = self._extract_data(file_path, sheet_name, col_idx, target_workshop)
                all_plans_by_date[d] = plans
                for p in plans:
                    all_wo_keys.add(p['wo_key'])
            
            if not all_wo_keys:
                messagebox.showinfo("无数据", "所选日期范围内没有排产计划。")
                return

            self._log("正在查询ERP BOM和库存...")
            static_wo_data = self._fetch_erp_data(list(all_wo_keys))
            
            all_parts = set()
            for w in static_wo_data.values():
                for b in w['bom']: all_parts.add(b['part'])
            
            static_inventory = self._fetch_inventory(list(all_parts))

            running_inv = copy.deepcopy(static_inventory)
            running_wo_issued = defaultdict(float)
            for k, v in static_wo_data.items():
                for b in v['bom']:
                    running_wo_issued[(k[0], k[1], b['part'])] = b['iss']

            for d in valid_dates:
                sheet_title = d.strftime("%Y-%m-%d")
                self._log(f"推演日期: {sheet_title}")
                
                plans = all_plans_by_date[d]
                
                if not plans:
                    new_ws = new_wb.create_sheet(title=sheet_title)
                    self._write_headers(new_ws)
                    continue

                results = self._simulate_logic_rolling_forced(
                    plans, static_wo_data, running_inv, running_wo_issued
                )
                
                new_ws = new_wb.create_sheet(title=sheet_title)
                self._write_sheet(new_ws, plans, results)
            
            new_wb.save(save_path)
            messagebox.showinfo("成功", f"文件已生成:\n{save_path}")
            self._log("全部完成。")

        except Exception as e:
            traceback.print_exc()
            self._log(f"错误: {e}")
            messagebox.showerror("运行错误", str(e))

    def _extract_data(self, file_path, sheet_name, col_idx, filter_ws):
        # 即使文件被筛选过，read_only=True 也会遍历所有物理行
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        ws = wb[sheet_name]
        c_ws = self.col_map_main.get(COL_NAME_WORKSHOP)
        c_type = self.col_map_main.get(COL_NAME_WO_TYPE)
        c_no = self.col_map_main.get(COL_NAME_WO_NO)
        
        data = []
        for row in ws.iter_rows(min_row=ROW_IDX_DATA_START):
            try:
                if col_idx > len(row): continue
                qty = row[col_idx-1].value
                
                if isinstance(qty, (int, float)) and qty > 0:
                    curr_ws = str(row[c_ws-1].value).strip() if (c_ws and row[c_ws-1].value) else "未分类"
                    if filter_ws != "全部车间" and curr_ws != filter_ws: continue
                    
                    row_dict = {}
                    for ti in KEEP_COL_INDICES:
                        # --- 关键修改 2: 强制将导出的“排产数量”列(O列)替换为当日排产数 qty ---
                        if ti == 15:
                            row_dict[ti] = float(qty)
                        else:
                            row_dict[ti] = row[ti-1].value if ti <= len(row) else None
                    
                    wt = row[c_type-1].value
                    wn = row[c_no-1].value
                    if wt and wn:
                        data.append({
                            'base': row_dict,
                            'wo_key': (str(wt).strip(), str(wn).strip()),
                            'qty': float(qty),
                            'ws': curr_ws
                        })
            except: continue
        wb.close()
        return data

    def _fetch_erp_data(self, keys):
        if not keys: return {}
        conditions = [f"(TA.TA001='{t}' AND TA.TA002='{n}')" for t, n in keys]
        data = defaultdict(lambda: {'total': 0, 'bom': []})
        batch_size = 200
        for i in range(0, len(conditions), batch_size):
            batch = conditions[i:i+batch_size]
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
            except: pass
        return data

    def _fetch_inventory(self, parts):
        if not parts: return {}
        inv = {}
        parts = list(set(parts))
        batch_size = 500
        for i in range(0, len(parts), batch_size):
            p_str = ",".join(f"'{p}'" for p in parts[i:i+batch_size])
            sql = f"SELECT RTRIM(MC001) p, SUM(MC007) q FROM INVMC WHERE MC001 IN ({p_str}) GROUP BY MC001"
            try:
                with pyodbc.connect(DB_CONN_STRING) as conn:
                    df = pd.read_sql(sql, conn)
                inv.update(pd.Series(df.q.values, index=df.p).to_dict())
            except: pass
        return inv

    def _simulate_logic_rolling_forced(self, plans, wo_data, running_inv, running_wo_issued):
        results = []

        for p in plans:
            key = p['wo_key']
            plan_qty = p['qty']
            info = wo_data.get(key)
            
            res = {
                'rate': 0.0, 'achievable': 0, 
                'net_demand': 0, 'excess': 0,
                'msg': "", 'status': 'normal'
            }

            if not info or not info['bom']:
                res['msg'] = "无ERP信息"; res['status'] = 'error'
                results.append(res); continue

            # 1. 工单净需求
            eff_demand = plan_qty 
            for b in info['bom']:
                unit_use = b['req'] / info['total'] if info['total'] > 0 else 0
                if unit_use <= 0: continue
                current_issued = running_wo_issued.get((key[0], key[1], b['part']), 0)
                theo_need = plan_qty * unit_use
                remain_issue = max(0, b['req'] - current_issued)
                if remain_issue < theo_need - 0.0001:
                    max_sets = remain_issue / unit_use
                    if max_sets < eff_demand:
                        eff_demand = max_sets
            
            # 2. 超出部分
            excess_qty = 0
            if eff_demand < plan_qty:
                excess_qty = plan_qty - eff_demand
                if excess_qty < 0.0001: excess_qty = 0
            
            # 3. 齐套率/缺料
            min_material_rate = 1.0 
            min_possible_sets = 999999
            short_details = []
            
            to_deduct_full = {} 

            for b in info['bom']:
                unit_use = b['req'] / info['total'] if info['total'] > 0 else 0
                if unit_use <= 0: continue
                
                part_net_demand = eff_demand * unit_use
                stock = running_inv.get(b['part'], 0)
                
                if part_net_demand > 0:
                    effective_stock = max(0, stock)
                    part_rate = effective_stock / part_net_demand
                    if part_rate > 1.0: part_rate = 1.0
                    if part_rate < min_material_rate:
                        min_material_rate = part_rate
                
                can_do = int(max(0, stock) // unit_use)
                min_possible_sets = min(min_possible_sets, can_do)
                
                if stock < part_net_demand - 0.0001:
                    diff = part_net_demand - stock
                    short_details.append(f"{b['name']}({b['part']})缺{diff:g}{b['unit']}")
                
                full_demand = plan_qty * unit_use
                to_deduct_full[b['part']] = full_demand

            achievable = min(int(eff_demand), min_possible_sets)
            
            # 状态判定
            if eff_demand < 0.001:
                res['rate'] = 1.0; res['achievable'] = 0
                res['net_demand'] = 0; res['excess'] = int(plan_qty)
                res['status'] = 'finished'; res['msg'] = "工单物料已领完/工单完结"
            else:
                res['rate'] = min_material_rate
                res['achievable'] = achievable
                res['net_demand'] = int(eff_demand)
                res['excess'] = int(excess_qty)
                
                fully_kitted = (min_material_rate >= 0.999)
                
                if not fully_kitted:
                    res['status'] = 'short'
                    msgs = []
                    if short_details: msgs.append("\n".join(short_details))
                    res['msg'] = "\n".join(msgs)
                elif excess_qty > 0:
                    res['status'] = 'warn'
                    res['msg'] = "此工单完结，排产超出工单数量"
                else:
                    res['status'] = 'ok'
                    res['msg'] = "齐套"
            
            results.append(res)
            
            # 扣减库存
            for part, qty in to_deduct_full.items():
                if part not in running_inv: running_inv[part] = 0.0
                running_inv[part] -= qty 
                running_wo_issued[(key[0], key[1], part)] += qty

        return results

    def _write_headers(self, ws):
        curr = 1
        for idx in KEEP_COL_INDICES:
            c = ws.cell(1, curr); c.value = self.header_names_map.get(idx,""); 
            c.fill = self.header_fill; c.border = self.thin_border
            curr += 1
        new_cols = ["齐套率", "可产数量", "工单净需求", "超出工单数量", "缺料信息"]
        for h in new_cols:
            c = ws.cell(1, curr); c.value = h; c.font = Font(bold=True)
            c.fill = self.header_fill; c.border = self.thin_border
            curr += 1

    def _write_sheet(self, ws, plans, results):
        self._write_headers(ws)
        font = Font(name="微软雅黑", size=9)
        align = Alignment(vertical="center", wrap_text=True)
        center = Alignment(vertical="center", horizontal="center")
        
        for i, (p, r) in enumerate(zip(plans, results)):
            ridx = i + 2
            curr = 1
            for idx in KEEP_COL_INDICES:
                c = ws.cell(ridx, curr); c.value = p['base'].get(idx)
                c.font = font; c.border = self.thin_border; c.alignment = Alignment(vertical="center")
                curr += 1
            
            c_rate = ws.cell(ridx, curr); c_rate.value = r['rate']; c_rate.number_format = '0%'
            c_qty = ws.cell(ridx, curr+1); c_qty.value = r['achievable']
            c_net = ws.cell(ridx, curr+2)
            if r['status'] == 'finished': c_net.value = "-"
            else: c_net.value = r['net_demand']
            c_excess = ws.cell(ridx, curr+3); c_excess.value = r['excess']
            c_msg = ws.cell(ridx, curr+4); c_msg.value = r['msg']
            c_msg.alignment = align
            
            for c in [c_rate, c_qty, c_net, c_excess]: c.border = self.thin_border; c.alignment = center
            c_msg.border = self.thin_border; c_msg.font = font
            
            fill = self.green_fill
            if r['status'] == 'short': fill = self.red_fill
            elif r['status'] == 'warn': fill = self.yellow_fill
            elif r['status'] == 'finished': fill = self.gray_fill
            
            for c in [c_rate, c_qty, c_net, c_excess, c_msg]: c.fill = fill
            
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions[openpyxl.utils.get_column_letter(len(KEEP_COL_INDICES)+5)].width = 50

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = DailyPlanAvailabilityApp(root)
        root.mainloop()
    except Exception as e:
        import tkinter.messagebox
        tkinter.messagebox.showerror("启动失败", str(e))
