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
from tkcalendar import DateEntry  

# ============== 1. 用户配置区 ==============

def get_best_sql_driver():
    """
    自动检测当前电脑安装了哪个 SQL Server 驱动。
    优先级：ODBC 18 > 17 > 13 > Native Client > 系统自带 SQL Server
    """
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

# 数据库连接字符串 (只读操作)
DB_CONN_STRING = (
    f"DRIVER={{{CURRENT_DRIVER}}};SERVER=192.168.0.117;DATABASE=FQD;"
    "UID=zhitan;PWD=Zt@forcome;TrustServerCertificate=yes;"
)

# Excel 结构配置
ROW_IDX_HEADER_MAIN = 2  
ROW_IDX_HEADER_DATE = 3  
ROW_IDX_DATA_START = 4   

COL_NAME_WORKSHOP = "车间"
COL_NAME_WO_TYPE = "单别"
COL_NAME_WO_NO = "工单单号"

# 需要保留的原始列索引 (B列到T列, 根据之前需求排除特定列)
KEEP_COL_INDICES = [2, 3, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14, 15, 16, 20]

# ============== 2. 应用程序类 ==============

class DailyPlanAvailabilityApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"每日排程齐套分析工具 v7.3 (最终安全版) - 驱动: {CURRENT_DRIVER}")
        self.root.geometry("1150x750")

        # 样式定义
        self.red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")     # 缺料
        self.green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")   # 齐套
        self.yellow_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")  # 警告(有超出)
        self.gray_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")    # 已结案
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

        # --- 区域1: 文件选择 ---
        file_frame = ttk.LabelFrame(main_frame, text="1. 数据源 (程序只读，不修改原文件)", padding="5")
        file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Entry(file_frame, textvariable=self.file_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="浏览Excel...", command=self._select_file).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(file_frame, text="   工作表:").pack(side=tk.LEFT)
        self.sheet_combo = ttk.Combobox(file_frame, textvariable=self.sheet_name, state="disabled", width=15)
        self.sheet_combo.pack(side=tk.LEFT, padx=5)
        self.sheet_combo.bind("<<ComboboxSelected>>", self._on_sheet_selected)

        # --- 区域2: 筛选设置 ---
        filter_frame = ttk.LabelFrame(main_frame, text="2. 计划筛选", padding="10")
        filter_frame.pack(fill=tk.X, pady=5)
        
        date_frame = ttk.Frame(filter_frame)
        date_frame.pack(side=tk.LEFT, fill=tk.X)
        
        ttk.Checkbutton(date_frame, text="选择日期范围", variable=self.is_date_range, command=self._toggle_date_mode).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(date_frame, text="开始日期:").pack(side=tk.LEFT)
        self.date_start = DateEntry(date_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy/mm/dd')
        self.date_start.pack(side=tk.LEFT, padx=5)
        
        self.lbl_end = ttk.Label(date_frame, text="结束日期:")
        self.date_end = DateEntry(date_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy/mm/dd')
        
        self._toggle_date_mode() # 初始化显示状态
        
        ttk.Label(filter_frame, text="选择车间:").pack(side=tk.LEFT, padx=(30, 5))
        self.workshop_combo = ttk.Combobox(filter_frame, textvariable=self.selected_workshop, state="disabled", width=20)
        self.workshop_combo.pack(side=tk.LEFT, padx=5)

        # --- 区域3: 执行按钮 ---
        action_frame = ttk.LabelFrame(main_frame, text="3. 执行 (结果将存为新文件)", padding="10")
        action_frame.pack(fill=tk.X, pady=10)
        ttk.Button(action_frame, text="生成缺料分析文件 (另存为)", command=self._run_analysis_batch).pack(fill=tk.X, padx=100)

        # --- 区域4: 日志窗口 ---
        self.log_text = tk.Text(main_frame, height=15, state="disabled", font=("Consolas", 9), bg="#F0F0F0")
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self._log(f"程序启动完成。当前数据库驱动: {CURRENT_DRIVER}")

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

    # ============== 3. 文件操作逻辑 (只读) ==============

    def _select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            self.file_path.set(path)
            try:
                # 关键：使用 read_only=True 确保不修改原文件
                wb = openpyxl.load_workbook(path, read_only=True)
                self.sheet_combo['values'] = wb.sheetnames
                if wb.sheetnames:
                    self.sheet_combo.current(0)
                    self._on_sheet_selected(None)
                self.sheet_combo.config(state="readonly")
                wb.close() # 立即释放文件占用
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件: {e}")

    def _on_sheet_selected(self, event):
        file_path = self.file_path.get()
        sheet_name = self.sheet_name.get()
        if not file_path or not sheet_name: return
        
        self._log("正在扫描Excel结构 (只读模式)...")
        try:
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            ws = wb[sheet_name]
            
            # 1. 扫描表头
            self.col_map_main = {}
            self.header_names_map = {}
            for idx, cell in enumerate(ws[ROW_IDX_HEADER_MAIN], start=1):
                val = str(cell.value).strip() if cell.value else ""
                if val: self.col_map_main[val] = idx
                if idx in KEEP_COL_INDICES: self.header_names_map[idx] = val
            
            # 2. 扫描日期
            self.date_column_map = {}
            for cell in ws[ROW_IDX_HEADER_DATE]:
                val = cell.value
                dt = self._parse_excel_date(val)
                if dt: self.date_column_map[dt] = cell.column
            
            # 3. 扫描车间
            col_ws_idx = self.col_map_main.get(COL_NAME_WORKSHOP)
            workshops = set()
            if col_ws_idx:
                for row in ws.iter_rows(min_row=ROW_IDX_DATA_START, min_col=col_ws_idx, max_col=col_ws_idx, values_only=True):
                    if row[0]: workshops.add(str(row[0]).strip())
            
            self.workshop_combo['values'] = ["全部车间"] + sorted(list(workshops))
            self.workshop_combo.current(0)
            self.workshop_combo.config(state="readonly")
            
            date_keys = sorted(list(self.date_column_map.keys()))
            if date_keys:
                self._log(f"找到 {len(date_keys)} 个有效日期列")
            else:
                self._log("警告: 未找到日期列，请检查表头(第3行)")
            
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

    # ============== 4. 核心执行逻辑 ==============

    def _run_analysis_batch(self):
        # 1. 获取日期
        target_dates = self._get_target_dates()
        if not target_dates: return
        
        file_path = self.file_path.get()
        sheet_name = self.sheet_name.get()
        target_workshop = self.selected_workshop.get()
        
        valid_dates = [d for d in target_dates if d in self.date_column_map]
        if not valid_dates:
            messagebox.showwarning("无日期", "所选日期在Excel表头中不存在。")
            return

        # 2. 选择保存路径 (生成新文件)
        date_str = valid_dates[0].strftime("%Y-%m-%d")
        default_name = f"{date_str}缺料分析.xlsx" if len(valid_dates)==1 else f"{date_str}等{len(valid_dates)}天缺料分析.xlsx"
        
        save_path = filedialog.asksaveasfilename(initialfile=default_name, defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not save_path: return

        try:
            self._log("="*50)
            self._log(f"开始批量分析... 共 {len(valid_dates)} 天")
            
            # 创建全新的 Workbook (内存对象)，物理隔离原文件
            new_wb = openpyxl.Workbook()
            if "Sheet" in new_wb.sheetnames: del new_wb["Sheet"]

            for d in valid_dates:
                sheet_title = d.strftime("%Y-%m-%d")
                self._log(f"正在处理: {sheet_title}")
                
                col_idx = self.date_column_map[d]
                
                # A. 读取原文件数据
                plans = self._extract_data(file_path, sheet_name, col_idx, target_workshop)
                if not plans:
                    self._log(f"  -> {sheet_title} 无排产数据")
                    new_ws = new_wb.create_sheet(title=sheet_title)
                    self._write_headers(new_ws)
                    continue

                # B. 获取ERP数据 (只读查询)
                wo_keys = list(set(p['wo_key'] for p in plans))
                wo_data = self._fetch_erp_data(wo_keys)
                
                parts = set()
                for w in wo_data.values():
                    for b in w['bom']: parts.add(b['part'])
                
                inventory = self._fetch_inventory(list(parts))

                # C. 内存计算 (不修改任何文件)
                results = self._simulate_logic(plans, wo_data, inventory)
                
                # D. 写入新文件对象
                new_ws = new_wb.create_sheet(title=sheet_title)
                self._write_sheet(new_ws, plans, results)
            
            new_wb.save(save_path)
            messagebox.showinfo("成功", f"文件已生成:\n{save_path}\n\n(原文件未被修改)")
            self._log("全部完成。")

        except Exception as e:
            traceback.print_exc()
            self._log(f"错误: {e}")
            messagebox.showerror("运行错误", str(e))

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

    def _extract_data(self, file_path, sheet_name, col_idx, filter_ws):
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
                        row_dict[ti] = row[ti-1].value if ti <= len(row) else None
                    
                    wt = row[c_type-1].value
                    wn = row[c_no-1].value
                    if wt and wn:
                        data.append({
                            'base': row_dict,
                            'wo_key': (str(wt).strip(), str(wn).strip()),
                            'qty': float(qty), # O列: 排产数
                            'ws': curr_ws
                        })
            except: continue
        wb.close()
        return data

    # ============== 5. 数据库查询 (只读) ==============

    def _fetch_erp_data(self, keys):
        if not keys: return {}
        conditions = [f"(TA.TA001='{t}' AND TA.TA002='{n}')" for t, n in keys]
        data = defaultdict(lambda: {'total': 0, 'bom': []})
        
        batch_size = 200
        for i in range(0, len(conditions), batch_size):
            batch = conditions[i:i+batch_size]
            # SELECT 只读
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
            # SELECT 只读
            sql = f"SELECT RTRIM(MC001) p, SUM(MC007) q FROM INVMC WHERE MC001 IN ({p_str}) GROUP BY MC001"
            try:
                with pyodbc.connect(DB_CONN_STRING) as conn:
                    df = pd.read_sql(sql, conn)
                inv.update(pd.Series(df.q.values, index=df.p).to_dict())
            except: pass
        return inv

    # ============== 6. 核心计算逻辑 (数学平衡版) ==============

    def _simulate_logic(self, plans, wo_data, inventory):
        """
        核心逻辑 v7.3:
        1. 遍历每个计划。
        2. 计算 eff_demand (工单净需求) = min(排产数, ERP余额)。
        3. 计算 excess_qty (超出工单数) = 排产数 - eff_demand。
        4. 基于 eff_demand 计算物料需求和齐套率。
        """
        running_inv = inventory.copy()
        results = []

        for p in plans:
            key = p['wo_key']
            plan_qty = p['qty'] # 排产数 (O列)
            info = wo_data.get(key)
            
            res = {
                'rate': 0.0, 'achievable': 0, 
                'net_demand': 0, 'excess': 0,
                'msg': "", 'status': 'normal'
            }

            if not info or not info['bom']:
                res['msg'] = "无ERP信息"; res['status'] = 'error'
                results.append(res); continue

            # --- 步骤A: 确定工单净需求 (受ERP余额限制) ---
            eff_demand = plan_qty 
            
            for b in info['bom']:
                unit_use = b['req'] / info['total'] if info['total'] > 0 else 0
                if unit_use <= 0: continue
                
                theo_need = plan_qty * unit_use # 理论物料需求
                remain_issue = max(0, b['req'] - b['iss']) # ERP剩余可领
                
                # 如果 ERP剩余 < 理论需求，说明这单被ERP限制了
                if remain_issue < theo_need - 0.0001:
                    max_sets = remain_issue / unit_use
                    if max_sets < eff_demand:
                        eff_demand = max_sets
            
            # --- 步骤B: 计算超出部分 ---
            excess_qty = 0
            if eff_demand < plan_qty:
                excess_qty = plan_qty - eff_demand
                if excess_qty < 0.0001: excess_qty = 0
            
            # 特殊情况: ERP已结案/无余额
            if eff_demand < 0.001:
                res['msg'] = "工单已领完/结案"
                res['rate'] = 1.0; res['achievable'] = 0; 
                res['net_demand'] = 0; res['excess'] = int(plan_qty)
                res['status'] = 'finished'
                results.append(res); continue

            # --- 步骤C: 计算齐套率和缺料 (基于有效需求 eff_demand) ---
            min_material_rate = 1.0 # 齐套率初始化
            min_possible_sets = 999999
            short_details = []
            deduct_list = {}
            fully_kitted = True

            for b in info['bom']:
                unit_use = b['req'] / info['total'] if info['total'] > 0 else 0
                if unit_use <= 0: continue
                
                # 物料净需求 (基于 eff_demand)
                part_net_demand = eff_demand * unit_use
                
                stock = running_inv.get(b['part'], 0)
                deduct_list[b['part']] = part_net_demand
                
                # 1. 计算该物料的满足率
                if part_net_demand > 0:
                    part_rate = stock / part_net_demand
                    if part_rate > 1.0: part_rate = 1.0
                    # 齐套率取所有物料中的最小值
                    if part_rate < min_material_rate:
                        min_material_rate = part_rate
                
                # 2. 计算可产数量 (木桶效应)
                can_do = int(stock // unit_use)
                min_possible_sets = min(min_possible_sets, can_do)
                
                # 3. 缺料判断
                if stock < part_net_demand - 0.0001:
                    fully_kitted = False
                    diff = part_net_demand - stock
                    short_details.append(f"{b['name']}缺{diff:g}{b['unit']}")

            # --- 步骤D: 汇总结果 ---
            achievable = min(int(eff_demand), min_possible_sets)
            
            res['rate'] = min_material_rate # 齐套率
            res['achievable'] = achievable
            res['net_demand'] = int(eff_demand) # 工单净需求
            res['excess'] = int(excess_qty)     # 超出部分
            
            msgs = []
            if short_details: msgs.append("\n".join(short_details))
            res['msg'] = "\n".join(msgs)
            
            if not fully_kitted: res['status'] = 'short' # 缺料
            elif excess_qty > 0: res['status'] = 'warn'  # 齐套但有超出
            else: res['status'] = 'ok'                   # 完美齐套
            
            results.append(res)
            
            # 只有齐套部分才预扣减内存库存
            if fully_kitted:
                for k, v in deduct_list.items(): running_inv[k] -= v

        return results

    # ============== 7. 写入Excel (新文件) ==============

    def _write_headers(self, ws):
        curr = 1
        for idx in KEEP_COL_INDICES:
            c = ws.cell(1, curr); c.value = self.header_names_map.get(idx,""); 
            c.fill = self.header_fill; c.border = self.thin_border
            curr += 1
        
        # 新增列头
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
            # 基础列
            for idx in KEEP_COL_INDICES:
                c = ws.cell(ridx, curr); c.value = p['base'].get(idx)
                c.font = font; c.border = self.thin_border; c.alignment = Alignment(vertical="center")
                curr += 1
            
            # 结果列
            c_rate = ws.cell(ridx, curr); c_rate.value = r['rate']; c_rate.number_format = '0%'
            c_qty = ws.cell(ridx, curr+1); c_qty.value = r['achievable']
            
            c_net = ws.cell(ridx, curr+2)
            if r['net_demand'] == 0 and r['status'] == 'finished': c_net.value = "-"
            else: c_net.value = r['net_demand']
            
            c_excess = ws.cell(ridx, curr+3); c_excess.value = r['excess']
            
            c_msg = ws.cell(ridx, curr+4); c_msg.value = r['msg']
            c_msg.alignment = align
            
            # 样式设置
            for c in [c_rate, c_qty, c_net, c_excess]: c.border = self.thin_border; c.alignment = center
            c_msg.border = self.thin_border; c_msg.font = font
            
            # 颜色标记
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
