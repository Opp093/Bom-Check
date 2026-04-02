import pandas as pd
import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# ==================== Module 1: 硬件抽象层 (HAL) 与数字滤波器 ====================

# 引脚重映射字典：包含你见过的所有奇葩表头
ALIAS_DICT = {
    'Designator': ['位号', 'designator', 'refdes', 'reference', '编号'],
    'Quantity': ['数量', 'quantity', 'qty', 'count', '用量'], 
    'DNP': ['贴片状态', 'dnp', 'nc', '是否贴片', 'status', '空贴'],
    'K3_Code': ['k3 no.', 'k3 no', 'k3编码', '物料代码', '物料编码', 'k3 code', '编码','k3'],
    'Value': ['value', '值', '参数', '规格型号', '容值', '阻值'],
    'Footprint': ['footprint', '封装', 'package']
}

def clean_param_for_match(val):
    """高级特征滤波器：剥离 AD 祖传的无用前缀，提取核心参数"""
    s = str(val).strip().upper()
    # 剥离 AD 常见的封装前缀，例如 'CAPC-0805' -> '0805', 'RESC-0402' -> '0402'
    s = re.sub(r'^(CAPC|CAPAE|RESC|IND|LED|DIOC|SOP|QFN|SOT)-?', '', s)
    return s

def get_footprint_core(val):
    """封装带通滤波器：提取核心尺寸码，滤除后缀如 -C, -R (例如 0402-C -> 0402)"""
    s = str(val).strip().upper()
    # 强制捕获标准的 4 位 SMD 封装码
    m = re.search(r'(0201|0402|0603|0805|1206|1210|2010|2512|3216|3225)', s)
    if m: return m.group(1)
    
    # 滤除 AD 祖传的杂乱前缀和后缀
    s = re.sub(r'^(CAPC|CAPAE|RESC|IND|LED|DIOC|SOP|QFN|SOT|SMA|SMB|SMC)-?', '', s)
    s = re.sub(r'-[CRLM]$', '', s)
    return s

def expand_value(val):
    """等效参数发生器：处理 100nF=0.1uF 等单位换算。极致防爆版"""
    val = str(val).strip().upper()
    
    # [空信号旁路] 如果传入的是空值，直接返回空列表
    if not val: 
        return []

    # ================= [核心修复：精度符号对称清洗] =================
    # 抹除 AD 参数中的 ± 和 +/-，使其与 K3 库实现绝对对称匹配
    val = val.replace('±', '').replace('+/-', '')
    # ================================================================

    val = val.replace('Ω', '') 
    val = re.sub(r'([KMG])R$', r'\1', val) # 1MR -> 1M
    
    eqs = set() # 绝对明确的空集合声明，杜绝 {} 语法糖异常
    eqs.add(val)
    
    if re.match(r'^\d+(\.\d+)?R$', val): 
        eqs.add(val.replace('R', ''))
    elif re.match(r'^\d+(\.\d+)?$', val): 
        eqs.add(val + 'R') 
        
    if re.match(r'^\d+(\.\d+)?[KMG]$', val): 
        eqs.add(val + 'Ω')

    cap_match = re.match(r'^([\d\.]+)(P|N|U|M)F?$', val)
    if cap_match:
        num, unit = float(cap_match.group(1)), cap_match.group(2)
        # 抛弃批量 update，全部采用单发 add 指令，极致防弹
        if unit == 'N':  
            eqs.add(f"{num/1000:g}U"); eqs.add(f"{num/1000:g}UF")
            eqs.add(f"{num*1000:g}P"); eqs.add(f"{num*1000:g}PF")
            eqs.add(f"{num:g}N"); eqs.add(f"{num:g}NF")
        elif unit == 'U': 
            eqs.add(f"{num*1000:g}N"); eqs.add(f"{num*1000:g}NF")
            eqs.add(f"{num:g}U"); eqs.add(f"{num:g}UF")
        elif unit == 'P': 
            eqs.add(f"{num/1000:g}N"); eqs.add(f"{num/1000:g}NF")
            eqs.add(f"{num:g}P"); eqs.add(f"{num:g}PF")
            
    # 全局无 F 后缀补充
    for e in list(eqs):
        if not e: continue 
        if e.endswith('F'): 
            eqs.add(e[:-1])
        elif e[-1] in ['P', 'N', 'U', 'M']: 
            eqs.add(e + 'F')
        
    return list(eqs)

def clean_value(val):
    """底层信号清洗：消除大小写和多余空格"""
    return str(val).strip().upper()

def is_dnp(val):
    """施密特触发器：精准识别空贴状态"""
    s = str(val).strip().upper()
    if s in ['DNP', 'NC', '空贴', 'TRUE']: return True
    for kw in ['DNP', 'NC', '空贴']:
        if f" {kw}" in s or f"/{kw}" in s or f"_{kw}" in s or f"-{kw}" in s: return True
    for kw in ['DNP ', 'NC ', '空贴 ']:
        if s.startswith(kw): return True
    return False

def get_base_dir():
    """物理绝对寻址：防止 EXE 打包后找不到路径"""
    if getattr(sys, 'frozen', False): 
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

# ==================== Module 2: 智能总线驱动 (带表头雷达) ====================

def normalize_columns(df):
    """引脚别名规范化引擎"""
    new_cols = {}
    for col in df.columns:
        col_clean = str(col).strip().lower().replace('\ufeff', '')
        mapped_name = col
        for std_name, aliases in ALIAS_DICT.items():
            if col_clean in aliases:
                mapped_name = std_name
                break
        new_cols[col] = mapped_name
    df.rename(columns=new_cols, inplace=True)
    return df

def load_bom_file(filepath):
    """多协议智能加载器 (搭载雷达穿透废话行，防死锁)"""
    ext = os.path.splitext(filepath)[1].lower()
    try:
        # 1. 强行以“无表头”模式加载，防止错位
        if ext == '.csv':
            df = pd.read_csv(filepath, dtype=str, encoding='gbk', encoding_errors='ignore', header=None).fillna('')
        elif ext in ['.xlsx', '.xls', '.xlsm']:
            df = pd.read_excel(filepath, dtype=str, header=None).fillna('')
        else:
            return pd.DataFrame()
            
        if df.empty: return df

        # 2. 启动智能雷达扫描表头 (向下扫描30行)
        trigger_words = [a for aliases in ALIAS_DICT.values() for a in aliases]
        header_idx = 0
        
        for i in range(min(30, len(df))):
            row_vals = [str(x).strip().lower().replace('\ufeff', '') for x in df.iloc[i].values]
            if any(cell in trigger_words for cell in row_vals):
                header_idx = i
                break
                
        # 3. 截断无效数据，重构矩阵
        df.columns = df.iloc[header_idx].astype(str).tolist()
        df = df[header_idx + 1:].reset_index(drop=True)
        
        return normalize_columns(df)
        
    except PermissionError:
        messagebox.showerror("总线占用", f"致命错误：文件被占用！\n\n【{os.path.basename(filepath)}】 正被 WPS 或 Excel 打开。\n请立即关闭该文件后重试！")
        return pd.DataFrame()
    except Exception as e:
        messagebox.showerror("读取异常", f"底层错误：\n{e}")
        return pd.DataFrame()

# ==================== [防线升级：深度拆包与矩阵摊平引擎] ====================
def flatten_bom_matrix(df):
    """
    底层算法：将压缩的位号字符串 (如 R1, R2, R3) 瞬间爆炸拆解为独立的器件行
    """
    if df.empty or 'Designator' not in df.columns:
        return df
        
    df = df.copy()
    
    # 1. 切割位号：兼容中英文逗号、空格，并过滤掉空值和 pandas 产生的 NAN
    df['Designator'] = df['Designator'].apply(
        lambda x: [d.strip().upper() for d in re.split(r'[,，\s]+', str(x)) 
                   if d.strip() and d.strip().upper() not in ['NAN', 'NONE', 'NULL']]
    )
    
    # 2. 核心大招：矩阵爆炸 (Explode)！把 1 行压缩包变成 100 行物理清单
    df = df.explode('Designator').reset_index(drop=True)
    
    # 3. 拦截空位号
    df = df[df['Designator'].notna()]
    
    # 4. [物理防呆] 强制销毁 Quantity 数量列
    # 防止原本 100 变成 99 时，剩下的 99 个器件因为这 1 个数字的变化被集体误判为“修改”
    if 'Quantity' in df.columns:
        df.drop(columns=['Quantity'], inplace=True)
        
    # 5. 绝对防撞墙：如果工程师在画板时手误留下了重复的位号，强行去重保命
    df.drop_duplicates(subset=['Designator'], keep='first', inplace=True)
    
    return df
# ============================================================================

# ==================== Module 3: 核心运算引擎 (双核异构) ====================

def process_diff(old_path, new_path, output_filename):
    """[模式 A]：新旧 BOM 差异对比核对"""
    old_df = load_bom_file(old_path)
    new_df = load_bom_file(new_path)

    if old_df.empty or new_df.empty: return
    if 'Designator' not in old_df.columns or 'Designator' not in new_df.columns:
        messagebox.showerror("位号丢失", "未能识别到'位号 (Designator)'列，请检查文件格式。")
        return

    # ================= [激活防线：注入摊平引擎] =================
    old_df = flatten_bom_matrix(old_df)
    new_df = flatten_bom_matrix(new_df)
    # ============================================================

    old_df.set_index('Designator', inplace=True)
    new_df.set_index('Designator', inplace=True)

    old_keys, new_keys = set(old_df.index), set(new_df.index)
    excel_data = []
    

    # 1. 移除与新增计算
    for k in old_keys - new_keys:
        row = {'变更类型': '[-] 移除', '位号': k}
        for col in old_df.columns: row[col] = str(old_df.loc[k, col])
        excel_data.append(row)

    for k in new_keys - old_keys:
        row = {'变更类型': '[+] 新增', '位号': k}
        for col in new_df.columns: row[col] = str(new_df.loc[k, col])
        excel_data.append(row)

    # 2. 修改项与 DNP 拦截计算
    for key in old_keys & new_keys:
        old_row, new_row = old_df.loc[key], new_df.loc[key]
        row_data = {'变更类型': '[*] 修改', '位号': key}
        has_diff = False
        
        for col in list(set(old_df.columns) | set(new_df.columns)):
            in_old, in_new = col in old_df.columns, col in new_df.columns
            if in_old and in_new:
                val_old, val_new = old_row[col], new_row[col]
                if clean_value(val_old) != clean_value(val_new):
                    was_dnp, is_now_dnp = is_dnp(val_old), is_dnp(val_new)
                    if not was_dnp and is_now_dnp: row_data['变更类型'] = '[!] 警报: 变为空贴'
                    elif was_dnp and not is_now_dnp: row_data['变更类型'] = '[!] 警报: 恢复贴片'
                    row_data[col] = f"[{val_old}] -> [{val_new}]"
                    has_diff = True
            elif in_old and not in_new: row_data[col], has_diff = "[新版缺失此列]", True
            elif not in_old and in_new: row_data[col], has_diff = "[旧版缺失此列]", True
                
        if has_diff: excel_data.append(row_data)

    if not excel_data:
        messagebox.showinfo("核对完成", "完美匹配！两份 BOM 没有任何差异。")
        return

    result_df = pd.DataFrame(excel_data).fillna('')
    
    # ================= [防线升级：结果逆向重组与合并同类项 (Implode)] =================
    import re
    # 1. 提取所有非位号的列作为“聚类指纹”
    group_cols = [col for col in result_df.columns if col != '位号']
    
    # 2. 工业级自然排序算法 (解决 C10 排在 C2 前面的痛点)
    def natural_sort_key(s):
        return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', str(s))]
        
    def join_designators(series):
        """将散落的位号重新排序并用逗号压缩拼装"""
        return ', '.join(sorted(series.astype(str).tolist(), key=natural_sort_key))

    # 3. 核心大招：执行聚合压缩！
    # 把具有相同变更类型、相同Value、相同封装的器件强行拍扁成一行
    compressed_df = result_df.groupby(group_cols, as_index=False).agg({'位号': join_designators})
    # ===================================================================================

    # 重构表头顺序 (把变更类型和位号提到最前面)
    cols = compressed_df.columns.tolist()
    cols.insert(0, cols.pop(cols.index('变更类型')))
    cols.insert(1, cols.pop(cols.index('位号')))
    
    # 输出到最终 Excel
    compressed_df[cols].sort_values(by=['变更类型']).to_excel(output_filename, index=False)
    render_excel(output_filename, 'diff')

def process_lib_check(ad_path, lib_path, output_filename):
    """[模式 B]：AD BOM 与 公司 K3 库首次比对校验"""
    ad_df = load_bom_file(ad_path)
    lib_df = load_bom_file(lib_path) 

    if ad_df.empty or lib_df.empty: return
    if 'Designator' not in ad_df.columns or 'K3_Code' not in ad_df.columns:
        messagebox.showerror("识别失败", "AD BOM 中必须包含'位号'与'K3编码'相关列！")
        return
    if 'K3_Code' not in lib_df.columns:
        messagebox.showerror("识别失败", "公司 K3 库中未识别到'编码'列，请检查表头。")
        return

    # 1. 构建高速哈希字典
    lib_dict = {}
    for _, row in lib_df.iterrows():
        k3_code = str(row['K3_Code']).strip()
        if k3_code and k3_code != 'nan':
            specs = [str(row[c]).strip() for c in lib_df.columns if c != 'K3_Code' and str(row[c]).strip() not in ['', 'nan']]
            lib_dict[k3_code] = " | ".join(specs)

    # 2. 链路状态指示灯：防御断路风险
    messagebox.showinfo("底层诊断", f"数据总线加载完毕！\n\nAD 待测物料数: {len(ad_df)} 项\nK3 标准库物料数: {len(lib_dict)} 项\n\n注：如果K3库数量极少，请确认导出的库是否包含了所有元件大类！")

    # 3. 三重寻址交叉碰撞匹配
    excel_data = []
    for _, row in ad_df.iterrows():
        k3_code = str(row.get('K3_Code', '')).strip()
        ad_val = str(row.get('Value', '')).strip()
        ad_foot = str(row.get('Footprint', '')).strip()
        ad_qty = str(row.get('Quantity', '')).strip() # <--- [新增] 提取数量数据
        
        row_data = {
            '位号 (Designator)': str(row.get('Designator', '')),
            '数量 (Qty)': ad_qty,  # <--- [新增] 将数量接入输出总线
            'AD K3编码': k3_code,
            'AD 原理图参数': f"Val: {ad_val} | Foot: {ad_foot}",
            '校验状态': '',
            '公司 K3 库标准参数': ''
        }
        # ... 后面的状态机和切片逻辑保持完全不变 ...

        # 状态机 A：缺件或彻底找不到
        if k3_code in ['', 'nan', 'NONE']:
            row_data['校验状态'] = '[!] 缺少编码'
        elif k3_code not in lib_dict:
            row_data['校验状态'] = '[x] 库中无此物料'
            
        # 状态机 B：编码存在 -> 启动深度交叉验证
        else:
            k3_specs_upper = str(lib_dict[k3_code]).upper()
            # ================= [核心终极修复：对称底噪抹平] =================
            # AD 的信号在底层被抹去了 Ω 和 ±，K3 库也必须同等抹去，否则会引发非对称断路！
            k3_specs_upper = k3_specs_upper.replace('Ω', '').replace('±', '').replace('+/-', '')
            # ================================================================
            designator_upper = str(row.get('Designator', '')).strip().upper()
            
            # ================= [防线升级 1] DNP 空贴噪声剥离与状态捕捉 =================
            is_row_dnp = False
            # 1. 检查是否存在专属的 DNP 列
            if 'DNP' in ad_df.columns and is_dnp(str(row.get('DNP', ''))):
                is_row_dnp = True
            # 2. 检查 Value 列是否夹杂了 DNP 字眼 (如 10K_DNP 或 0.1uF/NC)
            elif is_dnp(ad_val):
                is_row_dnp = True
                
            # 从 Value 中强行洗掉 DNP/NC/空贴 等噪音字眼，防止污染后续的切片校验
            clean_ad_val = ad_val.upper()
            
            # 【终极修复】应用严格的字母数字边界判定，彻底放过 NC7SZ 等芯片型号
            for kw in ['DNP', 'NC']:
                # (?<![A-Z0-9]) 确保前面不是字母或数字
                # (?![A-Z0-9])  确保后面不是字母或数字
                clean_ad_val = re.sub(r'(?<![A-Z0-9])' + kw + r'(?![A-Z0-9])', '', clean_ad_val)
            
            # 中文“空贴”不会与芯片型号重合，直接无脑切除即可
            clean_ad_val = re.sub(r'空贴', '', clean_ad_val)
            # =========================================================================

            # --- 通道 1：封装滤波器检测 (加装智能旁路路由) ---
            foot_core = get_footprint_core(ad_foot)
            is_rc_passive = re.match(r'^(R|C|RN|CN)\d+', designator_upper)
            
            if is_rc_passive:
                match_foot = (foot_core == '') or (foot_core in k3_specs_upper)
            else:
                match_foot = True
            
            # --- 通道 2：参数裂变器扫描 (多维交叉切片) ---
            # 【注意】这里使用洗净了 DNP 噪音的 clean_ad_val 进行切片
            sub_vals = [s.strip() for s in re.split(r'[_,\/\|\s]+', clean_ad_val) if s.strip()]
            
            match_val = True
            val_conflicts = [] 
            
            for sv in sub_vals:
                # [核心修正] 拆除 K3 编码套娃旁路，强制对 Value 列的所有切片进行物理参数校验
                sv_eqs = expand_value(sv) 
                if not any(eq in k3_specs_upper for eq in sv_eqs):
                    match_val = False
                    val_conflicts.append(sv)
            
            # --- 与门仲裁 ---
            # 如果捕捉到了 DNP 状态，准备好专属后缀标签
            dnp_tag = " (空贴/DNP)" if is_row_dnp else ""
            
            if match_val and match_foot:
                row_data['校验状态'] = f'[√] 完美匹配{dnp_tag}'
            else:
                conflicts = []
                if not match_val: conflicts.append(f"参数({','.join(val_conflicts)})冲突")
                if not match_foot: conflicts.append(f"封装({ad_foot})冲突")
                row_data['校验状态'] = f'[!] {", ".join(conflicts)}{dnp_tag}'
                
            row_data['公司 K3 库标准参数'] = lib_dict[k3_code]
            
        # ==================== [核心修复：焊接数据输出引脚] ====================
        excel_data.append(row_data) 
        # ======================================================================

    # 循环彻底结束后，再将收集到的数据转化为表格并排序
    result_df = pd.DataFrame(excel_data).fillna('')
    result_df.sort_values(by=['校验状态', '位号 (Designator)'], inplace=True)

    result_df = pd.DataFrame(excel_data).fillna('')
    result_df.sort_values(by=['校验状态', '位号 (Designator)'], inplace=True)
    result_df.to_excel(output_filename, index=False)
    render_excel(output_filename, 'check')

# ==================== Module 4: 渲染层 (莫兰迪护眼色彩驱动) ====================

def render_excel(filename, mode):
    wb = load_workbook(filename)
    ws = wb.active
    
    # 调色盘：启用低饱和度护眼色系 (莫兰迪色系)
    soft_red = PatternFill("solid", fgColor="FADBD8")    # 柔和粉红 (移除 / 找不到物料)
    soft_green = PatternFill("solid", fgColor="D5F5E3")  # 柔和薄荷绿 (新增 / 完美匹配)
    soft_yellow = PatternFill("solid", fgColor="FCF3CF") # 柔和奶黄 (数值修改)
    soft_orange = PatternFill("solid", fgColor="FDEBD0") # 柔和浅橙 (警报 / 参数冲突)
    
    # 柔和警报字体：使用砖红色替代刺眼的纯正红色
    alert_font = Font(bold=True, color="C0392B") 

    if mode == 'diff':
        fills = {'[-]': soft_red, '[+]': soft_green, '[!]': soft_orange, '[*]': soft_yellow}
        for row in range(2, ws.max_row + 1):
            val = str(ws.cell(row=row, column=1).value)
            if '移除' in val: [setattr(c, 'fill', fills['[-]']) for c in ws[row]]
            elif '新增' in val: [setattr(c, 'fill', fills['[+]']) for c in ws[row]]
            elif '警报' in val: 
                for c in ws[row]: c.fill, c.font = fills['[!]'], alert_font
            elif '修改' in val:
                for c in ws[row]:
                    if c.value and '->' in str(c.value): c.fill = fills['[*]']

    elif mode == 'check':
        fills = {'[x]': soft_red, '[!]': soft_orange, '[√]': soft_green}
        for row in range(2, ws.max_row + 1):
            # [核心修正] 因为插入了数量列，校验状态偏移到了第 5 列，寻址指针随之修改
            val = str(ws.cell(row=row, column=5).value) 
            for key in fills:
                if key in val:
                    for c in ws[row]: 
                        c.fill = fills[key]
                        if key == '[!]': c.font = alert_font
                    break

    wb.save(filename)
    messagebox.showinfo("执行完毕", f"🎉 报表生成成功！\n已应用护眼色彩渲染，保存在:\n{filename}")

# ==================== Module 5: 人机交互界面 (自研机械导航引擎 + 动态说明书版) ====================

class GradientHeader(tk.Canvas):
    """自研 CPU 软件渲染器：通过逐行扫描绘制高拟真渐变背景"""
    def __init__(self, parent, color1, color2, **kwargs):
        super().__init__(parent, **kwargs)
        self._color1 = color1
        self._color2 = color2
        self.bind("<Configure>", self._draw_gradient)

    def _draw_gradient(self, event=None):
        self.delete("gradient")
        width, height = self.winfo_width(), self.winfo_height()
        if width <= 1 or height <= 1: return 
        
        r1, g1, b1 = self.winfo_rgb(self._color1)
        r2, g2, b2 = self.winfo_rgb(self._color2)
        r_ratio, g_ratio, b_ratio = (r2-r1)/width, (g2-g1)/width, (b2-b1)/width

        for i in range(width):
            nr, ng, nb = int(r1 + (r_ratio * i)), int(g1 + (g_ratio * i)), int(b1 + (b_ratio * i))
            color = "#%4.4x%4.4x%4.4x" % (nr, ng, nb)
            self.create_line(i, 0, i, height, tags=("gradient",), fill=color)
        self.tag_lower("gradient")

def show_instructions(root):
    """编译期固化说明书引擎：防篡改，仅开发者可通过重新打包修改"""
    help_win = tk.Toplevel(root)
    help_win.title("📖 Bom-Check使用说明")
    help_win.geometry("580x460")
    help_win.configure(bg="#EAECEE")
    
    # 设置模态窗口（置顶，且必须关闭后才能操作主界面）
    help_win.transient(root)
    help_win.grab_set()

    # ==================== [核心重构：防篡改的内置固化存储] ====================
    # 👨‍💻 开发者专属修改区：直接在这里修改文字，然后用 PyInstaller 重新打包即可！
    developer_instructions = """欢迎使用 Excite Bom-check HW V4.1!

【模式 A:硬件改版差异核对】
1. 作用：对比新旧两份 BOM,提取新增、移除、修改(包括变为空贴)的器件。
2. 注：两份表头必须包含“位号”"k3 no"列。

【模式 B:首次发板 K3 校验】
1. 作用：将 EDA 导出的 BOM 与公司 K3 库进行三维交叉比对(K3 No.+Value+Footprint)。
2. 高级特性：
   - 自动等效换算:100nF=0.1uF,1MR=1M 等。
   - DNP 状态剥离：自动从参数中剔除 DNP/NC/空贴 字符。(需要将DNP参数写在Value中)
   - 智能旁路路由：芯片(U)、电感(L)、二极管(D) 等物料,由于K3中不会存在封装信息,所以屏蔽封装检测,检测value与K3编码

=========================================
具体使用详情查看内部飞书文档
https://s17weazhe5w.feishu.cn/wiki/QDXNwjkjriEpdmkqfD0cvmALn2b
   
=========================================
⚠️ 内部管控工具，请勿外传。
如有增加新物料规则或功能需求，请联系工具开发者。"""
    # ==========================================================================

    # 渲染滚动文本视窗
    txt_frame = tk.Frame(help_win, bd=2, relief=tk.SUNKEN)
    txt_frame.pack(expand=True, fill='both', padx=20, pady=20)

    txt = tk.Text(txt_frame, font=("Microsoft YaHei", 10), bg="#FDFEFE", fg="#2C3E50", wrap="word", padx=10, pady=10)
    scrollbar = ttk.Scrollbar(txt_frame, command=txt.yview)
    txt.configure(yscrollcommand=scrollbar.set)
    
    scrollbar.pack(side='right', fill='y')
    txt.pack(side='left', expand=True, fill='both')
    
    # 将内置的固化文字插入到界面中
    txt.insert('1.0', developer_instructions)
    # [安全锁] 彻底锁定文本框，禁止用户在界面上敲击键盘篡改内容
    txt.config(state='disabled')

def run_app():
    root = tk.Tk()
    root.title("BOM-Check V4.1 ")
    root.geometry("760x540") # 稍微拉宽以容纳帮助按钮
    root.configure(bg="#EAECEE") 
    
    # ---------------- 1. 顶部炫酷渐变横幅 ----------------
    header_frame = tk.Frame(root, height=60)
    header_frame.pack(fill='x')
    header_frame.pack_propagate(False) 
    
    gradient = GradientHeader(header_frame, color1="#1A2980", color2="#26D0CE", highlightthickness=0)
    gradient.place(relwidth=1, relheight=1)
    tk.Label(header_frame, text="Excite BOM-Check HW V4.1", font=("Microsoft YaHei", 18, "bold"), 
             fg="white", bg="#1A2980").place(relx=0.5, rely=0.5, anchor="center")

    # ---------------- 2. 引入基础样式 ----------------
    style = ttk.Style()
    if 'clam' in style.theme_names(): style.theme_use('clam')
    font_main = ("Microsoft YaHei", 10)
    font_bold = ("Microsoft YaHei", 10, "bold")
    
    style.configure("TLabelframe", background="#FFFFFF", font=font_bold, foreground="#2E4053")
    style.configure("TLabelframe.Label", background="#FFFFFF")

    # ==================== [核心重构：带有帮助按钮的导航矩阵] ====================
    nav_frame = tk.Frame(root, bg="#EAECEE")
    nav_frame.pack(fill='x', pady=15, padx=20)
    
    # 建立帮助按钮 (靠右放置，小巧精致)
    btn_help = tk.Button(nav_frame, text="📖 使用说明", font=("Microsoft YaHei", 10, "bold"), 
                         fg="#2980B9", bg="#EAECEE", relief=tk.GROOVE, bd=2, cursor="hand2",
                         command=lambda: show_instructions(root))
    btn_help.pack(side='right', fill='y', padx=(10, 0))
    
    # 建立两个宏大的物理按键
    btn_a = tk.Button(nav_frame, text="🔄 模式 A：差异核对", font=("Microsoft YaHei", 12, "bold"), cursor="hand2")
    btn_b = tk.Button(nav_frame, text="🔍 模式 B：K3 校验", font=("Microsoft YaHei", 12, "bold"), cursor="hand2")
    
    btn_a.pack(side='left', expand=True, fill='x', padx=(0, 10), ipady=8)
    btn_b.pack(side='left', expand=True, fill='x', padx=(0, 0), ipady=8)
    
    # 建立主工作区容器
    work_area = tk.Frame(root, bg="#FFFFFF", bd=2, relief=tk.GROOVE)
    work_area.pack(fill='both', expand=True, padx=20, pady=(0, 20))
    
    tab_a = tk.Frame(work_area, bg="#FFFFFF")
    tab_b = tk.Frame(work_area, bg="#FFFFFF")

    # ======================== 选项卡 1：差异核对面板内容 ========================
    tk.Label(tab_a, text=" 应用场景：版本迭代时，精准捕获元器件的新增、移除、数值修改及空贴变化 ", 
             fg="#85929E", bg="#FFFFFF", font=("Microsoft YaHei", 9)).pack(pady=10)
    frame_a = ttk.LabelFrame(tab_a, text=" 导入配置 (Mode A) ")
    frame_a.pack(fill='x', padx=30, pady=5)
    
    v_old, v_new = tk.StringVar(), tk.StringVar()
    f1 = tk.Frame(frame_a, bg="#FFFFFF"); f1.pack(fill='x', padx=15, pady=8)
    tk.Button(f1, text="选择基准旧版", width=12, font=font_main, bg="#ECF0F1", relief=tk.GROOVE, bd=2,
              command=lambda: v_old.set(filedialog.askopenfilename())).pack(side='left')
    ttk.Entry(f1, textvariable=v_old, state='readonly').pack(side='left', fill='x', expand=True, padx=10)
    
    f2 = tk.Frame(frame_a, bg="#FFFFFF"); f2.pack(fill='x', padx=15, pady=8)
    tk.Button(f2, text="选择待测新版", width=12, font=font_main, bg="#ECF0F1", relief=tk.GROOVE, bd=2,
              command=lambda: v_new.set(filedialog.askopenfilename())).pack(side='left')
    ttk.Entry(f2, textvariable=v_new, state='readonly').pack(side='left', fill='x', expand=True, padx=10)
    
    tk.Button(tab_a, text="⚡ 开始执行差异核对", font=("Microsoft YaHei", 12, "bold"),
              bg="#2980B9", fg="white", activebackground="#1F618D", activeforeground="white", 
              relief=tk.RAISED, bd=5, cursor="hand2", 
              command=lambda: process_diff(v_old.get(), v_new.get(), os.path.join(get_base_dir(), "BOM_差异对比报告.xlsx"))).pack(pady=20, fill='x', padx=100)

    # ======================== 选项卡 2：K3 主数据校验面板内容 ========================
    tk.Label(tab_b, text=" 应用场景：新板打样前，三维交叉验证图纸器件是否在 K3 ERP 库中合法存在 ", 
             fg="#85929E", bg="#FFFFFF", font=("Microsoft YaHei", 9)).pack(pady=10)
    frame_b = ttk.LabelFrame(tab_b, text=" 导入配置 (Mode B) ")
    frame_b.pack(fill='x', padx=30, pady=5)

    v_ad, v_lib = tk.StringVar(), tk.StringVar()
    f3 = tk.Frame(frame_b, bg="#FFFFFF"); f3.pack(fill='x', padx=15, pady=8)
    tk.Button(f3, text="AD 导出 BOM", width=12, font=font_main, bg="#ECF0F1", relief=tk.GROOVE, bd=2,
              command=lambda: v_ad.set(filedialog.askopenfilename())).pack(side='left')
    ttk.Entry(f3, textvariable=v_ad, state='readonly').pack(side='left', fill='x', expand=True, padx=10)
    
    f4 = tk.Frame(frame_b, bg="#FFFFFF"); f4.pack(fill='x', padx=15, pady=8)
    tk.Button(f4, text="公司 K3 总库", width=12, font=font_main, bg="#ECF0F1", relief=tk.GROOVE, bd=2,
              command=lambda: v_lib.set(filedialog.askopenfilename())).pack(side='left')
    ttk.Entry(f4, textvariable=v_lib, state='readonly').pack(side='left', fill='x', expand=True, padx=10)

    tk.Button(tab_b, text="⚡ 开始三维交叉比对", font=("Microsoft YaHei", 12, "bold"),
              bg="#27AE60", fg="white", activebackground="#1D8348", activeforeground="white", 
              relief=tk.RAISED, bd=5, cursor="hand2", 
              command=lambda: process_lib_check(v_ad.get(), v_lib.get(), os.path.join(get_base_dir(), "K3物料校验报告.xlsx"))).pack(pady=20, fill='x', padx=100)

    # ==================== [导航引擎的底层路由逻辑] ====================
    def switch_to_a():
        btn_a.config(relief=tk.SUNKEN, bg="#2980B9", fg="white", activebackground="#2980B9")
        btn_b.config(relief=tk.RAISED, bg="#E5E7E9", fg="#A6ACAF", activebackground="#E5E7E9")
        tab_b.pack_forget() 
        tab_a.pack(fill='both', expand=True, padx=10, pady=10) 
        
    def switch_to_b():
        btn_b.config(relief=tk.SUNKEN, bg="#27AE60", fg="white", activebackground="#27AE60")
        btn_a.config(relief=tk.RAISED, bg="#E5E7E9", fg="#A6ACAF", activebackground="#E5E7E9")
        tab_a.pack_forget() 
        tab_b.pack(fill='both', expand=True, padx=10, pady=10) 

    btn_a.config(command=switch_to_a)
    btn_b.config(command=switch_to_b)
    switch_to_a() 
    
    root.mainloop()

if __name__ == "__main__":
    run_app()