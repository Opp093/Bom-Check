import pandas as pd
import os

# [新增模块]：多协议自动协商加载器
def load_bom_file(filepath):
    # 提取文件后缀名并转换为小写，例如 '.csv', '.xlsx'
    ext = os.path.splitext(filepath)[1].lower()
    
    # 路由分支 1：处理 CSV 纯文本协议
    if ext == '.csv':
        # 加入 encoding_errors='ignore' 滤除脏数据噪声
        return pd.read_csv(filepath, dtype=str, encoding='gbk').fillna('')
    
    # 路由分支 2：处理 Excel 强格式协议
    elif ext in ['.xlsx', '.xls', '.xlsm']:
        # Excel 属于结构化二进制文件，没有 GBK 乱码问题，直接调用 read_excel
        return pd.read_excel(filepath, dtype=str).fillna('')
    
    # 路由分支 3：错误协议拦截
    else:
        print(f"\n[致命错误] 无法解析的文件格式: {filepath}")
        print(">>> 仅支持 .csv, .xlsx, .xls, .xlsm 格式！ <<<")
        return pd.DataFrame() # 返回空矩阵，触发后续的安全断电机制

def generate_bom_excel(old_bom_path, new_bom_path, output_filename="BOM_差异核对矩阵.xlsx"):
    # 1. 调用智能加载模块
    old_df = load_bom_file(old_bom_path)
    new_df = load_bom_file(new_bom_path)

    # 安全断电机制：如果加载出空矩阵（比如传错了文件格式），立刻终止运行
    if old_df.empty or new_df.empty:
        print("[系统提示] 数据加载失败，程序终止。")
        return

    # 2. 确立主键寻址
    old_df.set_index('Designator', inplace=True)
    new_df.set_index('Designator', inplace=True)

    # 3. 提取交差集
    old_keys = set(old_df.index)
    new_keys = set(new_df.index)

    added_keys = new_keys - old_keys
    removed_keys = old_keys - new_keys
    common_keys = old_keys & new_keys

    # 4. 构建结构化的数据缓存
    excel_data = []

    # 4.1 处理被移除的元件
    for k in removed_keys:
        old_row = old_df.loc[k]
        row_data = {'变更类型': '[-] 移除', '位号': k}
        for col in old_df.columns:
            if old_row[col] != '':
                row_data[col] = f"{old_row[col]}"
        excel_data.append(row_data)

    # 4.2 处理新增的元件
    for k in added_keys:
        new_row = new_df.loc[k]
        row_data = {'变更类型': '[+] 新增', '位号': k}
        for col in new_df.columns:
            if new_row[col] != '':
                row_data[col] = f"{new_row[col]}"
        excel_data.append(row_data)

    # 4.3 遍历处理被修改的元件
    for key in common_keys:
        old_row = old_df.loc[key]
        new_row = new_df.loc[key]
        row_data = {'变更类型': '[*] 修改', '位号': key}
        has_diff = False
        
        for col in old_df.columns:
            if col in new_df.columns:
                if str(old_row[col]) != str(new_row[col]):
                    row_data[col] = f"[{old_row[col]}] -> [{new_row[col]}]"
                    has_diff = True
            else:
                row_data[col] = "[新版缺失此列]"
                has_diff = True
        
        if has_diff: 
            excel_data.append(row_data)

    # 5. 驱动输出模块 (烧录为 Excel)
    if excel_data:
        result_df = pd.DataFrame(excel_data).fillna('')
        cols = result_df.columns.tolist()
        cols.insert(0, cols.pop(cols.index('变更类型')))
        cols.insert(1, cols.pop(cols.index('位号')))
        result_df = result_df[cols]
        
        result_df.sort_values(by=['变更类型', '位号'], inplace=True)
        result_df.to_excel(output_filename, index=False)
        print(f"\n[系统提示] >>> 完美！多列矩阵 Excel 报告已生成并保存在当前目录: {output_filename} <<<")
    else:
        print("\n[系统提示] >>> 完美！两份 BOM 完全一致，没有发现任何差异。 <<<")

def find_bom_file(prefix_name):
    # 1. 烧录优先级仲裁序列 (Priority Queue)
    # 越靠前，优先级越高。优先寻找新版的高精度 Excel
    priority_exts = ['.xlsx', '.xls', '.xlsm', '.csv']
    
    # 2. 依次轮询扫描引脚
    for ext in priority_exts:
        target_file = prefix_name + ext  # 组合出目标文件名，如 'old_bom.xlsx'
        if os.path.exists(target_file):  # 探针检测：如果物理硬盘上存在这个文件
            return target_file           # 立即锁定目标并返回，后面的低优先级格式不再理会
            
    # 3. 扫描完一圈都没找到，返回高阻态 (None)
    return None 

# 触发扫描雷达
print("\n[系统启动] 正在扫描当前目录下的 BOM 文件...")
old_bom_target = find_bom_file('old_bom')
new_bom_target = find_bom_file('new_bom')

# 状态机判断
if not old_bom_target or not new_bom_target:
    print("[致命故障] 扫描失败！引脚悬空！")
    print(">>> 请确保当前文件夹内至少包含一个 'old_bom' 和一个 'new_bom' 文件！ <<<")
else:
    print(f"[锁定目标] 读取旧版基准: {old_bom_target}")
    print(f"[锁定目标] 读取新版数据: {new_bom_target}")
    
    # 将锁定好的最高优先级文件，送入对比引擎
    generate_bom_excel(old_bom_target, new_bom_target)

# 阻塞程序挂起，防止界面闪退
input("\n[程序执行完毕] 请按回车键 (Enter) 退出窗口...")