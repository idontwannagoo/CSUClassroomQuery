import openpyxl
import csv
import re
import os  # 添加 os 模块用于文件操作
from datetime import datetime  # 添加 datetime 用于生成时间戳
from itertools import islice
from tqdm import tqdm
from openpyxl.styles import Alignment, Font  # 导入 Font
from openpyxl.utils import get_column_letter

def find_file_pairs(directory="source"):
    """在指定目录中查找成对的 .csv 和 .xlsx 文件"""
    pairs = {}
    try:
        files = os.listdir(directory)
        csv_files = {os.path.splitext(f)[0] for f in files if f.lower().endswith('.csv')}
        xlsx_files = {os.path.splitext(f)[0] for f in files if f.lower().endswith('.xlsx')}
        
        common_basenames = sorted(list(csv_files.intersection(xlsx_files)))
        
        if not common_basenames:
            print(f"在 '{directory}' 目录下未找到匹配的 .csv 和 .xlsx 文件对。")
            return {}
            
        print(f"在 '{directory}' 目录下找到以下文件对:")
        for i, basename in enumerate(common_basenames):
            pairs[i + 1] = basename
            print(f"  {i + 1}: {basename}")
        return pairs
        
    except FileNotFoundError:
        print(f"错误：目录 '{directory}' 不存在。")
        return {}
    except Exception as e:
        print(f"查找文件对时出错：{str(e)}")
        return {}

def load_classrooms(csv_filepath):
    """从指定的CSV文件加载教室列表"""
    classrooms = set()
    try:
        # 尝试用 gbk 解码，如果失败则尝试 utf-8
        try:
            with open(csv_filepath, 'r', encoding='gbk') as f:
                reader = csv.reader(f)
                for row in reader:
                    if row:
                        classrooms.add(row[0])
        except UnicodeDecodeError:
            print(f"警告：使用 gbk 解码 '{csv_filepath}' 失败，尝试使用 utf-8...")
            with open(csv_filepath, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                for row in reader:
                    if row:
                        classrooms.add(row[0])
        return classrooms
    except FileNotFoundError:
        print(f"错误：CSV文件 '{csv_filepath}' 未找到。")
        return set()
    except Exception as e:
        print(f"加载教室列表 '{csv_filepath}' 时出错：{str(e)}")
        return set()

def parse_time_periods(time_str):
    """解析时间字符串，返回(星期几, 节次列表)的元组"""
    if not time_str or not isinstance(time_str, str):
        return None
    
    # 提取星期几（第一个数字）
    day = int(time_str[0])
    if day == 6 or day == 7:  # 忽略周六周日
        return None
    
    # 提取课节
    periods = []
    for i in range(1, len(time_str), 2):
        if i + 1 < len(time_str):
            period = int(time_str[i:i+2])
            if 1 <= period <= 10:
                period_group = (period + 1) // 2  # 将1-10转换为1-5
                if period_group not in periods:
                    periods.append(period_group)
    
    return (day, sorted(periods)) if periods else None

def parse_weeks(week_range, single_double=None):
    """解析周次范围，返回周数集合"""
    weeks = set()
    
    # 如果week_range为空且single_double为空，返回空集合
    if not week_range and not single_double:
        return weeks
    
    # 如果week_range是空的但有single_double值，说明是单个周数
    if not week_range and single_double:
        try:
            week = int(single_double)
            weeks.add(week)
            return weeks
        except ValueError:
            return weeks
    
    # 处理多个部分（用逗号分隔）
    parts = str(week_range).split(',')
    
    for part in parts:
        part = part.strip()
        if '-' in part:
            # 处理范围（如"1-8"）
            try:
                start, end = map(int, part.split('-'))
                if single_double == '单周':
                    weeks.update(w for w in range(start, end + 1) if w % 2 == 1)
                elif single_double == '双周':
                    weeks.update(w for w in range(start, end + 1) if w % 2 == 0)
                else:  # 全周
                    weeks.update(range(start, end + 1))
            except ValueError:
                continue
        else:
            # 处理单个数字
            try:
                week = int(part)
                weeks.add(week)
            except ValueError:
                continue
    
    return weeks

def process_excel(basename, show_occupied=False, source_dir="source", output_dir="."):
    """处理指定基名的Excel文件并生成教室表
    
    Args:
        basename (str): 要处理的文件基名 (例如 "TianXin")
        show_occupied (bool): True显示占用教室，False显示空闲教室
        source_dir (str): 源文件目录
        output_dir (str): 输出文件目录
    """
    csv_filepath = os.path.join(source_dir, f"{basename}.csv")
    xlsx_filepath = os.path.join(source_dir, f"{basename}.xlsx")
    
    # 校区名称映射
    campus_map = {
        "TianXin": "天心校区",
        "XiaoXiang": "潇湘校区",
        "XingLin": "杏林校区",
        "YueLuShan": "岳麓山校区"
    }
    campus_name = campus_map.get(basename, "未知校区") # 获取校区名，默认为未知
    status_text = "占用" if show_occupied else "空闲"

    print(f"\n--- 开始处理 {basename} ({campus_name}) ---")
    print(f"源文件: {csv_filepath}, {xlsx_filepath}")
    
    # 加载所有教室
    all_classrooms = load_classrooms(csv_filepath)
    if not all_classrooms:
        print(f"错误：无法从 '{csv_filepath}' 加载教室列表，跳过 {basename}")
        return
    
    # 创建用于存储课程信息的数据结构
    # 格式：{(周数, 星期几, 节次): set(已用教室)}
    used_classrooms = {}
    
    workbook = None # 初始化 workbook 变量
    # 加载Excel文件
    try:
        workbook = openpyxl.load_workbook(xlsx_filepath, read_only=True, data_only=True)
        sheet = workbook.active
        
        # 获取总行数 (使用 max_row 替代迭代)
        total_rows = sheet.max_row
        
        print("正在读取课程信息...")
        
        with tqdm(total=total_rows - 3, desc=f"处理 {basename} 数据", ncols=100) as pbar: # 假设标题占3行
            # 从第4行开始读取 (openpyxl 行号是 1-based)
            for row in sheet.iter_rows(min_row=4):
                try:
                    # 使用 get_column_letter 获取列名可能更清晰，但索引也可以
                    time_str = str(row[6].value) if row[6].value else None  # 第7列 (G)
                    classroom = str(row[7].value) if row[7].value else None # 第8列 (H)
                    week_range = str(row[8].value) if row[8].value else None # 第9列 (I)
                    single_double = str(row[9].value) if row[9].value else None # 第10列 (J)
                    
                    if time_str and classroom:
                        # 解析时间
                        time_info = parse_time_periods(time_str)
                        if time_info:
                            day, periods = time_info
                            
                            # 解析周次
                            weeks = parse_weeks(week_range, single_double)
                            
                            # 更新已用教室信息
                            for week in weeks:
                                for period in periods:
                                    key = (week, day, period)
                                    if key not in used_classrooms:
                                        used_classrooms[key] = set()
                                    # 处理可能包含多个教室的情况，例如 "J01 101 J01 102"
                                    # 使用正则表达式查找所有符合模式的教室号 (例如 JXX XXX 或 XXX)
                                    found_classrooms = re.findall(r'[A-Za-z]\d{2}\s\d{3}|[A-Za-z]\d+\s*\d*|[A-Za-z]{1,3}\d{1,3}', classroom)
                                    if found_classrooms:
                                         used_classrooms[key].update(c.strip() for c in found_classrooms)
                                    elif classroom: # 如果正则没匹配到，但classroom字段不为空，直接添加
                                        used_classrooms[key].add(classroom.strip())
                
                except Exception as e:
                    print(f"\n处理 {basename} 的行 {row[0].row} 时出错：{str(e)}") # 打印行号
                
                pbar.update(1)
        
        # 创建输出Excel
        print(f"\n正在为 {basename} 生成{status_text}教室表...")
        output_workbook = openpyxl.Workbook()
        
        # 定义 Maple Mono 字体 (用于正文)
        maple_font = Font(name='Maple Mono Normal NL NF CN')
        
        # 定义标题字体
        title_font = Font(name='思源黑体 VF Medium', size=22)

        # 为每周创建工作表
        for week in range(1, 17): # 假设教学周为1-16周
            # 创建该周的工作表
            sheet_name = f"第{week}周"
            if week == 1:
                sheet = output_workbook.active
                sheet.title = sheet_name
            else:
                sheet = output_workbook.create_sheet(sheet_name)
            
            # --- 添加标题行 (Row 1) ---
            sheet.insert_rows(1) # 在顶部插入一行
            sheet.merge_cells('A1:F1') # 合并第一行的单元格 A1 到 F1
            title_cell = sheet['A1']
            title_string = f"第{week}周 {campus_name}{status_text}教室"
            title_cell.value = title_string
            title_cell.alignment = Alignment(horizontal='center', vertical='center')
            title_cell.font = title_font # 应用标题字体
            sheet.row_dimensions[1].height = 35 # 设置标题行的高度
            # --- 标题行结束 ---

            # 设置表头 (Row 2 and Column 1 from Row 3 onwards)
            days = ["", "周一", "周二", "周三", "周四", "周五"]
            periods = ["", "1-2节", "3-4节", "5-6节", "7-8节", "9-10节"]

            # Day headers (Row 2)
            for col_idx, day in enumerate(days, 1):
                cell = sheet.cell(row=2, column=col_idx, value=day)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                if col_idx > 1: # Apply Maple font to Mon-Fri headers
                    cell.font = maple_font
            
            # Period headers (Column 1, starting Row 3)
            for row_idx, period in enumerate(periods, 1):
                if row_idx > 1: # Skip the empty string placeholder
                    cell = sheet.cell(row=row_idx + 1, column=1, value=period) # row_idx starts at 1, need data from row 3
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.font = maple_font
            
            # 填充数据 (Starting Row 3, Column 2)
            data_start_row = 3
            data_end_row = data_start_row + len(periods) - 2 # -1 for zero-based, -1 for empty placeholder
            data_start_col = 2
            data_end_col = data_start_col + len(days) - 2 # -1 for zero-based, -1 for empty placeholder

            for day_idx in range(1, len(days)): # 1 to 5 (Mon to Fri)
                for period_idx in range(1, len(periods)): # 1 to 5 (1-2节 to 9-10节)
                    key = (week, day_idx, period_idx)
                    used = used_classrooms.get(key, set())
                    
                    # 从 all_classrooms 中移除 used 教室来获取空闲教室
                    available_classrooms = all_classrooms - used
                    
                    # 根据 show_occupied 选项决定显示哪些教室
                    classrooms_to_show = used if show_occupied else available_classrooms
                    
                    # 将教室号按字母数字排序后写入单元格
                    sorted_classrooms_raw = sorted(list(classrooms_to_show), key=lambda x: [int(t) if t.isdigit() else t.lower() for t in re.split('([0-9]+)', x)])
                    
                    # 应用文本替换规则
                    modified_classrooms = []
                    for name in sorted_classrooms_raw:
                        # 1. 去掉 '座'
                        modified_name = name.replace('座', '')
                        # 2. 在 '世'[ABCD] 和 数字 之间加空格
                        modified_name = re.sub(r'(世[ABCD])(\d+)', r'\1 \2', modified_name)
                        modified_classrooms.append(modified_name)
                        
                    cell_value = " ".join(modified_classrooms)
                    # Calculate correct row/col: +1 for 1-based index, +1 for title row, +1 for day header col
                    target_row = period_idx + 1 + 1 # +1 for 1-based, +1 for title row
                    target_col = day_idx + 1          # +1 for 1-based index
                    data_cell = sheet.cell(row=target_row, column=target_col, value=cell_value)
                    data_cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left') # 左上对齐，自动换行
                    data_cell.font = maple_font # 应用 Maple 字体
            
            # 调整列宽 (Column widths adjusted for content)
            sheet.column_dimensions[get_column_letter(1)].width = 15 # 时间段列 (A)
            for col_letter_idx in range(data_start_col, data_end_col + 1): # 周一到周五列 (B to F)
                sheet.column_dimensions[get_column_letter(col_letter_idx)].width = 50 # 增加宽度以容纳更多教室

            # 自动调整行高 (Row height - rely on wrap_text and Excel's auto-fit)
            # No explicit row height setting for data rows needed if wrap_text is True

        # 保存文件
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        status = "occupied" if show_occupied else "empty" # Define status for filename
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        output_filename = os.path.join(output_dir, f'{basename}_{status}_classrooms_{timestamp}.xlsx')
        
        print(f"\n正在保存 {basename} 的结果到 {output_filename}...")
        output_workbook.save(output_filename)
        print(f"{basename} 处理完成！")
        
    except FileNotFoundError:
        print(f"错误：文件 '{xlsx_filepath}' 未找到。")
    except Exception as e:
        print(f"处理 {basename} 的Excel文件时出错：{str(e)}")
    finally:
        if workbook: # 确保 workbook 已被赋值
            workbook.close() # 关闭工作簿释放资源

if __name__ == "__main__":
    source_directory = "source" # 定义源目录
    output_directory = "." # 定义输出目录 (当前目录)

    while True:
        print("\n请选择要生成的教室表类型：")
        print("1. 空闲教室表")
        print("2. 占用教室表")
        print("0. 退出程序")
        
        try:
            choice = int(input("请输入选项："))
            
            if choice == 0:
                break
            elif choice in [1, 2]:
                show_occupied = (choice == 2)
                status_text = "占用" if show_occupied else "空闲"
                
                # 查找可用的文件对
                available_pairs = find_file_pairs(source_directory)
                if not available_pairs:
                    continue # 如果没有找到文件对，则返回主菜单
                    
                print(f"\n请选择要处理的文件 ({status_text}表):")
                print("0. 处理所有找到的文件对")
                # 打印可用的文件对选项已在 find_file_pairs 中完成
                
                try:
                    file_choice = int(input("请输入文件选项 (输入数字): "))
                    
                    if file_choice == 0:
                        # 处理所有文件
                        print("\n将处理所有找到的文件对...")
                        for basename in available_pairs.values():
                            process_excel(basename, show_occupied=show_occupied, source_dir=source_directory, output_dir=output_directory)
                        print("\n所有文件处理完成。")
                    elif file_choice in available_pairs:
                        # 处理选定的文件
                        selected_basename = available_pairs[file_choice]
                        process_excel(selected_basename, show_occupied=show_occupied, source_dir=source_directory, output_dir=output_directory)
                    else:
                        print("无效的文件选项，请重新选择。")
                        
                except ValueError:
                    print("错误：请输入有效的数字作为文件选项。")
                except Exception as e:
                    print(f"处理文件选择时发生错误: {str(e)}")

            else:
                print("无效的操作选项，请重新输入。")
                
        except ValueError:
            print("错误：请输入有效的数字作为操作选项。")
        except Exception as e:
             print(f"主循环发生错误: {str(e)}")

    print("\n程序已退出。") 