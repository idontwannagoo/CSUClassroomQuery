import openpyxl
import csv
import re
from itertools import islice
from tqdm import tqdm
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

def load_classrooms():
    """从CSV文件加载教室列表"""
    classrooms = set()
    try:
        with open('source/TianXin.csv', 'r', encoding='gbk') as f:
            reader = csv.reader(f)
            for row in reader:
                if row:  # 确保行不为空
                    classrooms.add(row[0])
        return classrooms
    except Exception as e:
        print(f"加载教室列表时出错：{str(e)}")
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

def process_excel(show_occupied=False):
    """处理Excel文件并生成教室表
    
    Args:
        show_occupied (bool): True显示占用教室，False显示空闲教室
    """
    print("开始处理课程表...")
    
    # 加载所有教室
    all_classrooms = load_classrooms()
    if not all_classrooms:
        print("错误：无法加载教室列表")
        return
    
    # 创建用于存储课程信息的数据结构
    # 格式：{(周数, 星期几, 节次): set(已用教室)}
    used_classrooms = {}
    
    # 加载Excel文件
    try:
        workbook = openpyxl.load_workbook('source/TianXin.xlsx', read_only=True, data_only=True)
        sheet = workbook.active
        
        # 获取总行数
        total_rows = sum(1 for _ in sheet.rows)
        
        print("正在读取课程信息...")
        valid_count = 0  # 用于计数有效信息
        print("\n前10个有效课程信息：")
        print("-" * 50)
        
        with tqdm(total=total_rows, desc="处理数据", ncols=100) as pbar:
            for row in sheet.iter_rows(min_row=4):  # 从第4行开始读取
                try:
                    time_str = str(row[6].value) if row[6].value else None  # 第7列：时间
                    classroom = str(row[7].value) if row[7].value else None  # 第8列：教室
                    week_range = str(row[8].value) if row[8].value else None  # 第9列：周次
                    single_double = str(row[9].value) if row[9].value else None  # 第10列：单双周
                    
                    if time_str and classroom:
                        # 解析时间
                        time_info = parse_time_periods(time_str)
                        if time_info:
                            # 输出前10个有效信息
                            if valid_count < 10:
                                day, periods = time_info
                                print(f"第{valid_count + 1}条信息：")
                                print(f"时间：{time_str}（星期{day} 第{periods}节）")
                                print(f"教室：{classroom}")
                                print(f"周次：{week_range}")
                                print(f"单双周：{single_double}")
                                print("-" * 50)
                            valid_count += 1
                            
                            day, periods = time_info
                            
                            # 解析周次
                            weeks = parse_weeks(week_range, single_double)
                            
                            # 更新已用教室信息
                            for week in weeks:
                                for period in periods:
                                    key = (week, day, period)
                                    if key not in used_classrooms:
                                        used_classrooms[key] = set()
                                    used_classrooms[key].add(classroom)
                
                except Exception as e:
                    print(f"处理行时出错：{str(e)}")
                
                pbar.update(1)
        
        # 创建输出Excel
        print("\n正在生成{}教室表...".format("占用" if show_occupied else "空闲"))
        output_workbook = openpyxl.Workbook()
        
        # 为每周创建工作表
        for week in range(1, 17):
            # 创建该周的工作表
            sheet_name = f"第{week}周"
            if week == 1:
                sheet = output_workbook.active
                sheet.title = sheet_name
            else:
                sheet = output_workbook.create_sheet(sheet_name)
            
            # 设置表头
            days = ["", "周一", "周二", "周三", "周四", "周五"]
            for col, day in enumerate(days, 1):
                sheet.cell(row=1, column=col, value=day)
            
            periods = ["", "1-2节", "3-4节", "5-6节", "7-8节", "9-10节"]
            for row, period in enumerate(periods, 1):
                sheet.cell(row=row, column=1, value=period)
            
            # 填充数据
            for day in range(1, 6):
                for period in range(1, 6):
                    key = (week, day, period)
                    used = used_classrooms.get(key, set())
                    # 根据show_occupied选项决定显示哪些教室
                    classrooms_to_show = used if show_occupied else all_classrooms - used
                    
                    # 将教室号写入单元格
                    cell_value = " ".join(sorted(classrooms_to_show))
                    sheet.cell(row=period+1, column=day+1, value=cell_value)
            
            # 调整列宽和行高
            for col in range(1, 7):
                sheet.column_dimensions[get_column_letter(col)].width = 40
            for row in range(1, 7):
                sheet.row_dimensions[row].height = 40
            
            # 设置单元格对齐方式
            for row in sheet.iter_rows(min_row=1, max_row=6, min_col=1, max_col=6):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
        
        # 保存文件
        output_filename = 'TianXin_{}_classrooms.xlsx'.format("occupied" if show_occupied else "empty")
        print(f"\n正在保存结果到 {output_filename}...")
        output_workbook.save(output_filename)
        print("导出完成！")
        
    except Exception as e:
        print(f"处理Excel文件时出错：{str(e)}")
    finally:
        workbook.close()

if __name__ == "__main__":
    while True:
        print("\n请选择要生成的教室表类型：")
        print("1. 空闲教室表")
        print("2. 占用教室表")
        print("0. 退出程序")
        
        try:
            choice = int(input("请输入选项："))
            if choice == 0:
                break
            elif choice == 1:
                process_excel(show_occupied=False)
            elif choice == 2:
                process_excel(show_occupied=True)
            else:
                print("无效的选项，请重新输入")
        except ValueError:
            print("错误：请输入有效的数字") 