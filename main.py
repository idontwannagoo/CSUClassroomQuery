import openpyxl
import re
import json
from itertools import islice
from tqdm import tqdm
import time

def parse_classroom_info(text):
    if not isinstance(text, str):
        return None
        
    # 匹配周数信息和教室信息的正则表达式
    week_pattern = r'(\d+)-(\d+)周'
    classroom_pattern = r'[A-D]座\d+'
    
    # 查找所有符合模式的教室
    classroom = re.search(classroom_pattern, text)
    if not classroom:
        return None
        
    # 查找周数信息
    week_info = re.search(week_pattern, text)
    if not week_info:
        return None
        
    return {
        'start_week': int(week_info.group(1)),
        'end_week': int(week_info.group(2)),
        'classroom': classroom.group()
    }

def read_excel_cells(filename, base_row, col):
    try:
        print("\n正在读取Excel文件...")
        # 使用只读模式打开Excel文件，并禁用样式加载
        workbook = openpyxl.load_workbook(
            filename, 
            read_only=True, 
            data_only=True,
            keep_links=False
        )
        sheet = workbook.active
        
        results = []
        classroom_infos = []
        
        # 使用迭代器获取指定列的所有单元格
        column_cells = list(sheet.iter_rows(
            min_row=base_row,
            max_row=None,  # 读取到最后一行
            min_col=col,
            max_col=col,
            values_only=True
        ))
        
        # 创建进度条
        total_cells = len(column_cells)
        with tqdm(total=total_cells//10 + 1, desc="处理数据", ncols=100) as pbar:
            # 每隔10行取一个值
            for cell in islice(column_cells, 0, None, 10):
                cell_value = cell[0]  # 因为只取了一列，所以是第一个元素
                
                if cell_value:
                    if isinstance(cell_value, str):
                        cell_value = cell_value.replace('_x000D_', '')
                        # 处理多行课程信息
                        for course in cell_value.split('\n'):
                            info = parse_classroom_info(course)
                            if info:
                                results.append(
                                    f"第{info['start_week']}-{info['end_week']}周 {info['classroom']}"
                                )
                                classroom_infos.append(info)
                
                pbar.update(1)
                time.sleep(0.01)  # 添加小延迟使进度条更容易观察
        
        workbook.close()
        return results, classroom_infos if results else (["没有找到符合条件的教室信息"], [])
    
    except FileNotFoundError:
        return [f"错误：找不到文件 '{filename}'"], []
    except Exception as e:
        return [f"错误：{str(e)}"], []

def get_classrooms_by_week(classroom_infos, week):
    """获取指定周的所有教室（去重）"""
    return sorted(set(
        info['classroom'] 
        for info in classroom_infos 
        if info['start_week'] <= week <= info['end_week']
    ))

def load_classroom_info():
    """加载教室信息"""
    try:
        print("正在加载教室配置...")
        with open('classrooms.json', 'r', encoding='utf-8') as f:
            data = json.load(f)
            # 添加加载进度条
            result = {}
            for room in tqdm(data, desc="加载教室信息", ncols=100):
                result[room['room_number']] = room
                time.sleep(0.01)  # 添加小延迟使进度条更容易观察
            return result
    except FileNotFoundError:
        print("警告：找不到classrooms.json文件")
        return {}
    except json.JSONDecodeError:
        print("警告：classrooms.json文件格式错误")
        return {}

def get_available_classrooms(used_classrooms, all_classrooms):
    """获取可用教室列表"""
    available = []
    for room_number, room_info in all_classrooms.items():
        if room_number not in used_classrooms:
            available.append({
                'room_number': room_number,
                'room_type': room_info.get('room_type', '未知类型'),
                'capacity': room_info.get('capacity', '未知容量')
            })
    return sorted(available, key=lambda x: x['room_number'])

def get_row_col_by_time(day, period):
    """
    将周几和节次转换为对应的行列数
    day: 1-5 (周一到周五)
    period: 1-5 (1: 1-2节, 2: 3-4节, 3: 5-6节, 4: 7-8节, 5: 9-10节)
    返回: (row, col)
    """
    # 基准行数计算：每个时间段对应的基准行
    period_to_row = {
        1: 4,   # 1-2节
        2: 5,   # 3-4节
        3: 6,   # 5-6节
        4: 7,   # 7-8节
        5: 8    # 9-10节
    }
    
    # 列数计算：周一从第2列开始
    col = day + 1
    
    # 获取对应的行数
    row = period_to_row.get(period)
    
    if not row:
        raise ValueError("无效的节次")
    
    return row, col

def format_time_period(period):
    """将节次数转换为具体的节次范围"""
    period_map = {
        1: "1-2",
        2: "3-4",
        3: "5-6",
        4: "7-8",
        5: "9-10"
    }
    return period_map.get(period, "未知")

def main():
    print("Excel课程教室信息读取程序")
    print("请确保'input.xlsx'文件和classrooms.json在当前目录下")
    
    # 加载所有教室信息
    all_classrooms = load_classroom_info()
    
    while True:
        try:
            print("\n请输入查询信息（输入0退出）：")
            day = int(input("请输入星期（1-5，周一到周五）："))
            if day == 0:
                break
                
            if day < 1 or day > 5:
                print("错误：星期必须在1-5之间")
                continue
            
            period = int(input("请输入节次（1:1-2节, 2:3-4节, 3:5-6节, 4:7-8节, 5:9-10节）："))
            if period < 1 or period > 5:
                print("错误：节次必须在1-5之间")
                continue
            
            try:
                row, col = get_row_col_by_time(day, period)
                print(f"\n查询周{day} 第{format_time_period(period)}节的课程 (行{row}, 列{col})")
            except ValueError as e:
                print(f"错误：{str(e)}")
                continue
            
            results, classroom_infos = read_excel_cells('input1.xlsx', row, col)
            print("\n所有查询结果：")
            for result in results:
                print(result)
            
            if classroom_infos and all_classrooms:  # 确保有教室信息和配置文件
                while True:
                    try:
                        current_week = int(input("\n请输入当前周数（1-16，输入0返回）："))
                        if current_week == 0:
                            break
                        if current_week < 1 or current_week > 16:
                            print("错误：周数必须在1-16之间")
                            continue
                            
                        print(f"\n第{current_week}周 周{day} 第{format_time_period(period)}节的教室使用情况：")
                        used_classrooms = get_classrooms_by_week(classroom_infos, current_week)
                        if used_classrooms:
                            print(f"\n已使用教室（共{len(used_classrooms)}个）：")
                            for classroom in used_classrooms:
                                print(classroom)
                            
                            # 获取并显示可用教室
                            available = get_available_classrooms(used_classrooms, all_classrooms)
                            if available:
                                print(f"\n可用教室（共{len(available)}个）：")
                                print("教室号\t\t类型\t\t容量")
                                print("-" * 50)
                                for room in available:
                                    print(f"{room['room_number']:<10}\t{room['room_type']:<10}\t{room['capacity']}人")
                            else:
                                print("\n没有可用教室")
                        else:
                            print("本时段没有课程，所有教室都可用：")
                            print("教室号\t\t类型\t\t容量")
                            print("-" * 50)
                            for room_number, info in all_classrooms.items():
                                print(f"{room_number:<10}\t{info['room_type']:<10}\t{info['capacity']}人")
                            
                    except ValueError:
                        print("错误：请输入有效的数字")
            
        except ValueError:
            print("错误：请输入有效的数字")
        
        print("\n")

if __name__ == "__main__":
    main()
