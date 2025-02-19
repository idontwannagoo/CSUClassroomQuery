import openpyxl
import re
from itertools import islice

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
        column_cells = sheet.iter_rows(
            min_row=base_row,
            max_row=None,  # 读取到最后一行
            min_col=col,
            max_col=col,
            values_only=True
        )
        
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

def main():
    print("Excel课程教室信息读取程序")
    print("请确保'input.xlsx'文件在当前目录下")
    
    while True:
        try:
            base_row = int(input("请输入基准行号（1-10，输入0退出）："))
            if base_row == 0:
                break
            
            if base_row < 1 or base_row > 10:
                print("错误：行号必须在1-10之间")
                continue
                
            col = int(input("请输入列号："))
            
            results, classroom_infos = read_excel_cells('input1.xlsx', base_row, col)
            print("\n所有查询结果：")
            for result in results:
                print(result)
            
            if classroom_infos:
                while True:
                    try:
                        current_week = int(input("\n请输入当前周数（1-16，输入0返回）："))
                        if current_week == 0:
                            break
                        if current_week < 1 or current_week > 16:
                            print("错误：周数必须在1-16之间")
                            continue
                            
                        print(f"\n第{current_week}周的课程教室：")
                        classrooms = get_classrooms_by_week(classroom_infos, current_week)
                        if classrooms:
                            print(f"共{len(classrooms)}个教室：")
                            for classroom in classrooms:
                                print(classroom)
                        else:
                            print("本周没有课程")
                            
                    except ValueError:
                        print("错误：请输入有效的数字")
            
        except ValueError:
            print("错误：请输入有效的数字")
        
        print("\n")

if __name__ == "__main__":
    main()
