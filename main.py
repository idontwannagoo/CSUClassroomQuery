import openpyxl
import re

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

def is_classroom_available_this_week(info, current_week):
    """检查教室在当前周是否有课"""
    return info['start_week'] <= current_week <= info['end_week']

def read_excel_cells(filename, base_row, col):
    try:
        # 使用只读模式打开Excel文件
        workbook = openpyxl.load_workbook(filename, read_only=True, data_only=True)
        sheet = workbook.active
        
        # 获取最大行数
        max_row = sheet.max_row
        results = []
        classroom_infos = []  # 存储完整的教室信息
        
        # 遍历所有可能的行号（从base_row开始，每次加10）
        current_row = base_row
        while current_row <= max_row:
            cell_value = sheet.cell(row=current_row, column=col).value
            
            # 如果单元格有值，则添加到结果中
            if cell_value is not None:
                # 清理单元格值中的_x000D_
                if isinstance(cell_value, str):
                    cell_value = cell_value.replace('_x000D_', '')
                    # 处理可能包含多个课程信息的情况
                    for course in cell_value.split('\n'):
                        info = parse_classroom_info(course)
                        if info:
                            results.append(
                                f"第{info['start_week']}-{info['end_week']}周 {info['classroom']}"
                            )
                            classroom_infos.append(info)
            
            current_row += 10
        
        # 关闭工作簿
        workbook.close()    
        return results, classroom_infos if results else (["没有找到符合条件的教室信息"], [])
    
    except FileNotFoundError:
        return [f"错误：找不到文件 '{filename}'"], []
    except Exception as e:
        return [f"错误：{str(e)}"], []

def main():
    print("Excel课程教室信息读取程序")
    print("请确保'input.xlsx'文件在当前目录下")
    
    while True:
        try:
            # 获取用户输入
            base_row = int(input("请输入基准行号（1-10，输入0退出）："))
            if base_row == 0:
                break
            
            if base_row < 1 or base_row > 10:
                print("错误：行号必须在1-10之间")
                continue
                
            col = int(input("请输入列号："))
            
            # 读取符合模式的所有单元格值
            results, classroom_infos = read_excel_cells('input1.xlsx', base_row, col)
            print("\n所有查询结果：")
            for result in results:
                print(result)
            
            if classroom_infos:  # 如果有找到教室信息
                while True:
                    try:
                        current_week = int(input("\n请输入当前周数（1-16，输入0返回）："))
                        if current_week == 0:
                            break
                        if current_week < 1 or current_week > 16:
                            print("错误：周数必须在1-16之间")
                            continue
                            
                        print(f"\n第{current_week}周的课程教室：")
                        found = False
                        for info in classroom_infos:
                            if is_classroom_available_this_week(info, current_week):
                                print(f"{info['classroom']}")
                                found = True
                        
                        if not found:
                            print("本周没有课程")
                            
                    except ValueError:
                        print("错误：请输入有效的数字")
            
        except ValueError:
            print("错误：请输入有效的数字")
        
        print("\n")

if __name__ == "__main__":
    main()
