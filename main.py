import openpyxl

def read_excel_cell(filename, row, col):
    try:
        # 打开Excel文件
        workbook = openpyxl.load_workbook(filename)
        
        # 获取第一个工作表
        sheet = workbook.active
        
        # 获取指定单元格的值
        cell_value = sheet.cell(row=row, column=col).value
        
        # 清理单元格值中的_x000D_
        if isinstance(cell_value, str):
            cell_value = cell_value.replace('_x000D_', '')
        
        return cell_value
    
    except FileNotFoundError:
        return "错误：找不到文件 'input.xlsx'"
    except Exception as e:
        return f"错误：{str(e)}"

def main():
    print("Excel单元格读取程序")
    print("请确保'input.xlsx'文件在当前目录下")
    
    while True:
        try:
            # 获取用户输入
            row = int(input("请输入行号（输入0退出）："))
            if row == 0:
                break
                
            col = int(input("请输入列号："))
            
            # 读取单元格值
            result = read_excel_cell('input.xlsx', row, col)
            print(f"单元格内容：{result}")
            
        except ValueError:
            print("错误：请输入有效的数字")
        
        print("\n")

if __name__ == "__main__":
    main()
