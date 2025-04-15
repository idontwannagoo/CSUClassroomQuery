import os
import pandas as pd

def process_excel_files():
    # 获取source目录下的所有xlsx文件
    source_dir = 'source'
    if not os.path.exists(source_dir):
        os.makedirs(source_dir)
        print(f"创建了{source_dir}目录")
        return

    # 遍历所有xlsx文件
    for filename in os.listdir(source_dir):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(source_dir, filename)
            
            try:
                # 读取Excel文件
                df = pd.read_excel(file_path)
                
                # 获取第八列数据（索引为7）
                if len(df.columns) >= 8:
                    classroom_data = df.iloc[2:, 7] # 从第三行开始选取第八列数据
                    
                    # 去重
                    unique_classrooms = pd.Series(classroom_data.unique())
                    
                    # 创建输出文件名（将.xlsx替换为.csv）
                    output_filename = filename.replace('.xlsx', '.csv')
                    output_path = os.path.join(source_dir, output_filename)
                    
                    # 保存到CSV文件
                    unique_classrooms.to_csv(output_path, index=False, header=False)
                    print(f"已处理文件: {filename}")
                else:
                    print(f"警告: {filename} 没有第八列数据")
                    
            except Exception as e:
                print(f"处理文件 {filename} 时出错: {str(e)}")

if __name__ == "__main__":
    process_excel_files() 