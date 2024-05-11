import os
from datetime import datetime
import shutil
import pandas as pd
# 读取Excel文件
file_path = r'D:\Personal\Desktop\世界时.xls'
df = pd.read_excel(file_path)

# 1.假设日期时间列名为 '世界时'
datetime_column = '世界时'
# 向后推2小时
df['end_time'] = pd.to_datetime(df[datetime_column]) + pd.Timedelta('2 hours')
# 保存修改后的数据为Excel文件
output_file_path = os.path.join(os.path.dirname(file_path), 'modified_data.xlsx')
df.to_excel(output_file_path, index=False)
print(f"修改后的数据已保存在：{output_file_path}")

file_path = r'D:\Personal\Desktop\modified_data.xlsx'
modified_data = pd.read_excel(file_path)

# 将日期时间数据列转换为datetime对象
modified_data['开始时间'] = pd.to_datetime(modified_data['世界时'])
modified_data['结束时间'] = pd.to_datetime(modified_data['end_time'])
# 删除 '世界时' 和 'end_time' 列
modified_data = modified_data.drop(columns=['世界时', 'end_time'])
# 保存修改后的数据回原Excel文件
modified_data.to_excel(file_path, index=False)

print("数据已转换为datetime对象并保存回原Excel文件：", file_path)


# 指定目录路径
directory = r'F:\陈卓6-10月fmt'
output_directory = r'D:\Personal\Desktop\导出文件夹'
file_path1 = r'D:\Personal\Desktop\modified_data.xlsx'

# 读取 Excel 文件
modified_data = pd.read_excel(file_path1)

# 创建新的文件夹用于存放符合条件的文件
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# 指定要处理的压缩格式列表
valid_extensions = [".rar", ".zip", ".7z"]  # 添加您希望处理的压缩格式

# 遍历指定目录下的文件
for filename in os.listdir(directory):
    file_extension = os.path.splitext(filename)[1]
    if file_extension in valid_extensions:
        # 提取第四个和第五个下划线之间的文本作为时间
        parts = filename.split('_')
        if len(parts) >= 5:
            time_text = parts[4]
            try:
                time_obj = datetime.strptime(time_text, '%Y%m%d%H%M%S')
                print(f"文件名: {filename}, 提取的时间文本: {time_text}, 转换后的时间: {time_obj}")

                # 遍历每行时间范围
                for index, row in modified_data.iterrows():
                    start_time = row['开始时间']
                    end_time = row['结束时间']

                    # 判断文件名的时间是否在时间范围内
                    if start_time <= time_obj <= end_time:
                        # 如果时间在范围内，则复制文件到新的文件夹中
                        shutil.copy(os.path.join(directory, filename), output_directory)
                        print(f"文件 {filename} 已导出到 {output_directory}")
                        break  # 找到匹配的时间范围后退出循环

            except ValueError:
                pass