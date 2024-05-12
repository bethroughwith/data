import os
from datetime import datetime
import shutil
import pandas as pd
import pytz

# 读取Excel文件
file_path = r'D:\Personal\Desktop\2021-2024冰雹时间.xls'
df = pd.read_excel(file_path,sheet_name=1)

# 假设开始时间和结束时间列名分别为 '开始时间' 和 '结束时间'
df['开始时间'] = pd.to_datetime(df['开始时间']).dt.tz_localize('Asia/Shanghai').dt.tz_convert(pytz.utc)
df['结束时间'] = pd.to_datetime(df['结束时间']).dt.tz_localize('Asia/Shanghai').dt.tz_convert(pytz.utc)
print(df)
# 指定目录路径
directory = r'D:\陈卓'
output_directory = r'D:\Personal\Desktop\导出文件夹'

# 创建新的文件夹用于存放符合条件的文件
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# 指定要处理的压缩格式列表
valid_extensions = [".rar", ".zip", ".7z", ".bz2"]  # 添加您希望处理的压缩格式

# 遍历指定目录及其所有子目录下的文件
for root, dirs, files in os.walk(directory):
    for filename in files:
        parts = filename.split('_')
        if len(parts) >= 5:
            time_text = parts[4].split('.')[0]  # 假设时间信息在文件名中的第四个和第五个下划线之间的文本作为时间
            try:
                time_obj = datetime.strptime(time_text, '%Y%m%d%H%M%S').replace(tzinfo=None)  # 将时间对象转换为没有时区信息的对象
                print(f"文件名: {filename}, 提取的时间文本: {time_text}, 转换后的时间: {time_obj}")
                # 遍历每行时间范围
                for index, row in df.iterrows():
                    start_time = row['开始时间']
                    end_time = row['结束时间']
                    # 为time_obj添加时区信息（假设时区为UTC）
                    time_obj = time_obj.replace(tzinfo=pytz.UTC)
                    # 判断文件名的时间是否在时间范围内
                    if start_time <= time_obj <= end_time:
                        # 如果时间在范围内，则复制文件到新的文件夹中
                        shutil.copy(os.path.join(root, filename), output_directory)
                        print(f"文件 {filename} 已导出到 {output_directory}")
                        break  # 找到匹配的时间范围后退出循环

            except ValueError:
                pass
