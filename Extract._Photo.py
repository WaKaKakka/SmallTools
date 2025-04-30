
#此程序是用来提取excel表格中的图片

import os
from openpyxl import load_workbook
from PIL import Image
from io import BytesIO

# 配置路径
excel_path = ''#这里是excel表格的路径
output_folder = ''#这里是输出图片的路径
# 创建输出文件夹
os.makedirs(output_folder, exist_ok=True)

# 加载工作簿和第一个工作表
wb = load_workbook(excel_path)
ws = wb.active

# 遍历表格中的图片
for image in ws._images:
    # 获取图片所在的单元格位置，比如 "B2"
    anchor = image.anchor._from
    row = anchor.row + 1  # openpyxl从0开始
    col = anchor.col + 1

    # 假设名字在图片同一行的A列（即第1列）
    name_cell = ws.cell(row=row, column=3)#这里是名字所在的列,column=3是第3列
    name = str(name_cell.value).strip()

    if not name:
        print(f"第 {row} 行没有名字，跳过")
        continue

    # 提取图片数据并保存
    img_data = image._data()
    img = Image.open(BytesIO(img_data))

    # 构造保存路径，默认保存为 PNG
    filename = f"{name}.png"
    save_path = os.path.join(output_folder, filename)
    img.save(save_path)
    print(f"已保存图片：{save_path}")
