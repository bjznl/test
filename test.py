import os
import sys
from docx import Document

# 检查并创建输出文件夹
output_dir = "output"
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# 打开输出文件
output_file_path = os.path.join(output_dir, "output.txt")
try:
    with open(output_file_path, "w") as f:
        f.write("是否吸烟,诊断\n")
except IOError:
    print("无法打开输出文件")
    sys.exit(1)

# 遍历当前目录中的Word文件
for file_name in os.listdir("."):
    if not file_name.endswith(".docx"):
        continue

    # 打开Word文件
    try:
        document = Document(file_name)
    except:
        print(f"无法打开文件 {file_name}")
        continue

    # 遍历Word文件中的所有表格
    for table in document.tables:
        # 查找需要的信息
        cy = zd = None
        for row in table.rows:
            for i, cell in enumerate(row.cells):
                text = cell.text.strip()
                if text == "是否吸烟":
                    cy = row.cells[i+1].text.strip()
                elif text == "诊断":
                    zd = row.cells[i+1].text.strip()

        # 输出结果到文件
        if cy and zd:
            try:
                with open(output_file_path, "a") as f:
                    f.write(f"{cy},{zd}\n")
            except IOError:
                print("无法写入输出文件")
                sys.exit(1)

    #关闭Word文件
    try:
        document.close()
    except:
        print(f"无法关闭文件 {file_name}")
        continue

# 输出完成信息
print("输出完成")