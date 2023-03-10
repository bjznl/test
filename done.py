import os
import sys
import re
from docx import Document

# 检查并创建输出文件夹
output_dir = "output"
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# 打开输出文件
output_file_path = os.path.join(output_dir, "output.txt")
try:
    with open(output_file_path, "w") as f:
        f.write("姓名,性别,出生日期,籍贯,身份证号,民族,婚姻情况,血型,护理级别,文化程度,是否吸烟,诊断\n")
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
        name = gender = dob = hometown = id = mz = hy = xx = hljb = whcd = cy =  zd = None
        for row in table.rows:
            for i, cell in enumerate(row.cells):
                text = cell.text.strip()
                if text == "姓名":
                    name = row.cells[i+1].text.strip()
                elif text == "性别":
                    gender = row.cells[i+1].text.strip()
                elif text == "出生日期":
                    dob = row.cells[i+1].text.strip()
                elif text == "籍贯":
                    hometown = row.cells[i+1].text.strip()
                elif text == "身份证号":
                    id = row.cells[i+1].text.strip()
                elif text == "民族":
                    mz = row.cells[i+1].text.strip()
                elif text == "婚姻情况":
                    hy = row.cells[i+1].text.strip()
                elif text == "血型":
                    xx = row.cells[i+1].text.strip()
                elif text == "护理级别":
                    hljb = row.cells[i+1].text.strip()
                elif text == "文化程度":
                    whcd = row.cells[i+1].text.strip()
                elif text == "是否吸烟":
                    cy = row.cells[i+1].text.strip()
                # elif text == "疾病":
                #     jb = row.cells[i+1].text.strip()
                elif text == "诊断":
                    zd = row.cells[i+1].text.strip()
        # 输出结果到文件
        
        if name and gender and dob and hometown and id and mz and hy and xx and hljb and whcd and cy and  None :
            try:
                with open(output_file_path, "a") as f:
                    f.write(f"{name},{gender},{dob},{hometown},{id},{mz},{hy},{xx},{hljb},{whcd},{cy},{zd}\n")
            except IOError:
                print("无法写入输出文件")
                sys.exit(1)

    # 关闭Word文件
    try:
        document.close()
    except:
        print(f"无法关闭文件 {file_name}")
        continue

# 输出完成信息
print("输出完成")
