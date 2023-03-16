import os
import re
import sys
import docx
import openpyxl
from tkinter import filedialog
from tkinter import messagebox
from tkinter import *

# 提取信息函数
# 提取信息函数
def extract_information(file_path):
    return personal_information
      # 遍历Word文件中的所有表格
for table in document.tables:
        # 查找需要的信息
        name = sex = id = jg = mz = blood = jb = whcd = cy = zd = None
        for row in table.rows:
            for i, cell in enumerate(row.cells):
                text = cell.text.strip()
                if text == "姓名":
                    name = row.cells[i+1].text.strip()
                elif text == "性别":
                    sex = row.cells[i+1].text.strip()
                elif text == "身份证号":
                    id = row.cells[i+1].text.strip()
                elif text == "籍贯":
                    jg = row.cells[i+1].text.strip()
                elif text == "民族":
                    mz = row.cells[i+1].text.strip()
                elif text == "血型":
                    blood = row.cells[i+1].text.strip()
                elif text == "护理级别":
                    jb = row.cells[i+1].text.strip()
                elif text == "文化程度":
                    whcd = row.cells[i+1].text.strip()
                elif text == "是否吸烟":
                    cy = row.cells[i+1].text.strip()
                elif text == "诊断":
                    zd = row.cells[i+1].text.strip()
                    zd = re.sub(r'\s+', ' ', zd)
                    

        # 输出结果到文件
        if name or zd:
            ws.append([name, sex, id, jg, mz, blood, jb, whcd, cy, zd])

    #关闭Word文件
try:
    document.close()
except:
        #print(f"无法关闭文件 {file_name}")
continue

# 写入 Excel
def write_to_excel(data, output_path):
    for row in ws.iter_rows(min_row=3, min_col=10, max_col=10):
        if row[0].value:  # 如果I3单元格不为空
        # 将I3单元格的值复制到I2单元格
            ws.cell(row=row[0].row - 1, column=10).value = row[0].value
    # 删除第三行
        ws.delete_rows(row[0].row)
    wb.save(output_file_path)

def main(input_directory, output_directory):
    # 遍历文件夹中的所有文件
    for file in os.listdir(input_directory):
        # 如果文件是 Word 文档
        if file.endswith(".docx"):
            file_path = os.path.join(input_directory, file)
            try:
                personal_information = extract_information(file_path)
                write_to_excel(personal_information, output_directory)
                messagebox.showinfo("成功", f"处理文件 {file} 成功！")
            except Exception as e:
                messagebox.showerror("错误", f"处理文件 {file} 时出错: {str(e)}")

# 创建一个简单的 Tkinter GUI
def create_gui():
    # 初始化 tkinter
    root = Tk()
    root.title("江福一体机健康档案提取转换工具")

    input_directory = StringVar()
    output_directory = StringVar()

    def choose_input_directory():
        dir_path = filedialog.askdirectory()
        if dir_path:
            input_directory.set(dir_path)

    def choose_output_directory():
        dir_path = filedialog.askdirectory()
        if dir_path:
            output_directory.set(dir_path)

    def start():
        if not input_directory.get() or not output_directory.get():
            messagebox.showerror("错误", "请先选择输入和输出目录！")
            return
        main(input_directory.get(), output_directory.get())

    Button(root, text="说明", command=lambda: messagebox.showinfo("说明", "仅供江福一体机健康档案提取转换用")).grid(row=0, column=0, padx=10, pady=10)
    Button(root, text="选择档案目录", command=choose_input_directory).grid(row=1, column=0, padx=10, pady=10)
    Button(root, text="选择导出目录", command=choose_output_directory).grid(row=2, column=0, padx=10, pady=10)
    Button(root, text="启动", command=start).grid(row=3, column=0, padx=10, pady=10)

    Label(root, text="输入目录:").grid(row=1, column=1, sticky=W)
    Label(root, text="输出目录:").grid(row=2, column=1, sticky=W)
    Label(root, textvariable=input_directory).grid(row=1, column=2, sticky=W)
    Label(root, textvariable=output_directory).grid(row=2, column=2, sticky=W)

    root.mainloop()

if __name__ == "__main__":
    create_gui()


