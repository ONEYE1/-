import os
import sys
import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox
import configparser


def sanitize_folder_name(folder_name):
    """清理文件夹名称，移除非法字符并替换空格"""
    # 定义非法字符集合
    illegal_chars = set(r'<>:"/\|?*')
    
    # 确保 folder_name 是字符串类型
    folder_name = str(folder_name).strip() if folder_name is not None else ""
    
    # 去除非法字符
    folder_name = ''.join(char for char in folder_name if char not in illegal_chars)
    
    # 去除空格和制表符
    folder_name = folder_name.replace(" ", "_").replace("\t", "_")
    
    return folder_name


def create_folders(base_path, folder_names):
    """批量创建文件夹"""
    for folder_name in folder_names:
        # 清理文件夹名称
        folder_name = sanitize_folder_name(folder_name)
        
        # 校验文件夹名称
        if len(folder_name) > 255:
            print(f"错误: 文件夹名称 '{folder_name}' 超过255个字符限制")
            continue
        
        if not folder_name:
            print("错误: 文件夹名称在去除非法字符后为空")
            continue
        
        folder_path = os.path.join(base_path, folder_name)
        if not os.path.exists(folder_path):
            try:
                os.makedirs(folder_path)
                print(f"已创建文件夹: {folder_path}")
            except OSError as e:
                print(f"错误: 无法创建文件夹 '{folder_path}'。详情: {e}")
        else:
            print(f"文件夹已存在: {folder_path}")


def read_txt_file(file_path):
    """读取 .txt 文件并返回文件夹名称列表"""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            return [line.strip() for line in file if line.strip()]
    except Exception as e:
        print(f"错误: 读取文件时发生异常: {e}")
        sys.exit(1)


def read_xlsx_file(file_path):
    """读取 .xlsx 文件并返回文件夹名称列表"""
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        return [cell.value for row in sheet.iter_rows(min_row=1, max_col=1) for cell in row if cell.value]
    except Exception as e:
        print(f"错误: 读取文件时发生异常: {e}")
        sys.exit(1)


class FolderCreatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("批量创建文件夹")
        
        # 创建配置文件对象
        self.config = configparser.ConfigParser()
        self.config_file = 'config.ini'
        
        # 读取配置文件（如果存在）
        if os.path.exists(self.config_file):
            try:
                self.config.read(self.config_file, encoding='utf-8')
            except Exception as e:
                print(f"警告: 无法读取配置文件: {e}，将使用默认设置")
        
        # 获取上次使用的基础路径，默认为空
        self.last_base_path = self.config.get('DEFAULT', 'last_base_path', fallback='')
        
        # 获取上次使用的选择文件路径，默认为空
        self.last_selected_file = self.config.get('DEFAULT', 'last_selected_file', fallback='')
        
        # 创建界面元素
        self.create_widgets()
        
        # 设置默认基础路径
        self.base_path_entry.insert(0, self.last_base_path)
        
        # 设置上次选择的文件路径
        self.file_entry.insert(0, self.last_selected_file)

    def create_widgets(self):
        # 文件路径选择
        tk.Label(self.root, text="选择文件:").grid(row=0, column=0, padx=10, pady=10)
        self.file_entry = tk.Entry(self.root, width=50)
        self.file_entry.grid(row=0, column=1, padx=10, pady=10)
        tk.Button(self.root, text="浏览", command=self.browse_file).grid(row=0, column=2, padx=10, pady=10)

        # 基础路径选择
        tk.Label(self.root, text="基础路径:").grid(row=1, column=0, padx=10, pady=10)
        self.base_path_entry = tk.Entry(self.root, width=50)
        self.base_path_entry.grid(row=1, column=1, padx=10, pady=10)
        tk.Button(self.root, text="浏览", command=self.browse_base_path).grid(row=1, column=2, padx=10, pady=10)

        # 创建文件夹按钮
        tk.Button(self.root, text="创建文件夹", command=self.create_folders_gui).grid(row=2, column=1, padx=10, pady=10)

    def browse_file(self):
        """浏览并选择文件"""
        file_path = filedialog.askopenfilename(filetypes=[("文本文件", "*.txt"), ("Excel 文件", "*.xlsx")])
        self.file_entry.delete(0, tk.END)
        self.file_entry.insert(0, file_path)

    def browse_base_path(self):
        """浏览并选择基础路径"""
        base_path = filedialog.askdirectory()
        self.base_path_entry.delete(0, tk.END)
        self.base_path_entry.insert(0, base_path)

    def create_folders_gui(self):
        """创建文件夹的GUI处理函数"""
        base_path = self.base_path_entry.get()
        folder_names_file = self.file_entry.get()

        if not os.path.exists(folder_names_file):
            messagebox.showerror("错误", f"文件 '{folder_names_file}' 不存在。")
            return

        file_extension = os.path.splitext(folder_names_file)[1].lower()
        if file_extension == '.txt':
            folder_names = read_txt_file(folder_names_file)
        elif file_extension == '.xlsx':
            folder_names = read_xlsx_file(folder_names_file)
        else:
            messagebox.showerror("错误", f"不支持的文件格式 '{file_extension}'。支持的格式为 .txt 和 .xlsx。")
            return

        create_folders(base_path, folder_names)
        messagebox.showinfo("成功", "文件夹创建成功！")

        # 更新配置文件中的基础路径和选择的文件路径
        if 'DEFAULT' not in self.config:
            self.config['DEFAULT'] = {}
        self.config['DEFAULT']['last_base_path'] = base_path
        self.config['DEFAULT']['last_selected_file'] = folder_names_file
        try:
            with open(self.config_file, 'w', encoding='utf-8') as configfile:
                self.config.write(configfile)
        except Exception as e:
            print(f"警告: 无法保存配置文件: {e}")


def get_last_selected_file_path():
    """
    获取上次使用的选择文件路径
    
    Returns:
        str: 上次选择的文件路径，如果不存在则返回空字符串
    """
    config = configparser.ConfigParser()
    config_file = 'config.ini'
    
    # 读取配置文件（如果存在）
    if os.path.exists(config_file):
        try:
            config.read(config_file, encoding='utf-8')
            # 获取上次使用的选择文件路径，默认为空
            last_selected_file = config.get('DEFAULT', 'last_selected_file', fallback='')
            return last_selected_file
        except Exception as e:
            print(f"警告: 无法读取配置文件: {e}")
            return ""
    
    return ""


def main():
    root = tk.Tk()
    app = FolderCreatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()