import pandas as pd
import os
import subprocess
import sys
import re
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

# --- 配置 ---
# Excel 文件名
EXCEL_FILE = '科研课题管理总表.xlsx'
# 课题文件夹的根目录
PROJECTS_ROOT_DIR = os.path.join(os.path.abspath('.'), '科研课题管理')
# Excel 中的工作表名称
SHEET_NAME = '课题列表'
# 标准的课题子文件夹结构，包含 03_过程管理 的子文件夹
FOLDER_STRUCTURE = [
    '01_申报材料',
    '02_立项文件',
    {
        'name': '03_过程管理',
        'subfolders': ['01_开题', '02_中期', '03_变更']
    },
    '04_结题材料',
    '05_财务资料'
]
# Excel 表格的列名
EXCEL_COLUMNS = ['课题编号', '课题名称', '负责人', '课题状态', '申报日期', '立项日期', '计划开始日期', '计划结题日期',
                 '实际结题日期', '课题文件夹路径', '备注']


# --- 函数 ---

def load_projects_data():
    """从 Excel 加载课题数据"""
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, engine='openpyxl', dtype={'课题编号': str})
        print(f"成功从 '{EXCEL_FILE}' 加载 {len(df)} 条课题数据。")
        for col in EXCEL_COLUMNS:
            if col not in df.columns:
                df[col] = None
                print(f"警告: 文件中缺少列 '{col}'，已添加。")
        df = df[EXCEL_COLUMNS]
        return df
    except FileNotFoundError:
        print(f"信息: Excel 文件 '{EXCEL_FILE}' 未找到。将创建一个新的空数据表结构。")
        return pd.DataFrame(columns=EXCEL_COLUMNS)
    except Exception as e:
        print(f"加载 Excel 文件 '{EXCEL_FILE}' 时出错: {e}")
        return pd.DataFrame(columns=EXCEL_COLUMNS)


def save_projects_data(df):
    """将课题数据保存回 Excel"""
    try:
        df_to_save = df.reindex(columns=EXCEL_COLUMNS)
        df_to_save.to_excel(EXCEL_FILE, sheet_name=SHEET_NAME, index=False, engine='openpyxl')
        print(f"数据已成功保存到 '{EXCEL_FILE}'。")
        return True
    except Exception as e:
        print(f"保存数据到 Excel 文件 '{EXCEL_FILE}' 时出错: {e}")
        return False


def sanitize_foldername(name):
    """清理文件名，移除或替换不适用于文件夹名称的字符"""
    forbidden_chars = r'[\<\>\:\"\/\\\|\?\*]'
    sanitized_name = re.sub(forbidden_chars, '_', name)
    sanitized_name = sanitized_name.strip()
    return sanitized_name


def create_project_folders(project_id, project_name):
    """为新课题创建文件夹结构"""
    sanitized_name = sanitize_foldername(project_name)
    folder_name = f"{project_id}_{sanitized_name}"
    project_path = os.path.join(PROJECTS_ROOT_DIR, folder_name)
    try:
        if not os.path.exists(project_path):
            os.makedirs(project_path)
            print(f"创建课题主文件夹: {project_path}")
        else:
            print(f"信息: 文件夹 '{project_path}' 已存在，检查子文件夹。")

        created_subfolder = False
        for item in FOLDER_STRUCTURE:
            if isinstance(item, dict):
                subfolder_name = item['name']
                subfolder_path = os.path.join(project_path, subfolder_name)
                if not os.path.exists(subfolder_path):
                    os.makedirs(subfolder_path)
                    print(f"创建子文件夹: {subfolder_path}")
                    created_subfolder = True
                for sub_subfolder in item['subfolders']:
                    sub_subfolder_path = os.path.join(subfolder_path, sub_subfolder)
                    if not os.path.exists(sub_subfolder_path):
                        os.makedirs(sub_subfolder_path)
                        print(f"创建子子文件夹: {sub_subfolder_path}")
                        created_subfolder = True
            else:
                subfolder_path = os.path.join(project_path, item)
                if not os.path.exists(subfolder_path):
                    os.makedirs(subfolder_path)
                    print(f"创建子文件夹: {subfolder_path}")
                    created_subfolder = True

        if created_subfolder:
            print(f"已在 '{project_path}' 中补全标准子文件夹。")
        return project_path
    except OSError as e:
        print(f"创建文件夹 '{project_path}' 时发生 OS 错误: {e}")
        return None
    except Exception as e:
        print(f"创建文件夹时发生未知错误: {e}")
        return None


def add_new_project(df, project_id, project_name, pi_name, start_date=None, end_date=None, notes=""):
    """添加新课题到 DataFrame 并创建文件夹"""
    if not df['课题编号'].astype(str).str.fullmatch(str(project_id)).empty:
        print(f"错误: 课题编号 '{project_id}' 已存在，无法添加。")
        return df, False

    folder_path = create_project_folders(project_id, project_name)
    if folder_path:
        new_project_data = {
            '课题编号': str(project_id),
            '课题名称': project_name,
            '负责人': pi_name,
            '课题状态': '申报中',
            '申报日期': datetime.now().strftime('%Y-%m-%d'),
            '立项日期': None,
            '计划开始日期': start_date,
            '计划结题日期': end_date,
            '实际结题日期': None,
            '课题文件夹路径': folder_path,
            '备注': notes
        }
        new_project_df = pd.DataFrame([new_project_data], columns=EXCEL_COLUMNS)
        df = pd.concat([df, new_project_df], ignore_index=True)
        print(f"课题 '{project_name}' (编号: {project_id}) 添加成功。")
        return df, True
    else:
        print(f"错误: 未能为课题 '{project_name}' 创建文件夹，添加失败。")
        return df, False


def update_project_status(df, project_id, new_status):
    """更新指定课题的状态"""
    project_id_str = str(project_id)
    project_index = df.index[df['课题编号'].astype(str) == project_id_str].tolist()

    if not project_index:
        print(f"错误: 找不到课题编号 '{project_id_str}'。")
        return df, False

    idx = project_index[0]
    df.loc[idx, '课题状态'] = new_status
    print(f"课题 '{project_id_str}' 的状态已更新为 '{new_status}'。")

    if new_status == '已立项' and pd.isna(df.loc[idx, '立项日期']):
        df.loc[idx, '立项日期'] = datetime.now().strftime('%Y-%m-%d')
        print(f"已记录课题 '{project_id_str}' 的立项日期。")

    if new_status == '已结题' and pd.isna(df.loc[idx, '实际结题日期']):
        df.loc[idx, '实际结题日期'] = datetime.now().strftime('%Y-%m-%d')
        print(f"已记录课题 '{project_id_str}' 的实际结题日期。")

    return df, True


def find_project(df, query, column='课题名称'):
    """根据指定列和查询词查找课题 (大小写不敏感)"""
    if column not in df.columns:
        print(f"错误: 列名 '{column}' 不存在于数据表中。可选列: {', '.join(df.columns)}")
        return pd.DataFrame(columns=df.columns)

    try:
        results = df[df[column].astype(str).str.contains(query, case=False, na=False)]
        if results.empty:
            print(f"未找到 '{column}' 中包含 '{query}' 的课题。")
        else:
            print(f"\n--- 查询结果 ('{column}' 包含 '{query}') ---")
            print(results.to_string(index=False))
            print("--- 查询结束 ---")
        return results
    except Exception as e:
        print(f"查询课题时出错: {e}")
        return pd.DataFrame(columns=df.columns)


def open_folder(folder_path):
    """在文件资源管理器中打开文件夹"""
    if not folder_path or not isinstance(folder_path, str):
        print(f"错误: 提供的文件夹路径 '{folder_path}' 无效。")
        return False
    if not os.path.isdir(folder_path):
        print(f"错误: 文件夹路径 '{folder_path}' 不是一个有效的目录或不存在。")
        return False
    try:
        print(f"尝试打开文件夹: {folder_path}")
        if sys.platform == 'win32':
            os.startfile(os.path.normpath(folder_path))
        elif sys.platform == 'darwin':
            subprocess.run(['open', folder_path], check=True)
        else:
            subprocess.run(['xdg-open', folder_path], check=True)
        return True
    except FileNotFoundError:
        print(f"错误: 无法找到用于打开文件夹的程序 (如 'open' 或 'xdg-open')。")
    except subprocess.CalledProcessError as e:
        print(f"错误: 执行打开文件夹命令时出错: {e}")
    except Exception as e:
        print(f"打开文件夹 '{folder_path}' 时发生未知错误: {e}")
    return False


def get_project_folder_path(df, project_id):
    """获取指定课题的文件夹路径"""
    project_id_str = str(project_id)
    project_row = df[df['课题编号'].astype(str) == project_id_str]
    if project_row.empty:
        print(f"错误: 找不到课题编号 '{project_id_str}'。")
        return None
    folder_path = project_row['课题文件夹路径'].iloc[0]
    if pd.isna(folder_path):
        print(f"警告: 课题 '{project_id_str}' 的文件夹路径为空。")
        return None
    return folder_path


# --- GUI 类 ---
class ProjectManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("科研课题管理系统")
        self.root.geometry("1200x700")

        if not os.path.exists(PROJECTS_ROOT_DIR):
            try:
                os.makedirs(PROJECTS_ROOT_DIR)
                print(f"已创建课题根目录: {PROJECTS_ROOT_DIR}")
            except Exception as e:
                messagebox.showerror("错误", f"无法创建根目录 '{PROJECTS_ROOT_DIR}': {e}")
                self.root.destroy()
                return

        self.projects_df = load_projects_data()
        self.data_changed = False

        self.create_widgets()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.refresh_treeview()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(button_frame, text="添加新课题", command=self.add_project_dialog).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="更新课题状态", command=self.update_status_dialog).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="打开课题文件夹", command=self.open_selected_folder).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="刷新数据", command=self.refresh_treeview).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="保存数据", command=self.save_data).pack(side=tk.LEFT, padx=5)

        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(search_frame, text="搜索列:").pack(side=tk.LEFT, padx=(0, 5))
        self.search_column = ttk.Combobox(search_frame, values=EXCEL_COLUMNS, width=15)
        self.search_column.current(1)
        self.search_column.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(search_frame, text="搜索内容:").pack(side=tk.LEFT, padx=(0, 5))
        self.search_entry = ttk.Entry(search_frame, width=30)
        self.search_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.search_entry.bind("<Return>", lambda e: self.search_projects())

        ttk.Button(search_frame, text="搜索", command=self.search_projects).pack(side=tk.LEFT)
        ttk.Button(search_frame, text="清除搜索", command=self.clear_search).pack(side=tk.LEFT, padx=(5, 0))

        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(table_frame, columns=EXCEL_COLUMNS, show="headings")

        for col in EXCEL_COLUMNS:
            self.tree.heading(col, text=col)
            if col in ['课题编号', '负责人', '课题状态', '申报日期', '立项日期', '计划开始日期', '计划结题日期',
                       '实际结题日期']:
                self.tree.column(col, width=100, anchor=tk.CENTER)
            elif col == '课题名称':
                self.tree.column(col, width=200, anchor=tk.W)
            elif col == '课题文件夹路径':
                self.tree.column(col, width=250, anchor=tk.W)
            elif col == '备注':
                self.tree.column(col, width=150, anchor=tk.W)

        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="打开课题文件夹", command=self.open_selected_folder)
        self.context_menu.add_command(label="更新课题状态", command=self.update_status_dialog)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="编辑课题信息", command=self.edit_project_dialog)
        self.context_menu.add_command(label="删除课题", command=self.delete_project)

        self.tree.bind("<Button-3>", self.show_context_menu)
        self.tree.bind("<Double-1>", lambda e: self.open_selected_folder())

        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def refresh_treeview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for _, row in self.projects_df.iterrows():
            values = [row[col] if pd.notna(row[col]) else "" for col in EXCEL_COLUMNS]
            self.tree.insert("", tk.END, values=values)

        self.status_var.set(f"共 {len(self.projects_df)} 条课题记录")

    def search_projects(self):
        column = self.search_column.get()
        query = self.search_entry.get().strip()

        if not query:
            self.refresh_treeview()
            return

        for item in self.tree.get_children():
            self.tree.delete(item)

        results = find_project(self.projects_df, query, column)

        for _, row in results.iterrows():
            values = [row[col] if pd.notna(row[col]) else "" for col in EXCEL_COLUMNS]
            self.tree.insert("", tk.END, values=values)

        self.status_var.set(f"找到 {len(results)} 条匹配记录")

    def clear_search(self):
        self.search_entry.delete(0, tk.END)
        self.refresh_treeview()

    def add_project_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("添加新课题")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()

        form_frame = ttk.Frame(dialog, padding="20")
        form_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(form_frame, text="课题编号:").grid(row=0, column=0, sticky=tk.W, pady=5)
        id_entry = ttk.Entry(form_frame, width=30)
        id_entry.grid(row=0, column=1, sticky=tk.W, pady=5)

        ttk.Label(form_frame, text="课题名称:").grid(row=1, column=0, sticky=tk.W, pady=5)
        name_entry = ttk.Entry(form_frame, width=30)
        name_entry.grid(row=1, column=1, sticky=tk.W, pady=5)

        ttk.Label(form_frame, text="负责人:").grid(row=2, column=0, sticky=tk.W, pady=5)
        pi_entry = ttk.Entry(form_frame, width=30)
        pi_entry.grid(row=2, column=1, sticky=tk.W, pady=5)

        ttk.Label(form_frame, text="计划开始日期:").grid(row=3, column=0, sticky=tk.W, pady=5)
        start_entry = ttk.Entry(form_frame, width=30)
        start_entry.grid(row=3, column=1, sticky=tk.W, pady=5)
        ttk.Label(form_frame, text="(YYYY-MM-DD 格式)").grid(row=3, column=2, sticky=tk.W, pady=5)

        ttk.Label(form_frame, text="计划结题日期:").grid(row=4, column=0, sticky=tk.W, pady=5)
        end_entry = ttk.Entry(form_frame, width=30)
        end_entry.grid(row=4, column=1, sticky=tk.W, pady=5)
        ttk.Label(form_frame, text="(YYYY-MM-DD 格式)").grid(row=4, column=2, sticky=tk.W, pady=5)

        ttk.Label(form_frame, text="备注:").grid(row=5, column=0, sticky=tk.W, pady=5)
        notes_text = tk.Text(form_frame, width=30, height=5)
        notes_text.grid(row=5, column=1, sticky=tk.W, pady=5)

        button_frame = ttk.Frame(form_frame)
        button_frame.grid(row=6, column=0, columnspan=3, pady=20)

        def submit():
            project_id = id_entry.get().strip()
            project_name = name_entry.get().strip()
            pi_name = pi_entry.get().strip()
            start_date = start_entry.get().strip() or None
            end_date = end_entry.get().strip() or None
            notes = notes_text.get("1.0", tk.END).strip()

            if not project_id or not project_name or not pi_name:
                messagebox.showerror("错误", "课题编号、课题名称和负责人为必填项！", parent=dialog)
                return

            date_format = r'^\d{4}-\d{2}-\d{2}$'
            if start_date and not re.match(date_format, start_date):
                messagebox.showerror("错误", "计划开始日期格式不正确，请使用 YYYY-MM-DD 格式！", parent=dialog)
                return
            if end_date and not re.match(date_format, end_date):
                messagebox.showerror("错误", "计划结题日期格式不正确，请使用 YYYY-MM-DD 格式！", parent=dialog)
                return

            self.projects_df, success = add_new_project(
                self.projects_df, project_id, project_name, pi_name, start_date, end_date, notes
            )

            if success:
                self.data_changed = True
                messagebox.showinfo("成功", f"课题 '{project_name}' 添加成功！", parent=dialog)
                dialog.destroy()
                self.refresh_treeview()
            else:
                messagebox.showerror("错误", f"添加课题 '{project_name}' 失败！", parent=dialog)

        ttk.Button(button_frame, text="提交", command=submit).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT)

    def update_status_dialog(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("警告", "请先选择一个课题！")
            return

        project_id = self.tree.item(selected_item[0], "values")[0]

        dialog = tk.Toplevel(self.root)
        dialog.title("更新课题状态")
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.grab_set()

        form_frame = ttk.Frame(dialog, padding="20")
        form_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(form_frame, text="课题编号:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Label(form_frame, text=project_id).grid(row=0, column=1, sticky=tk.W, pady=5)

        ttk.Label(form_frame, text="新状态:").grid(row=1, column=0, sticky=tk.W, pady=5)
        status_combo = ttk.Combobox(form_frame, values=["申报中", "已立项", "进行中", "已结题", "中止"], width=15)
        status_combo.grid(row=1, column=1, sticky=tk.W, pady=5)
        status_combo.current(0)

        button_frame = ttk.Frame(form_frame)