import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import os
from datetime import datetime
import re
from tkcalendar import DateEntry
from data_manager import load_projects_data, save_projects_data, add_new_project, update_project, delete_project
from file_manager import create_project_folders, open_folder


# ... (Other imports and code unchanged)

class ProjectDialog:
    def __init__(self, parent, app, project_id=None):
        self.app = app
        self.project_id = project_id
        self.data_changed = False

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("编辑课题信息" if project_id else "添加新课题")
        self.dialog.geometry("600x700")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # Create a canvas and scrollbar
        content_frame = ttk.Frame(self.dialog, padding="10")
        content_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(content_frame)
        scrollbar = ttk.Scrollbar(content_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Store canvas for reference
        self.canvas = canvas

        # Bind mouse wheel to canvas (use bind, not bind_all, to limit scope)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)

        # Form fields
        form_frame = ttk.Frame(self.scrollable_frame)
        form_frame.pack(fill=tk.X, padx=10, pady=10)

        fields = [
            ("课题编号", "entry"),
            ("课题名称", "entry"),
            ("负责人", "entry"),
            ("课题状态", "combobox", ["申报", "已立项", "在研", "中期已过", "已结题", "延期", "中止"]),
            ("申报日期", "date"),
            ("立项日期", "date"),
            ("计划开始日期", "date"),
            ("计划结题日期", "date"),
            ("实际结题日期", "date"),
            ("备注", "text"),
        ]

        self.entries = {}
        row = 0
        for field_name, field_type, *args in fields:
            label = ttk.Label(form_frame, text=f"{field_name}:")
            label.grid(row=row, column=0, sticky=tk.W, pady=5)

            if field_type == "entry":
                entry = ttk.Entry(form_frame, width=50)
                entry.grid(row=row, column=1, sticky=tk.W, pady=5)
                self.entries[field_name] = entry
            elif field_type == "combobox":
                entry = ttk.Combobox(form_frame, values=args[0], width=47)
                entry.grid(row=row, column=1, sticky=tk.W, pady=5)
                self.entries[field_name] = entry
            elif field_type == "date":
                entry = DateEntry(form_frame, width=47, date_pattern="yyyy-mm-dd")
                entry.grid(row=row, column=1, sticky=tk.W, pady=5)
                self.entries[field_name] = entry
            elif field_type == "text":
                entry = tk.Text(form_frame, width=50, height=4)
                entry.grid(row=row, column=1, sticky=tk.W, pady=5)
                self.entries[field_name] = entry

            row += 1

        # Buttons
        button_frame = ttk.Frame(form_frame)
        button_frame.grid(row=row, column=0, columnspan=2, pady=20)

        ttk.Button(button_frame, text="保存", command=self.save).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=self.dialog.destroy).pack(side=tk.LEFT)

        # Load existing project data if editing
        if project_id:
            project = app.projects_df[app.projects_df['课题编号'].astype(str) == str(project_id)]
            if not project.empty:
                row = project.iloc[0]
                for field in self.entries:
                    value = row.get(field, "")
                    if pd.notna(value):
                        if isinstance(self.entries[field], tk.Text):
                            self.entries[field].delete("1.0", tk.END)
                            self.entries[field].insert("1.0", str(value))
                        elif isinstance(self.entries[field], DateEntry):
                            try:
                                self.entries[field].set_date(value)
                            except:
                                self.entries[field].set_date("")
                        else:
                            self.entries[field].delete(0, tk.END)
                            self.entries[field].insert(0, str(value))

        # Ensure dialog is centered
        self.dialog.update_idletasks()
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        dialog_width = self.dialog.winfo_width()
        dialog_height = self.dialog.winfo_height()
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        self.dialog.geometry(f"+{x}+{y}")

    def _on_mousewheel(self, event):
        """Handle mouse wheel scrolling for the canvas."""
        if self.canvas.winfo_exists():
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def save(self):
        """Save the project data and close the dialog."""
        data = {}
        for field, entry in self.entries.items():
            if isinstance(entry, tk.Text):
                value = entry.get("1.0", tk.END).strip()
            elif isinstance(entry, DateEntry):
                value = entry.get() or None
            else:
                value = entry.get().strip() or None
            data[field] = value

        if not data["课题编号"] or not data["课题名称"] or not data["负责人"]:
            messagebox.showerror("错误", "课题编号、课题名称和负责人为必填项！", parent=self.dialog)
            return

        date_fields = ["申报日期", "立项日期", "计划开始日期", "计划结题日期", "实际结题日期"]
        date_format = r'^\d{4}-\d{2}-\d{2}$'
        for field in date_fields:
            if data[field] and not re.match(date_format, data[field]):
                messagebox.showerror("错误", f"{field} 格式不正确，请使用 YYYY-MM-DD 格式！", parent=self.dialog)
                return

        project_id = data["课题编号"]
        if not self.project_id:  # New project
            if not self.app.projects_df['课题编号'].astype(str).str.fullmatch(str(project_id)).empty:
                messagebox.showerror("错误", f"课题编号 '{project_id}' 已存在！", parent=self.dialog)
                return
            start_year = data["申报日期"][:4] if data["申报日期"] else str(datetime.now().year)
            folder_path = create_project_folders(project_id, data["课题名称"], data["课题状态"], start_year)
            if not folder_path:
                messagebox.showerror("错误", f"无法创建课题文件夹！", parent=self.dialog)
                return
            data["课题文件夹路径"] = folder_path
            self.app.projects_df, success = add_new_project(
                self.app.projects_df,
                project_id,
                data["课题名称"],
                data["负责人"],
                data["课题状态"],
                data["申报日期"],
                data["立项日期"],
                data["计划开始日期"],
                data["计划结题日期"],
                data["实际结题日期"],
                data["备注"],
                folder_path
            )
        else:  # Update existing project
            self.app.projects_df, success = update_project(
                self.app.projects_df,
                project_id,
                data["课题名称"],
                data["负责人"],
                data["课题状态"],
                data["申报日期"],
                data["立项日期"],
                data["计划开始日期"],
                data["计划结题日期"],
                data["实际结题日期"],
                data["备注"]
            )

        if success:
            self.data_changed = True
            self.app.refresh_treeview()
            save_projects_data(self.app.projects_df)
            messagebox.showinfo("成功", "课题信息已保存！", parent=self.dialog)
            self.dialog.destroy()
        else:
            messagebox.showerror("错误", "保存课题信息失败！", parent=self.dialog)

    def destroy(self):
        """Clean up bindings before destroying the dialog."""
        # Unbind mouse wheel to prevent callbacks on destroyed canvas
        self.canvas.unbind("<MouseWheel>")
        self.dialog.destroy()


class Application:
    # ... (Rest of Application class unchanged)
    def __init__(self, root):
        self.root = root
        self.root.title("科研课题管理系统")
        self.root.geometry("1200x800")

        self.projects_df = load_projects_data()
        self.data_changed = False

        self.create_widgets()
        self.refresh_treeview()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        # ... (Unchanged widget creation code)
        pass

    def refresh_treeview(self):
        # ... (Unchanged treeview refresh code)
        pass

    def add_project(self):
        dialog = ProjectDialog(self.root, self)
        self.root.wait_window(dialog.dialog)
        if dialog.data_changed:
            self.data_changed = True
            self.refresh_treeview()

    def edit_project(self):
        # ... (Unchanged edit project code)
        pass

    def on_closing(self):
        # ... (Unchanged closing handler)
        pass


# ... (Rest of gui.py unchanged)

if __name__ == "__main__":
    root = tk.Tk()
    app = Application(root)
    root.mainloop()