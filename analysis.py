import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import seaborn as sns
from wordcloud import WordCloud
import os
import matplotlib

# 设置matplotlib支持中文显示
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'SimSun', 'Arial Unicode MS']
matplotlib.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题


class AnalysisDialog(tk.Toplevel):
    def __init__(self, parent, projects_df, display_columns):
        super().__init__(parent)
        self.transient(parent)
        self.title("课题数据分析")
        self.geometry("1000x700")
        self.parent = parent
        self.projects_df = projects_df
        self.display_columns = display_columns
        self.selected_dimensions = []

        self.main_frame = ttk.Frame(self, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.setup_widgets()
        self.grab_set()

    def setup_widgets(self):
        # Dimension selection frame
        dim_frame = ttk.LabelFrame(self.main_frame, text="选择分析维度", padding="10")
        dim_frame.pack(fill=tk.X, pady=5)

        ttk.Label(dim_frame, text="可用维度:").pack(side=tk.LEFT, padx=5)
        self.dim_listbox = tk.Listbox(dim_frame, selectmode=tk.MULTIPLE, height=5, width=30)
        for col in self.display_columns:
            if col not in ['序号', '课题编号']:  # 只排除不适合聚合的字段，保留'课题名称'字段
                self.dim_listbox.insert(tk.END, col)
        self.dim_listbox.pack(side=tk.LEFT, padx=5)

        ttk.Button(dim_frame, text="添加维度", command=self.add_dimensions).pack(side=tk.LEFT, padx=5)

        # Selected dimensions display
        self.selected_dims_var = tk.StringVar(value="已选维度: 无")
        ttk.Label(dim_frame, textvariable=self.selected_dims_var).pack(side=tk.LEFT, padx=10)

        # Visualization type selection
        vis_frame = ttk.LabelFrame(self.main_frame, text="选择可视化类型", padding="10")
        vis_frame.pack(fill=tk.X, pady=5)

        self.vis_type = tk.StringVar(value="饼状图")
        vis_types = ["饼状图", "柱状图", "词云", "趋势图"]
        ttk.Label(vis_frame, text="可视化类型:").pack(side=tk.LEFT, padx=5)
        ttk.Combobox(vis_frame, textvariable=self.vis_type, values=vis_types, state="readonly").pack(side=tk.LEFT,
                                                                                                     padx=5)

        ttk.Button(vis_frame, text="生成可视化", command=self.generate_visualization).pack(side=tk.LEFT, padx=10)

        # Plot display frame
        self.plot_frame = ttk.Frame(self.main_frame)
        self.plot_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # Status bar
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(self.main_frame, textvariable=self.status_var, relief=tk.SUNKEN).pack(side=tk.BOTTOM, fill=tk.X)

    def add_dimensions(self):
        selected_indices = self.dim_listbox.curselection()
        self.selected_dimensions = [self.dim_listbox.get(i) for i in selected_indices]
        if not self.selected_dimensions:
            self.status_var.set("请至少选择一个维度")
            return
        self.selected_dims_var.set(f"已选维度: {', '.join(self.selected_dimensions)}")
        self.status_var.set(f"已选择 {len(self.selected_dimensions)} 个维度")

    def generate_visualization(self):
        if not self.selected_dimensions:
            messagebox.showwarning("警告", "请先选择分析维度！", parent=self)
            return

        vis_type = self.vis_type.get()

        # Clear previous plot
        for widget in self.plot_frame.winfo_children():
            widget.destroy()

        try:
            if vis_type == "饼状图":
                self.create_pie_chart()
            elif vis_type == "柱状图":
                self.create_bar_chart()
            elif vis_type == "词云":
                self.create_word_cloud()
            elif vis_type == "趋势图":
                self.create_trend_chart()
        except Exception as e:
            messagebox.showerror("错误", f"生成可视化失败: {e}", parent=self)
            self.status_var.set("生成可视化失败")

    def create_pie_chart(self):
        if len(self.selected_dimensions) > 1:
            messagebox.showwarning("警告", "饼状图仅支持单一维度分析", parent=self)
            return

        dim = self.selected_dimensions[0]
        counts = self.projects_df[dim].value_counts()

        fig, ax = plt.subplots(figsize=(6, 4))
        ax.pie(counts, labels=counts.index, autopct='%1.1f%%', startangle=90)
        ax.axis('equal')
        plt.title(f"{dim} 分布")

        self.embed_plot(fig)
        self.status_var.set(f"已生成 {dim} 的饼状图")

    def create_bar_chart(self):
        dim = self.selected_dimensions[0]
        if len(self.selected_dimensions) > 1:
            # For multiple dimensions, use groupby
            group_cols = self.selected_dimensions
            counts = self.projects_df.groupby(group_cols).size().unstack(fill_value=0)

            fig, ax = plt.subplots(figsize=(8, 5))
            counts.plot(kind='bar', stacked=False, ax=ax)
            ax.set_xlabel(group_cols[0])
            ax.set_ylabel("数量")
            ax.set_title("多维度课题分布")
            ax.legend(title=" | ".join(group_cols[1:]), bbox_to_anchor=(1.05, 1), loc='upper left')
            plt.tight_layout()
        else:
            counts = self.projects_df[dim].value_counts()
            fig, ax = plt.subplots(figsize=(6, 4))
            counts.plot(kind='bar', ax=ax)
            ax.set_xlabel(dim)
            ax.set_ylabel("数量")
            ax.set_title(f"{dim} 课题分布")

        self.embed_plot(fig)
        self.status_var.set(f"已生成 {', '.join(self.selected_dimensions)} 的柱状图")

    def create_word_cloud(self):
        if len(self.selected_dimensions) != 1 or self.selected_dimensions[0] != "课题名称":
            messagebox.showwarning("警告", "词云仅支持‘课题名称’维度", parent=self)
            return

        text = " ".join(self.projects_df["课题名称"].dropna().astype(str))
        if not text.strip():
            messagebox.showwarning("警告", "课题名称数据为空，无法生成词云", parent=self)
            return

        # 尝试查找系统中可用的中文字体
        font_path = None
        possible_fonts = [
            "C:\\Windows\\Fonts\\simhei.ttf",  # 黑体
            "C:\\Windows\\Fonts\\msyh.ttc",   # 微软雅黑
            "C:\\Windows\\Fonts\\simsun.ttc"  # 宋体
        ]
        
        for font in possible_fonts:
            if os.path.exists(font):
                font_path = font
                break
                
        wordcloud = WordCloud(
            width=800, 
            height=400, 
            background_color='white', 
            font_path=font_path,  # 使用找到的中文字体
            max_words=200,
            max_font_size=100,
            random_state=42
        ).generate(text)

        fig, ax = plt.subplots(figsize=(8, 4))
        ax.imshow(wordcloud, interpolation='bilinear')
        ax.axis('off')
        plt.title("课题名称词云")

        self.embed_plot(fig)
        self.status_var.set("已生成课题名称词云")

    def create_trend_chart(self):
        if len(self.selected_dimensions) != 1:
            messagebox.showwarning("警告", "趋势图仅支持单一维度分析", parent=self)
            return

        dim = self.selected_dimensions[0]
        # 检查所选维度是否可以作为趋势图的数据
        if pd.api.types.is_numeric_dtype(self.projects_df[dim]) or dim in ["开始年份", "开始日期", "计划结束日期", "实际结题时间"]:
            # 对于日期或年份类型的数据，使用value_counts并排序
            counts = self.projects_df[dim].value_counts().sort_index()
            if counts.empty:
                messagebox.showwarning("警告", f"{dim}数据为空，无法生成趋势图", parent=self)
                return

            fig, ax = plt.subplots(figsize=(8, 4))
            counts.plot(kind='line', marker='o', ax=ax)
            ax.set_xlabel(dim)
            ax.set_ylabel("课题数量")
            ax.set_title(f"课题数量随{dim}趋势")
            ax.grid(True)

            self.embed_plot(fig)
            self.status_var.set(f"已生成{dim}趋势图")
        else:
            messagebox.showwarning("警告", f"{dim}不适合生成趋势图，请选择日期或数值类型的维度", parent=self)
            return

    def embed_plot(self, fig):
        canvas = FigureCanvasTkAgg(fig, master=self.plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        plt.close(fig)  # Close to free memory