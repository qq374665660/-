# main.py
import tkinter as tk
from gui import ProjectManagerApp

if __name__ == "__main__":
    # 确保项目根目录存在
    from config import PROJECTS_ROOT_DIR
    import os
    if not os.path.exists(PROJECTS_ROOT_DIR):
        try:
            os.makedirs(PROJECTS_ROOT_DIR)
            print(f"已创建课题根目录: {PROJECTS_ROOT_DIR}")
        except OSError as e:
            print(f"创建课题根目录 '{PROJECTS_ROOT_DIR}' 时出错: {e}")
            # 如果根目录创建失败，可能无法继续，可以考虑退出或提示用户

    # 创建 Tkinter 主窗口
    root = tk.Tk()
    # 实例化应用程序类
    app = ProjectManagerApp(root)
    # 启动 Tkinter 事件循环
    root.mainloop()