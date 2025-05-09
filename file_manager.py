# file_manager.py
import os
import re
import subprocess
import sys
from datetime import datetime

# 从 config 模块导入配置
from config import PROJECTS_ROOT_DIR

# 定义文件夹结构，包含 03_过程管理 的子文件夹
FOLDER_STRUCTURE = [
    '01_申报',
    '02_立项',
    {
        'name': '03_过程管理',
        'subfolders': ['01_开题', '02_中期', '03_变更']
    },
    '04_结题',
    '05_财务',
    '06_其他'
]

def sanitize_foldername(name):
    """清理文件名，移除或替换不适用于文件夹名称的字符"""
    forbidden_chars = r'[\<\>\:\"\/\\\|\?\*]'
    sanitized_name = re.sub(forbidden_chars, '_', name)
    sanitized_name = sanitized_name.strip()
    return sanitized_name

def create_project_folders(project_id, project_name, status, start_year, custom_path=None):
    """为新课题创建文件夹结构，使用命名规则：年度-课题状态-课题编号-课题名称"""
    sanitized_name = sanitize_foldername(project_name)
    sanitized_id = sanitize_foldername(str(project_id))
    # Use provided start_year or current year if None
    year = str(start_year) if start_year and str(start_year) != '' else str(datetime.now().year)
    # Use provided status or default to '申报'
    status = status if status and status in ['申报', '已立项', '在研', '中期已过', '已结题', '延期', '中止', '其他'] else '申报'

    folder_name = f"{year}-{status}-{sanitized_id}-{sanitized_name}"
    # Use custom_path if provided, else default to PROJECTS_ROOT_DIR
    base_path = custom_path if custom_path and os.path.isdir(os.path.dirname(custom_path)) else PROJECTS_ROOT_DIR
    project_path = os.path.join(base_path, folder_name)

    try:
        if not os.path.exists(project_path):
            os.makedirs(project_path)
            print(f"创建课题主文件夹: {project_path}")
        else:
            print(f"信息: 文件夹 '{project_path}' 已存在，检查子文件夹。")

        created_subfolder = False
        for item in FOLDER_STRUCTURE:
            if isinstance(item, dict):
                # Handle nested folder structure (e.g., 03_过程管理)
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
                # Handle regular folder
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

def rename_project_folder(old_path, project_id, project_name, new_status, start_year):
    """根据新状态重命名课题文件夹"""
    if not old_path or not os.path.exists(old_path):
        print(f"错误: 原文件夹路径 '{old_path}' 不存在或无效。")
        return None

    sanitized_name = sanitize_foldername(project_name)
    sanitized_id = sanitize_foldername(str(project_id))
    year = str(start_year) if start_year and str(start_year) != '' else str(datetime.now().year)
    new_folder_name = f"{year}-{new_status}-{sanitized_id}-{sanitized_name}"
    new_path = os.path.join(os.path.dirname(old_path), new_folder_name)

    try:
        if old_path == new_path:
            print(f"信息: 文件夹名称未更改，无需重命名: {old_path}")
            return old_path
        if os.path.exists(new_path):
            print(f"错误: 目标文件夹 '{new_path}' 已存在，无法重命名。")
            return old_path
        os.rename(old_path, new_path)
        print(f"文件夹已从 '{old_path}' 重命名为 '{new_path}'")
        return new_path
    except OSError as e:
        print(f"重命名文件夹从 '{old_path}' 到 '{new_path}' 时发生 OS 错误: {e}")
        return old_path
    except Exception as e:
        print(f"重命名文件夹时发生未知错误: {e}")
        return old_path

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