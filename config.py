# config.py
import os

# --- 配置 ---
# Excel 文件名
EXCEL_FILE = '科研课题管理总表.xlsx'
# 课题文件夹的根目录
PROJECTS_ROOT_DIR = os.path.join(os.path.abspath('.'), '科研课题管理')
# Excel 中的工作表名称
SHEET_NAME = '课题列表'
# 标准的课题子文件夹结构
FOLDER_STRUCTURE = ['01_申报', '02_立项', '03_过程管理', '04_结题', '05_财务', '06_其他']
# Excel 表格的列名 (移除 '课题文件夹路径')
EXCEL_COLUMNS = [
    '序号', '归口单位', '承担单位', '课题名称', '课题级别', '课题类型', '开始年份',
    '参与角色', '课题状态', '课题编号', '课题联系人', '课题负责人', 
    '开始日期', '计划结束日期', '延期时间', '实际结题时间', 
    '总预算', '外部专项经费', '院自筹经费', '所属单位自筹经费'
]

# --- 下拉选项 --- 
PROJECT_TYPES = ['应用研究', '试验发展', '其他']
PROJECT_LEVELS = ['国家级', '省部级', '公司级']
PROJECT_STATUSES = ['申报', '已立项', '在研', '中期已过', '已结题', '延期', '中止', '其他']
PROJECT_CHARACTER = ['牵头','参与']
PROJECT_AUTHOR =['西勘院','地下空间']
