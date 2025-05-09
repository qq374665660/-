# data_manager.py
import pandas as pd
from datetime import datetime
import os
import file_manager

# 从 config 模块导入配置
from config import EXCEL_FILE, SHEET_NAME, EXCEL_COLUMNS, PROJECT_STATUSES

def load_projects_data():
    """从 Excel 加载课题数据，处理日期和特定类型，并为新项目创建文件夹"""
    try:
        date_columns = ['开始日期', '计划结束日期', '延期时间', '实际结题时间']
        string_columns = {
            '课题编号': str, '课题联系人': str, '课题负责人': str,
            '承担单位': str, '参与角色': str, '课题名称': str, '归口单位': str,
            '课题级别': str, '课题类型': str, '课题状态': str,
        }
        string_columns_to_use = {k: v for k, v in string_columns.items() if k in EXCEL_COLUMNS}

        df = pd.read_excel(
            EXCEL_FILE,
            sheet_name=SHEET_NAME,
            engine='openpyxl',
            dtype=string_columns_to_use,
        )
        print(f"成功从 '{EXCEL_FILE}' 加载 {len(df)} 条课题数据。")

        # Validate project IDs
        if '课题编号' in df.columns:
            project_ids = df['课题编号'].astype(str)
            # Check for missing project IDs
            missing_ids = project_ids.isna() | (project_ids == '')
            if missing_ids.any():
                print(f"警告: 发现 {missing_ids.sum()} 条记录缺少课题编号，将为这些记录分配临时编号。")
                df.loc[missing_ids, '课题编号'] = [f"temp_id_{i}" for i in df.index[missing_ids]]
            # Check for duplicate project IDs
            duplicates = project_ids.duplicated(keep=False)
            if duplicates.any():
                duplicate_ids = df.loc[duplicates, '课题编号'].unique()
                print(f"警告: 发现重复的课题编号: {', '.join(map(str, duplicate_ids))}。请确保课题编号唯一。")
        else:
            print("警告: Excel 文件中缺少 '课题编号' 列，将为所有记录分配临时编号。")
            df['课题编号'] = [f"temp_id_{i}" for i in df.index]

        missing_cols_added = False
        for col in EXCEL_COLUMNS:
            if col not in df.columns:
                df[col] = None
                missing_cols_added = True
                print(f"警告: 文件中缺少列 '{col}'，已添加。")

        df = df[EXCEL_COLUMNS]

        if '序号' in EXCEL_COLUMNS:
            df['序号'] = range(1, len(df) + 1)

        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].dt.strftime('%Y-%m-%d').fillna('')
            else:
                df[col] = ''

        numeric_cols = ['外部专项经费', '院自筹经费', '所属单位自筹经费', '总预算']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            else:
                df[col] = 0.0

        if all(c in df.columns for c in ['外部专项经费', '院自筹经费', '所属单位自筹经费']):
            df['总预算'] = df['外部专项经费'] + df['院自筹经费'] + df['所属单位自筹经费']
        elif '总预算' in df.columns:
            print("警告: 缺少部分经费列，无法重新计算总预算。将使用文件中读取的值。")

        if '开始年份' in df.columns and '开始日期' in df.columns:
            mask = df['开始日期'] != ''
            df.loc[mask, '开始年份'] = pd.to_datetime(df.loc[mask, '开始日期']).dt.year
            df['开始年份'] = df['开始年份'].astype('Int64').fillna(pd.NA).astype(str).replace('<NA>', '')

        df.fillna('', inplace=True)

        # Check for projects without valid folders and create them
        folder_cache = {}
        for idx, row in df.iterrows():
            project_id = str(row['课题编号'])
            project_name = row.get('课题名称', '')
            status = row.get('课题状态', '申报')
            start_year = row.get('开始年份', '')
            # Check if folder needs to be created
            folder_path = None
            # Assuming folder path isn't stored in Excel (as per EXCEL_COLUMNS), we create folders for all projects
            if not folder_path or not os.path.isdir(folder_path):
                folder_path = file_manager.create_project_folders(project_id, project_name, status, start_year)
                if folder_path:
                    print(f"自动为课题 '{project_name}' (编号: {project_id}) 创建文件夹: {folder_path}")
                    folder_cache[project_id] = folder_path
                else:
                    print(f"警告: 无法为课题 '{project_name}' (编号: {project_id}) 创建文件夹")
            else:
                folder_cache[project_id] = folder_path

        if missing_cols_added:
            print("提示：由于添加了缺失列，建议检查数据并保存。")

        return df, folder_cache

    except FileNotFoundError:
        print(f"信息: Excel 文件 '{EXCEL_FILE}' 未找到。将创建一个新的空 DataFrame。")
        empty_df = pd.DataFrame(columns=EXCEL_COLUMNS)
        for col in EXCEL_COLUMNS:
            if col in ['外部专项经费', '院自筹经费', '所属单位自筹经费', '总预算']:
                empty_df[col] = 0.0
            elif col in ['开始日期', '计划结束日期', '延期时间', '实际结题时间']:
                empty_df[col] = ''
            elif col == '开始年份':
                empty_df[col] = ''
            elif col == '序号':
                empty_df[col] = pd.Series(dtype='int64')
            else:
                empty_df[col] = ''
        return empty_df, {}

    except Exception as e:
        print(f"加载 Excel 文件 '{EXCEL_FILE}' 时发生严重错误: {e}")
        return pd.DataFrame(columns=EXCEL_COLUMNS).fillna(''), {}

def save_projects_data(df):
    """将课题数据保存回 Excel，确保数据类型正确"""
    try:
        df_to_save = df.reindex(columns=EXCEL_COLUMNS)
        numeric_cols = ['外部专项经费', '院自筹经费', '所属单位自筹经费', '总预算']
        for col in numeric_cols:
            if col in df_to_save.columns:
                df_to_save[col] = pd.to_numeric(df_to_save[col].replace('', 0), errors='coerce').fillna(0)

        date_columns = ['开始日期', '计划结束日期', '延期时间', '实际结题时间']
        for col in date_columns:
            if col in df_to_save.columns:
                df_to_save[col] = pd.to_datetime(df_to_save[col].replace('', None), errors='coerce')

        if '开始年份' in df_to_save.columns:
            df_to_save['开始年份'] = pd.to_numeric(df_to_save['开始年份'], errors='coerce').astype('Int64').fillna(pd.NA)

        df_to_save.to_excel(EXCEL_FILE, sheet_name=SHEET_NAME, index=False, engine='openpyxl')
        print(f"数据已成功保存到 '{EXCEL_FILE}'。")
        return True
    except PermissionError:
        print(f"保存错误: 无法写入文件 '{EXCEL_FILE}'。请确保文件未被其他程序打开，并且您有写入权限。")
        return False
    except Exception as e:
        print(f"保存数据到 Excel 文件 '{EXCEL_FILE}' 时发生未知错误: {e}")
        return False

def add_project_record(df, data, custom_folder_path=None):
    """添加新课题记录到 DataFrame，并创建文件夹"""
    project_id = data.get('课题编号')
    if project_id is None or str(project_id).strip() == "":
        print(f"错误: 未提供有效的课题编号，无法添加记录。")
        return df, False, None

    project_id_str = str(project_id).strip()
    if project_id_str.lower() in df['课题编号'].astype(str).str.lower().values:
        print(f"错误: 课题编号 '{project_id_str}' 已存在，无法添加。")
        return df, False, None

    # Create folder using file_manager
    project_name = data.get('课题名称', '')
    status = data.get('课题状态', '申报')
    start_year = data.get('开始年份', '')
    folder_path = file_manager.create_project_folders(project_id, project_name, status, start_year, custom_folder_path)
    if not folder_path:
        print(f"错误: 未能为课题 '{project_name}' 创建文件夹，添加失败。")
        return df, False, None

    new_record = {col: None for col in EXCEL_COLUMNS}
    for key, value in data.items():
        if key in EXCEL_COLUMNS:
            cleaned_value = value.strip() if isinstance(value, str) else value
            new_record[key] = cleaned_value if cleaned_value is not None else None

    budget_fields = ['外部专项经费', '院自筹经费', '所属单位自筹经费']
    for field in budget_fields:
        if field in new_record:
            try:
                new_record[field] = float(new_record[field] or 0)
            except (ValueError, TypeError):
                print(f"警告: 添加时字段 '{field}' 的值 '{new_record[field]}' 无效，已设为 0。")
                new_record[field] = 0.0

    new_record['总预算'] = sum(new_record.get(f, 0.0) for f in budget_fields)

    if '开始日期' in new_record and new_record['开始日期']:
        try:
            start_date = pd.to_datetime(new_record['开始日期'])
            new_record['开始年份'] = start_date.year
        except (ValueError, TypeError):
            print(f"警告: 添加时开始日期 '{new_record['开始日期']}' 格式无效，无法计算开始年份。")
            new_record['开始年份'] = None
    else:
        new_record['开始年份'] = None

    date_columns = ['开始日期', '计划结束日期', '延期时间', '实际结题时间']
    for col in date_columns:
        if col in new_record and new_record[col]:
            try:
                new_record[col] = pd.to_datetime(new_record[col]).strftime('%Y-%m-%d')
            except (ValueError, TypeError):
                print(f"警告: 添加时日期字段 '{col}' 的值 '{new_record[col]}' 格式无效，已清空。")
                new_record[col] = None
        elif col in new_record:
            new_record[col] = None

    new_df_row = pd.DataFrame([new_record], columns=EXCEL_COLUMNS)
    df = pd.concat([df, new_df_row], ignore_index=True)

    if '序号' in df.columns:
        df['序号'] = range(1, len(df) + 1)

    print(f"课题 '{project_name}' (编号: {project_id_str}) 添加成功。")
    return df, True, folder_path

def update_project_status(df, project_id, new_status, folder_path=None):
    """更新指定课题的状态，并重命名文件夹"""
    project_id_str = str(project_id)
    project_index = df.index[df['课题编号'].astype(str) == project_id_str].tolist()

    if not project_index:
        print(f"错误: 找不到课题编号 '{project_id_str}'。")
        return df, False, folder_path

    if new_status not in PROJECT_STATUSES:
        print(f"错误: 无效的课题状态 '{new_status}'。可选状态: {', '.join(PROJECT_STATUSES)}")
        return df, False, folder_path

    idx = project_index[0]
    old_status = df.loc[idx, '课题状态']
    df.loc[idx, '课题状态'] = new_status
    print(f"课题 '{project_id_str}' 的状态已更新为 '{new_status}'。")

    if folder_path and old_status != new_status:
        project_name = df.loc[idx, '课题名称']
        start_year = df.loc[idx, '开始年份']
        new_folder_path = file_manager.rename_project_folder(folder_path, project_id, project_name, new_status, start_year)
        if new_folder_path != folder_path:
            folder_path = new_folder_path

    return df, True, folder_path

def find_project(df, query, column='课题名称'):
    """根据指定列和查询词查找课题 (大小写不敏感)"""
    if column not in df.columns:
        print(f"错误: 列名 '{column}' 不存在于数据表中。可选列: {', '.join(df.columns)}")
        return pd.DataFrame(columns=df.columns)

    if query is None or query.strip() == "":
        print("信息: 查询词为空，返回所有数据。")
        return df

    try:
        results = df[df[column].astype(str).str.contains(query, case=False, na=False)]
        if results.empty:
            print(f"未找到 '{column}' 中包含 '{query}' 的课题。")
        return results
    except Exception as e:
        print(f"查询课题时出错: {e}")
        return pd.DataFrame(columns=df.columns)

def update_project_record(df, project_id, updated_data, folder_path=None):
    """更新指定课题的记录信息"""
    project_id_str = str(project_id)
    project_index = df.index[df['课题编号'].astype(str) == project_id_str].tolist()

    if not project_index:
        print(f"错误: 找不到课题编号 '{project_id_str}' 无法更新。")
        return df, False, folder_path

    idx = project_index[0]
    try:
        budget_changed = False
        start_date_changed = False

        for key, value in updated_data.items():
            if key in df.columns:
                if key in ['课题编号', '总预算', '开始年份', '序号']:
                    continue

                original_value = df.loc[idx, key]
                cleaned_value = value.strip() if isinstance(value, str) else value
                original_compare = '' if pd.isna(original_value) else str(original_value)
                new_compare = '' if cleaned_value is None else str(cleaned_value)

                if key in ['外部专项经费', '院自筹经费', '所属单位自筹经费']:
                    budget_changed = True
                    try:
                        num_value = float(cleaned_value or 0)
                        df.loc[idx, key] = num_value
                    except (ValueError, TypeError):
                        print(f"警告: 更新时数值字段 '{key}' 的值 '{cleaned_value}' 无效，已设为 0。")
                        df.loc[idx, key] = 0.0

                elif key == '开始日期':
                    start_date_changed = True
                    if cleaned_value:
                        try:
                            df.loc[idx, key] = pd.to_datetime(cleaned_value).strftime('%Y-%m-%d')
                        except (ValueError, TypeError):
                            print(f"警告: 更新时日期字段 '{key}' 的值 '{cleaned_value}' 格式无效，已清空。")
                            df.loc[idx, key] = ''
                    else:
                        df.loc[idx, key] = ''

                elif key in ['计划结束日期', '延期时间', '实际结题时间']:
                    if cleaned_value:
                        try:
                            df.loc[idx, key] = pd.to_datetime(cleaned_value).strftime('%Y-%m-%d')
                        except (ValueError, TypeError):
                            print(f"警告: 更新时日期字段 '{key}' 的值 '{cleaned_value}' 格式无效，已清空。")
                            df.loc[idx, key] = ''
                    else:
                        df.loc[idx, key] = ''

                else:
                    df.loc[idx, key] = str(cleaned_value or '')

            else:
                print(f"警告: 尝试更新的字段 '{key}' 不存在于数据表中，已忽略。")

        if start_date_changed:
            start_date_str = df.loc[idx, '开始日期']
            if start_date_str:
                try:
                    start_dt = datetime.strptime(start_date_str, '%Y-%m-%d')
                    df.loc[idx, '开始年份'] = start_dt.year
                except ValueError:
                    df.loc[idx, '开始年份'] = ''
            else:
                df.loc[idx, '开始年份'] = ''

        if budget_changed:
            ext_fund = float(df.loc[idx, '外部专项经费'] or 0)
            inst_fund = float(df.loc[idx, '院自筹经费'] or 0)
            dept_fund = float(df.loc[idx, '所属单位自筹经费'] or 0)
            df.loc[idx, '总预算'] = ext_fund + inst_fund + dept_fund

        print(f"课题 '{project_id_str}' 的信息已更新。")
        return df, True, folder_path
    except Exception as e:
        print(f"更新课题 '{project_id_str}' 信息时发生严重错误: {e}")
        import traceback
        traceback.print_exc()
        return df, False, folder_path

def delete_project_record(df, project_id):
    """从 DataFrame 删除指定课题的记录 (不处理文件夹)"""
    project_id_str = str(project_id)
    initial_len = len(df)
    df = df[df['课题编号'].astype(str) != project_id_str]
    if len(df) < initial_len:
        print(f"课题 '{project_id_str}' 的记录已从数据表中删除。")
        if '序号' in df.columns:
            df['序号'] = range(1, len(df) + 1)
        return df, True
    else:
        print(f"错误: 找不到课题编号 '{project_id_str}'，无法删除记录。")
        return df, False

def get_project_folder_path(df, project_id, folder_cache):
    """获取指定课题的文件夹路径，从缓存中获取"""
    project_id_str = str(project_id)
    folder_path = folder_cache.get(project_id_str)
    if folder_path and os.path.isdir(folder_path):
        return folder_path
    print(f"信息: 课题 '{project_id_str}' 的文件夹路径未在缓存中找到或无效。")
    return None