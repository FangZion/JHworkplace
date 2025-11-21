import os
import pandas as pd
import glob
import sys

def get_files_in_directory():
    """获取脚本所在文件夹内的所有数据文件"""
    # 获取脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 支持的文件格式
    file_patterns = ['*.csv', '*.xlsx', '*.xls']
    all_files = []
    
    for pattern in file_patterns:
        all_files.extend(glob.glob(os.path.join(script_dir, pattern)))
    
    # 过滤掉输出文件（如果已存在）
    output_file = os.path.join(script_dir, '已去重.xlsx')
    all_files = [f for f in all_files if '已去重' not in f]
    
    return script_dir, all_files

def display_file_choices(files):
    """显示文件选择菜单"""
    print("\n发现以下数据文件:")
    print("-" * 50)
    for i, file in enumerate(files, 1):
        file_name = os.path.basename(file)
        file_size = os.path.getsize(file) / 1024  # KB
        print(f"{i}. {file_name} ({file_size:.1f} KB)")
    print("-" * 50)
    
    while True:
        try:
            choice = int(input("请选择要处理的文件编号: "))
            if 1 <= choice <= len(files):
                return files[choice - 1]
            else:
                print(f"请输入 1-{len(files)} 之间的数字")
        except ValueError:
            print("请输入有效的数字")

def get_duplicate_columns(df):
    """获取数据框的列供用户选择（使用abcd字母编号）"""
    print("\n可用的列:")
    print("-" * 30)
    
    # 生成字母编号（a, b, c, ...）
    letters = [chr(97 + i) for i in range(len(df.columns))]  # 97是'a'的ASCII码
    
    for i, (letter, col) in enumerate(zip(letters, df.columns)):
        print(f"{letter}. {col}")
    print("-" * 30)
    
    while True:
        try:
            choice = input("请选择用于去重的列字母(多个列用逗号分隔，如 a,c): ").strip().lower()
            selected_letters = [x.strip() for x in choice.split(',') if x.strip()]
            
            # 验证选择
            valid_letters = all(letter in letters for letter in selected_letters)
            if valid_letters and selected_letters:
                selected_columns = []
                for letter in selected_letters:
                    index = ord(letter) - 97  # 将字母转换为索引
                    selected_columns.append(df.columns[index])
                
                print(f"选择的去重列: {', '.join(selected_columns)}")
                return selected_columns
            else:
                available_letters = ', '.join(letters)
                print(f"请输入有效的字母，可用选项: {available_letters}")
        except Exception as e:
            print(f"输入错误: {e}")
            available_letters = ', '.join(letters)
            print(f"请输入有效的字母，如: a 或 a,b,c，可用选项: {available_letters}")

def load_file(file_path):
    """根据文件类型加载数据"""
    file_ext = os.path.splitext(file_path)[1].lower()
    
    try:
        if file_ext == '.csv':
            # 尝试不同的编码
            encodings = ['utf-8', 'gbk', 'gb2312', 'latin1']
            for encoding in encodings:
                try:
                    df = pd.read_csv(file_path, encoding=encoding)
                    print(f"成功读取文件，使用编码: {encoding}")
                    return df
                except UnicodeDecodeError:
                    continue
            # 如果所有编码都失败，使用默认编码并忽略错误
            df = pd.read_csv(file_path, encoding='utf-8', errors='ignore')
            print("使用utf-8编码并忽略错误字符")
            return df
            
        elif file_ext in ['.xlsx', '.xls']:
            df = pd.read_excel(file_path)
            print("成功读取Excel文件")
            return df
            
    except Exception as e:
        print(f"读取文件时出错: {e}")
        return None

def remove_duplicates(df, columns):
    """根据指定列去除重复值，保留第一次出现的行"""
    print(f"\n去重前数据形状: {df.shape}")
    print(f"去重前总行数: {len(df)}")
    
    # 检查重复行
    duplicates = df.duplicated(subset=columns, keep='first')
    duplicate_count = duplicates.sum()
    
    print(f"发现重复行数量: {duplicate_count}")
    
    if duplicate_count > 0:
        # 显示一些重复的示例
        duplicate_samples = df[duplicates].head(3)
        print("\n重复行示例:")
        for col in columns:
            if col in duplicate_samples.columns:
                sample_values = duplicate_samples[col].head(3).tolist()
                print(f"  {col}: {sample_values}")
    
    # 去除重复，保留第一次出现的行
    df_cleaned = df.drop_duplicates(subset=columns, keep='first')
    
    print(f"去重后数据形状: {df_cleaned.shape}")
    print(f"去重后总行数: {len(df_cleaned)}")
    print(f"移除重复行数量: {duplicate_count}")
    
    return df_cleaned

def save_result(df, script_dir, original_filename):
    """保存去重后的结果为Excel文件"""
    # 生成输出文件名
    original_name = os.path.splitext(original_filename)[0]
    output_filename = f"已去重_{original_name}.xlsx"
    output_path = os.path.join(script_dir, output_filename)
    
    try:
        # 保存为Excel文件
        df.to_excel(output_path, index=False, engine='openpyxl')
        print(f"\n去重结果已保存为: {output_filename}")
        print(f"文件路径: {output_path}")
        return True
    except Exception as e:
        print(f"保存Excel文件时出错: {e}")
        
        # 如果保存Excel失败，尝试保存为CSV
        try:
            csv_filename = f"已去重_{original_name}.csv"
            csv_path = os.path.join(script_dir, csv_filename)
            df.to_csv(csv_path, index=False, encoding='utf-8-sig')
            print(f"已改为保存为CSV文件: {csv_filename}")
            return True
        except Exception as e2:
            print(f"保存CSV文件也失败: {e2}")
            return False

def main():
    """主函数"""
    print("=" * 60)
    print("            CSV/Excel 文件去重工具")
    print("=" * 60)
    
    # 获取文件列表
    script_dir, files = get_files_in_directory()
    
    if not files:
        print("在脚本所在文件夹中未找到任何CSV或Excel文件")
        print("请确保文件位于同一文件夹中")
        input("按回车键退出...")
        return
    
    # 选择文件
    selected_file = display_file_choices(files)
    file_name = os.path.basename(selected_file)
    print(f"\n已选择文件: {file_name}")
    
    # 加载文件
    print("正在读取文件...")
    df = load_file(selected_file)
    
    if df is None:
        print("无法读取文件，请检查文件格式或权限")
        input("按回车键退出...")
        return
    
    # 显示数据基本信息
    print(f"\n文件基本信息:")
    print(f"总行数: {len(df)}")
    print(f"总列数: {len(df.columns)}")
    print(f"列名: {list(df.columns)}")
    
    # 选择去重列（使用字母选择）
    duplicate_columns = get_duplicate_columns(df)
    
    # 验证选择的列是否存在空值
    missing_info = {}
    for col in duplicate_columns:
        missing_count = df[col].isnull().sum()
        if missing_count > 0:
            missing_info[col] = missing_count
    
    if missing_info:
        print("\n警告: 以下去重列存在空值:")
        for col, count in missing_info.items():
            print(f"  {col}: {count} 个空值")
        print("包含空值的行在去重时可能会被单独处理")
    
    # 确认操作
    confirm = input("\n确认开始去重？(y/n): ").strip().lower()
    if confirm not in ['y', 'yes', '是']:
        print("操作已取消")
        input("按回车键退出...")
        return
    
    # 执行去重
    print("\n开始去重处理...")
    df_cleaned = remove_duplicates(df, duplicate_columns)
    
    if len(df_cleaned) == len(df):
        print("\n未发现重复数据，文件保持不变")
        # 即使没有重复，也保存文件以便用户查看
        save_anyway = input("是否仍然保存文件？(y/n): ").strip().lower()
        if save_anyway not in ['y', 'yes', '是']:
            print("操作已取消")
            input("按回车键退出...")
            return
    else:
        print("\n发现重复数据，正在处理...")
    
    # 保存结果
    success = save_result(df_cleaned, script_dir, file_name)
    if success:
        print("\n去重完成！")
    else:
        print("\n去重失败，请检查文件权限或磁盘空间")
    
    # 显示统计信息
    print("\n统计信息:")
    print(f"原始文件: {len(df)} 行")
    print(f"去重后: {len(df_cleaned)} 行")
    print(f"移除: {len(df) - len(df_cleaned)} 行")
    
    input("\n按回车键退出...")

if __name__ == "__main__":
    main()