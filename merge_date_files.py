import os
import sys

# --- 配置区 ---
# 1. 设置源文件夹：从这个文件夹读取原始txt文件
SOURCE_FOLDER = "Outlook_Emails_TXT"

# 2. 设置输出文件夹：合并后的文件将存放在这里
OUTPUT_FOLDER = "Daily-Outlook_Emails_TXT"
# --- 配置区结束 ---

def merge_files_for_date(target_date, source_folder, output_folder):
    """
    从源文件夹中合并文件名以特定日期开头的所有txt文件，
    并将结果保存到输出文件夹中。
    """
    # 1. 检查源文件夹是否存在
    if not os.path.isdir(source_folder):
        print(f"错误：源文件夹 '{source_folder}' 不存在。")
        print("请确保脚本旁边有这个文件夹，并且里面存放了原始txt文件。")
        return

    # 2. 检查并创建输出文件夹（如果不存在）
    try:
        os.makedirs(output_folder, exist_ok=True)
    except OSError as e:
        print(f"错误：无法创建输出文件夹 '{output_folder}'。原因: {e}")
        return

    print(f"正在从文件夹 '{source_folder}' 中搜索日期为 '{target_date}' 的文件...")

    # 根据日期构造输出文件名
    output_filename = f"{target_date}.txt"
    output_file_path = os.path.join(output_folder, output_filename)

    # 获取源文件夹下所有文件的列表
    try:
        all_files = os.listdir(source_folder)
    except OSError as e:
        print(f"错误：无法访问源文件夹 '{source_folder}'。原因: {e}")
        return

    # 筛选出所有以指定日期开头并且是 .txt 结尾的文件
    txt_files_for_date = [f for f in all_files if f.startswith(target_date) and f.lower().endswith('.txt')]

    if not txt_files_for_date:
        print(f"在文件夹 '{source_folder}' 下没有找到以 '{target_date}' 开头的 .txt 文件。")
        return

    # 按文件名排序，确保文件按时间顺序合并
    txt_files_for_date.sort()
    print(f"找到了 {len(txt_files_for_date)} 个文件，准备合并到 '{output_file_path}'...")

    try:
        # 使用 'w' (写入) 模式打开输出文件
        with open(output_file_path, 'w', encoding='utf-8') as outfile:
            # 遍历所有找到的txt文件
            for i, filename in enumerate(txt_files_for_date):
                # 构造每个源文件的完整路径
                file_path = os.path.join(source_folder, filename)
                
                # 添加分隔符
                separator = f"\n\n{'='*20} [内容来源: {filename}] {'='*20}\n\n"
                outfile.write(separator)
                
                try:
                    # 读取源文件内容并写入
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as infile:
                        content = infile.read()
                        outfile.write(content)
                    
                    print(f"({i+1}/{len(txt_files_for_date)}) 已合并: {filename}")

                except Exception as e:
                    print(f"读取文件 '{filename}' 时出错: {e}")
                    outfile.write(f"--- 读取文件 {filename} 失败 ---\n")
            
        print(f"\n合并完成！所有内容已保存到文件: '{output_file_path}'")

    except IOError as e:
        print(f"写入到输出文件 '{output_file_path}' 时发生错误: {e}")

def main():
    """
    主函数，用于获取用户输入并调用合并函数。
    """
    print("--- TXT文件按日期合并工具 ---")
    
    if len(sys.argv) > 1:
        date_to_merge = sys.argv[1]
        print(f"从命令行参数获取到日期: {date_to_merge}")
    else:
        date_to_merge = input("请输入要合并的日期 (格式 YYYY-MM-DD): ")

    if not (len(date_to_merge) == 10 and date_to_merge[4] == '-' and date_to_merge[7] == '-'):
        print("错误：日期格式不正确。请输入类似 '2025-08-08' 的格式。")
        return
        
    # 调用合并函数，传入源文件夹和输出文件夹的路径
    merge_files_for_date(date_to_merge, SOURCE_FOLDER, OUTPUT_FOLDER)

if __name__ == '__main__':
    main()