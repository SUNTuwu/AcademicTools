import os

# --- 配置区 ---
# 1. 设置包含txt文件的文件夹路径
#    可以直接写绝对路径，例如: r"C:\Users\YourName\Documents\MyTxtFiles"
#    如果文件夹和这个python脚本在同一个目录下，可以直接写文件夹名，例如: "Outlook_Emails_TXT"
folder_path = "Outlook_Emails_TXT"

# 2. 设置合并后输出的文件名
output_file_path = "merged_content.txt"
# --- 配置区结束 ---

def merge_txt_files(source_folder, output_file):
    """
    合并指定文件夹下所有txt文件的内容到一个输出文件中。
    """
    # 检查源文件夹是否存在
    if not os.path.isdir(source_folder):
        print(f"错误：文件夹 '{source_folder}' 不存在。")
        return

    print(f"开始从文件夹 '{source_folder}' 读取txt文件...")

    # 获取文件夹下所有文件的列表
    try:
        all_files = os.listdir(source_folder)
    except OSError as e:
        print(f"错误：无法访问文件夹 '{source_folder}'。原因: {e}")
        return

    # 筛选出所有以 .txt 结尾的文件
    txt_files = [f for f in all_files if f.lower().endswith('.txt')]

    if not txt_files:
        print("在该文件夹下没有找到任何 .txt 文件。")
        return

    print(f"找到了 {len(txt_files)} 个txt文件，准备合并...")

    try:
        # 使用 'w' (写入) 模式打开输出文件，如果文件已存在则会覆盖
        # 使用 utf-8 编码以支持多种语言
        with open(output_file, 'w', encoding='utf-8') as outfile:
            # 遍历所有找到的txt文件
            for i, filename in enumerate(txt_files):
                file_path = os.path.join(source_folder, filename)
                
                # 添加一个清晰的分隔符，并注明源文件名
                separator = f"\n\n{'='*20} [内容来源: {filename}] {'='*20}\n\n"
                outfile.write(separator)
                
                try:
                    # 使用 'r' (读取) 模式打开每个txt文件
                    # 使用 errors='ignore' 来跳过可能出现的编码错误字符
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as infile:
                        # 读取文件内容并写入到输出文件中
                        content = infile.read()
                        outfile.write(content)
                    
                    print(f"({i+1}/{len(txt_files)}) 已合并: {filename}")

                except Exception as e:
                    print(f"读取文件 '{filename}' 时出错: {e}")
                    outfile.write(f"--- 读取文件 {filename} 失败 ---\n")
        
        print(f"\n合并完成！所有内容已保存到文件: '{output_file}'")

    except IOError as e:
        print(f"写入到输出文件 '{output_file}' 时发生错误: {e}")

if __name__ == '__main__':
    merge_txt_files(folder_path, output_file_path)