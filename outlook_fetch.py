import win32com.client
import os
import re

def clean_filename(name):
    """
    移除文件名中的非法字符，避免保存文件时出错。
    """
    return re.sub(r'[\\/*?:"<>|]', "", name)

def list_all_emails():
    """
    列出收件箱中所有邮件的标题和基本信息
    """
    try:
        # 连接到Outlook应用程序
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        
        print(f"正在扫描收件箱 '{inbox.Name}'...")
        
        # 获取所有邮件
        messages = inbox.Items
        
        # 将所有邮件转换为列表以获得准确计数
        email_list = []
        for message in messages:
            try:
                email_info = {
                    'subject': message.Subject if message.Subject else "(无主题)",
                    'sender': message.SenderName if hasattr(message, 'SenderName') else "(未知发件人)",
                    'received_time': message.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S") if hasattr(message, 'ReceivedTime') else "(无时间信息)"
                }
                email_list.append(email_info)
            except Exception as e:
                email_list.append({
                    'subject': f"(读取失败: {str(e)})",
                    'sender': "(未知)",
                    'received_time': "(未知)"
                })
        
        print(f"\n总共找到 {len(email_list)} 封邮件：")
        print("=" * 80)
        
        for i, email in enumerate(email_list, 1):
            clean_subject = clean_filename(email['subject'])
            print(f"{i:3d}. [{email['received_time']}] {clean_subject}")
            print(f"     发件人: {email['sender']}")
            print("-" * 80)
        
        print(f"\n扫描完成！收件箱共有 {len(email_list)} 封邮件。")
        return len(email_list)
        
    except Exception as e:
        print(f"扫描邮件时发生错误: {e}")
        return 0

def export_emails_to_txt():
    # 先列出所有邮件
    total_emails = list_all_emails()
    
    if total_emails == 0:
        print("没有找到邮件，退出导出程序。")
        return
    
    # 询问用户是否继续导出
    user_input = input(f"\n发现 {total_emails} 封邮件，是否继续导出到TXT文件？(y/n): ")
    if user_input.lower() not in ['y', 'yes', '是']:
        print("取消导出。")
        return
    
    # 指定保存TXT文件的文件夹路径
    output_dir = os.path.join(os.getcwd(), "Outlook_Emails_TXT")

    # 如果文件夹不存在，则创建它
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    try:
        # 重新连接到Outlook应用程序
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)

        print(f"\n开始导出邮件到: {output_dir}")

        # 获取文件夹中的所有邮件
        messages = inbox.Items
        
        # 将消息转换为列表以避免计数问题
        message_list = list(messages)
        total_count = len(message_list)

        # 遍历每一封邮件
        for i, message in enumerate(message_list):
            try:
                # 邮件主题
                subject = message.Subject if message.Subject else "(无主题)"
                # 邮件正文
                body = message.Body if hasattr(message, 'Body') else "(无正文内容)"
                # 邮件接收时间
                received_time = message.ReceivedTime.strftime("%Y-%m-%d_%H-%M-%S") if hasattr(message, 'ReceivedTime') else "未知时间"

                # 创建一个安全的文件名 (格式：时间-主题.txt)
                safe_subject = clean_filename(subject)
                if len(safe_subject) > 100: # 限制文件名长度
                    safe_subject = safe_subject[:100]

                filename = f"{received_time}-{safe_subject}.txt"
                filepath = os.path.join(output_dir, filename)

                # 将邮件正文写入TXT文件，使用UTF-8编码
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(f"主题: {subject}\n")
                    f.write(f"时间: {message.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S') if hasattr(message, 'ReceivedTime') else '未知时间'}\n")
                    f.write(f"发件人: {message.SenderName if hasattr(message, 'SenderName') else '未知发件人'}\n")
                    f.write("--------------------正文--------------------\n\n")
                    f.write(body)

                print(f"({i+1}/{total_count}) 成功导出: {filename}")

            except Exception as e:
                print(f"处理第 {i+1} 封邮件时出错: {e}")

        print(f"\n导出完成！所有文件保存在: {output_dir}")

    except Exception as e:
        print(f"连接Outlook或读取邮件时发生错误: {e}")
        print("请确保Outlook正在运行，并且脚本有权限访问它。")

if __name__ == '__main__':
    # 先运行邮件列表功能
    export_emails_to_txt()