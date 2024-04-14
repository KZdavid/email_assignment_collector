import copy
import json
import os
import re
import shutil
from datetime import datetime
from email import policy
from email.parser import BytesParser

import pandas as pd

# 定义常量
ATTACHMENT_FOLDER = '附件'
EMAIL_ARCHIVE_FOLDER = '邮件存档'

def excel_col_to_index(col):
    """Convert Excel-style column letter to zero-based column index."""
    index = 0
    for char in col:
        index = index * 26 + (ord(char.upper()) - ord('A') + 1)
    return index - 1

class EmailProcessor:
    def __init__(self, course_name ='Course' ,assignment_name='Assignment', config=None):
        '''
        course_name: 课程名称
        assignment_name: 作业名称
        config: 配置信息
        '''
        # 确保配置包含所有必要的项
        if config is None:
            raise ValueError("Config is required.")
        required_keys = ['email_dir', 'output_dir', 'processed_log_path', 'roster_config']
        missing_keys = [key for key in required_keys if key not in config]
        if missing_keys:
            raise ValueError(f"Missing config keys: {', '.join(missing_keys)}")
        
        self.course_name = course_name
        self.assignment_name = assignment_name
        self.config = config
        # 读取名册信息
        self.load_roster()
        # 读取已处理的邮件信息
        self.load_processed_emails()
        
    def load_roster(self):
        roster_config = self.config['roster_config']
        # 检查配置是否包含所有必要的项
        required_keys = ['path', 'student_id_column', 'name_column', 'start_row']
        missing_keys = [key for key in required_keys if key not in roster_config]
        if missing_keys:
            raise ValueError(f"Missing roster config keys: {', '.join(missing_keys)}")
        # 读取名册信息
        usecols = [excel_col_to_index(roster_config['student_id_column']), excel_col_to_index(roster_config['name_column'])]
        # 跳过标题行
        skiprows = roster_config['start_row'] - 1  # 转换为0-based index
        self.roster_df = pd.read_excel(
            roster_config['path'],
            usecols=usecols,
            skiprows=skiprows,
            header=None,
            names=['学号', '姓名'],
            dtype={'学号': str}
        )

    def load_processed_emails(self):
        try:
            with open(self.config['processed_log_path'], 'r', encoding='utf-8') as file:
                self.processed_emails = json.load(file)
        except FileNotFoundError:
            self.processed_emails = {}

    def save_processed_emails(self):
        with open(self.config['processed_log_path'], 'w') as file:
            json.dump(self.processed_emails, file, indent=4)

    def process_emails(self):
        for email_file in os.listdir(self.config['email_dir']):
            if email_file.endswith('.eml'):
                file_path = os.path.join(self.config['email_dir'], email_file)

                subject, attachments, sender, timestamp = self.parse_email(file_path)
                
                # Generate a unique key for each email based on sender, subject, and timestamp
                email_key = f"{sender}-{subject}-{timestamp}"
                
                if email_key in self.processed_emails:
                    continue  # Skip already processed emails

                student_id, name = self.find_student_info(subject + " " + " ".join([a[0] for a in attachments]))
                
                if student_id and name:
                    new_email_path = self.process_email(file_path, student_id, name, attachments)
                    attachment_paths = self.save_attachments(student_id, name, attachments)
                    self.record_email(email_key, student_id, name, new_email_path, attachment_paths)
                else:
                    print(f"No valid student info found in email: {file_path}")
                    
    def record_email(self, email_key, student_id, name, new_email_path, attachment_paths):
        self.processed_emails[email_key] = {
            'student_id': student_id,
            'name': name,
            'email_path': new_email_path,
            'attachments': attachment_paths
        }
        self.save_processed_emails()

    def process_email(self, file_path, student_id, name, attachments):
        new_name = f"{self.course_name} - {self.assignment_name} - {student_id} - {name}.eml"
        new_path = os.path.join(self.config['output_dir'], EMAIL_ARCHIVE_FOLDER, new_name)
        self.save_attachments(student_id, name, attachments)
        # 确保目录存在
        os.makedirs(os.path.dirname(new_path), exist_ok=True)
        shutil.move(file_path, new_path)
        return new_path

    def find_student_info(self, text):
        for _, row in self.roster_df.iterrows():
            student_id = str(row['学号'])
            name = row['姓名']
            # 创建一个正则表达式，确保学号前后不是字母或数字
            student_id_pattern = r'(?<![A-Za-z\d])' + re.escape(student_id) + r'(?!\d)'
            # 搜索学号和姓名
            if re.search(student_id_pattern, text) and re.search(re.escape(name), text):
                return student_id, name
        return None, None

    def save_attachments(self, student_id, name, attachments):
        folder_path = os.path.join(self.config['output_dir'], ATTACHMENT_FOLDER, f"{self.course_name} - {self.assignment_name} - {student_id} - {name}")
        os.makedirs(folder_path, exist_ok=True)
        attachment_paths = []

        for filename, data in attachments:
            save_path = os.path.join(folder_path, filename)
            with open(save_path, 'wb') as f:
                f.write(data)
            attachment_paths.append(save_path)
            print(f"Attachment saved: {save_path}")
            
        return attachment_paths

    def parse_email(self, file_path):
        with open(file_path, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)
        subject = msg['subject']
        sender = msg['from']
        date = msg['date']
        attachments = []
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart' or part.get('Content-Disposition') is None:
                continue
            filename = part.get_filename()
            if filename:
                attachments.append((filename, part.get_payload(decode=True)))
        return subject, attachments, sender, date
    
    import pandas as pd

    def generate_report(self):
        processed_emails = self.processed_emails
        if processed_emails is None:
            raise ValueError("No processed emails found. Please run process_emails() first.")

        # 读取名册信息
        report_df = copy.deepcopy(self.roster_df)

        # 初始化新的列
        report_df['是否提交作业'] = '否'
        report_df['邮件路径'] = ''
        report_df['附件文件夹路径'] = ''

        # 更新名册信息
        for email_info in processed_emails.values():
            student_id = email_info['student_id']
            report_df.loc[report_df['学号'] == student_id, '是否提交作业'] = '是'
            report_df.loc[report_df['学号'] == student_id, '邮件路径'] = email_info['email_path']
            report_df.loc[report_df['学号'] == student_id, '附件文件夹路径'] = os.path.dirname(email_info['attachments'][0]) if email_info['attachments'] else ''

        # 保存更新后的名册信息
        report_path = os.path.join(self.config['output_dir'], f"{self.course_name} - {self.assignment_name} - 提交情况.xlsx")
        report_df.to_excel(report_path, index=False)

# Example usage:
'''
config = {
    'email_dir': '/path/to/emails',
    'output_dir': '/path/to/processed_emails',
    'processed_log_path': '/path/to/processed_emails.json'
}
processor = EmailProcessor(config)
processor.process_emails()
'''