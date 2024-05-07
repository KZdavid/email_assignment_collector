import os
import shutil
import sys
import unittest

# 添加主模块路径到系统路径，以便可以导入主模块
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from email_processor import EmailProcessor
from email_processor.email_processor import ATTACHMENT_FOLDER, EMAIL_ARCHIVE_FOLDER


class TestEmailProcessor(unittest.TestCase):
    def setUp(self):
        # 测试环境配置
        self.email_dir = os.path.normpath(os.path.join(os.path.dirname(__file__), 'data', 'emails'))
        self.test_email_dir = os.path.normpath(os.path.join(os.path.dirname(__file__), 'data', 'test_emails'))
        self.output_dir = os.path.normpath(os.path.join(os.path.dirname(__file__), 'output'))
        self.processed_log_path = os.path.normpath(os.path.join(os.path.dirname(__file__), 'output', 'processed_emails.json'))
        self.roster_path = os.path.normpath(os.path.join(os.path.dirname(__file__), 'data', 'roster_test.xlsx'))
        self.roster_config = {
            'path': self.roster_path,
            'student_id_column': 'B',
            'name_column': 'C',
            'start_row': 6
        }

        # 确保测试邮件目录和输出目录存在
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.email_dir, exist_ok=True)
        
        # 清理output_dir和email_dir
        if os.path.exists(self.output_dir):
            for file in os.listdir(self.output_dir):
                if os.path.isdir(os.path.join(self.output_dir, file)):
                    shutil.rmtree(os.path.join(self.output_dir, file))
                else:
                    os.remove(os.path.join(self.output_dir, file))
        if os.path.exists(self.email_dir):
            for file in os.listdir(self.email_dir):
                if os.path.isdir(os.path.join(self.email_dir, file)):
                    shutil.rmtree(os.path.join(self.email_dir, file))
                else:
                    os.remove(os.path.join(self.email_dir, file))

        # 复制测试邮件到邮箱目录
        for file_name in os.listdir(self.test_email_dir):
            shutil.copy(os.path.join(self.test_email_dir, file_name), self.email_dir)

        course_name = 'test_course_name'
        assignment_name = 'test_assignment_name'
        
        self.processor = EmailProcessor(course_name,assignment_name,config={
            'email_dir': self.email_dir,
            'output_dir': self.output_dir,
            'processed_log_path': self.processed_log_path,
            'roster_config': self.roster_config
        })

    def tearDown(self):
        pass

    def test_email_processing(self):
        # 统计待处理邮件数量
        email_count = len(os.listdir(self.email_dir))
        # 运行邮件处理
        self.processor.process_emails()
        # 检查是否有输出生成
        self.assertTrue(os.path.exists(self.output_dir))
        self.assertTrue(any(os.listdir(self.output_dir)))
        # 检查处理后的邮件数量是否等于待处理邮件数量
        # 统计.eml文件数量
        output_eml_path = os.path.join(self.output_dir, EMAIL_ARCHIVE_FOLDER)
        processed_count = len([file for file in os.listdir(output_eml_path) if file.endswith('.eml')])
        # 统计附件文件夹数量
        output_attachment_path = os.path.join(self.output_dir, ATTACHMENT_FOLDER)
        processed_folder_count = len([file for file in os.listdir(output_attachment_path) if os.path.isdir(os.path.join(output_attachment_path, file))])
                
        self.assertEqual(processed_count, email_count)
        self.assertEqual(processed_folder_count, email_count)
        # 检查是否有日志生成
        self.assertTrue(os.path.exists(self.processed_log_path))
        # 检查日志中记录的邮件数量是否等于待处理邮件数量
        self.assertEqual(len(self.processor.processed_emails_list), email_count)
        
    def test_skip_processed_emails(self):
        # 运行邮件处理
        self.processor.process_emails()
        # 记录处理后的邮件数量
        output_eml_path = os.path.join(self.output_dir, EMAIL_ARCHIVE_FOLDER)
        processed_count = len([file for file in os.listdir(output_eml_path) if file.endswith('.eml')])
        # 再次运行邮件处理
        self.processor.process_emails()
        # 检查处理后的邮件数量是否没有变化
        self.assertEqual(processed_count, len([file for file in os.listdir(output_eml_path) if file.endswith('.eml')]))        
    
    def test_email_processing_report(self):
        # 运行邮件处理
        self.processor.process_emails()
        # 生成报告
        self.processor.generate_report()
        # 检查是否有报告生成
        self.assertTrue(os.path.exists(os.path.join(self.output_dir, f"{self.processor.course_name} - {self.processor.assignment_name} - 提交情况.xlsx")))

if __name__ == '__main__':
    unittest.main()
