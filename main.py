import os
import yaml

from email_processor import EmailProcessor


def main():
    # 环境配置 导入config
    try:
        with open('config.yaml', 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
    except FileNotFoundError:
        raise FileNotFoundError("没有找到配置文件 config.yaml 请参考 config.example.yaml 创建配置文件")
    
    # 使用配置中的数据
    # course_name = config['course_name']
    # assignment_name = config['assignment_name']
    email_dir = config['email_dir']
    output_dir = config['output_dir']
    # processed_log_path = config['processed_log_path']
    # roster_config = config['roster_config']
    
    # 确保邮件目录和输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(email_dir, exist_ok=True)

    # 创建并运行 EmailProcessor
    processor = EmailProcessor(config)
    processor.process_emails()
    processor.generate_report()

if __name__ == "__main__":
    main()