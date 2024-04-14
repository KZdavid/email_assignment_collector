# 作业邮件收集脚本

这个项目是一个作业邮件收集脚本，用于处理某课程学生通过邮件提交的作业。

## 使用方法

1. 创建一个 [`config.yaml`] 配置文件，参考 [`config.example.yaml`] 文件。在配置文件中，你需要指定以下参数：
   - [`course_name`]\: 课程名称
   - [`assignment_name`]\: 作业名称
   - [`email_dir`]\: 存放邮件的目录
   - [`output_dir`]\: 输出目录
   - [`processed_log_path`]\: 处理过的邮件的日志文件路径
   - [`roster_config`]\: 学生名单文件设置
      - [`path`]：学生名单文件路径
      - [`student_id_column`]\: 学号所在列的（Excel列的字母表示）
      - [`name_column`]\: 学生姓名所在的列
      - [`start_row`]\: 点名册开始读取的行号
   
   例如，假设学生名单在`"data/点名册.xlsx"`，第一个学生的学号在B3单元格，姓名在C3，则
   - [`roster_config`]\: "data/点名册.xlsx"
      - [`student_id_column`]\: 'B'
      - [`name_column`]\: 'C'
      - [`start_row`]\: 3
2. 通过邮箱下载所有学生某次作业的`.eml`格式的作业邮件到[`email_dir`]
3. 运行 [`main.py`] 文件来处理邮件。处理过程包括提取邮件中的附件，将邮件存档，以及生成处理报告

## 做了什么
1. 读取学生名单里的学号、姓名信息，在邮件的主题、附件文件名里匹配，如果同时匹配到了对应学生的两个信息，则将该邮件视为该学生此次提交的作业，并把该邮件移动到输出目录[`output_dir`]下的[`邮件存档`]文件夹里，并将邮件按照[课程名称 - 作业名称 - 学号 - 姓名]的格式重命名
2. 下载该邮件的附件并存到[`output_dir`]下的[`附件`]文件夹里，按照[课程名称 - 作业名称 - 学号 - 姓名]的格式创建文件夹，把该学生的所有附件下载到该文件夹里
3. 把处理得到的信息存入日志文件[`processed_log_path`]，包括所有下载附件的路径
4. 生成一份[课程名称 - 作业名称 - 提交情况.xlsx]根据学生名单信息，增加名为“是否提交作业”的列，并列出作业邮件路径和附件路径

## 测试

你可以运行 [``tests/test_email_processor.py``] 文件来进行单元测试。

## 注意事项

- 请确保邮件目录和输出目录存在，否则程序将无法运行。
- 请确保邮件文件是 `.eml` 格式。
- 请确保学生名单文件是 `.xlsx` 格式，并且包含学生的姓名和学号的列。
