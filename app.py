#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import re
import sys
import locale
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, jsonify
from openai import OpenAI
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from dotenv import load_dotenv
import markdown
from io import BytesIO, StringIO
import uuid
import tempfile
import json
from datetime import datetime, timedelta
import sqlite3
from flask_mail import Mail, Message
from apscheduler.schedulers.background import BackgroundScheduler
import threading
import time
import configparser

# 设置系统编码为UTF-8
if sys.platform.startswith('win'):
    # Windows系统设置UTF-8编码
    os.environ['PYTHONIOENCODING'] = 'utf-8'
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')

# 设置默认编码
try:
    locale.setlocale(locale.LC_ALL, 'zh_CN.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_ALL, 'Chinese_China.UTF-8')
    except locale.Error:
        pass

# 加载环境变量
load_dotenv()

app = Flask(__name__)
app.secret_key = os.urandom(24)

# 配置文件管理
def load_config():
    """加载配置文件"""
    config = configparser.ConfigParser()
    config_file = 'config.ini'
    
    if not os.path.exists(config_file):
        # 创建默认配置文件
        config['DEFAULT'] = {
            'mail_server': 'smtp.gmail.com',
            'mail_port': '587',
            'mail_use_tls': 'True',
            'mail_username': '',
            'mail_password': '',
            'mail_default_sender': '',
            'deepseek_api_key': '',
            'daily_practice_time': '07:00'
        }
        with open(config_file, 'w', encoding='utf-8') as f:
            config.write(f)
    
    config.read(config_file, encoding='utf-8')
    return config

# 加载配置
config = load_config()

# 邮件配置
app.config['MAIL_SERVER'] = config.get('DEFAULT', 'mail_server', fallback='smtp.gmail.com')
app.config['MAIL_PORT'] = config.getint('DEFAULT', 'mail_port', fallback=587)
app.config['MAIL_USE_TLS'] = config.getboolean('DEFAULT', 'mail_use_tls', fallback=True)
app.config['MAIL_USERNAME'] = config.get('DEFAULT', 'mail_username', fallback='')
app.config['MAIL_PASSWORD'] = config.get('DEFAULT', 'mail_password', fallback='')
app.config['MAIL_DEFAULT_SENDER'] = config.get('DEFAULT', 'mail_default_sender', fallback='')

mail = Mail(app)

# 初始化DeepSeek客户端
deepseek_api_key = config.get('DEFAULT', 'deepseek_api_key', fallback='')
# 从环境变量或配置中获取代理
# 假设你从配置中读取代理
client = OpenAI(
    api_key=config['DEFAULT']['deepseek_api_key']
)

# 数据库初始化（改进UTF-8支持）
def init_db():
    """初始化数据库，确保UTF-8编码支持"""
    # 创建数据库连接，确保UTF-8编码
    conn = sqlite3.connect('students.db', 
                          detect_types=sqlite3.PARSE_DECLTYPES,
                          check_same_thread=False)
    
    # 设置数据库连接的编码
    conn.execute("PRAGMA encoding = 'UTF-8'")
    conn.execute("PRAGMA foreign_keys = ON")
    
    c = conn.cursor()
    
    # 学生订阅表（添加详细信息字段）
    c.execute('''CREATE TABLE IF NOT EXISTS student_subscriptions
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  email TEXT UNIQUE NOT NULL,
                  name TEXT,
                  grade INTEGER DEFAULT 7,
                  subject TEXT,
                  knowledge_scope TEXT,
                  learning_goals TEXT,
                  special_requirements TEXT,
                  difficulty TEXT DEFAULT '中等',
                  daily_questions INTEGER DEFAULT 5,
                  preferred_time TEXT DEFAULT '07:00',
                  subscription_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                  is_active BOOLEAN DEFAULT 1)''')
    
    # 练习记录表
    c.execute('''CREATE TABLE IF NOT EXISTS practice_records
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  email TEXT NOT NULL,
                  practice_date DATE NOT NULL,
                  questions TEXT,
                  score INTEGER,
                  completed BOOLEAN DEFAULT 0,
                  feedback TEXT,
                  FOREIGN KEY (email) REFERENCES student_subscriptions (email))''')
    
    conn.commit()
    conn.close()

# 改进的数据库连接函数
def get_db_connection():
    """获取数据库连接，确保UTF-8支持"""
    conn = sqlite3.connect('students.db', 
                          detect_types=sqlite3.PARSE_DECLTYPES,
                          check_same_thread=False)
    conn.execute("PRAGMA encoding = 'UTF-8'")
    conn.row_factory = sqlite3.Row
    return conn

# 初始化数据库
init_db()

def generate_daily_practice(subject, difficulty, question_count=5, grade=None, knowledge_scope=None, learning_goals=None):
    """生成每日练习题（支持详细信息定制）"""
    grade_info = f"{grade}年级" if grade else ""
    knowledge_info = f"，知识范围：{knowledge_scope}" if knowledge_scope else ""
    goals_info = f"，学习目标：{learning_goals}" if learning_goals else ""
    
    system_prompt = f"""你是一位专业的{subject}教师，负责为{grade_info}学生生成每日练习题{knowledge_info}{goals_info}。
    
    生成要求：
    1. 难度级别：{difficulty}
    2. 题目数量：{question_count}道
    3. 题目类型：选择题、填空题、简答题混合
    4. 题目内容要符合学生的认知水平和学习进度
    5. 包含详细的答案解析和学习建议
    6. 使用Markdown格式输出
    
    输出格式：
    # {subject}每日练习题 - {datetime.now().strftime('%Y年%m月%d日')}
    
    ## 今日学习目标
    [根据学习目标编写]
    
    ## 练习题
    
    ### 选择题
    1. 题目...
       A. 选项A
       B. 选项B
       C. 选项C
       D. 选项D
       **答案：** A
       **解析：** 详细解析...
       **学习建议：** 相关学习建议...
    
    ### 填空题
    2. 题目...
       **答案：** 正确答案
       **解析：** 详细解析...
       **学习建议：** 相关学习建议...
    
    ### 简答题
    3. 题目...
       **参考答案：** 详细解答
       **学习建议：** 相关学习建议...
    
    ## 今日总结
    [根据练习内容编写总结]
    """
    
    user_prompt = f"""
    请为{grade_info if grade_info else ''}{subject}科目生成{difficulty}难度的每日练习题，共{question_count}道题目。
    {f'学生年级：{grade}年级，请根据该年级的认知水平出题。' if grade else ''}
    {f'重点知识范围：{knowledge_scope}' if knowledge_scope else ''}
    {f'学习目标：{learning_goals}' if learning_goals else ''}
题目应该：
    1. 覆盖该科目的重点知识点
    2. 难度适中，适合学生日常练习
    3. 包含详细的解析和学习建议
    4. 帮助学生巩固知识和提高能力
    """
    
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            stream=False
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"生成练习题失败: {str(e)}"

def send_daily_practice_email():
    """发送每日练习题邮件（支持详细信息定制）"""
    conn = sqlite3.connect('students.db')
    c = conn.cursor()
    
    # 获取所有活跃订阅的学生（包含详细信息）
    c.execute('''SELECT email, name, grade, subject, knowledge_scope, learning_goals, 
                        difficulty, daily_questions 
                 FROM student_subscriptions 
                 WHERE is_active = 1''')
    students = c.fetchall()
    
    for student in students:
        email, name, grade, subject, knowledge_scope, learning_goals, difficulty, daily_questions = student
        
        # 生成每日练习题（使用详细信息）
        practice_content = generate_daily_practice(
            subject, difficulty, daily_questions, grade, knowledge_scope, learning_goals
        )
        practice_html = markdown.markdown(practice_content)
        
        # 发送邮件
        try:
            msg = Message(
                subject=f"{subject}每日练习题 - {datetime.now().strftime('%Y年%m月%d日')}",
                recipients=[email],
                html=f"""
                <html>
                <body>
                    <h2>亲爱的{name}同学：</h2>
                    <p>这是为您定制的{subject}每日练习题，请认真完成！</p>
                    {f'<p><strong>年级：</strong>{grade}年级</p>' if grade else ''}
                    {f'<p><strong>知识范围：</strong>{knowledge_scope}</p>' if knowledge_scope else ''}
                    {f'<p><strong>学习目标：</strong>{learning_goals}</p>' if learning_goals else ''}
                    <div style="background-color: #f8f9fa; padding: 20px; border-radius: 10px;">
                        {practice_html}
</div>
                    <p style="color: #666; font-size: 12px; margin-top: 20px;">
                        此邮件由智能试卷生成系统自动发送，如需取消订阅请访问系统设置。
                    </p>
                </body>
                </html>
                """
            )
            mail.send(msg)
            
            # 记录发送历史
            c.execute('''INSERT INTO practice_records (email, practice_date, questions) 
                         VALUES (?, ?, ?)''', 
                     (email, datetime.now().date(), practice_content))
            conn.commit()
            
            print(f"已发送个性化练习题给 {email}")
        except Exception as e:
            print(f"发送邮件失败 {email}: {str(e)}")
    
    conn.close()

def schedule_daily_practice():
    """安排每日练习发送任务"""
    scheduler = BackgroundScheduler()
    
    # 从配置文件中获取发送时间
    practice_time = config.get('DEFAULT', 'daily_practice_time', fallback='07:00')
    hour, minute = map(int, practice_time.split(':'))
    
    scheduler.add_job(
        func=send_daily_practice_email,
        trigger='cron',
        hour=hour,
        minute=minute,
        id='daily_practice'
    )
    
    scheduler.start()

# 启动定时任务
try:
    schedule_daily_practice()
    print(f"每日练习发送任务已启动，发送时间：{config.get('DEFAULT', 'daily_practice_time', fallback='07:00')}")
except Exception as e:
    print(f"定时任务启动失败: {str(e)}")

def generate_exam(material, exam_type, difficulty, total_score, time_limit, question_types, instructions):
    """调用DeepSeek API生成试卷"""
    system_prompt = """你是一位专业的试卷生成专家。请根据提供的资料和配置参数生成一份高质量试卷。
    
    试卷生成要求：
    1. 根据试卷类型、难度级别、总分和考试时间合理分配题目
    2. 题目类型多样，包括选择题、填空题、简答题、应用题等
    3. 题目难度分布合理，符合指定的难度级别
    4. 包含清晰的题目描述、分值和参考答案
    5. 使用规范的Markdown格式输出
    
    输出格式要求：
    # 试卷名称
    ## 考试信息（类型、难度、总分、时间）
    ### 一、选择题（每题分值）
    1. 题目内容...
       A. 选项A
       B. 选项B
       C. 选项C
       D. 选项D
       **答案：** A
       
    ### 二、填空题
    1. 题目内容...
       **答案：** 正确答案
       
    ### 三、简答题
    1. 题目内容...
       **参考答案：** 详细解答
       
    ## 参考答案汇总
    """
    
    # 构建详细的用户提示
    user_prompt = f"""
    试卷生成配置：
    - 试卷类型：{exam_type}
    - 难度级别：{difficulty}
    - 总分：{total_score}分
    - 考试时间：{time_limit}分钟
    - 题目类型要求：{question_types if question_types else '自动分配'}
    - 自定义指令：{instructions if instructions else '无'}
    
    资料内容：
    {material}
    
    请根据以上配置和资料生成一份完整的试卷，严格按照Markdown格式输出。
    """
    
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            stream=False
        )
        return response.choices[0].message.content
    except Exception as e:
        app.logger.error(f"API调用错误: {str(e)}")
        return f"生成试卷失败: {str(e)}"

def convert_markdown_to_docx(markdown_content):
    """将Markdown内容转换为DOCX文档"""
    doc = Document()
    
    # 设置文档样式
    styles = doc.styles
    title_style = styles['Heading 1']
    title_style.font.size = Pt(18)
    title_style.font.bold = True
    
    heading_style = styles['Heading 2']
    heading_style.font.size = Pt(14)
    heading_style.font.bold = True
    
    # 处理Markdown内容
    lines = markdown_content.split('\n')
    current_paragraph = None
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        if line.startswith('# '):
            # 一级标题
            doc.add_heading(line[2:], level=1)
        elif line.startswith('## '):
            # 二级标题
            doc.add_heading(line[3:], level=2)
        elif line.startswith('### '):
            # 三级标题
            doc.add_heading(line[4:], level=3)
        elif line.startswith('- ') or line.startswith('* '):
            # 列表项
            if current_paragraph is None:
                current_paragraph = doc.add_paragraph(style='List Bullet')
            else:
                current_paragraph = doc.add_paragraph(style='List Bullet')
            current_paragraph.add_run(line[2:])
        elif line.startswith('1. ') or line.startswith('2. ') or line.startswith('3. '):
            # 有序列表
            if current_paragraph is None:
                current_paragraph = doc.add_paragraph(style='List Number')
            else:
                current_paragraph = doc.add_paragraph(style='List Number')
            current_paragraph.add_run(re.sub(r'^\d+\.\s*', '', line))
        elif line.startswith('**答案：**') or line.startswith('**参考答案：**'):
            # 答案部分
            if current_paragraph is None:
                current_paragraph = doc.add_paragraph()
            current_paragraph.add_run(line).bold = True
        else:
            # 普通段落
            if current_paragraph is None:
                current_paragraph = doc.add_paragraph()
            else:
                current_paragraph.add_run(' ')
            current_paragraph.add_run(line)
    
    # 保存到内存
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def convert_markdown_to_pdf(markdown_content):
    """将Markdown内容转换为PDF（简化版，实际需要安装weasyprint或其他PDF库）"""
    # 这里先返回HTML内容，实际部署时需要安装PDF生成库
    html_content = markdown.markdown(markdown_content)
    buffer = StringIO()
    buffer.write(f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            body {{ font-family: Arial, sans-serif; line-height: 1.6; }}
            h1 {{ color: #2c3e50; }}
            h2, h3 {{ color: #34495e; }}
        </style>
    </head>
    <body>
        {html_content}
    </body>
    </html>
    """)
    buffer.seek(0)
    return buffer

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # 获取表单数据
        material_file = request.files.get('material')
        material_text = request.form.get('material_text', '')
        exam_type = request.form.get('exam_type', '考试')
        difficulty = request.form.get('difficulty', '中等')
        total_score = request.form.get('total_score', '100')
        time_limit = request.form.get('time_limit', '90')
        question_types = request.form.get('question_types', '')
        instructions = request.form.get('instructions', '')
        
        # 处理资料内容（改进UTF-8编码处理）
        material = ""
        if material_file and material_file.filename != '':
            try:
                if material_file.filename.endswith('.txt') or material_file.filename.endswith('.md'):
                    # 尝试多种编码格式读取文件
                    file_content = material_file.read()
                    encodings = ['utf-8', 'gbk', 'gb2312', 'latin1']
                    
                    for encoding in encodings:
                        try:
                            material = file_content.decode(encoding)
                            break
                        except UnicodeDecodeError:
                            continue
                    else:
                        # 如果所有编码都失败，使用utf-8并忽略错误
                        material = file_content.decode('utf-8', errors='ignore')
                else:
                    flash("目前仅支持TXT和MD格式文件")
                    return redirect(url_for('index'))
            except Exception as e:
                flash(f"文件读取失败: {str(e)}")
                return redirect(url_for('index'))
        
        # 如果没有上传文件，使用文本框内容
        if not material and material_text:
            material = material_text
        
        if not material:
            flash("请提供资料内容（上传文件或输入文本）")
            return redirect(url_for('index'))
        
        # 验证数值参数
        try:
            total_score = int(total_score)
            time_limit = int(time_limit)
            if total_score < 10 or total_score > 200:
                flash("总分应在10-200分之间")
                return redirect(url_for('index'))
            if time_limit < 10 or time_limit > 300:
                flash("考试时间应在10-300分钟之间")
                return redirect(url_for('index'))
        except ValueError:
            flash("请输入有效的数值参数")
            return redirect(url_for('index'))
        
        # 生成试卷
        exam_content = generate_exam(material, exam_type, difficulty, total_score, time_limit, question_types, instructions)
        
        # 转换为HTML
        exam_html = markdown.markdown(exam_content)
        
        # 保存内容到临时文件（确保UTF-8编码）
        temp_dir = tempfile.gettempdir()
        file_id = str(uuid.uuid4())
        md_path = os.path.join(temp_dir, f"{file_id}.md")
        
        # 使用UTF-8编码保存文件，并添加BOM以支持Windows系统
        with open(md_path, 'w', encoding='utf-8-sig') as f:
            f.write(exam_content)
        
        return render_template('index.html', 
                             exam_content=exam_html,
                             file_id=file_id)
    
    return render_template('index.html', exam_content=None, file_id=None)

@app.route('/student', methods=['GET', 'POST'])
def student_portal():
    """学生门户页面（支持详细信息输入）"""
    if request.method == 'POST':
        email = request.form.get('email')
        name = request.form.get('name')
        grade = request.form.get('grade')
        subject = request.form.get('subject')
        knowledge_scope = request.form.get('knowledge_scope')
        learning_goals = request.form.get('learning_goals')
        special_requirements = request.form.get('special_requirements')
        difficulty = request.form.get('difficulty', '中等')
        daily_questions = request.form.get('daily_questions', 5)
        preferred_time = request.form.get('preferred_time', '07:00')
        action = request.form.get('action')
        
        conn = get_db_connection()
        c = conn.cursor()
        
        if action == 'subscribe':
            # 订阅每日练习（包含详细信息）
            try:
                c.execute('''INSERT OR REPLACE INTO student_subscriptions 
                             (email, name, grade, subject, knowledge_scope, learning_goals, 
                              special_requirements, difficulty, daily_questions, preferred_time, is_active) 
                             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)''',
                         (email, name, grade, subject, knowledge_scope, learning_goals, 
                          special_requirements, difficulty, int(daily_questions), preferred_time))
                conn.commit()
                flash("订阅成功！您将每天收到个性化练习题。")
            except Exception as e:
                flash(f"订阅失败: {str(e)}")
        
        elif action == 'unsubscribe':
            # 取消订阅
            c.execute("UPDATE student_subscriptions SET is_active = 0 WHERE email = ?", (email,))
            conn.commit()
            flash("已取消订阅每日练习服务")
        
        elif action == 'send_test':
            # 发送测试练习题（使用详细信息）
            practice_content = generate_daily_practice(
                subject, difficulty, int(daily_questions), grade, knowledge_scope, learning_goals
            )
            practice_html = markdown.markdown(practice_content)
            
            try:
                msg = Message(
                    subject=f"{subject}测试练习题",
                    recipients=[email],
                    html=f"""
                    <html>
                    <head>
                        <meta charset="UTF-8">
                    </head>
                    <body>
                        <h2>测试练习题</h2>
                        {f'<p><strong>年级：</strong>{grade}年级</p>' if grade else ''}
                        {f'<p><strong>知识范围：</strong>{knowledge_scope}</p>' if knowledge_scope else ''}
                        <div style="background-color: #f8f9fa; padding: 20px; border-radius: 10px;">
                            {practice_html}
                        </div>
                    </body>
                    </html>
                    """
                )
                mail.send(msg)
                flash("测试练习题已发送到您的邮箱")
            except Exception as e:
                flash(f"发送测试邮件失败: {str(e)}")
        
        conn.close()
        return redirect(url_for('student_portal'))
    
    return render_template('student.html')

@app.route('/admin/config', methods=['GET', 'POST'])
def admin_config():
    """系统配置管理页面"""
    if request.method == 'POST':
        # 更新配置文件
        config['DEFAULT']['mail_server'] = request.form.get('mail_server', '')
        config['DEFAULT']['mail_port'] = request.form.get('mail_port', '587')
        config['DEFAULT']['mail_username'] = request.form.get('mail_username', '')
        config['DEFAULT']['mail_password'] = request.form.get('mail_password', '')
        config['DEFAULT']['mail_default_sender'] = request.form.get('mail_default_sender', '')
        config['DEFAULT']['deepseek_api_key'] = request.form.get('deepseek_api_key', '')
        config['DEFAULT']['daily_practice_time'] = request.form.get('daily_practice_time', '07:00')
        
        with open('config.ini', 'w', encoding='utf-8') as f:
            config.write(f)
        
        flash("配置已保存！需要重启应用使配置生效。")
        return redirect(url_for('admin_config'))
    
    return render_template('admin_config.html', config=config['DEFAULT'])

@app.route('/api/send_practice', methods=['POST'])
def api_send_practice():
    """API接口：立即发送练习题"""
    data = request.get_json()
    email = data.get('email')
    subject = data.get('subject', '数学')
    difficulty = data.get('difficulty', '中等')
    question_count = data.get('question_count', 5)
    
    practice_content = generate_daily_practice(subject, difficulty, question_count)
    
    try:
        msg = Message(
            subject=f"{subject}即时练习题",
            recipients=[email],
            html=markdown.markdown(practice_content)
        )
        mail.send(msg)
        return jsonify({'success': True, 'message': '练习题已发送'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/download/<file_type>/<file_id>')
def download(file_type, file_id):
    temp_dir = tempfile.gettempdir()
    md_path = os.path.join(temp_dir, f"{file_id}.md")
    
    if not os.path.exists(md_path):
        flash("文件不存在或已过期")
        return redirect(url_for('index'))
    
    # 使用UTF-8编码读取文件
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    if file_type == 'docx':
        buffer = convert_markdown_to_docx(content)
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f'exam_paper_{datetime.now().strftime("%Y年%m月%d日")}.docx',
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    elif file_type == 'html':
        html_content = markdown.markdown(content)
        buffer = StringIO()
        buffer.write(f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <title>试卷 - {datetime.now().strftime("%Y年%m月%d日")}</title>
    <style>
        body {{ font-family: "Microsoft YaHei", "SimHei", sans-serif; line-height: 1.6; }}
        h1, h2, h3 {{ color: #2c3e50; }}
    </style>
</head>
<body>
    {html_content}
</body>
</html>''')
        buffer.seek(0)
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f'exam_paper_{datetime.now().strftime("%Y%m%d_%H%M")}.html',
            mimetype='text/html; charset=utf-8'
        )
    elif file_type == 'pdf':
        buffer = convert_markdown_to_pdf(content)
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f'exam_paper_{datetime.now().strftime("%Y%m%d_%H%M")}.pdf',
            mimetype='application/pdf'
        )
    else:
        flash("不支持的文件类型")
        return redirect(url_for('index'))

@app.route('/api/generate', methods=['POST'])
def api_generate():
    """API接口，用于前端异步调用"""
    data = request.get_json()
    
    material = data.get('material', '')
    exam_type = data.get('exam_type', '考试')
    difficulty = data.get('difficulty', '中等')
    total_score = data.get('total_score', 100)
    time_limit = data.get('time_limit', 90)
    question_types = data.get('question_types', '')
    instructions = data.get('instructions', '')
    
    if not material:
        return jsonify({'error': '请提供资料内容'}), 400
    
    exam_content = generate_exam(material, exam_type, difficulty, total_score, time_limit, question_types, instructions)
    
    return jsonify({
        'exam_content': exam_content,
        'exam_html': markdown.markdown(exam_content)
    })

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)