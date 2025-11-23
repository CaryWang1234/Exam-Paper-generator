# test_app.py
import os
import sqlite3
from app import app, init_db, load_config, get_db_connection
from run import setup_utf8_environment

def test_utf8_environment_setup():
    """测试UTF-8编码环境设置"""
    setup_utf8_environment()
    assert os.environ.get('PYTHONIOENCODING') == 'utf-8'
    assert os.environ.get('LANG') == 'zh_CN.UTF-8'

def test_config_loading():
    """测试配置文件加载"""
    config = load_config()
    assert config is not None
    # 验证默认配置项存在
    assert 'mail_server' in config['DEFAULT']
    assert 'deepseek_api_key' in config['DEFAULT']

def test_database_initialization():
    """测试数据库初始化"""
    init_db()
    # 检查数据库文件是否创建
    assert os.path.exists('students.db')
    # 检查表是否创建成功
    conn = get_db_connection()
    cursor = conn.cursor()
    # 检查学生订阅表
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='student_subscriptions'")
    assert cursor.fetchone() is not None
    # 检查练习记录表
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='practice_records'")
    assert cursor.fetchone() is not None
    conn.close()

def test_flask_app_initialization():
    """测试Flask应用初始化"""
    assert app is not None
    # 测试应用是否处于调试模式（根据run.py配置，应为False）
    assert app.debug is False