# test_app.py
import os
import time
import uuid
import tempfile

# 测试隔离：必须在导入 app 之前将配置文件与数据库指向临时文件
_TMP_DIR = tempfile.mkdtemp(prefix='exam_test_')
os.environ['APP_CONFIG_FILE'] = os.path.join(_TMP_DIR, 'test_config.ini')
os.environ['STUDENTS_DB'] = os.path.join(_TMP_DIR, 'test_students.db')
os.environ.pop('DEEPSEEK_API_KEY', None)

import pytest
import app as app_module
from app import app, init_db, load_config, get_db_connection
from run import setup_utf8_environment


@pytest.fixture
def client():
    """Flask 测试客户端"""
    app.config['TESTING'] = True
    with app.test_client() as c:
        yield c


def test_utf8_environment_setup():
    """测试UTF-8编码环境设置"""
    setup_utf8_environment()
    assert os.environ.get('PYTHONIOENCODING') == 'utf-8'
    assert os.environ.get('LANG') == 'zh_CN.UTF-8'


def test_config_loading():
    """测试配置文件加载（不存在时自动创建默认配置）"""
    config = load_config()
    assert config is not None
    # 验证默认配置项存在
    assert 'mail_server' in config['DEFAULT']
    assert 'deepseek_api_key' in config['DEFAULT']
    assert 'deepseek_model' in config['DEFAULT']


def test_database_initialization():
    """测试数据库初始化"""
    init_db()
    # 检查数据库文件是否创建
    assert os.path.exists(app_module.DB_PATH)
    # 检查表是否创建成功
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='student_subscriptions'")
    assert cursor.fetchone() is not None
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='practice_records'")
    assert cursor.fetchone() is not None
    conn.close()


def test_database_migration_from_legacy_schema(tmp_path, monkeypatch):
    """测试旧版数据库表结构自动升级"""
    import sqlite3
    legacy_db = str(tmp_path / 'legacy.db')
    conn = sqlite3.connect(legacy_db)
    conn.execute('''CREATE TABLE student_subscriptions
                    (id INTEGER PRIMARY KEY AUTOINCREMENT,
                     email TEXT UNIQUE NOT NULL,
                     name TEXT,
                     subject TEXT,
                     difficulty TEXT,
                     daily_questions INTEGER,
                     is_active BOOLEAN DEFAULT 1)''')
    conn.execute('''CREATE TABLE practice_records
                    (id INTEGER PRIMARY KEY AUTOINCREMENT,
                     email TEXT NOT NULL,
                     practice_date DATE NOT NULL,
                     questions TEXT)''')
    conn.commit()
    conn.close()

    monkeypatch.setattr(app_module, 'DB_PATH', legacy_db)
    app_module.init_db()

    conn = sqlite3.connect(legacy_db)
    columns = {row[1] for row in conn.execute('PRAGMA table_info(student_subscriptions)')}
    record_columns = {row[1] for row in conn.execute('PRAGMA table_info(practice_records)')}
    conn.close()
    assert {'grade', 'knowledge_scope', 'learning_goals',
            'special_requirements', 'preferred_time'} <= columns
    assert 'feedback' in record_columns


def test_flask_app_initialization():
    """测试Flask应用初始化"""
    assert app is not None
    # 测试应用是否处于调试模式（根据run.py配置，应为False）
    assert app.debug is False
    # secret_key 应已持久化配置
    assert app.secret_key


def test_index_page(client):
    """测试首页正常渲染"""
    response = client.get('/')
    assert response.status_code == 200
    assert '智能试卷生成系统'.encode('utf-8') in response.data


def test_student_page(client):
    """测试学生门户页面正常渲染"""
    response = client.get('/student')
    assert response.status_code == 200


def test_admin_config_local_access(client):
    """测试管理页面本机可访问"""
    response = client.get('/admin/config')
    assert response.status_code == 200


def test_admin_config_remote_forbidden(client):
    """测试管理页面拒绝非本机访问"""
    response = client.get('/admin/config', environ_overrides={'REMOTE_ADDR': '203.0.113.5'})
    assert response.status_code == 403


def test_api_send_practice_requires_local(client):
    """测试即时练习API拒绝非本机访问"""
    response = client.post('/api/send_practice',
                           json={'email': 'test@example.com'},
                           environ_overrides={'REMOTE_ADDR': '203.0.113.5'})
    assert response.status_code == 403


def test_subscribe_rejects_invalid_email(client):
    """测试订阅时非法邮箱被拒绝"""
    response = client.post('/student', data={
        'action': 'subscribe',
        'email': 'not-an-email',
        'name': '测试',
        'subject': '数学',
    }, follow_redirects=True)
    assert response.status_code == 200
    conn = get_db_connection()
    row = conn.execute(
        "SELECT * FROM student_subscriptions WHERE email = ?", ('not-an-email',)
    ).fetchone()
    conn.close()
    assert row is None


def test_subscribe_and_unsubscribe(client):
    """测试订阅与取消订阅流程"""
    email = 'student@example.com'
    response = client.post('/student', data={
        'action': 'subscribe',
        'email': email,
        'name': '小明',
        'grade': '7',
        'subject': '数学',
        'difficulty': '中等',
        'daily_questions': '5',
        'preferred_time': '19:00',
    }, follow_redirects=True)
    assert response.status_code == 200

    conn = get_db_connection()
    row = conn.execute(
        "SELECT * FROM student_subscriptions WHERE email = ?", (email,)
    ).fetchone()
    assert row is not None
    assert row['is_active'] == 1
    assert row['grade'] == 7
    conn.close()

    # 取消订阅
    client.post('/student', data={'action': 'unsubscribe', 'email': email},
                follow_redirects=True)
    conn = get_db_connection()
    row = conn.execute(
        "SELECT is_active FROM student_subscriptions WHERE email = ?", (email,)
    ).fetchone()
    conn.close()
    assert row['is_active'] == 0


def test_generate_exam_without_api_key():
    """测试未配置API密钥时返回友好错误提示"""
    result = app_module.generate_exam('资料', '考试', '中等', 100, 90, '', '')
    assert result.startswith('生成试卷失败')
    assert 'API' in result


def test_generate_exam_with_mock_client(monkeypatch):
    """测试mock AI客户端时试卷生成正常返回"""

    class FakeMessage:
        content = '# 模拟试卷\n## 考试信息'

    class FakeChoice:
        message = FakeMessage()

    class FakeResponse:
        choices = [FakeChoice()]

    class FakeCompletions:
        def create(self, **kwargs):
            assert kwargs.get('model') == 'deepseek-chat'
            return FakeResponse()

    class FakeChat:
        completions = FakeCompletions()

    class FakeClient:
        chat = FakeChat()

    monkeypatch.setattr(app_module, 'get_ai_client', lambda: FakeClient())
    result = app_module.generate_exam('资料内容', '考试', '中等', 100, 90, '', '')
    assert '模拟试卷' in result


def test_email_validation():
    """测试邮箱格式校验"""
    assert app_module.is_valid_email('user@example.com')
    assert app_module.is_valid_email('stu.name+01@school.edu.cn')
    assert not app_module.is_valid_email('')
    assert not app_module.is_valid_email(None)
    assert not app_module.is_valid_email('invalid-email')
    assert not app_module.is_valid_email('a@b')


def test_time_normalization():
    """测试偏好时间格式规范化"""
    assert app_module._normalize_time('07:00') == '07:00'
    assert app_module._normalize_time('7:05') == '07:05'
    assert app_module._normalize_time('bad-time') == '07:00'
    assert app_module._normalize_time('', '19:00') == '19:00'
    assert app_module._normalize_time('25:00') == '07:00'


def test_download_html_and_docx(client):
    """测试HTML与DOCX下载流程"""
    file_id = str(uuid.uuid4())
    md_path = os.path.join(tempfile.gettempdir(), f'{file_id}.md')
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write('# Exam Title\n## Section One\n1. Question one\n**答案：** A\n')

    try:
        response = client.get(f'/download/html/{file_id}')
        assert response.status_code == 200
        assert 'Exam Title'.encode('utf-8') in response.data

        response = client.get(f'/download/docx/{file_id}')
        assert response.status_code == 200
        # DOCX 文件以 zip 魔数开头
        assert response.data[:2] == b'PK'
    finally:
        os.remove(md_path)


def test_download_pdf(client):
    """测试PDF下载流程（真实PDF输出）"""
    file_id = str(uuid.uuid4())
    md_path = os.path.join(tempfile.gettempdir(), f'{file_id}.md')
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write('# Exam Title\n1. Question one\n')

    try:
        response = client.get(f'/download/pdf/{file_id}')
        assert response.status_code == 200
        assert response.data[:4] == b'%PDF'
    finally:
        os.remove(md_path)


def test_download_missing_file(client):
    """测试下载不存在的文件时重定向到首页"""
    response = client.get('/download/docx/nonexistent-id')
    assert response.status_code == 302


def test_api_generate_requires_material(client):
    """测试生成API缺少资料时返回400"""
    response = client.post('/api/generate', json={'material': ''})
    assert response.status_code == 400


def test_cleanup_temp_files():
    """测试过期临时试卷文件清理"""
    old_file = os.path.join(tempfile.gettempdir(), f'{uuid.uuid4()}.md')
    new_file = os.path.join(tempfile.gettempdir(), f'{uuid.uuid4()}.md')
    for path in (old_file, new_file):
        with open(path, 'w', encoding='utf-8') as f:
            f.write('test')

    # 将 old_file 的修改时间设置为 48 小时前
    old_time = time.time() - 48 * 3600
    os.utime(old_file, (old_time, old_time))

    try:
        app_module.cleanup_temp_files(max_age_hours=24)
        assert not os.path.exists(old_file)
        assert os.path.exists(new_file)
    finally:
        for path in (old_file, new_file):
            if os.path.exists(path):
                os.remove(path)


def test_convert_markdown_to_docx_structure():
    """测试DOCX转换器基本结构（标题/列表/答案不串段）"""
    buffer = app_module.convert_markdown_to_docx(
        '# 标题\n### 一、选择题\n1. 第一题\n**答案：** A\n2. 第二题\n'
    )
    from docx import Document
    from io import BytesIO
    doc = Document(BytesIO(buffer.read()))
    texts = [p.text for p in doc.paragraphs]
    assert '标题' in texts
    assert '第一题' in texts
    assert '答案： A' in texts
    # 答案应为独立段落，而非拼接进上一段
    assert all('答案' not in t or t.startswith('答案') for t in texts)
