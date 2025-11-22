# 智能试卷生成系统

基于 AI 技术的专业级试卷生成工具，支持自定义配置、多格式导出和学生个性化练习推送。

[![Python](https://img.shields.io/badge/python-3.7%2B-blue)](https://www.python.org/)
[![Flask](https://img.shields.io/badge/flask-2.3.3-green)](https://palletsprojects.com/p/flask/)
[![License](https://img.shields.io/badge/license-MIT-blue)](LICENSE)

## 🌟 功能特点

### 📝 智能试卷生成
- **AI驱动**：基于 DeepSeek API 分析资料内容，自动生成高质量试卷
- **多种题型**：支持选择题、填空题、简答题、应用题等多种题型
- **灵活配置**：可自定义试卷类型、难度、总分、时间限制和题目类型
- **智能排版**：自动生成规范格式的试卷，包含题目、答案和解析

### 📤 多格式导出
- **DOCX格式**：可直接用于打印和编辑的Word文档
- **HTML格式**：便于在线浏览和分享的网页格式
- **PDF格式**：便于分发和存档的便携式文档格式

### 👨‍🎓 学生个性化练习
- **定制化服务**：根据学生年级、科目、知识范围和学习目标定制练习题
- **定时推送**：支持每日定时向订阅学生发送个性化练习题
- **详细反馈**：每道题都配有详细解析和学习建议
- **灵活订阅**：学生可随时订阅或取消练习服务

### ⚙️ 系统管理
- **可视化配置**：友好的Web界面配置邮件服务和API密钥
- **多邮箱支持**：支持Gmail、QQ邮箱、163邮箱等多种邮件服务商
- **定时任务**：可自定义每日练习发送时间

## 🚀 快速开始

### 系统要求
- Python 3.7 或更高版本
- pip 包管理工具
- DeepSeek API 密钥

### 安装步骤

1. 克隆仓库
```bash
git clone https://github.com/your-repo/exam-paper-generator.git
cd exam-paper-generator
```

2. 创建并激活虚拟环境
```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

3. 安装依赖
```bash
pip install -r requirements.txt
```

4. 配置环境变量
```bash
# 复制配置文件模板
cp .env.example .env
```

在 `.env` 文件中设置您的 DeepSeek API 密钥：
```env
DEEPSEEK_API_KEY=sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

5. 初始化数据库
```bash
# 第一次运行会自动创建数据库
python app.py
```

6. 启动应用
```bash
python run.py
```

7. 访问系统
打开浏览器访问：`http://localhost:5000`

## 📖 使用指南

### 生成试卷

1. 在首页选择或输入资料（支持 TXT、MD 格式）
2. 配置试卷参数：
   - 选择试卷类型（小练、考试、模拟考、期中期末等）
   - 设置难度级别（简单、中等、困难、竞赛级）
   - 填写总分和考试时间
   - 配置题目类型及数量（可选）
   - 添加自定义指令（可选）
3. 点击 "生成试卷" 按钮
4. 生成完成后可预览并选择导出格式

### 学生练习门户

1. 点击首页 "学生练习门户" 进入
2. 填写个人信息：姓名、邮箱、年级、科目
3. 设置学习信息：
   - 知识范围：指定希望重点练习的知识点
   - 学习目标：明确学习方向和期望
   - 特殊要求：其他个性化需求
4. 配置练习参数：
   - 难度级别
   - 每日题量（3-20题）
   - 偏好发送时间
5. 点击 "订阅个性化练习" 完成订阅
6. 系统将在指定时间发送练习题到您的邮箱

### 系统配置

1. 点击首页 "系统配置" 进入管理页面
2. 配置邮件服务器：
   - 填写邮件服务器地址和端口
   - 输入邮箱用户名和密码/应用密码
   - 设置发件人邮箱
3. 配置 API：
   - 输入 DeepSeek API 密钥
4. 设置每日练习发送时间
5. 点击 "保存配置" 完成设置

## 🔧 技术架构

### 后端技术栈
- **核心框架**：Flask
- **AI接口**：DeepSeek API (deepseek-chat模型)
- **文档处理**：python-docx (DOCX生成)、markdown (Markdown解析)
- **数据库**：SQLite3 (轻量级数据库存储)
- **邮件服务**：Flask-Mail
- **任务调度**：APScheduler (定时任务)
- **配置管理**：configparser

### 前端技术栈
- **模板引擎**：Jinja2 (Flask内置)
- **样式设计**：原生CSS + 响应式布局
- **交互增强**：原生JavaScript





## 📧 邮件服务器配置指南

### Gmail配置
邮件服务器: smtp.gmail.com 端口: 587 加密方式: TLS 用户名: your-email@gmail.com 密码: 应用专用密码（需开启两步验证）


plainText

### QQ邮箱配置
邮件服务器: smtp.qq.com 端口: 587 或 465 加密方式: TLS 或 SSL 用户名: your-email@qq.com 密码: 授权码（非邮箱登录密码）




### 163邮箱配置
邮件服务器: smtp.163.com 端口: 25 或 465 或 994 加密方式: 非加密 或 SSL 用户名: your-email@163.com 密码: 授权码（非邮箱登录密码）




## ⚠️ 注意事项

1. **API密钥**：需从 [DeepSeek官网](https://platform.deepseek.com/) 获取API密钥
2. **邮件服务**：确保所使用的邮箱已开启SMTP服务
3. **编码问题**：系统已优化UTF-8编码支持，确保中文正常显示
4. **配置生效**：修改系统配置后需重启应用才能完全生效
5. **安全性**：
   - 不要在代码中硬编码敏感信息
   - 定期更换API密钥和邮箱密码
   - 生产环境中应使用HTTPS协议

## 🤝 贡献指南

欢迎提交 Issue 和 Pull Request 来帮助改进这个项目！

### 开发环境搭建
```bash
# Fork项目后克隆
git clone https://github.com/your-username/exam-paper-generator.git
cd exam-paper-generator

# 安装开发依赖
pip install -r requirements.txt

# 创建开发分支
git checkout -b feature/your-feature-name

# 提交更改
git commit -am "Add some feature"

# 推送到GitHub
git push origin feature/your-feature-name
```

## 📄 许可证

本项目采用 MIT 许可证，详情请见 [LICENSE](LICENSE) 文件。

## 📞 联系方式

如有任何问题或建议，请通过以下方式联系：
- 提交 GitHub Issue
- 发送邮件至项目维护者邮箱

---