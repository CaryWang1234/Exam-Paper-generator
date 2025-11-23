#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智能试卷生成系统启动脚本
确保UTF-8编码支持
"""

import os
import sys
import locale

# 设置UTF-8编码环境
def setup_utf8_environment():
    """设置UTF-8编码环境"""
    # 设置环境变量
    os.environ['PYTHONIOENCODING'] = 'utf-8'
    os.environ['LANG'] = 'zh_CN.UTF-8'
    os.environ['LC_ALL'] = 'zh_CN.UTF-8'
    
    # 设置标准输出编码
    if sys.platform.startswith('win'):
        # Windows系统特殊处理
        if hasattr(sys.stdout, 'reconfigure'):
            sys.stdout.reconfigure(encoding='utf-8')
        if hasattr(sys.stderr, 'reconfigure'):
            sys.stderr.reconfigure(encoding='utf-8')
    else:
        # Linux/Mac系统
        try:
            locale.setlocale(locale.LC_ALL, 'zh_CN.UTF-8')
        except locale.Error:
            try:
                locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')
            except locale.Error:
                pass

if __name__ == '__main__':
    # 设置UTF-8环境
    setup_utf8_environment()
    
    # 导入并运行应用
    from app import app
    
    print("=" * 50)
    print("智能试卷生成系统")
    print("系统编码:", sys.getdefaultencoding())
    print("文件系统编码:", sys.getfilesystemencoding())
    print("=" * 50)
    
    # 运行应用
    app.run(host='0.0.0.0', port=5000, debug=False)