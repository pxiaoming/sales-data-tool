#!/bin/bash
cd "$(dirname "$0")"

if [ -x ".venv/bin/python" ]; then
    .venv/bin/python scripts/run.py
else
    echo "错误：未找到 Python 环境，请联系维护人员安装依赖。"
    echo "维护人员请在项目目录执行：pip install -r requirements.txt"
fi

echo
read -p "按回车键关闭..."
