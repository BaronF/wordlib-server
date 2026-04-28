#!/bin/bash
# 词库管理系统一键部署脚本
set -e

echo "========================================="
echo "  词库管理系统 - 一键部署"
echo "========================================="

# 安装依赖
apt update && apt install -y python3 python3-pip git

# 克隆代码
cd /opt
rm -rf wordlib-server
git clone https://github.com/BaronF/wordlib-server.git
cd wordlib-server

# 安装 Python 依赖
pip3 install xlsxwriter

# 创建系统服务（开机自启）
cat > /etc/systemd/system/wordlib.service << 'EOF'
[Unit]
Description=WordLib Server
After=network.target

[Service]
WorkingDirectory=/opt/wordlib-server
ExecStart=/usr/bin/python3 server.py
Restart=always
RestartSec=3
Environment=PORT=8080

[Install]
WantedBy=multi-user.target
EOF

# 启动服务
systemctl daemon-reload
systemctl enable wordlib
systemctl start wordlib

echo ""
echo "========================================="
echo "  部署完成！"
echo "  访问地址: http://$(curl -s ifconfig.me):8080"
echo "  默认账号: admin / admin123"
echo "========================================="
