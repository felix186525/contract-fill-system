#!/bin/bash
set -e
SERVER="root@139.224.225.188"
APP_DIR="/opt/contract-platform"

echo "=== 安装服务器依赖 ==="
ssh $SERVER "
  # 安装LibreOffice（用于doc/docx转PDF）和ImageMagick（PDF转PNG）
  yum install -y libreoffice libreoffice-headless ghostscript ImageMagick 2>/dev/null || \
  dnf install -y libreoffice libreoffice-headless ghostscript ImageMagick 2>/dev/null || true

  # 安装中文字体
  yum install -y wqy-zenhei-fonts wqy-microhei-fonts 2>/dev/null || \
  dnf install -y wqy-zenhei-fonts 2>/dev/null || true
  fc-cache -f 2>/dev/null || true

  # pdftoppm备选
  yum install -y poppler-utils 2>/dev/null || dnf install -y poppler-utils 2>/dev/null || true

  # Python3环境
  python3 --version
  pip3 install --upgrade pip 2>/dev/null || true
"

echo "=== 部署应用 ==="
ssh $SERVER "mkdir -p $APP_DIR/uploads $APP_DIR/output $APP_DIR/templates $APP_DIR/static"
scp app.py requirements.txt $SERVER:$APP_DIR/
scp templates/index.html $SERVER:$APP_DIR/templates/

echo "=== 安装Python依赖 ==="
ssh $SERVER "pip3 install -r $APP_DIR/requirements.txt"

echo "=== 配置systemd服务 ==="
ssh $SERVER "cat > /etc/systemd/system/contract-platform.service << 'EOF'
[Unit]
Description=Contract Platform
After=network.target

[Service]
Type=simple
User=root
WorkingDirectory=$APP_DIR
ExecStart=/usr/bin/python3 app.py
Restart=always
RestartSec=3
Environment=PYTHONUNBUFFERED=1

[Install]
WantedBy=multi-user.target
EOF
systemctl daemon-reload
systemctl enable contract-platform
systemctl restart contract-platform
sleep 2
systemctl status contract-platform --no-pager"

echo "=== 开放防火墙端口 ==="
ssh $SERVER "
  firewall-cmd --permanent --add-port=5000/tcp 2>/dev/null || true
  firewall-cmd --reload 2>/dev/null || true
  # 阿里云安全组需手动开放5000端口
"

echo ""
echo "✅ 部署完成！访问地址：http://139.224.225.188:5000"
