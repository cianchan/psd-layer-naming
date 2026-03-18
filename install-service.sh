#!/bin/bash
# PSD Layer Naming Tool - 一键安装脚本
# 适用于 macOS 12+，需要已安装 Python3 和 Adobe Photoshop

set -e

REPO_DIR="$HOME/psd-layer-naming"
PLIST="$HOME/Library/LaunchAgents/com.psd-layer-naming.plist"
PORT=7861

echo ""
echo "════════════════════════════════════════"
echo "  PSD 图层重命名工具 - 安装程序"
echo "════════════════════════════════════════"
echo ""

# 1. Clone or update
if [ -d "$REPO_DIR/.git" ]; then
  echo "▶ 更新代码..."
  git -C "$REPO_DIR" pull
else
  echo "▶ 克隆仓库..."
  git clone https://github.com/cianchan/psd-layer-naming.git "$REPO_DIR"
fi

# 2. Install Python dependencies
echo "▶ 安装 Python 依赖..."
pip3 install --quiet flask python-docx

# 3. Find python3 path
PYTHON3=$(which python3)

# 4. Write launchd plist
echo "▶ 创建开机自启服务..."
cat > "$PLIST" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.psd-layer-naming</string>
    <key>ProgramArguments</key>
    <array>
        <string>${PYTHON3}</string>
        <string>${REPO_DIR}/app.py</string>
    </array>
    <key>WorkingDirectory</key>
    <string>${REPO_DIR}</string>
    <key>RunAtLoad</key>
    <true/>
    <key>KeepAlive</key>
    <true/>
    <key>StandardOutPath</key>
    <string>${HOME}/Library/Logs/psd-layer-naming.log</string>
    <key>StandardErrorPath</key>
    <string>${HOME}/Library/Logs/psd-layer-naming-error.log</string>
</dict>
</plist>
EOF

# 5. Load (or reload) the service
launchctl unload "$PLIST" 2>/dev/null || true
launchctl load "$PLIST"

sleep 2

echo ""
echo "════════════════════════════════════════"
echo "  ✅ 安装完成！"
echo ""
echo "  访问地址: http://127.0.0.1:${PORT}"
echo ""
echo "  服务已设置为开机自动启动。"
echo "  如需停止服务，运行："
echo "  launchctl unload ~/Library/LaunchAgents/com.psd-layer-naming.plist"
echo "════════════════════════════════════════"
echo ""

# Open in browser
open "http://127.0.0.1:${PORT}" 2>/dev/null || true
