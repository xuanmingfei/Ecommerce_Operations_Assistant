#!/bin/zsh

PROJECT_DIR="$HOME/Desktop/Codex/04-GMV诊断系统"
PORT=8780
PYTHON_BIN="$HOME/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
LOG_FILE="$PROJECT_DIR/用户数据/gmv-server.log"

echo "正在重启电商智能AI助手..."
echo ""

if [ ! -d "$PROJECT_DIR" ]; then
  echo "没有找到项目目录：$PROJECT_DIR"
  echo "请确认桌面 Codex 文件夹里还有 04-GMV诊断系统。"
  echo ""
  read "?按回车退出..."
  exit 1
fi

if [ ! -x "$PYTHON_BIN" ]; then
  PYTHON_BIN="$(command -v python3)"
fi

if [ -z "$PYTHON_BIN" ]; then
  echo "没有找到 Python，无法启动服务。"
  echo ""
  read "?按回车退出..."
  exit 1
fi

OLD_PIDS="$(lsof -tiTCP:$PORT -sTCP:LISTEN 2>/dev/null)"
if [ -n "$OLD_PIDS" ]; then
  echo "发现旧服务，正在关闭..."
  echo "$OLD_PIDS" | xargs kill 2>/dev/null
  sleep 1
fi

mkdir -p "$PROJECT_DIR/用户数据"
cd "$PROJECT_DIR" || exit 1

echo "正在启动本地服务..."
"$PYTHON_BIN" "GMV诊断服务.py" > "$LOG_FILE" 2>&1 &
SERVER_PID=$!

for i in {1..30}; do
  if lsof -nP -iTCP:$PORT -sTCP:LISTEN >/dev/null 2>&1; then
    echo "启动成功，正在打开浏览器..."
    open "http://127.0.0.1:$PORT/"
    echo ""
    echo "如果以后打不开网页，双击桌面这个文件即可重启。"
    echo "请保持这个 Terminal 窗口打开；关闭窗口会停止本地服务。"
    echo "日志文件：$LOG_FILE"
    echo ""
    wait "$SERVER_PID"
    exit 0
  fi
  sleep 1
done

echo "启动失败。最近日志如下："
echo ""
tail -40 "$LOG_FILE" 2>/dev/null
kill "$SERVER_PID" 2>/dev/null
echo ""
read "?按回车退出..."
