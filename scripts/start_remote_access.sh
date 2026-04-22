#!/bin/zsh
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
SERVER_LOG="$ROOT_DIR/web_service/server.log"
TUNNEL_LOG="$ROOT_DIR/web_service/tunnel.log"
SERVER_PID_FILE="$ROOT_DIR/web_service/server.pid"
SCREEN_SESSION="fitness_tunnel"
PORT="${FITNESS_PORT:-8123}"

cd "$ROOT_DIR"

start_server() {
  if lsof -nP -iTCP:"$PORT" -sTCP:LISTEN >/dev/null 2>&1; then
    echo "Server already listening on port $PORT."
    return
  fi

  nohup ./.venv/bin/python web_service/server.py >"$SERVER_LOG" 2>&1 &
  echo $! >"$SERVER_PID_FILE"
  sleep 2
  echo "Server started on port $PORT."
}

start_tunnel() {
  if screen -list | grep -q "[.]$SCREEN_SESSION"; then
    echo "Tunnel already running in screen session '$SCREEN_SESSION'."
    return
  fi

  : >"$TUNNEL_LOG"
  screen -dmS "$SCREEN_SESSION" zsh -lc \
    "ssh -o StrictHostKeyChecking=no -o ServerAliveInterval=30 -R 80:127.0.0.1:$PORT nokey@localhost.run | tee -a '$TUNNEL_LOG'"

  for _ in {1..30}; do
    if grep -Eo 'https://[a-z0-9.-]+\.lhr\.life' "$TUNNEL_LOG" | tail -n 1 >/dev/null; then
      break
    fi
    sleep 1
  done
}

start_server
start_tunnel

PUBLIC_URL="$(grep -Eo 'https://[a-z0-9.-]+\.lhr\.life' "$TUNNEL_LOG" | tail -n 1 || true)"

if [[ -n "$PUBLIC_URL" ]]; then
  echo "Public URL: $PUBLIC_URL"
else
  echo "Tunnel started, but URL is not visible yet. Check $TUNNEL_LOG"
fi

echo "Server log: $SERVER_LOG"
echo "Tunnel log: $TUNNEL_LOG"
