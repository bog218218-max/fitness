#!/bin/zsh
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
SERVER_PID_FILE="$ROOT_DIR/web_service/server.pid"
SCREEN_SESSION="fitness_tunnel"

cd "$ROOT_DIR"

if screen -list | grep -q "[.]$SCREEN_SESSION"; then
  screen -S "$SCREEN_SESSION" -X quit
  echo "Stopped tunnel screen session '$SCREEN_SESSION'."
fi

if [[ -f "$SERVER_PID_FILE" ]]; then
  SERVER_PID="$(cat "$SERVER_PID_FILE")"
  if kill -0 "$SERVER_PID" 2>/dev/null; then
    kill "$SERVER_PID"
    echo "Stopped server process $SERVER_PID."
  fi
  rm -f "$SERVER_PID_FILE"
fi
