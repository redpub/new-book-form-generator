#!/usr/bin/env bash
set -e

APP_DIR="/workspaces/new-book-form-generator"
VENV="$APP_DIR/.venv/bin/streamlit"
LOG="/tmp/streamlit.log"

# Ensure venv exists
if [ ! -f "$VENV" ]; then
    echo "❌ Streamlit venv not found at $VENV"
    exit 0
fi

# Avoid duplicate processes
if pgrep -f "streamlit run streamlit_app.py" > /dev/null; then
    echo "Streamlit already running"
    exit 0
fi

echo "Starting Streamlit..."
cd "$APP_DIR"
nohup $VENV run streamlit_app.py \
    --server.address 0.0.0.0 \
    --server.port 8501 \
    --server.headless true \
    --server.showEmailPrompt false \
    --server.enableCORS false \
    --browser.gatherUsageStats false \
    --browser.gatherUsageStats false \
    > "$LOG" 2>&1 &
