#!/bin/bash
# PCIC Form Studio Single-Click Launcher

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# Check if server is already running on port 5000
if pgrep -f "python.*app.py" > /dev/null; then
    echo "PCIC Form Studio server is already active."
else
    echo "Starting PCIC Form Studio backend server..."
    nohup ./venv/bin/python3 app.py >> app.log 2>&1 &
    disown
fi

# Wait until the server is actively responding on port 5000 (up to 15 seconds)
echo "Waiting for server to become ready..."
for i in {1..30}; do
    if curl -s -m 1 http://127.0.0.1:5000/ > /dev/null 2>&1; then
        echo "PCIC Form Studio server is ready!"
        break
    fi
    sleep 0.5
done

# Automatically launch the web UI in default browser
echo "Opening PCIC Form Studio in web browser..."
if command -v xdg-open &> /dev/null; then
    xdg-open "http://localhost:5000"
elif command -v google-chrome &> /dev/null; then
    google-chrome "http://localhost:5000"
elif command -v firefox &> /dev/null; then
    firefox "http://localhost:5000"
else
    echo "Please open http://localhost:5000 in your browser."
fi
