#!/bin/bash
# PCIC Form Studio Single-Click Launcher

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# Check if server is already running on port 5000
if pgrep -f "python.*app.py" > /dev/null; then
    echo "PCIC Form Studio server is already active."
else
    echo "Starting PCIC Form Studio backend server..."
    ./venv/bin/python3 app.py &
    sleep 2
fi

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
