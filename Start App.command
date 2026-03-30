#!/bin/bash

# Find the directory this script lives in
DIR="$(cd "$(dirname "$0")" && pwd)"

# Kill any existing instance on port 5001
lsof -ti :5001 | xargs kill -9 2>/dev/null

# Start the Flask server in the background
cd "$DIR"
python3 "$DIR/start.py" > "$DIR/server.log" 2>&1 &
SERVER_PID=$!
echo "Started server (PID $SERVER_PID)"

# Wait for it to be ready
for i in {1..20}; do
  if curl -s http://localhost:5001/ > /dev/null 2>&1; then
    break
  fi
  sleep 0.5
done

# Open the browser
open http://localhost:5001

echo "App is running at http://localhost:5001"
echo "Close this window to stop the server."

# Keep running until the window is closed, then clean up
trap "kill $SERVER_PID 2>/dev/null; echo 'Server stopped.'" EXIT
wait $SERVER_PID
