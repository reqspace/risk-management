#!/bin/bash
# Risk Management - Mac/Linux Start Script
# Double-click this file to start the application

cd "$(dirname "$0")"

echo "=========================================="
echo "   Risk Management Dashboard"
echo "=========================================="
echo ""

# Check if .env exists
if [ ! -f .env ]; then
    echo "ERROR: .env file not found!"
    echo "Please copy .env.example to .env and fill in your settings."
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

# Check if node_modules exists
if [ ! -d "dashboard/node_modules" ]; then
    echo "Installing dashboard dependencies (first run only)..."
    cd dashboard && npm install && cd ..
    echo ""
fi

# Kill any existing processes on our ports
lsof -ti:5001 | xargs kill -9 2>/dev/null
lsof -ti:3000 | xargs kill -9 2>/dev/null

echo "Starting Flask API server..."
python3 server.py &
SERVER_PID=$!
sleep 2

echo "Starting Dashboard..."
cd dashboard
npm run dev &
DASHBOARD_PID=$!
cd ..

sleep 3

echo ""
echo "=========================================="
echo "   Application Started!"
echo "=========================================="
echo ""
echo "Dashboard: http://localhost:3000"
echo "API Server: http://localhost:5001"
echo ""
echo "Press Ctrl+C to stop all services"
echo ""

# Open browser
open http://localhost:3000 2>/dev/null || xdg-open http://localhost:3000 2>/dev/null

# Wait and cleanup on exit
trap "kill $SERVER_PID $DASHBOARD_PID 2>/dev/null; exit" INT TERM
wait
