#!/bin/bash

echo "Starting Photo Mapper Application..."

# Function to cleanup processes on exit
cleanup() {
    echo "Stopping services..."
    kill $BACKEND_PID 2>/dev/null
    kill $FRONTEND_PID 2>/dev/null
    exit
}

# Set trap for cleanup
trap cleanup SIGINT SIGTERM

# Start backend
echo "Starting Python backend on port 5001..."
cd backend
python3 app.py &
BACKEND_PID=$!

# Wait a moment for backend to start
sleep 3

# Start frontend
echo "Starting NextJS frontend on port 3000..."
cd ../frontend
npm run dev &
FRONTEND_PID=$!

echo ""
echo "🚀 Photo Mapper is running:"
echo "   Frontend: http://localhost:3000"
echo "   Backend:  http://localhost:5001"
echo "   Main App: http://localhost:3000/map-images-to-players"
echo ""
echo "Press Ctrl+C to stop both services"

# Wait for processes
wait