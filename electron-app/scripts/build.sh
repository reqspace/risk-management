#!/bin/bash
# Build script for Risk Management Desktop App
# Run this on macOS to build the Mac .dmg

set -e

echo "=========================================="
echo "  Risk Management - Build Script"
echo "=========================================="

cd "$(dirname "$0")/.."
ROOT_DIR="$(pwd)/.."

echo ""
echo "[1/5] Installing Electron dependencies..."
npm install

echo ""
echo "[2/5] Building React dashboard..."
cd "$ROOT_DIR/dashboard"
npm install
npm run build
mkdir -p "$ROOT_DIR/electron-app/build/webapp"
cp -r dist/* "$ROOT_DIR/electron-app/build/webapp/"

echo ""
echo "[3/5] Installing PyInstaller..."
pip3 install pyinstaller

echo ""
echo "[4/5] Building Python server..."
cd "$ROOT_DIR/electron-app/python"
pyinstaller --clean --noconfirm risk_server.spec

echo ""
echo "[5/5] Building Electron app..."
cd "$ROOT_DIR/electron-app"
npm run build:mac

echo ""
echo "=========================================="
echo "  Build Complete!"
echo "=========================================="
echo ""
echo "Mac app: electron-app/dist/Risk Management.dmg"
echo ""
