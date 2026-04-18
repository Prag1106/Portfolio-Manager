#!/bin/bash
echo "========================================="
echo "  Portfolio Manager — Avvio server"
echo "========================================="
echo ""
echo "Aprire nel browser: http://localhost:5000"
echo "Per fermare il server: CTRL + C"
echo ""
cd "$(dirname "$0")"
python3 app.py
