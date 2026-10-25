#!/bin/zsh
# Doppio click per aggiornare i prezzi (Excel -> DB locale -> PythonAnywhere)
cd "$(dirname "$0")"
/opt/anaconda3/bin/python3 refresh_prices.py
echo ""
read -k 1 "?Premi un tasto per chiudere..."
