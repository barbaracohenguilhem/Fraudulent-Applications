#!/bin/bash
cd "$(dirname "$0")"
echo "Iniciando Dashboard Python Dash em http://127.0.0.1:8050 ..."
(sleep 2 && open "http://127.0.0.1:8050") &
./.venv/bin/python app_dash.py

