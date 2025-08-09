#!/usr/bin/env python3
import http.server
import socketserver
import os

# 確保在正確的目錄中
os.chdir('/Users/user/employee-selector-app')

PORT = 3000

Handler = http.server.SimpleHTTPRequestHandler

with socketserver.TCPServer(("", PORT), Handler) as httpd:
    print(f"伺服器運行在 http://localhost:{PORT}")
    print("按 Ctrl+C 停止伺服器")
    httpd.serve_forever()
