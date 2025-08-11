#!/bin/bash
echo "部署開箱工具系統到生產環境..."
echo "1. 清理 dist 資料夾"
rm -rf dist/*
echo "2. 複製生產檔案"
cp index.html styles.css script.js dist/
cp function-menu.html function-menu.js dist/
cp po-header.html po-header.js dist/
cp receiving-confirm.html receiving-confirm.js dist/
echo "3. 創建生產版本 README"
echo "# 開箱工具系統 - 生產版本" > dist/README.md
echo "部署完成！"
