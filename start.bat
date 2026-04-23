@echo off
chcp 65001 >nul
title 词库词根管理系统 v3（含数据资产整改）
echo ========================================
echo   词库词根管理系统 v3
echo   含数据资产整改功能
echo ========================================
echo.

REM 检查Python环境
py --version >nul 2>&1
if errorlevel 1 (
    python --version >nul 2>&1
    if errorlevel 1 (
        echo [错误] 未检测到Python环境
        pause
        exit /b 1
    )
    set PY=python
) else (
    set PY=py
)

REM 检查依赖
echo [检查] 正在检查依赖库...
%PY% -c "import docx" >nul 2>&1
if errorlevel 1 (
    echo [安装] 正在安装 python-docx...
    %PY% -m pip install python-docx -q
)
%PY% -c "import xlsxwriter" >nul 2>&1
if errorlevel 1 (
    echo [安装] 正在安装 xlsxwriter...
    %PY% -m pip install xlsxwriter -q
)
echo [检查] 依赖库就绪
echo.

echo [启动] 正在启动服务...
echo [启动] 访问地址: http://localhost:8080
echo [启动] 功能: 词条管理 + 词根管理 + 数据资产整改
echo [启动] 操作日志: server.log
echo [启动] 按 Ctrl+C 停止服务
echo.
%PY% -u server.py 2>&1
pause
