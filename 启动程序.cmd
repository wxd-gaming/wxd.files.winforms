@echo off
rem 以普通用户权限启动（不要从管理员/沙箱进程里启动，否则 TortoiseSVN 右键会失效）
cd /d "%~dp0wxd.files.winforms\bin\Debug\net10.0-windows7.0"
start "" wxd.files.winforms.exe
