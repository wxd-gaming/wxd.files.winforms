#!/bin/bash
# 逐个把对话框顶出来截图：每轮都要先杀掉上一轮的 gridshot，否则 PrintWindow 会抓到旧窗口
set -u
export THEMELAB_BIN='D:\work\wxd.files.winforms\wxd.files.winforms\bin\Debug-verify'
GS='D:\work\wxd.files.winforms\tools\gridshot\bin\Debug\net10.0-windows7.0\gridshot.exe'
OUT='D:\work\wxd.files.winforms\tools\restyle-shots'
PS='C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe'

for combo in "0 0 skyl" "0 1 skyd" "1 0 densl" "1 1 dense"; do
  set -- $combo
  for d in alert err confirm3 celledit colmgr tcolsel linksel cmdmgr tmpl tmpls linksetup dupdetail; do
    taskkill //IM gridshot.exe //F >/dev/null 2>&1
    sleep 1
    cmd //c start "" "$GS" "$1" "$2" x dlg "$d"
    sleep 4
    "$PS" -NoProfile -ExecutionPolicy Bypass -File "$OUT/winsht.ps1" \
      -ProcName gridshot -Out "$OUT/dlg-$d-$3.png" -Index 0 | tr -d '\r'
  done
done
taskkill //IM gridshot.exe //F >/dev/null 2>&1
echo DONE
