#!/bin/bash
set -u
export THEMELAB_BIN='D:\work\wxd.files.winforms\wxd.files.winforms\bin\Debug-verify'
GS='D:\work\wxd.files.winforms\tools\gridshot\bin\Debug\net10.0-windows7.0\gridshot.exe'
OUT='D:\work\wxd.files.winforms\tools\restyle-shots'
PS='C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe'
FILE='D:\策划配置\cfg_monster.csv'

for combo in "0 0 skyl" "0 1 skyd" "1 0 densl" "1 1 dense"; do
  set -- $combo
  for m in main reader; do
    taskkill //IM gridshot.exe //F >/dev/null 2>&1
    sleep 1
    if [ "$m" = main ]; then
      cmd //c start "" "$GS" "$1" "$2" "D:\策划配置" main
    else
      cmd //c start "" "$GS" "$1" "$2" "D:\策划配置" reader "$FILE"
    fi
    sleep 7
    "$PS" -NoProfile -File "$OUT/winsht.ps1" -ProcName gridshot -Out "$OUT/$m-$3.png" -Index 0 | tr -d '\r'
  done
done
taskkill //IM gridshot.exe //F >/dev/null 2>&1
echo DONE
