#!/bin/bash
# 行号映射改版后，把阅读器在四种皮肤/深浅色组合下各截一张真窗口像素
set -u
BIN=/d/work/wxd.files.winforms/tools/gridshot/bin/Debug/net10.0-windows7.0
SHOTS=/d/work/wxd.files.winforms/tools/restyle-shots
PS=C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe
FILE='D:\策划配置\cfg_monster.csv'
export THEMELAB_BIN='D:\work\wxd.files.winforms\wxd.files.winforms\bin\Debug-verify'

cd "$SHOTS" || exit 1
for combo in "0 0" "1 1" "1 0"; do
  set -- $combo
  skin=$1; dark=$2
  "$BIN/gridshot.exe" "$skin" "$dark" "$FILE" reader "$FILE" &
  pid=$!
  sleep 9
  "$PS" -NoProfile -File winsht.ps1 -ProcName gridshot -Out "$SHOTS/rm-s${skin}d${dark}.png" -Index 0 2>&1 | tail -1
  taskkill //PID $pid //F >/dev/null 2>&1 || true
  wait $pid 2>/dev/null
  sleep 1
done
ls -l "$SHOTS"/rm-s*.png
