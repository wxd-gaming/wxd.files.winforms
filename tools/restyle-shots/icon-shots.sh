#!/bin/bash
# 图标验收：抓窗口后裁出标题栏左上角并放大，直接看图标像素。
# 图标不随皮肤变色，所以只挑明暗两种标题栏背景各验一遍。
set -u
export THEMELAB_BIN='D:\work\wxd.files.winforms\wxd.files.winforms\bin\Debug-verify'
GS='D:\work\wxd.files.winforms\tools\gridshot\bin\Debug\net10.0-windows7.0\gridshot.exe'
OUT='D:\work\wxd.files.winforms\tools\restyle-shots'
PS='C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe'
DIR='D:\策划配置'
MON="$DIR\cfg_monster.csv"
LOG="$DIR\cfg_stress_log.csv"

run() {
    local name=$1 wait=$2 idx=$3
    shift 3
    taskkill //IM gridshot.exe //F >/dev/null 2>&1
    sleep 1
    "$GS" "$@" >/dev/null 2>&1 &
    sleep "$wait"
    "$PS" -NoProfile -File "$OUT/winsht.ps1" -ProcName gridshot -Out "$OUT/$name.png" -Index "$idx" 2>&1 | tr -d '\r' | grep -o 'size=[0-9x]*'
    "$PS" -NoProfile -File "$OUT/crop2.ps1" -In "$OUT/$name.png" -Out "$OUT/icon-$name.png" -X 0 -Y 0 -W 420 -H 40 -Scale 3 >/dev/null
    taskkill //IM gridshot.exe //F >/dev/null 2>&1
    echo "icon-$name.png"
}

run cap-main-skyl      8 0 "0" "0" "$DIR" main
run cap-main-skyd      8 0 "0" "1" "$DIR" main
run cap-reader-skyl   12 0 "0" "0" "$MON" reader "$MON"
run cap-reader-skyd   12 0 "0" "1" "$MON" reader "$MON"
run cap-logs-skyd     26 0 "0" "1" "$DIR" logs "$LOG"
run cap-logsdetail    30 0 "0" "0" "$DIR" logs "$LOG" detail
for x in celledit colmgr tcolsel linksel cmdmgr tmpl linksetup dupdetail note vcmp jump; do
    run "cap-dlg-$x" 5 0 "0" "0" x dlg "$x"
done
for x in colmgr linksetup note; do
    run "cap-dlg-$x-skyd" 5 0 "0" "1" x dlg "$x"
done
taskkill //IM gridshot.exe //F >/dev/null 2>&1
echo DONE
