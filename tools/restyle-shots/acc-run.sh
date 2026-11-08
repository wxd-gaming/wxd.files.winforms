#!/bin/bash
# 完整验收矩阵：2 皮肤 × 深浅色 = 4 组合，每组跑主界面 / 表格阅读器 / 日志阅读器 / 全部对话框。
# 每轮必须先杀掉上一轮的 gridshot，否则 PrintWindow 会抓到残留窗口（-Index 0 取的是最顶层）。
set -u
export THEMELAB_BIN='D:\work\wxd.files.winforms\wxd.files.winforms\bin\Debug-verify'
GS='D:\work\wxd.files.winforms\tools\gridshot\bin\Debug\net10.0-windows7.0\gridshot.exe'
OUT='D:\work\wxd.files.winforms\tools\restyle-shots'
PS='C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe'
DIR='D:\策划配置'
MON="$DIR\cfg_monster.csv"
LOG="$DIR\cfg_stress_log.csv"
DLGS=(alert err confirm3 celledit colmgr tcolsel linksel cmdmgr tmpl tmpls linksetup dupdetail help update updn pbar note noteitem conflict vcmp)

# run <产物名> <等待秒> <gridshot 参数...>
run() {
    local name=$1 wait=$2
    shift 2
    taskkill //IM gridshot.exe //F >/dev/null 2>&1
    sleep 1
    "$GS" "$@" >/dev/null 2>&1 &
    sleep "$wait"
    local got
    got=$("$PS" -NoProfile -File "$OUT/winsht.ps1" -ProcName gridshot -Out "$OUT/$name.png" -Index 0 2>&1 | tr -d '\r' | grep -o 'size=[0-9x]*')
    echo "$name $got"
    taskkill //IM gridshot.exe //F >/dev/null 2>&1
}

for combo in "0 0 skyl" "0 1 skyd" "1 0 densl" "1 1 dense"; do
    set -- $combo
    s=$1; d=$2; tag=$3
    run "acc-main-$tag"    8  "$s" "$d" "$DIR" main
    run "acc-reader-$tag" 12  "$s" "$d" "$MON" reader "$MON"
    run "acc-logs-$tag"   26  "$s" "$d" "$DIR" logs "$LOG"
    for x in "${DLGS[@]}"; do
        run "acc-dlg-$x-$tag" 5 "$s" "$d" x dlg "$x"
    done
done
taskkill //IM gridshot.exe //F >/dev/null 2>&1
echo DONE
