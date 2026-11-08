#!/bin/bash
# 运行期翻主题这条路径单独验：静态矩阵里每个界面都是「开窗前就定好主题」，
# 而用户实际会开着窗口切皮肤/明暗，只有运行时翻转才走 Theme.Changed + ThemeRecolor。
# 截图落在翻转后的那一套配色上，所以 acccheck 用的是目的组合，不是起始组合。
set -u
export THEMELAB_BIN='D:\work\wxd.files.winforms\wxd.files.winforms\bin\Debug-verify'
GS='D:\work\wxd.files.winforms\tools\gridshot\bin\Debug\net10.0-windows7.0\gridshot.exe'
OUT='D:\work\wxd.files.winforms\tools\restyle-shots'
PS='C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe'
DIR='D:\策划配置'
MON="$DIR\cfg_monster.csv"
LOG="$DIR\cfg_stress_log.csv"

opp() { if [ "$1" = "skyl" ]; then echo skyd; elif [ "$1" = "skyd" ]; then echo skyl;
        elif [ "$1" = "densl" ]; then echo dense; else echo densl; fi; }

# run <产物名> <等待秒> <gridshot 参数...>，截完立刻按「翻转后」的调色板查一遍配色
check() {
    local name=$1 dest=$2
    "$PS" -NoProfile -File "$OUT/acccheck.ps1" -In "$OUT/$name.png" -Combo "$dest" \
        2>&1 | tr -d '\r' | grep -E "OFFENDERS|OFF " | head -8 | sed "s/^/$name[$dest] /"
}

shoot() {
    local name=$1 wait=$2 flip=$3
    shift 3
    taskkill //IM gridshot.exe //F >/dev/null 2>&1
    sleep 1
    if [ "$flip" = "1" ]; then
        GRIDSHOT_FLIP=1 "$GS" "$@" >/dev/null 2>&1 &
    else
        "$GS" "$@" >/dev/null 2>&1 &
    fi
    sleep "$wait"
    local got
    got=$("$PS" -NoProfile -File "$OUT/winsht.ps1" -ProcName gridshot -Out "$OUT/$name.png" -Index 0 2>&1 | tr -d '\r' | grep -o 'size=[0-9x]*')
    echo "$name $got"
    taskkill //IM gridshot.exe //F >/dev/null 2>&1
}

for combo in "0 0 skyl" "0 1 skyd" "1 0 densl" "1 1 dense"; do
    set -- $combo
    s=$1; d=$2; tag=$3
    dst=$(opp "$tag")

    shoot "flip-main-$tag" 10 0 "$s" "$d" "$DIR" main switch
    check "flip-main-$tag" "$dst"

    shoot "flip-logs-$tag" 28 0 "$s" "$d" "$DIR" logs "$LOG" switch
    check "flip-logs-$tag" "$dst"

    shoot "flip-reader-$tag" 16 0 "$s" "$d" "$MON" reader "$MON" switch
    check "flip-reader-$tag" "$dst"

    shoot "flip-note-$tag" 9 1 "$s" "$d" x dlg note
    check "flip-note-$tag" "$dst"
done
taskkill //IM gridshot.exe //F >/dev/null 2>&1
echo DONE
