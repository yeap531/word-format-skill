#!/usr/bin/env bash
# 把渲染好的 HTML 通过浏览器复制 → 粘贴进 Word → 保存为 .docx。
#
# 两种模式：
#
#   模式 A（新建空白文档 + 粘贴 + save as）：
#     bash render_and_paste.sh <render.html> [output.docx]
#     · 只保留字符 / 段落级排版，页面设置 / 样式表 / 主题用 Word 默认
#
#   模式 B（在参考 docx 副本末尾续写，强烈推荐）：
#     bash render_and_paste.sh --append-to <reference.docx> <render.html> <output.docx>
#     · 100% 继承参考 docx 的页面设置 / 样式表 / 主题 / 页眉页脚
#     · 新内容粘贴到副本末尾
#
# 脚本结尾会把焦点还给运行前的前台应用，避免打断用户。
# 注意：脚本运行期间（约 5~7 秒）不要使用键盘 / 鼠标，UI 自动化在跑。

set -euo pipefail

# -------- 参数解析 --------
APPEND_TO=""
INPUT_HTML=""
OUTPUT_DOCX=""

while [ $# -gt 0 ]; do
    case "$1" in
        --append-to)
            APPEND_TO="${2:-}"
            [ -z "$APPEND_TO" ] && { echo "ERROR: --append-to 需要参数" >&2; exit 2; }
            shift 2
            ;;
        -h|--help)
            cat <<USAGE
usage:
  $0 <render.html> [output.docx]                                # 模式 A：新建文档
  $0 --append-to <reference.docx> <render.html> <output.docx>   # 模式 B：副本末尾续写
USAGE
            exit 0
            ;;
        --)
            shift
            ;;
        *)
            if [ -z "$INPUT_HTML" ]; then
                INPUT_HTML="$1"
            elif [ -z "$OUTPUT_DOCX" ]; then
                OUTPUT_DOCX="$1"
            else
                echo "ERROR: 多余参数 $1" >&2
                exit 2
            fi
            shift
            ;;
    esac
done

if [ -z "$INPUT_HTML" ]; then
    echo "usage: $0 <render.html> [output.docx]   或   $0 --append-to <ref.docx> <render.html> <output.docx>" >&2
    exit 2
fi

# 路径标准化
ABS_HTML=$(python3 -c "import sys; from pathlib import Path; print(Path(sys.argv[1]).expanduser().resolve())" "$INPUT_HTML")
ABS_URI=$(python3  -c "import sys; from pathlib import Path; print(Path(sys.argv[1]).expanduser().resolve().as_uri())" "$INPUT_HTML")
if [ ! -f "$ABS_HTML" ]; then
    echo "ERROR: not found: $ABS_HTML" >&2
    exit 1
fi

if [ -n "$OUTPUT_DOCX" ]; then
    OUTPUT_DOCX=$(python3 -c "import sys; from pathlib import Path; print(Path(sys.argv[1]).expanduser().resolve())" "$OUTPUT_DOCX")
    mkdir -p "$(dirname "$OUTPUT_DOCX")"
fi

# 模式 B：先把参考 docx 复制到 OUTPUT_DOCX
if [ -n "$APPEND_TO" ]; then
    if [ -z "$OUTPUT_DOCX" ]; then
        echo "ERROR: --append-to 模式必须提供 <output.docx>" >&2
        exit 2
    fi
    APPEND_TO=$(python3 -c "import sys; from pathlib import Path; print(Path(sys.argv[1]).expanduser().resolve())" "$APPEND_TO")
    if [ ! -f "$APPEND_TO" ]; then
        echo "ERROR: 参考 docx 不存在: $APPEND_TO" >&2
        exit 1
    fi
    cp "$APPEND_TO" "$OUTPUT_DOCX"
fi

# 浏览器选择
if   [ -d "/Applications/Google Chrome.app" ]; then BROWSER="Google Chrome"
elif [ -d "/Applications/Safari.app" ];        then BROWSER="Safari"
else
    echo "ERROR: 需要 Google Chrome 或 Safari" >&2
    exit 1
fi

echo "Browser: $BROWSER"
echo "Render:  $ABS_HTML"
if [ -n "$APPEND_TO" ]; then
    echo "Mode:    续写（继承模板）"
    echo "Append:  $APPEND_TO  →  $OUTPUT_DOCX"
else
    echo "Mode:    新建空白文档"
    [ -n "$OUTPUT_DOCX" ] && echo "Save as: $OUTPUT_DOCX"
fi

# bash → AppleScript 布尔
if [ -n "$APPEND_TO" ]; then APPEND_FLAG="true"; else APPEND_FLAG="false"; fi

# -------- 自动化主体 --------
set +e
osascript_output=$(osascript <<APPLESCRIPT 2>&1

set targetUrl   to "$ABS_URI"
set browserName to "$BROWSER"
set appendMode  to $APPEND_FLAG
set outputDocx  to "$OUTPUT_DOCX"

-- 记录起始 frontmost 进程，结尾还原焦点
set origFront to ""
try
    tell application "System Events"
        set origFront to name of first process whose frontmost is true
    end tell
end try

-- 1) 浏览器加载渲染页
tell application browserName to activate
delay 0.6

if browserName is "Google Chrome" then
    tell application "Google Chrome"
        if (count of windows) is 0 then make new window
        set URL of active tab of front window to targetUrl
    end tell
else
    tell application "Safari"
        if not (exists front document) then make new document
        set URL of front document to targetUrl
    end tell
end if

-- 等待页面渲染（含字体加载）
delay 2.5

-- 全选 + 复制
tell application browserName to activate
delay 0.4
tell application "System Events"
    tell process browserName
        keystroke "a" using command down
        delay 0.4
        keystroke "c" using command down
    end tell
end tell
delay 0.6

-- 2) Word：续写 → 打开副本；新建 → 新建空文档
tell application "Microsoft Word"
    activate
    if appendMode then
        open POSIX file outputDocx
    else
        make new document
    end if
end tell
delay 1.5

-- 3) 确认 Word 已成为前台进程
tell application "Microsoft Word" to activate
delay 0.5
repeat 10 times
    tell application "System Events"
        set frontProcess to name of first process whose frontmost is true
    end tell
    if frontProcess is "Microsoft Word" then exit repeat
    delay 0.2
end repeat

-- 4) 粘贴
tell application "System Events"
    tell process "Microsoft Word"
        if appendMode then
            -- 续写：Cmd+A 选全文 → 右方向键塌缩到文档末尾 → Cmd+V 粘到末尾
            keystroke "a" using command down
            delay 0.3
            key code 124
            delay 0.3
        end if
        keystroke "v" using command down
    end tell
end tell
delay 2.0

-- 5) 保存
tell application "Microsoft Word"
    if appendMode then
        save active document
    else
        if outputDocx is not "" then
            set theDoc to active document
            save as theDoc file name outputDocx file format format document
        end if
    end if
end tell
delay 0.4

-- 6) 还原起始焦点（让浏览器/Word 不留在前台打扰用户）
if origFront is not "" then
    try
        tell application "System Events"
            set frontmost of process origFront to true
        end tell
    end try
end if

APPLESCRIPT
)
osascript_rc=$?
set -e

if [ $osascript_rc -ne 0 ]; then
    echo "$osascript_output" >&2
    if echo "$osascript_output" | grep -q "1002"; then
        cat >&2 <<'HINT'

======================================================================
错误：osascript 没有发送按键的权限。

解决：开启「辅助功能」权限：
  系统设置 → 隐私与安全性 → 辅助功能
  在列表里找到当前终端宿主（Terminal / iTerm / Ghostty / Claude Code）并打开开关。
  如果列表里没有，点 + 号把它加进去。

授权完成后重跑本脚本即可。
======================================================================
HINT
    fi
    exit $osascript_rc
fi

[ -n "$OUTPUT_DOCX" ] && echo "Saved: $OUTPUT_DOCX"
echo "Done."
