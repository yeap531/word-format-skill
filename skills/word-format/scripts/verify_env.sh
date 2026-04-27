#!/usr/bin/env bash
# 检查 word-format skill 运行所需的环境与权限。
set -e

ok()   { echo "  ✓ $1"; }
miss() { echo "  ✗ $1" >&2; missing=1; }

missing=0

echo "[ 应用 ]"
[ -d "/Applications/Microsoft Word.app" ] && ok "Microsoft Word" || miss "Microsoft Word（未安装）"

if   [ -d "/Applications/Google Chrome.app" ]; then ok "Google Chrome (主选)"
elif [ -d "/Applications/Safari.app" ];        then ok "Safari (回退)"
else                                                miss "需要 Google Chrome 或 Safari"
fi

echo
echo "[ 命令 ]"
command -v osascript >/dev/null && ok "osascript" || miss "osascript"
command -v python3   >/dev/null && ok "python3"   || miss "python3"

echo
echo "[ 首次使用须开启的系统授权 ]"
echo "  系统设置 → 隐私与安全性 → 自动化"
echo "    勾选「终端 / Claude」控制：Microsoft Word、Google Chrome（或 Safari）、System Events"
echo "  系统设置 → 隐私与安全性 → 辅助功能"
echo "    启用「终端 / Claude」"
echo
[ "$missing" = "1" ] && { echo "环境不完整。" >&2; exit 1; }
echo "环境就绪。"
