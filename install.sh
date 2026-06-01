#!/bin/bash
# 安装脚本 - 将 convert 命令添加到系统 PATH

echo "===================================="
echo "  文档引号转换工具 - 安装程序"
echo "===================================="
echo ""

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# 检测 shell
if [ -n "$ZSH_VERSION" ]; then
    SHELL_RC="$HOME/.zshrc"
elif [ -n "$BASH_VERSION" ]; then
    SHELL_RC="$HOME/.bashrc"
else
    SHELL_RC="$HOME/.bashrc"
fi

echo "检测到配置文件: $SHELL_RC"
echo ""

# 检查是否已安装
if grep -q "convert_quote" "$SHELL_RC" 2>/dev/null; then
    echo "✓ 似乎已安装过，跳过重复安装"
    echo ""
else
    # 添加到 PATH
    echo "# 文档引号转换工具路径 (convert_quote)" >> "$SHELL_RC"
    echo "export PATH=\"\$PATH:$SCRIPT_DIR\"" >> "$SHELL_RC"
    echo ""
    echo "✓ 已将程序路径添加到 $SHELL_RC"
fi

# 设置可执行权限
chmod +x "$SCRIPT_DIR/convert_quotes.py"
chmod +x "$SCRIPT_DIR/convert"

echo "✓ 已设置可执行权限"
echo ""
echo "安装完成！请运行以下命令使配置生效:"
echo ""
echo "  source $SHELL_RC"
echo ""
echo "之后你可以在任何位置使用:"
echo ""
echo "  convert -f 文件.md         # 转换单个文件"
echo "  convert -d 文件夹路径     # 批量转换文件夹"
echo "  convert -i                 # 进入交互模式"
echo ""
