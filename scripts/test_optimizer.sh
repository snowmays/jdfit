#!/bin/bash
# 简历优化工具测试脚本

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "📋 简历优化小工具 - 测试脚本"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

# 检查 Python
if ! command -v python3 &> /dev/null; then
    echo "❌ 错误: 未找到 python3"
    echo "   请先安装 Python 3.7+"
    exit 1
fi

echo "✅ Python 版本: $(python3 --version)"
echo ""

# 检查依赖
echo "📦 检查依赖..."
if ! python3 -c "import docx" 2>/dev/null; then
    echo "⚠️  未安装 python-docx，正在安装..."
    pip3 install -r requirements.txt
    echo "✅ 依赖安装完成"
else
    echo "✅ 依赖已安装"
fi
echo ""

# 创建测试目录
TEST_DIR="/tmp/resume-optimizer-test"
mkdir -p "$TEST_DIR"
echo "📁 测试目录: $TEST_DIR"
echo ""

# 检查是否有示例文件
if [ ! -f "$TEST_DIR/test_resume.docx" ]; then
    echo "⚠️  未找到测试简历文件"
    echo "   请将测试简历放置在: $TEST_DIR/test_resume.docx"
    echo ""
    echo "💡 提示: 你也可以手动运行:"
    echo "   python3 docx_optimizer_advanced.py \\"
    echo "     --original /path/to/resume.docx \\"
    echo "     --optimizations example_optimizations.json \\"
    echo "     --output-dir /path/to/output \\"
    echo "     --job-title '产品经理'"
    exit 1
fi

# 运行优化
echo "🚀 开始优化简历..."
echo ""

python3 docx_optimizer_advanced.py \
    --original "$TEST_DIR/test_resume.docx" \
    --optimizations example_optimizations.json \
    --output-dir "$TEST_DIR/output" \
    --job-title "测试职位"

echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "✅ 测试完成！"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
echo "📁 输出文件位于: $TEST_DIR/output/"
echo "   ├─ resume_original_backup.docx"
echo "   └─ resume_optimized_测试职位.docx"
echo ""
