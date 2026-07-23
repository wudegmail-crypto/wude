#!/bin/bash
# 麒麟系统 PyInstaller 打包脚本
# 在容器内执行

set -e

SCRIPT_NAME="cx_new.py"
ICON_NAME="myicon.ico"
OUTPUT_NAME="合同数据合并工具_v12_kylin"

echo "=== 开始打包 ${OUTPUT_NAME} ==="
echo "Python 版本: $(python3.8 --version)"
echo "glibc 版本: $(ldd --version | head -1)"

cd /build

# 检查源文件
if [ ! -f "${SCRIPT_NAME}" ]; then
    echo "错误: 未找到 ${SCRIPT_NAME}"
    exit 1
fi

ICON_ARG=""
if [ -f "${ICON_NAME}" ]; then
    ICON_ARG="--icon=${ICON_NAME}"
    echo "使用图标: ${ICON_NAME}"
fi

# PyInstaller 打包
python3.8 -m PyInstaller \
    --onefile \
    --name="${OUTPUT_NAME}" \
    --noconsole \
    --clean \
    ${ICON_ARG} \
    "${SCRIPT_NAME}"

# 校验输出
OUTPUT_PATH="dist/${OUTPUT_NAME}"
if [ -f "${OUTPUT_PATH}" ]; then
    echo "=== 打包成功 ==="
    echo "输出文件: ${OUTPUT_PATH}"
    ls -lh "${OUTPUT_PATH}"
else
    echo "错误: 打包失败，未生成输出文件"
    ls -la dist/ 2>/dev/null || echo "dist 目录为空"
    exit 1
fi
