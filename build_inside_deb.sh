#!/bin/bash
# ============================================
# 工程物资一本账 - PyInstaller 单文件打包
# ============================================
set -e

SCRIPT_FILE="cx_new_v6.py"
APP_NAME="cx_new_v6"

echo "========================================"
echo "  工程物资一本账 - 单文件打包"
echo "========================================"
echo "  Python 版本: $(python3 --version)"
echo "  glibc 版本: $(ldd --version 2>/dev/null | head -1 || echo 'unknown')"
echo ""

# 清理旧构建
rm -rf build dist *.spec 2>/dev/null || true

# PyInstaller --onefile 打包
python3 -m PyInstaller \
    --onefile \
    --name "${APP_NAME}" \
    --hidden-import tkinter \
    --hidden-import _tkinter \
    --hidden-import pandas \
    --hidden-import openpyxl \
    --hidden-import numpy \
    --hidden-import python_dateutil \
    --hidden-import six \
    --hidden-import et_xmlfile \
    --hidden-import tzdata \
    --hidden-import openpyxl.styles \
    --hidden-import openpyxl.utils \
    --hidden-import openpyxl.cell \
    --clean \
    --noconfirm \
    "${SCRIPT_FILE}"

# 验证输出
ONE_FILE="dist/${APP_NAME}"
if [ ! -f "${ONE_FILE}" ]; then
    echo "错误: PyInstaller 打包失败"
    ls -la dist/ 2>/dev/null || echo "dist 目录为空"
    exit 1
fi

# 复制到输出目录
OUTPUT_DIR="/work/dist_deb"
mkdir -p "${OUTPUT_DIR}"
cp "${ONE_FILE}" "${OUTPUT_DIR}/${APP_NAME}"

echo ""
echo "========================================"
echo "  打包完成!"
echo "  输出: dist_deb/${APP_NAME}"
echo "  大小: $(du -sh ${ONE_FILE} | cut -f1)"
echo "  用法: chmod +x ${APP_NAME} && ./${APP_NAME}"
echo "========================================"
