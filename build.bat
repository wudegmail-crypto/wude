@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ================================================
echo   合同数据合并工具 - Linux 麒麟系统打包脚本
echo   目标 glibc: ^>= 2.17 (覆盖麒麟 V10 全系列)
echo ================================================
echo.

:: 切换到脚本所在目录
cd /d "%~dp0"

:: ========================================
:: [1/4] 构建 Docker 镜像
:: ========================================
echo [1/4] 构建 CentOS 7 打包镜像...
docker build -f Dockerfile.centos7 -t py-builder .
if %errorlevel% neq 0 (
    echo [错误] 镜像构建失败！
    pause
    exit /b 1
)
echo [OK] 镜像构建完成
echo.

:: ========================================
:: [2/4] 查看容器环境信息
:: ========================================
echo [2/4] 检查容器环境...
echo.
echo --- Python 版本 ---
docker run --rm py-builder python --version
echo.
echo --- 已安装的包 ---
docker run --rm py-builder pip list 2^>nul | findstr /I "numpy pandas openpyxl pyinstaller tk"
echo.

:: ========================================
:: [3/4] 清理旧输出
:: ========================================
echo [3/4] 清理旧打包输出...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "cx_new_pkg.spec" del /q "cx_new_pkg.spec"
echo.

:: ========================================
:: [4/4] 在容器中打包
:: ========================================
echo [4/4] 开始 PyInstaller 打包...
docker run --rm ^
    -v "%cd%:/work" ^
    -w /work ^
    py-builder ^
    bash -c " \
        pyinstaller \
            --onedir \
            --name cx_new \
            --add-data '汇总表模板:汇总表模板' \
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
            cx_new_pkg.py"

if %errorlevel% neq 0 (
    echo [错误] 打包失败！请检查上方的错误信息。
    pause
    exit /b 1
)

echo.
echo ================================================
echo   打包完成！
echo ================================================
echo.
echo 输出文件: dist\cx_new\cx_new
echo 包大小:
dir /s "dist\cx_new" 2>nul | findstr "个文件"

echo.
echo ================================================
echo   检查 glibc 版本需求（最低兼容版本）
echo ================================================
docker run --rm -v "%cd%:/work" -w /work py-builder ^
    bash -c " \
        echo '--- 启动引导器需求 ---' && \
        objdump -T dist/cx_new/cx_new 2>/dev/null | grep GLIBC | sed 's/.*GLIBC_/  GLIBC_/' | sort -V -u && \
        echo '' && \
        echo '--- numpy/pandas .so 文件需求 ---' && \
        find dist/cx_new/_internal -name '*.so' -exec objdump -T {} \; 2>/dev/null | grep GLIBC | sed 's/.*GLIBC_/  GLIBC_/' | sort -V -u"

echo.
echo ================================================
echo   部署到麒麟系统
echo ================================================
echo.
echo 1. 将 dist\cx_new\ 整个目录复制到麒麟服务器
echo    例如: scp -r dist\cx_new user@kylin-server:/opt/app/
echo.
echo 2. 在麒麟系统上安装 Tcl/Tk 运行时:
echo     sudo yum install -y tk tcl
echo.
echo 3. 赋予执行权限并运行:
echo     cd /opt/app/cx_new
echo     chmod +x cx_new
echo     ./cx_new
echo.
echo 注意: 如果无图形界面, 程序会报错 "no display".
echo       需要通过 SSH -X 转发或在桌面环境下运行.
echo ================================================

pause
