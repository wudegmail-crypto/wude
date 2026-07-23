# ============================================
# 麒麟系统打包脚本 (PowerShell)
# 使用方法:
#   1. 安装 Docker Desktop 并启动
#   2. 在此目录打开 PowerShell
#   3. 运行: .\build_kylin.ps1
# ============================================

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  麒麟系统 v12 打包工具" -ForegroundColor Cyan
Write-Host "  目标: 合同数据合并工具_v12_kylin" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 检查 Docker
Write-Host "[1/4] 检查 Docker..." -ForegroundColor Yellow
try {
    docker --version 2>$null
    if ($LASTEXITCODE -ne 0) {
        throw "Docker 未运行"
    }
    Write-Host "  Docker 就绪" -ForegroundColor Green
} catch {
    Write-Host "  错误: Docker 未安装或未启动，请先安装 Docker Desktop" -ForegroundColor Red
    exit 1
}

# 构建镜像
Write-Host "[2/4] 构建打包镜像 (almalinux:8, glibc 2.28)..." -ForegroundColor Yellow
docker build -t kylin-packager:v12 -f "$ScriptDir\Dockerfile" "$ScriptDir"
if ($LASTEXITCODE -ne 0) {
    Write-Host "  镜像构建失败" -ForegroundColor Red
    exit 1
}
Write-Host "  镜像构建完成" -ForegroundColor Green

# 运行打包
Write-Host "[3/4] 开始打包..." -ForegroundColor Yellow
docker run --rm `
    -v "$ScriptDir\cx_new.py:/build/cx_new.py:ro" `
    -v "$ScriptDir\myicon.ico:/build/myicon.ico:ro" `
    -v "$ScriptDir\dist_kylin:/build/dist" `
    kylin-packager:v12

if ($LASTEXITCODE -ne 0) {
    Write-Host "  打包失败" -ForegroundColor Red
    exit 1
}

# 验证输出
Write-Host "[4/4] 验证输出..." -ForegroundColor Yellow
$OutputFile = Join-Path $ScriptDir "dist_kylin\合同数据合并工具_v12_kylin"
if (Test-Path $OutputFile) {
    $Size = (Get-Item $OutputFile).Length / 1MB
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  打包成功!" -ForegroundColor Green
    Write-Host "  输出文件: dist_kylin\合同数据合并工具_v12_kylin" -ForegroundColor Green
    Write-Host "  文件大小: $([math]::Round($Size, 1)) MB" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
} else {
    Write-Host "  错误: 未找到输出文件" -ForegroundColor Red
    exit 1
}
