# ============================================
# 工程物资一本账更新工具 - .deb 打包脚本
# 目标: 生成可在统信 UOS / 麒麟 V10 安装的 .deb 包
# 原理: Docker 启动 Debian 容器 → PyInstaller 打包 → dpkg-deb 构建
# ============================================
param(
    [string]$Version = "1.0.0"
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# 颜色函数
function Write-Color($Text, $Color = "White") {
    Write-Host $Text -ForegroundColor $Color
}

Write-Color "========================================" "Cyan"
Write-Color "  工程物资一本账 - .deb 打包工具" "Cyan"
Write-Color "  版本: $Version" "Cyan"
Write-Color "  目标: cx_new_v6_${Version}_amd64.deb" "Cyan"
Write-Color "========================================" "Cyan"
Write-Host ""

# --------------------------------------------------
# [1/4] 检查 Docker
# --------------------------------------------------
Write-Color "[1/4] 检查 Docker 环境..." "Yellow"
try {
    $dockerVersion = docker --version 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "docker 命令执行失败"
    }
    Write-Color "  $dockerVersion" "Green"

    # 检查 Docker 是否在运行
    $dockerInfo = docker info 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Docker 未运行"
    }
    Write-Color "  Docker 运行正常" "Green"
} catch {
    Write-Color "  错误: Docker 未安装或未启动，请先安装 Docker Desktop" "Red"
    Write-Color "  下载地址: https://www.docker.com/products/docker-desktop/" "Yellow"
    pause
    exit 1
}

# --------------------------------------------------
# [2/4] 检查源文件
# --------------------------------------------------
Write-Color "[2/4] 检查源文件..." "Yellow"
$ScriptFile = Join-Path $ScriptDir "cx_new_v6.py"
if (-not (Test-Path $ScriptFile)) {
    Write-Color "  错误: 未找到 ${ScriptFile}" "Red"
    pause
    exit 1
}
Write-Color "  ${ScriptFile}  OK" "Green"

$BuildScript = Join-Path $ScriptDir "build_inside_deb.sh"
if (-not (Test-Path $BuildScript)) {
    Write-Color "  错误: 未找到 ${BuildScript}" "Red"
    pause
    exit 1
}
Write-Color "  ${BuildScript}  OK" "Green"

$Dockerfile = Join-Path $ScriptDir "Dockerfile.deb"
if (-not (Test-Path $Dockerfile)) {
    Write-Color "  错误: 未找到 ${Dockerfile}" "Red"
    pause
    exit 1
}
Write-Color "  ${Dockerfile}  OK" "Green"

Write-Host ""

# --------------------------------------------------
# [3/4] 构建 Docker 镜像
# --------------------------------------------------
Write-Color "[3/4] 构建 Docker 打包镜像 (python:3.9-slim-bullseye, glibc 2.31)..." "Yellow"
Write-Color "  基于 Debian Bullseye，兼容统信 UOS / 麒麟 V10 SP1+" "DarkYellow"

docker build -t cx-deb-builder:v1 -f $Dockerfile $ScriptDir

if ($LASTEXITCODE -ne 0) {
    Write-Color "  镜像构建失败！请检查上方的错误信息。" "Red"
    pause
    exit 1
}

Write-Color "  镜像构建完成" "Green"
Write-Host ""

# --------------------------------------------------
# [4/4] 在容器中打包 .deb
# --------------------------------------------------
Write-Color "[4/4] 在容器中执行 .deb 打包..." "Yellow"
Write-Color "  步骤: PyInstaller 打包 - 组装 deb 目录 - dpkg-deb 构建" "DarkYellow"

# 清理旧的输出
$DistDir = Join-Path $ScriptDir "dist_deb"
if (Test-Path $DistDir) {
    Remove-Item -Recurse -Force $DistDir
    Write-Color "  已清理旧输出目录" "DarkYellow"
}

# 运行容器
docker run --rm `
    -v "${ScriptDir}:/work" `
    -w /work `
    cx-deb-builder:v1 `
    bash /work/build_inside_deb.sh

if ($LASTEXITCODE -ne 0) {
    Write-Color "  .deb 构建失败！请检查上方的错误信息。" "Red"
    pause
    exit 1
}

# --------------------------------------------------
# 验证输出
# --------------------------------------------------
Write-Host ""
Write-Color "========================================" "Green"
Write-Color "  打包完成!" "Green"
Write-Color "========================================" "Green"
Write-Host ""

$DebFile = Join-Path $DistDir "cx_new_v6_${Version}_amd64.deb"
if (Test-Path $DebFile) {
    $Size = [math]::Round((Get-Item $DebFile).Length / 1MB, 1)
    Write-Color "  输出文件: dist_deb\cx_new_v6_${Version}_amd64.deb" "Green"
    Write-Color "  文件大小: ${Size} MB" "Green"
    Write-Host ""
    Write-Color "  安装方法 (在统信 UOS / 麒麟系统上):" "Cyan"
    Write-Color "    sudo dpkg -i cx_new_v6_${Version}_amd64.deb" "White"
    Write-Color "    sudo apt install -f" "White"
    Write-Host ""
    Write-Color "  启动方法:" "Cyan"
    Write-Color "    终端: cx_new_v6" "White"
    Write-Color "    启动器: 搜索 '一本账'" "White"
} else {
    Write-Color "  错误: 未找到输出文件 dist_deb\cx_new_v6_${Version}_amd64.deb" "Red"
    Write-Color "  dist_deb 目录内容:" "Red"
    Get-ChildItem $DistDir -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "    $($_.Name)" }
    pause
    exit 1
}

Write-Host ""
Write-Color "========================================" "Cyan"
Write-Color "  兼容性说明" "Cyan"
Write-Color "========================================" "Cyan"
Write-Color "  统信 UOS 20+    (Debian 10+, glibc >= 2.28)  支持" "Green"
Write-Color "  麒麟 V10 SP1+    (glibc >= 2.28)  支持" "Green"
Write-Color "  需要图形界面环境 (tkinter GUI)" "Yellow"
Write-Host ""

pause
