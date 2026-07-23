# 麒麟系统打包 - Docker 配置
# 基于 AlmaLinux 8 (glibc 2.28)，兼容麒麟 V10

FROM almalinux:8

# 启用 EPEL 和 PowerTools 仓库
RUN dnf install -y epel-release && \
    dnf config-manager --set-enabled powertools 2>/dev/null || \
    dnf config-manager --set-enabled PowerTools 2>/dev/null || true && \
    dnf install -y python38 python38-pip python38-tkinter gcc python38-devel && \
    dnf clean all && \
    python3.8 -m pip install --upgrade pip

# 安装 Python 依赖
RUN pip3.8 install --no-cache-dir \
    pandas \
    openpyxl \
    pyinstaller

# 创建工作目录
WORKDIR /build

# 拷贝入口脚本
COPY entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

ENTRYPOINT ["/entrypoint.sh"]
