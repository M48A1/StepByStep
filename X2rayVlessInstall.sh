#!/bin/bash

# ==================== 颜色定义 ====================
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[0;33m'
BLUE='\033[0;34m'
PURPLE='\033[0;35m'
CYAN='\033[0;36m'
NC='\033[0m'

clear
echo -e "${BLUE}==================================================${NC}"
echo -e "${GREEN}   VLESS-Reality Ultimate Final 完美修复版   ${NC}"
echo -e "${BLUE}==================================================${NC}"

if [ "$EUID" -ne 0 ]; then
    echo -e "${RED}错误：请使用 root 权限运行！${NC}"
    exit 1
fi

XRAY_CONFIG="/usr/local/etc/xray/config.json"

if [ -f "$XRAY_CONFIG" ]; then
    echo -e "${RED}⚠️ 检测到旧配置${NC}"
    read -p "是否删除旧配置继续？(y/n，默认 y): " REMOVE_OLD
    [[ "$REMOVE_OLD" != "n" && "$REMOVE_OLD" != "N" ]] && rm -f "$XRAY_CONFIG"
fi

# 用户输入
echo -e "${YELLOW}👉 节点名称 [默认: My_Reality]: ${NC}"
read -p "" NODE_NAME
NODE_NAME=${NODE_NAME:-My_Reality}

echo -e "${YELLOW}👉 SNI 域名 [默认: www.sony.com]: ${NC}"
read -p "" CUSTOM_SNI
CUSTOM_SNI=${CUSTOM_SNI:-www.sony.com}

echo -e "${YELLOW}👉 端口 [默认: 443]: ${NC}"
read -p "" CUSTOM_PORT
PORT=${CUSTOM_PORT:-443}

# 安装依赖 (加入 uuid-runtime，确保依赖完整)
echo -e "${YELLOW}正在安装系统依赖...${NC}"
apt-get update -qq && apt-get install -y curl jq qrencode ufw uuid-runtime

# 安装 Xray 核心
bash -c "$(curl -Ls https://github.com/XTLS/Xray-install/raw/main/install-release.sh)" @ install

# 密钥与标识生成 (【核心修复】彻底解决空密钥导致校验失败的问题)
echo -e "${YELLOW}生成密钥...${NC}"
XRAY_KEYS=$(/usr/local/bin/xray x25519)
PRIVATE_KEY=$(echo "$XRAY_KEYS" | grep -i "private" | awk -F':' '{print $2}' | sed 's/ //g')
PUBLIC_KEY=$(echo "$XRAY_KEYS" | grep -i "public" | awk -F':' '{print $2}' | sed 's/ //g')

SHORT_ID=$(openssl rand -hex 8)

# 优先使用 xray 自身生成 UUID
if /usr/local/bin/xray uuid >/dev/null 2>&1; then
    UUID=$(/usr/local/bin/xray uuid)
else
    UUID=$(uuidgen)
fi

# 【新增安全校验】若密钥依然提取失败，则中止脚本，防止写入破坏的 json 文件
if [ -z "$PRIVATE_KEY" ] || [ -z "$PUBLIC_KEY" ]; then
    echo -e "${RED}❌ 错误：未能从 Xray 成功提取密钥！原始输出如下：${NC}"
    echo "$XRAY_KEYS"
    exit 1
fi

# 配置写入
cat > "$XRAY_CONFIG" <<EOF
{
    "log": { "loglevel": "warning" },
    "dns": {
        "servers": ["1.1.1.1", "1.0.0.1"],
        "queryStrategy": "UseIPv4"
    },
    "inbounds": [{
        "port": ${PORT},
        "protocol": "vless",
        "settings": {
            "clients": [{ "id": "${UUID}", "flow": "xtls-rprx-vision" }],
            "decryption": "none"
        },
        "streamSettings": {
            "network": "tcp",
            "security": "reality",
            "realitySettings": {
                "show": false,
                "dest": "${CUSTOM_SNI}:443",
                "serverNames": ["${CUSTOM_SNI}"],
                "privateKey": "${PRIVATE_KEY}",
                "shortIds": ["${SHORT_ID}"]
            }
        }
    }],
    "outbounds": [{ "protocol": "freedom" }]
}
EOF

# 校验
echo -e "${YELLOW}校验配置...${NC}"
if /usr/local/bin/xray -test -config "$XRAY_CONFIG" > /dev/null 2>&1; then
    echo -e "${GREEN}✅ 配置校验通过${NC}"
else
    echo -e "${RED}❌ 配置校验失败！${NC}"
    echo -e "${YELLOW}详细错误：${NC}"
    /usr/local/bin/xray -test -config "$XRAY_CONFIG"
    exit 1
fi

# 安全配置 UFW 防火墙 (防止断开 SSH)
echo -e "${YELLOW}配置 UFW 防火墙...${NC}"
if command -v ufw >/dev/null; then
    # 自动探测当前 SSH 端口，强制允许，防止用户被挡在服务器外
    SSH_PORT=$(ss -tlnp | grep sshd | awk '{print $4}' | awk -F':' '{print $2}' | head -n 1)
    SSH_PORT=${SSH_PORT:-22}
    ufw allow "${SSH_PORT}"/tcp > /dev/null 2>&1
    
    ufw allow "${PORT}"/tcp > /dev/null 2>&1
    ufw --force enable
    echo -e "${GREEN}UFW 已自动放行 SSH 端口(${SSH_PORT}) 及 节点端口(${PORT})${NC}"
fi

# BBR 开启（增强了容器化环境的兼容性）
echo -e "${YELLOW}配置 BBR 加速...${NC}"
grep -q "net.core.default_qdisc=fq" /etc/sysctl.conf || {
    cat >> /etc/sysctl.conf <<EOF
net.core.default_qdisc=fq
net.ipv4.tcp_congestion_control=bbr
EOF
}
sysctl -p > /dev/null 2>&1 || echo -e "${YELLOW}提示: 当前环境无法直接更新内核 sysctl 参数，已优雅跳过 BBR 开启步骤。${NC}"

# 启动服务
systemctl daemon-reload
systemctl enable --now xray

# IP 获取
SERVER_IP=$(curl -4s -m 5 https://api.ipify.org 2>/dev/null || \
            curl -4s -m 5 https://ipv4.icanhazip.com 2>/dev/null || \
            curl -4s -m 5 https://ifconfig.me 2>/dev/null || \
            hostname -I | awk '{print $1}' 2>/dev/null || echo "IP获取失败")

# 节点名称 URL 编码处理（防止因中文或空格引发导入解析错误）
ENCODED_NODE_NAME=$(echo -n "${NODE_NAME}" | jq -sRr @uri)

VLESS_LINK="vless://${UUID}@${SERVER_IP}:${PORT}?security=reality&encryption=none&pbk=${PUBLIC_KEY}&fp=chrome&type=tcp&sni=${CUSTOM_SNI}&sid=${SHORT_ID}#${ENCODED_NODE_NAME}"

clear
echo -e "${PURPLE}==========================================================${NC}"
echo -e "${GREEN}             🎉 安装完成！${NC}"
echo -e "${PURPLE}==========================================================${NC}"

echo -e "${GREEN}节点名称: ${NODE_NAME}${NC}"
echo -e "IP: ${SERVER_IP}"
echo -e "公钥: ${PUBLIC_KEY}"
echo -e "短ID: ${SHORT_ID}"

echo -e "\n${GREEN}Loon / 通用 导入链接：${NC}"
echo -e "${CYAN}${VLESS_LINK}${NC}\n"

qrencode -t UTF8 -m 2 "${VLESS_LINK}"

echo -e "${PURPLE}==========================================================${NC}"
echo -e "${GREEN}✅ 最终版完成${NC}"
echo -e "${PURPLE}==========================================================${NC}"
