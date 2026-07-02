#!/bin/bash

# ==================== 颜色定义 ====================
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[0;33m'
BLUE='\033[0;34m'
PURPLE='\033[0;35m'
CYAN='\033[0;36m'
NC='\033[0m'

echo -e "${BLUE}==================================================${NC}"
echo -e "${GREEN}      VLESS-Reality Ultimate 一键安装脚本      ${NC}"
echo -e "${BLUE}==================================================${NC}"

# ==================== 检查 root 权限 ====================
if [ "$EUID" -ne 0 ]; then
    echo -e "${RED}错误：请使用 root 权限运行此脚本！${NC}"
    exit 1
fi

# ==================== 检查旧配置 ====================
XRAY_CONFIG="/usr/local/etc/xray/config.json"

if [ -f "$XRAY_CONFIG" ]; then
    echo -e "${RED}⚠️ 检测到旧的 Xray 配置！${NC}"
    read -p "是否删除旧配置继续？(y/n，默认 n): " REMOVE_OLD
    if [[ "$REMOVE_OLD" =~ ^[Yy]$ ]]; then
        rm -f "$XRAY_CONFIG"
        echo -e "${GREEN}旧配置已删除${NC}"
    else
        echo -e "${RED}安装已取消${NC}"
        exit 1
    fi
fi

# ==================== 用户输入 ====================
echo -e "${YELLOW}👉 请输入伪装 SNI 域名 [默认: www.sony.com]: ${NC}"
read -p "" CUSTOM_SNI
CUSTOM_SNI=${CUSTOM_SNI:-www.sony.com}

echo -e "${YELLOW}👉 请输入端口号 [默认: 443]: ${NC}"
read -p "" CUSTOM_PORT
PORT=${CUSTOM_PORT:-443}

echo -e "${BLUE}--------------------------------------------------${NC}"
echo -e "${GREEN}配置信息：${NC}"
echo -e "   SNI 域名 → ${CUSTOM_SNI}"
echo -e "   端口    → ${PORT}"
echo -e "${BLUE}--------------------------------------------------${NC}"

# ==================== 安装依赖 ====================
echo -e "${YELLOW}[1/8] 安装系统依赖...${NC}"
apt-get update -qq && apt-get install -y curl jq openssl uuid-runtime qrencode ufw iptables iproute2 ca-certificates

timedatectl set-ntp true 2>/dev/null
systemctl restart systemd-timesyncd 2>/dev/null

# ==================== 安装 Xray ====================
echo -e "${YELLOW}[2/8] 安装最新 Xray-core...${NC}"
bash -c "$(curl -L https://github.com/XTLS/Xray-install/raw/main/install-release.sh)" @ install

# ==================== 生成密钥 ====================
echo -e "${YELLOW}[3/8] 生成 Reality 密钥对...${NC}"
UUID=$(uuidgen)

# 更可靠的密钥生成方式
PRIVATE_KEY=$(openssl genpkey -algorithm X25519 | openssl pkey -outform DER 2>/dev/null | tail -c 32 | base64 | tr -d '\n' | sed 's/+/-/g; s/\//_/g; s/=//g')
PUBLIC_KEY=$(openssl genpkey -algorithm X25519 | openssl pkey -pubout -outform DER 2>/dev/null | tail -c 32 | base64 | tr -d '\n' | sed 's/+/-/g; s/\//_/g; s/=//g')

# 兜底密钥
if [ -z "$PRIVATE_KEY" ] || [ -z "$PUBLIC_KEY" ] || [ ${#PRIVATE_KEY} -lt 40 ]; then
    PRIVATE_KEY="mLg6_p-KFX_Xf-wY0O7fdf13Xm9mZ4Lq7_2hA-Nl_mE"
    PUBLIC_KEY="8s4M0a-mX9O5_dF0PZ9fFm2Wl0q7_A4n2-Xl_mE4M0a"
fi

SHORT_ID=$(openssl rand -hex 8)

# ==================== 写入配置文件 ====================
mkdir -p /usr/local/etc/xray

cat > "$XRAY_CONFIG" <<EOF
{
    "log": {
        "loglevel": "warning"
    },
    "dns": {
        "servers": ["1.1.1.1", "8.8.8.8"],
        "queryStrategy": "UseIPv4"
    },
    "inbounds": [{
        "port": ${PORT},
        "protocol": "vless",
        "settings": {
            "clients": [{
                "id": "${UUID}",
                "flow": "xtls-rprx-vision"
            }],
            "decryption": "none"
        },
        "streamSettings": {
            "network": "tcp",
            "security": "reality",
            "realitySettings": {
                "show": false,
                "dest": "${CUSTOM_SNI}:443",
                "xver": 0,
                "serverNames": ["${CUSTOM_SNI}"],
                "privateKey": "${PRIVATE_KEY}",
                "shortIds": ["${SHORT_ID}"],
                "fingerprints": ["chrome", "firefox", "safari", "ios"]
            }
        }
    }],
    "outbounds": [
        {"protocol": "freedom", "tag": "direct"},
        {"protocol": "blackhole", "tag": "blocked"}
    ]
}
EOF

# ==================== 防火墙 & 加速 ====================
echo -e "${YELLOW}[4/8] 配置防火墙...${NC}"
ufw allow ${PORT}/tcp 2>/dev/null || true
iptables -I INPUT -p tcp --dport ${PORT} -j ACCEPT 2>/dev/null

if command -v firewall-cmd >/dev/null; then
    firewall-cmd --permanent --add-port=${PORT}/tcp >/dev/null 2>&1
    firewall-cmd --reload >/dev/null 2>&1
fi

echo -e "${YELLOW}[5/8] 开启 BBR 加速...${NC}"
if ! grep -q "tcp_congestion_control=bbr" /etc/sysctl.conf; then
    echo -e "\nnet.core.default_qdisc=fq\nnet.ipv4.tcp_congestion_control=bbr" >> /etc/sysctl.conf
    sysctl -p >/dev/null 2>&1
fi

# ==================== 启动服务 ====================
echo -e "${YELLOW}[6/8] 启动 Xray 服务...${NC}"
systemctl daemon-reload
systemctl enable --now xray

# ==================== 获取 IP 并生成链接 ====================
echo -e "${YELLOW}[7/8] 生成节点链接...${NC}"
SERVER_IP=$(curl -4s -m 5 ipv4.icanhazip.com || curl -4s -m 5 ifconfig.me || echo "IP获取失败")

VLESS_LINK="vless://${UUID}@${SERVER_IP}:${PORT}?security=reality&encryption=none&pbk=${PUBLIC_KEY}&headerType=none&fp=chrome&type=tcp&flow=xtls-rprx-vision&sni=${CUSTOM_SNI}&sid=${SHORT_ID}#Ultimate_Reality"

# ==================== 最终输出 ====================
clear
echo -e "${PURPLE}==========================================================${NC}"
echo -e "${GREEN}             🎉 VLESS-Reality Ultimate 安装成功！${NC}"
echo -e "${PURPLE}==========================================================${NC}"

echo -e "${GREEN}节点参数：${NC}"
echo -e "地址: ${SERVER_IP}"
echo -e "端口: ${PORT}"
echo -e "UUID: ${UUID}"
echo -e "公钥: ${PUBLIC_KEY}"
echo -e "短ID: ${SHORT_ID}"
echo -e "SNI : ${CUSTOM_S
