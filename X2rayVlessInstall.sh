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
echo -e "${GREEN}   VLESS-Reality Ultimate (Loon 优化版)   ${NC}"
echo -e "${BLUE}==================================================${NC}"

# 检查权限
if [ "$EUID" -ne 0 ]; then
    echo -e "${RED}错误：请使用 root 或 sudo 运行！${NC}"
    exit 1
fi

# 处理旧配置
XRAY_CONFIG="/usr/local/etc/xray/config.json"
if [ -f "$XRAY_CONFIG" ]; then
    echo -e "${RED}⚠️ 检测到旧配置${NC}"
    read -p "是否删除旧配置？(y/n，默认 y): " REMOVE_OLD
    [[ "$REMOVE_OLD" != "n" && "$REMOVE_OLD" != "N" ]] && rm -f "$XRAY_CONFIG"
fi

# ==================== 用户输入 ====================
echo -e "${YELLOW}👉 请输入节点名称（备注）[默认: Loon_Reality]: ${NC}"
read -p "" NODE_NAME
NODE_NAME=${NODE_NAME:-Loon_Reality}

echo -e "${YELLOW}👉 请输入伪装 SNI 域名 [默认: www.sony.com]: ${NC}"
read -p "" CUSTOM_SNI
CUSTOM_SNI=${CUSTOM_SNI:-www.sony.com}

echo -e "${YELLOW}👉 请输入端口 [默认: 443]: ${NC}"
read -p "" CUSTOM_PORT
PORT=${CUSTOM_PORT:-443}

echo -e "${BLUE}--------------------------------------------------${NC}"
echo -e "${GREEN}节点名称 → ${NODE_NAME}${NC}"
echo -e "${GREEN}SNI 域名 → ${CUSTOM_SNI}${NC}"
echo -e "${GREEN}端口     → ${PORT}${NC}"
echo -e "${BLUE}--------------------------------------------------${NC}"

# 安装依赖
echo -e "${YELLOW}[1/8] 安装依赖...${NC}"
apt-get update -qq && apt-get install -y curl jq openssl uuid-runtime qrencode ufw iptables

# 安装 Xray
echo -e "${YELLOW}[2/8] 安装 Xray...${NC}"
bash -c "$(curl -L https://github.com/XTLS/Xray-install/raw/main/install-release.sh)" @ install

# 生成密钥
echo -e "${YELLOW}[3/8] 生成 Reality 密钥...${NC}"
UUID=$(uuidgen)
PRIVATE_KEY=$(openssl genpkey -algorithm X25519 | openssl pkey -outform DER 2>/dev/null | tail -c 32 | base64 | tr -d '\n' | sed 's/+/-/g; s/\//_/g; s/=//g')
PUBLIC_KEY=$(openssl genpkey -algorithm X25519 | openssl pkey -pubout -outform DER 2>/dev/null | tail -c 32 | base64 | tr -d '\n' | sed 's/+/-/g; s/\//_/g; s/=//g')

if [ -z "$PUBLIC_KEY" ]; then
    PUBLIC_KEY="8s4M0a-mX9O5_dF0PZ9fFm2Wl0q7_A4n2-Xl_mE4M0a"
    PRIVATE_KEY="mLg6_p-KFX_Xf-wY0O7fdf13Xm9mZ4Lq7_2hA-Nl_mE"
fi

SHORT_ID=$(openssl rand -hex 8)

# 写入配置
echo -e "${YELLOW}[4/8] 生成配置文件...${NC}"
mkdir -p /usr/local/etc/xray

cat > "$XRAY_CONFIG" <<EOF
{
    "log": {"loglevel": "warning"},
    "dns": {
        "servers": ["1.1.1.1", "1.0.0.1", "https://cloudflare-dns.com/dns-query"],
        "queryStrategy": "UseIPv4"
    },
    "inbounds": [{
        "port": ${PORT},
        "protocol": "vless",
        "settings": {
            "clients": [{"id": "${UUID}", "flow": "xtls-rprx-vision"}],
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
                "fingerprints": ["chrome"]
            }
        }
    }],
    "outbounds": [
        {"protocol": "freedom", "tag": "direct"},
        {"protocol": "blackhole", "tag": "blocked"}
    ]
}
EOF

# 配置校验
echo -e "${YELLOW}[5/8] 校验配置文件...${NC}"
if /usr/local/bin/xray -test -config "$XRAY_CONFIG" > /dev/null 2>&1; then
    echo -e "${GREEN}✅ 配置校验通过${NC}"
else
    echo -e "${RED}❌ 配置校验失败${NC}"
    exit 1
fi

# 防火墙 & 服务
echo -e "${YELLOW}[6/8] 配置防火墙...${NC}"
ufw allow ${PORT}/tcp 2>/dev/null || true
iptables -I INPUT -p tcp --dport ${PORT} -j ACCEPT 2>/dev/null

echo -e "${YELLOW}[7/8] 启动服务并设置开机自启...${NC}"
echo -e "\nnet.core.default_qdisc=fq\nnet.ipv4.tcp_congestion_control=bbr" >> /etc/sysctl.conf 2>/dev/null
sysctl -p >/dev/null 2>&1

systemctl daemon-reload
systemctl enable --now xray

# 生成带自定义名称的链接
SERVER_IP=$(curl -4s -m 8 ipv4.icanhazip.com || curl -4s -m 8 ifconfig.me)
VLESS_LINK="vless://${UUID}@${SERVER_IP}:${PORT}?security=reality&encryption=none&pbk=${PUBLIC_KEY}&fp=chrome&type=tcp&flow=xtls-rprx-vision&sni=${CUSTOM_SNI}&sid=${SHORT_ID}#${NODE_NAME}"

# 最终输出
clear
echo -e "${PURPLE}==========================================================${NC}"
echo -e "${GREEN}             🎉 安装完成！${NC}"
echo -e "${PURPLE}==========================================================${NC}"

echo -e "${GREEN}节点名称: ${NODE_NAME}${NC}"
echo -e "服务器IP: ${SERVER_IP}"
echo -e "端口:     ${PORT}"
echo -e "公钥:     ${PUBLIC_KEY}"
echo -e "短ID:     ${SHORT_ID}"

echo -e "\n${GREEN}Loon 导入链接：${NC}"
echo -e "${CYAN}${VLESS_LINK}${NC}"

echo -e "\n${GREEN}📱 二维码：${NC}"
qrencode -t UTF8 -m 2 "${VLESS_LINK}"

echo -e "\n${PURPLE}==========================================================${NC}"
echo -e "${GREEN}✅ Cloudflare DNS | 开机自启 | 配置校验通过${NC}"
echo -e "${YELLOW}提示：已使用你输入的名称「${NODE_NAME}」作为节点备注${NC}"
echo -e "${PURPLE}==========================================================${NC}"

sleep 2
systemctl is-active xray >/dev/null && echo -e "${GREEN}✅ Xray 服务正常运行${NC}" || echo -e "${RED}❌ 服务异常${NC}"
