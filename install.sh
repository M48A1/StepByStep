#!/bin/bash

# 字体颜色定义
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[0;33m'
BLUE='\033[0;34m'
NC='\033[0m'

echo -e "${BLUE}==================================================${NC}"
echo -e "${GREEN}      欢迎使用自定义 VLESS-Reality 一键安装脚本    ${NC}"
echo -e "${BLUE}==================================================${NC}"

# 1. 基础环境检查与安装依赖
echo -e "${YELLOW}[1/4] 正在安装系统依赖环境...${NC}"
apt-get update -y && apt-get install -y curl jq openssl uuid-runtime

# 2. 调用官方脚本安装最新的 Xray-core
echo -e "${YELLOW}[2/4] 正在下载并安装官方 Xray-core...${NC}"
bash -c "$(curl -L https://github.com/XTLS/Xray-install/raw/main/install-release.sh)" @ install

# 3. 自动生成 Reality 所需的各种随机密钥和配置参数
echo -e "${YELLOW}[3/4] 正在自动生成高性能安全配置...${NC}"
UUID=$(uuidgen)
PORT=443

# 生成 Reality 专属的公钥和私钥
XRAY_KEYS=$(xray x25519)
PRIVATE_KEY=$(echo "$XRAY_KEYS" | grep "Private key:" | awk '{print $3}')
PUBLIC_KEY=$(echo "$XRAY_KEYS" | grep "Public key:" | awk '{print $3}')

# 生成 8 字节的十六进制随机 ShortID
SHORT_ID=$(openssl rand -hex 8)

# 4. 写入 Xray 配置文件 (config.json)
mkdir -p /usr/local/etc/xray

cat <<EOF > /usr/local/etc/xray/config.json
{
    "log": {
        "loglevel": "warning"
    },
    "inbounds": [
        {
            "port": ${PORT},
            "protocol": "vless",
            "settings": {
                "clients": [
                    {
                        "id": "${UUID}",
                        "flow": "xtls-rprx-vision"
                    }
                ],
                "decryption": "none"
            },
            "streamSettings": {
                "network": "tcp",
                "security": "reality",
                "realitySettings": {
                    "show": false,
                    "dest": "17.253.145.10:443",
                    "xver": 0,
                    "serverNames": [
                        "www.apple.com",
                        "images.apple.com"
                    ],
                    "privateKey": "${PRIVATE_KEY}",
                    "minClientVer": "",
                    "maxClientVer": "",
                    "timedALPN": [
                        "h2",
                        "http/1.1"
                    ],
                    "shortIds": [
                        "${SHORT_ID}"
                    ]
                }
            }
        }
    ],
    "outbounds": [
        {
            "protocol": "freedom",
            "tag": "direct"
        },
        {
            "protocol": "blackhole",
            "tag": "blocked"
        }
    ]
}
EOF

# 5. 启动服务并设置开机自启
echo -e "${YELLOW}[4/4] 正在启动服务并设置开机自启...${NC}"
systemctl daemon-reload
systemctl enable xray
systemctl restart xray

# 获取服务器公网 IP
SERVER_IP=$(curl -s ifconfig.me)

# 6. 打印精美的安装成功菜单与节点链接
echo -e "${BLUE}==================================================${NC}"
echo -e "${GREEN}             🎉 VLESS-Reality 搭建成功！            ${NC}"
echo -e "${BLUE}==================================================${NC}"
echo -e "${YELLOW}服务器 IP:${NC} ${SERVER_IP}"
echo -e "${YELLOW}端口 (Port):${NC} ${PORT}"
echo -e "${YELLOW}用户 UUID:${NC} ${UUID}"
echo -e "${YELLOW}流控 (Flow):${NC} xtls-rprx-vision"
echo -e "${YELLOW}公开密钥 (PublicKey):${NC} ${PUBLIC_KEY}"
echo -e "${YELLOW}Short ID:${NC} ${SHORT_ID}"
echo -e "${YELLOW}伪装域名 (SNI):${NC} www.apple.com"
echo -e "${BLUE}--------------------------------------------------${NC}"
echo -e "${GREEN}您的客户端通用 VLESS 链接 (复制即可导入):${NC}"
echo -e "${CYAN}vless://${UUID}@${SERVER_IP}:${PORT}?security=reality&encryption=none&pbk=${PUBLIC_KEY}&headerType=none&fp=chrome&spx=%2F&type=tcp&flow=xtls-rprx-vision&sni=www.apple.com&sid=${SHORT_ID}#My_Custom_Reality${NC}"
echo -e "${BLUE}==================================================${NC}"
