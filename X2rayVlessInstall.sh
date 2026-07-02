#!/bin/bash

# 字体颜色定义
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[0;33m'
BLUE='\033[0;34m'
PURPLE='\033[0;35m'
CYAN='\033[0;36m'
NC='\033[0m'

echo -e "${BLUE}==================================================${NC}"
echo -e "${GREEN}      欢迎使用自定义 VLESS-Reality 一键安装脚本    ${NC}"
echo -e "${BLUE}==================================================${NC}"

# 1. 检查现有配置文件
XRAY_CONFIG_PATH="/usr/local/etc/xray/config.json"

if [ -f "$XRAY_CONFIG_PATH" ]; then
    echo -e "${RED}⚠️ 检测到系统内已存在旧的 Xray 配置文件！${NC}"
    read -p "是否删除旧配置文件并继续安装？(y/n, 默认 n): " REMOVE_OLD
    if [ "$REMOVE_OLD" = "y" ] || [ "$REMOVE_OLD" = "Y" ]; then
        echo -e "${YELLOW}正在删除旧配置文件...${NC}"
        rm -f "$XRAY_CONFIG_PATH"
    else
        echo -e "${RED}安装已取消，未做任何修改。${NC}"
        exit 1
    fi
else
    echo -e "${GREEN}检查通过：未发现冲突的旧配置文件，开始安装。${NC}"
fi
echo -e "${BLUE}--------------------------------------------------${NC}"

# 2. 动态输入：伪装域名 (SNI)
echo -e "${YELLOW}👉 请输入你想要伪装的 SNI 域名（例如: www.sony.com）${NC}"
read -p "直接敲回车默认使用 [www.sony.com]: " CUSTOM_SNI
if [ -z "$CUSTOM_SNI" ]; then
    CUSTOM_SNI="www.sony.com"
fi

# 3. 动态输入：自定义端口 (Port)
echo -e "${YELLOW}👉 请输入你想要运行的端口号（1-65535）${NC}"
read -p "直接敲回车默认使用 [443]: " CUSTOM_PORT
if [ -z "$CUSTOM_PORT" ]; then
    PORT=443
else
    PORT=$CUSTOM_PORT
fi

echo -e "${BLUE}--------------------------------------------------${NC}"
echo -e "${GREEN}已选择伪装域名: ${CUSTOM_SNI}${NC}"
echo -e "${GREEN}已选择运行端口: ${PORT}${NC}"
echo -e "${BLUE}--------------------------------------------------${NC}"

# 4. 基础环境检查与安装依赖
echo -e "${YELLOW}[1/5] 正在安装系统依赖环境及二维码组件...${NC}"
apt-get update -y && apt-get install -y curl jq openssl uuid-runtime ufw iptables qrencode

# 5. 调用官方脚本安装最新的 Xray-core
echo -e "${YELLOW}[2/5] 正在下载并安装官方 Xray-core...${NC}"
bash -c "$(curl -L https://github.com/XTLS/Xray-install/raw/main/install-release.sh)" @ install

# 6. 自动生成 Reality 所需的各种随机密钥（修复重点：采用绝对路径与健壮截取）
echo -e "${YELLOW}[3/5] 正在自动生成高性能安全配置...${NC}"
UUID=$(uuidgen)

# 使用绝对路径调用 xray 生成密钥对
XRAY_KEYS=$(/usr/local/bin/xray x25519)
PRIVATE_KEY=$(echo "$XRAY_KEYS" | awk -F': ' '/Private key/{print $2}' | tr -d ' ')
PUBLIC_KEY=$(echo "$XRAY_KEYS" | awk -F': ' '/Public key/{print $2}' | tr -d ' ')

# 健壮性检查：如果依然为空，进行二次兜底尝试
if [ -z "$PUBLIC_KEY" ]; then
    XRAY_KEYS=$(xray x25519 2>/dev/null)
    PRIVATE_KEY=$(echo "$XRAY_KEYS" | awk -F': ' '/Private key/{print $2}' | tr -d ' ')
    PUBLIC_KEY=$(echo "$XRAY_KEYS" | awk -F': ' '/Public key/{print $2}' | tr -d ' ')
fi

# 生成 8 字节的十六进制随机 ShortID
SHORT_ID=$(openssl rand -hex 8)

# 7. 写入 Xray 配置文件 (config.json)
mkdir -p /usr/local/etc/xray

cat <<EOF > /usr/local/etc/xray/config.json
{
    "log": {
        "loglevel": "warning"
    },
    "dns": {
        "servers": [
            "8.8.8.8",
            "1.1.1.1"
        ],
        "queryStrategy": "UseIPv4"
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
                    "dest": "${CUSTOM_SNI}:443",
                    "xver": 0,
                    "serverNames": [
                        "${CUSTOM_SNI}"
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
            "tag": "direct",
            "settings": {
                "domainStrategy": "UseIPv4"
            }
        },
        {
            "protocol": "blackhole",
            "tag": "blocked"
        }
    ]
}
EOF

# 8. 自动放行 VPS 系统防火墙端口
echo -e "${YELLOW}[4/5] 正在自动配置防火墙放行端口 ${PORT}...${NC}"
if command -v ufw > /dev/null 2>&1; then
    ufw allow ${PORT}/tcp > /dev/null 2>&1
    ufw allow ${PORT}/udp > /dev/null 2>&1
fi
iptables -I INPUT -p tcp --dport ${PORT} -j ACCEPT > /dev/null 2>&1
iptables -I INPUT -p udp --dport ${PORT} -j ACCEPT > /dev/null 2>&1

# 9. 启动服务并设置开机自启
echo -e "${YELLOW}[5/5] 正在启动服务并设置开机自启...${NC}"
systemctl daemon-reload
systemctl enable xray
systemctl restart xray

# 强制获取服务器的 IPv4 地址
SERVER_IP=$(curl -4 -s ipv4.icanhazip.com)
if [ -z "$SERVER_IP" ]; then
    SERVER_IP=$(curl -s ifconfig.me)
fi

# 拼接完整的标准通用一键导入链接
VLESS_LINK="vless://${UUID}@${SERVER_IP}:${PORT}?security=reality&encryption=none&pbk=${PUBLIC_KEY}&headerType=none&fp=chrome&spx=%2F&type=tcp&flow=xtls-rprx-vision&sni=${CUSTOM_SNI}&sid=${SHORT_ID}#My_Custom_Reality"

# 10. 打印超完整、格式化的节点参数面板
clear
echo -e "${PURPLE}==========================================================${NC}"
echo -e "${GREEN}             🎉 VLESS-Reality 搭建成功！            ${NC}"
echo -e "${PURPLE}==========================================================${NC}"
echo -e "${GREEN}👉 完整的节点明细参数清单 (供手动录入参考)：${NC}"
echo -e "----------------------------------------------------------"
echo -e "${YELLOW}协议类型 (Protocol):${NC}  VLESS"
echo -e "${YELLOW}服务器地址 (Address):${NC} ${SERVER_IP}"
echo -e "${YELLOW}端口号 (Port):${NC}        ${PORT}"
echo -e "${YELLOW}用户UUID (ID):${NC}        ${UUID}"
echo -e "${YELLOW}流控参数 (Flow):${NC}      xtls-rprx-vision"
echo -e "${YELLOW}加密方式 (Encryption):${NC}none"
echo -e "${YELLOW}传输层网络 (Network):${NC} tcp"
echo -e "${YELLOW}安全传输 (Security):${NC}  reality"
echo -e "${YELLOW}目标伪装 (Dest):${NC}      ${CUSTOM_SNI}:443"
echo -e "${YELLOW}伪装域名 (SNI):${NC}       ${CUSTOM_SNI}"
echo -e "${YELLOW}公钥 (PublicKey/pbk):${NC} ${PUBLIC_KEY}"
echo -e "${YELLOW}短ID (ShortID/sid):${NC}   ${SHORT_ID}"
echo -e "${YELLOW}客户端指纹 (Finger):${NC}  chrome"
echo -e "${YELLOW}应用层协议 (ALPN):${NC}    h2, http/1.1"
echo -e "${YELLOW}DNS解析策略:${NC}          强制仅使用 IPv4"
echo -e "${PURPLE}----------------------------------------------------------${NC}"
echo -e "${GREEN}👉 🔗 客户端通用一键导入链接 (直接全选复制)：${NC}"
echo -e "${CYAN}${VLESS_LINK}${NC}"
echo -e "${PURPLE}----------------------------------------------------------${NC}"

# 11. 调用 qrencode 动态渲染终端二维码
echo -e "${GREEN}👉 📱 手机客户端扫码快捷导入：${NC}"
echo ""
qrencode -t UTF8 "${VLESS_LINK}"
echo ""
echo -e "${PURPLE}==========================================================${NC}"
