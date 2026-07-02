#!/bin/bash

# 全局错误处理
set -e
trap 'echo -e "${RED}❌ 脚本执行失败，已停止${NC}" >&2' ERR

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
echo -e "${GREEN}   VLESS-Reality Ultimate Final    ${NC}"
echo -e "${BLUE}==================================================${NC}"

if [ "$EUID" -ne 0 ]; then
    echo -e "${RED}错误：请使用 root 权限运行！${NC}"
    exit 1
fi

XRAY_CONFIG="/usr/local/etc/xray/config.json"

if [ -f "$XRAY_CONFIG" ]; then
    echo -e "${RED}⚠️ 检测到旧配置${NC}"
    read -p "是否删除旧配置继续？(y/n，默认 y): " REMOVE_OLD
    if [[ "$REMOVE_OLD" != "n" && "$REMOVE_OLD" != "N" ]]; then
        BACKUP_FILE="${XRAY_CONFIG}.bak.$(date +%s)"
        cp "$XRAY_CONFIG" "$BACKUP_FILE"
        echo -e "${YELLOW}已备份到: $BACKUP_FILE${NC}"
        rm -f "$XRAY_CONFIG"
    fi
fi

# ====================================================================
# VLESS Vision Reality 搭建函数库
# ====================================================================

# 1. Xray - VLESS Vision Reality 配置文件生成
generate_xray_vless_vision_reality() {
    local configPath=$1
    local realityPort=$2
    local realityServerName=$3
    local realityDomainPort=$4
    local realityPrivateKey=$5
    local realityPublicKey=$6
    local realityMldsa65Seed=$7
    local realityMldsa65Verify=$8
    
    # 确保目录存在
    mkdir -p "$(dirname "$configPath")"
    
    cat <<EOF > "$configPath"
{
    "log": { "loglevel": "warning" },
    "dns": {
        "servers": ["1.1.1.1", "1.0.0.1"],
        "queryStrategy": "UseIPv4"
    },
    "inbounds": [{
        "tag": "dokodemo-in-VLESSReality",
        "port": ${realityPort},
        "protocol": "dokodemo-door",
        "settings": {
            "address": "127.0.0.1",
            "port": 45987,
            "network": "tcp"
        },
        "sniffing": {
            "enabled": true,
            "destOverride": ["tls"],
            "routeOnly": true
        }
    },
    {
        "listen": "127.0.0.1",
        "port": 45987,
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
                "target": "${realityServerName}:${realityDomainPort}",
                "xver": 0,
                "serverNames": ["${realityServerName}"],
                "privateKey": "${realityPrivateKey}",
                "publicKey": "${realityPublicKey}",
                "mldsa65Seed": "${realityMldsa65Seed}",
                "mldsa65Verify": "${realityMldsa65Verify}",
                "maxTimeDiff": 70000,
                "shortIds": ["", "${SHORT_ID}"]
            }
        },
        "sniffing": {
            "enabled": true,
            "destOverride": ["http", "tls", "quic"],
            "routeOnly": true
        }
    }],
    "outbounds": [{ "protocol": "freedom" }]
}
EOF
}

# 2. Reality 密钥生成 (使用 Xray)
generate_reality_key_xray() {
    local privateKey=$1
    
    if [[ -n "${privateKey}" ]]; then
        /usr/local/bin/xray x25519 -i "${privateKey}"
    else
        /usr/local/bin/xray x25519
    fi
}

# 3. VLESS Vision Reality 客户端配置生成
generate_vless_reality_client_config() {
    local uuid=$1
    local server_ip=$2
    local port=$3
    local public_key=$4
    local server_name=$5
    local short_id=$6
    local node_name=$7
    
    local client_url="vless://${uuid}@${server_ip}:${port}?type=tcp&security=reality&flow=xtls-rprx-vision&pbk=${public_key}&sni=${server_name}&sid=${short_id}#${node_name}"
    echo "${client_url}"
}

# 4. 完整的搭建过程
setup_vless_vision_reality() {
    # ==================== 推荐的 Reality 目标域名 ====================
    local RECOMMENDED_DOMAINS=(
        "download-installer.cdn.mozilla.net"
        "addons.mozilla.org"
        "s0.awsstatic.com"
        "d1.awsstatic.com"
        "images-na.ssl-images-amazon.com"
        "m.media-amazon.com"
        "player.live-video.net"
        "one-piece.com"
        "lol.secure.dyn.riotcdn.net"
        "www.lovelive-anime.jp"
        "academy.nvidia.com"
        "dl.google.com"
        "www.google-analytics.com"
        "www.caltech.edu"
        "www.calstatela.edu"
        "www.suny.edu"
        "www.suffolk.edu"
        "www.python.org"
        "vuejs-jp.org"
        "vuejs.org"
        "zh-hk.vuejs.org"
        "react.dev"
        "www.java.com"
        "www.oracle.com"
        "www.mysql.com"
        "www.mongodb.com"
        "redis.io"
        "cname.vercel-dns.com"
        "vercel-dns.com"
        "www.swift.com"
        "www.cisco.com"
        "www.asus.com"
        "www.samsung.com"
        "www.amd.com"
        "www.umcg.nl"
        "www.fom-international.com"
        "www.u-can.co.jp"
        "github.io"
    )
    
    # 用户输入
    echo -e "${YELLOW}👉 节点名称 [默认: My_Reality]: ${NC}"
    read -p "" NODE_NAME
    NODE_NAME=${NODE_NAME:-My_Reality}
    
    echo -e "${YELLOW}👉 端口 [默认: 443]: ${NC}"
    read -p "" CUSTOM_PORT
    PORT=${CUSTOM_PORT:-443}
    
    echo -e "${YELLOW}👉 选择 SNI 域名 (输入序号或自定义): ${NC}"
    echo "快速选择:"
    echo -e "  ${CYAN}1${NC}. download-installer.cdn.mozilla.net"
    echo -e "  ${CYAN}2${NC}. addons.mozilla.org"
    echo -e "  ${CYAN}3${NC}. s0.awsstatic.com"
    echo -e "  ${CYAN}4${NC}. react.dev"
    echo -e "  ${CYAN}5${NC}. github.io"
    echo -e "  ${CYAN}0${NC}. 查看完整列表"
    echo -e "  ${CYAN}-1${NC}. 自定义输入"
    read -p "请选择 [默认: 1]: " DOMAIN_CHOICE
    DOMAIN_CHOICE=${DOMAIN_CHOICE:-1}
    
    if [ "$DOMAIN_CHOICE" -eq -1 ]; then
        echo -e "${YELLOW}请输入自定义 SNI 域名 [默认: www.sony.com]: ${NC}"
        read -p "" CUSTOM_SNI
        CUSTOM_SNI=${CUSTOM_SNI:-www.sony.com}
    elif [ "$DOMAIN_CHOICE" -eq 0 ]; then
        echo "完整域名列表:"
        for i in "${!RECOMMENDED_DOMAINS[@]}"; do
            echo -e "  ${CYAN}$((i+1))${NC}. ${RECOMMENDED_DOMAINS[$i]}"
        done
        read -p "请选择序号 [默认: 1]: " DOMAIN_CHOICE
        DOMAIN_CHOICE=${DOMAIN_CHOICE:-1}
        if [ "$DOMAIN_CHOICE" -ge 1 ] && [ "$DOMAIN_CHOICE" -le ${#RECOMMENDED_DOMAINS[@]} ]; then
            CUSTOM_SNI=${RECOMMENDED_DOMAINS[$((DOMAIN_CHOICE-1))]}
        else
            CUSTOM_SNI="${RECOMMENDED_DOMAINS[0]}"
        fi
    elif [ "$DOMAIN_CHOICE" -ge 1 ] && [ "$DOMAIN_CHOICE" -le 5 ]; then
        case $DOMAIN_CHOICE in
            1) CUSTOM_SNI="${RECOMMENDED_DOMAINS[0]}" ;;
            2) CUSTOM_SNI="${RECOMMENDED_DOMAINS[1]}" ;;
            3) CUSTOM_SNI="${RECOMMENDED_DOMAINS[2]}" ;;
            4) CUSTOM_SNI="${RECOMMENDED_DOMAINS[21]}" ;;
            5) CUSTOM_SNI="${RECOMMENDED_DOMAINS[36]}" ;;
        esac
    else
        CUSTOM_SNI="${RECOMMENDED_DOMAINS[0]}"
    fi
    
    echo -e "${GREEN}✅ 已选择 SNI 域名: ${CUSTOM_SNI}${NC}"
    
    # 1. 生成 Reality 密钥
    echo -e "${YELLOW}生成 Reality 密钥...${NC}"
    REALITY_KEYS=$(generate_reality_key_xray "")
    PRIVATE_KEY=$(echo "$REALITY_KEYS" | grep "PrivateKey" | awk '{print $2}')
    PUBLIC_KEY=$(echo "$REALITY_KEYS" | grep "Password" | awk '{print $3}')
    
    if [ -z "$PRIVATE_KEY" ] || [ -z "$PUBLIC_KEY" ]; then
        echo -e "${RED}❌ 错误：未能生成 Reality 密钥${NC}"
        echo "$REALITY_KEYS"
        return 1
    fi
    echo -e "${GREEN}✅ Reality 密钥生成成功${NC}"
    
    # 2. 生成 mldsa65 密钥 (可选)
    echo -e "${YELLOW}检查目标域名是否支持 ML-DSA-65...${NC}"
    MLDSA65_SEED=""
    MLDSA65_VERIFY=""
    
    if /usr/local/bin/xray tls ping "${CUSTOM_SNI}:443" 2>/dev/null | grep -q "X25519MLKEM768"; then
        LENGTH=$(/usr/local/bin/xray tls ping "${CUSTOM_SNI}:443" 2>/dev/null | grep "Certificate chain's total length:" | awk '{print $5}' | head -1)
        if [ -n "$LENGTH" ] && [ "$LENGTH" -gt 3500 ]; then
            echo -e "${YELLOW}生成 ML-DSA-65 密钥...${NC}"
            MLDSA65=$(/usr/local/bin/xray mldsa65)
            MLDSA65_SEED=$(echo "$MLDSA65" | head -1 | awk '{print $2}')
            MLDSA65_VERIFY=$(echo "$MLDSA65" | tail -1 | awk '{print $2}')
            echo -e "${GREEN}✅ ML-DSA-65 密钥生成成功${NC}"
        else
            echo -e "${YELLOW}⚠️ 证书链长度不足，跳过 ML-DSA-65${NC}"
        fi
    else
        echo -e "${YELLOW}⚠️ 目标域名不支持 X25519MLKEM768，跳过 ML-DSA-65${NC}"
    fi
    
    # 3. 生成 UUID 和 Short ID
    echo -e "${YELLOW}生成 UUID 和 Short ID...${NC}"
    UUID=$(/usr/local/bin/xray uuid)
    SHORT_ID=$(openssl rand -hex 8)
    echo -e "${GREEN}✅ UUID 和 Short ID 生成成功${NC}"
    
    # 4. 生成配置文件
    echo -e "${YELLOW}生成 Xray 配置文件...${NC}"
    generate_xray_vless_vision_reality \
        "$XRAY_CONFIG" \
        "${PORT}" \
        "${CUSTOM_SNI}" \
        "443" \
        "${PRIVATE_KEY}" \
        "${PUBLIC_KEY}" \
        "${MLDSA65_SEED}" \
        "${MLDSA65_VERIFY}"
    echo -e "${GREEN}✅ 配置文件生成成功${NC}"
    
    # 5. 设置配置文件权限
    chmod 600 "$XRAY_CONFIG"
    echo -e "${GREEN}✅ 配置文件权限已设置${NC}"
    
    # 6. 验证配置
    echo -e "${YELLOW}验证配置...${NC}"
    if /usr/local/bin/xray -test -config "$XRAY_CONFIG" > /dev/null 2>&1; then
        echo -e "${GREEN}✅ 配置校验通过${NC}"
    else
        echo -e "${RED}❌ 配置校验失败！${NC}"
        echo -e "${YELLOW}详细错误：${NC}"
        /usr/local/bin/xray -test -config "$XRAY_CONFIG"
        return 1
    fi
    
    # 7. 配置防火墙
    echo -e "${YELLOW}配置 UFW 防火墙...${NC}"
    if command -v ufw >/dev/null; then
        SSH_PORT=$(ss -tlnp | grep sshd | awk '{print $4}' | awk -F':' '{print $2}' | head -n 1)
        SSH_PORT=${SSH_PORT:-22}
        ufw allow "${SSH_PORT}"/tcp > /dev/null 2>&1
        ufw allow "${PORT}"/tcp > /dev/null 2>&1
        ufw --force enable
        echo -e "${GREEN}UFW 已自动放行 SSH 端口(${SSH_PORT}) 及 节点端口(${PORT})${NC}"
    fi
    
    # 8. 配置 BBR
    echo -e "${YELLOW}配置 BBR 加速...${NC}"
    grep -q "net.core.default_qdisc=fq" /etc/sysctl.conf || {
        cat >> /etc/sysctl.conf <<EOF
net.core.default_qdisc=fq
net.ipv4.tcp_congestion_control=bbr
EOF
    }
    sysctl -p > /dev/null 2>&1 || echo -e "${YELLOW}提示: 当前环境无法直接更新内核 sysctl 参数${NC}"
    
    # 9. 启动服务
    echo -e "${YELLOW}启动 Xray 服务...${NC}"
    systemctl daemon-reload
    systemctl enable --now xray
    
    # 10. 验证服务状态
    sleep 2
    if systemctl is-active --quiet xray; then
        echo -e "${GREEN}✅ Xray 服务已成功启动${NC}"
    else
        echo -e "${RED}❌ 错误：Xray 服务启动失败！${NC}"
        systemctl status xray
        return 1
    fi
    
    # 11. 获取服务器 IP
    SERVER_IP=$(curl -4s -m 5 https://api.ipify.org 2>/dev/null || \
                curl -4s -m 5 https://ipv4.icanhazip.com 2>/dev/null || \
                curl -4s -m 5 https://ifconfig.me 2>/dev/null || \
                hostname -I | awk '{print $1}' 2>/dev/null || echo "IP获取失败")
    
    # 12. 生成客户端配置
    echo -e "${YELLOW}生成客户端配置...${NC}"
    ENCODED_NODE_NAME=$(echo -n "${NODE_NAME}" | jq -sRr @uri)
    VLESS_LINK=$(generate_vless_reality_client_config \
        "${UUID}" \
        "${SERVER_IP}" \
        "${PORT}" \
        "${PUBLIC_KEY}" \
        "${CUSTOM_SNI}" \
        "${SHORT_ID}" \
        "${ENCODED_NODE_NAME}")
    
    # 13. 输出结果
    clear
    echo -e "${PURPLE}==========================================================${NC}"
    echo -e "${GREEN}             🎉 安装完成！${NC}"
    echo -e "${PURPLE}==========================================================${NC}"
    
    echo -e "${YELLOW}📋 关键参数说明：${NC}"
    echo -e "  • 节点名称: ${NODE_NAME}"
    echo -e "  • Target Domain: ${CUSTOM_SNI}"
    echo -e "  • 端口: ${PORT}"
    echo -e "  • UUID: ${UUID}"
    echo -e "  • Private Key: ${PRIVATE_KEY}"
    echo -e "  • Public Key: ${PUBLIC_KEY}"
    echo -e "  • Short ID: ${SHORT_ID}"
    if [ -n "$MLDSA65_SEED" ]; then
        echo -e "  • ML-DSA-65 Seed: ${MLDSA65_SEED}"
    else
        echo -e "  • ML-DSA-65: ${YELLOW}不支持或未启用${NC}"
    fi
    
    echo -e "\n${YELLOW}📝 客户端导入链接：${NC}"
    echo -e "${CYAN}${VLESS_LINK}${NC}\n"
    
    # 生成二维码
    if command -v qrencode >/dev/null; then
        qrencode -t UTF8 -m 2 "${VLESS_LINK}"
    else
        echo -e "${YELLOW}⚠️ qrencode 未安装，跳过二维码生成${NC}"
    fi
    
    echo -e "\n${YELLOW}📌 常用命令：${NC}"
    echo -e "  • 查看日志: journalctl -u xray -f"
    echo -e "  • 配置文件: $XRAY_CONFIG"
    echo -e "  • 配置备份: ${XRAY_CONFIG}.bak.*"
    echo -e "  • 重启服务: systemctl restart xray"
    echo -e "  • 测试配置: /usr/local/bin/xray -test -config $XRAY_CONFIG"
    
    echo -e "\n${YELLOW}🔐 VLESS Reality 配置详解：${NC}"
    echo -e "  • target: 伪装的目标域名，需支持 HTTPS"
    echo -e "  • serverNames: 客户端使用的 SNI 值"
    echo -e "  • privateKey: 服务端私钥（必须保密）"
    echo -e "  • publicKey: 客户端使用的公钥"
    echo -e "  • mldsa65: 后量子密码学支持（支持则启用）"
    echo -e "  • shortIds: 短标识符，增加混淆程度"
    echo -e "  • flow: xtls-rprx-vision (VLESS Vision 流控)"
    
    echo -e "\n${PURPLE}==========================================================${NC}"
    echo -e "${GREEN}✨ 相关资源：${NC}"
    echo -e "  • Xray 官方: https://xtls.github.io/"
    echo -e "  • Reality 文档: https://github.com/XTLS/Xray-core"
    echo -e "  • 社区讨论: https://github.com/XTLS"
    echo -e "${PURPLE}==========================================================${NC}"
}

# ==================== 主程序开始 ====================
# 安装依赖 (加入 uuid-runtime，确保依赖完整)
echo -e "${YELLOW}正在安装系统依赖...${NC}"
apt-get update -qq && apt-get install -y curl jq qrencode ufw uuid-runtime

# 安装 Xray 核心（带重试机制）
echo -e "${YELLOW}正在安装 Xray 核心...${NC}"
XRAY_INSTALL_SCRIPT=$(mktemp)
INSTALL_RETRIES=3
INSTALL_COUNT=0

while [ $INSTALL_COUNT -lt $INSTALL_RETRIES ]; do
    if curl -Ls --connect-timeout 10 --max-time 60 https://github.com/XTLS/Xray-install/raw/main/install-release.sh -o "$XRAY_INSTALL_SCRIPT" 2>/dev/null && [ -s "$XRAY_INSTALL_SCRIPT" ]; then
        if bash "$XRAY_INSTALL_SCRIPT" @ install >/dev/null 2>&1; then
            break
        fi
    fi
    INSTALL_COUNT=$((INSTALL_COUNT + 1))
    if [ $INSTALL_COUNT -lt $INSTALL_RETRIES ]; then
        echo -e "${YELLOW}⚠️ 安装失败，5秒后重试 ($INSTALL_COUNT/$INSTALL_RETRIES)...${NC}"
        sleep 5
    fi
done

rm -f "$XRAY_INSTALL_SCRIPT"

# 验证安装
if [ ! -f "/usr/local/bin/xray" ] || [ ! -x "/usr/local/bin/xray" ]; then
    echo -e "${RED}❌ 错误：Xray 核心安装失败！${NC}"
    exit 1
fi

# 验证 Xray 可用性
if ! /usr/local/bin/xray -version >/dev/null 2>&1; then
    echo -e "${RED}❌ 错误：Xray 核心不可用！${NC}"
    exit 1
fi

XRAY_VERSION=$(/usr/local/bin/xray -version 2>/dev/null | head -n 1)
echo -e "${GREEN}✅ Xray 核心安装成功：$XRAY_VERSION${NC}"

# 调用搭建函数
setup_vless_vision_reality || exit 1
