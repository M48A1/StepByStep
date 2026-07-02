#!/bin/bash

# ==================== 版本信息 ====================
VERSION="1.0.6"
BUILD_DATE="2026-07-02"

# 全局错误处理
set -e
trap 'echo -e "${RED}❌ 脚本在第 ${LINENO} 行失败：${BASH_COMMAND}${NC}" >&2' ERR

# ==================== 颜色定义 ====================
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[0;33m'
BLUE='\033[0;34m'
PURPLE='\033[0;35m'
CYAN='\033[0;36m'
NC='\033[0m'

clear
echo -e "${CYAN}═══════════════════════════════════════════════════${NC}"
echo -e "${CYAN}VLESS-Reality Installation Script${NC}"
echo -e "${CYAN}Version: ${VERSION} | Build: ${BUILD_DATE}${NC}"
echo -e "${CYAN}═══════════════════════════════════════════════════${NC}"
echo ""

echo -e "${BLUE}==================================================${NC}"
echo -e "${GREEN}   VLESS-Reality Ultimate Final    ${NC}"
echo -e "${BLUE}==================================================${NC}"

if [ "$EUID" -ne 0 ]; then
    echo -e "${RED}错误：请使用 root 权限运行！${NC}"
    exit 1
fi

if ! command -v apt-get >/dev/null 2>&1; then
    echo -e "${RED}错误：当前脚本仅支持 Debian/Ubuntu 系统（需要 apt-get）。${NC}"
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
    else
        echo -e "${YELLOW}已保留旧配置，脚本退出。${NC}"
        exit 0
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
    local configDir=$(dirname "$configPath")
    if ! mkdir -p "$configDir" 2>/dev/null; then
        echo -e "${RED}❌ 错误：无法创建目录 $configDir${NC}" >&2
        return 1
    fi
    
    # 使用 jq 构建 Reality Settings
    local realitySettings=$(jq -n \
        --arg show "false" \
        --arg target "${realityServerName}:${realityDomainPort}" \
        --arg xver "0" \
        --arg serverName "${realityServerName}" \
        --arg privateKey "${realityPrivateKey}" \
        --arg maxTimeDiff "70000" \
        --arg shortId "${SHORT_ID}" \
        '{
            show: ($show == "true"),
            target: $target,
            xver: ($xver | tonumber),
            serverNames: [$serverName],
            privateKey: $privateKey,
            maxTimeDiff: ($maxTimeDiff | tonumber),
            shortIds: ["", $shortId]
        }')
    
    # 如果有 ML-DSA-65 参数，添加到 realitySettings
    if [ -n "$realityMldsa65Seed" ] && [ -n "$realityMldsa65Verify" ]; then
        realitySettings=$(echo "$realitySettings" | jq \
            --arg seed "$realityMldsa65Seed" \
            --arg verify "$realityMldsa65Verify" \
            '. + {mldsa65Seed: $seed, mldsa65Verify: $verify}')
    fi
    
    # 使用标准的单入站 VLESS + Reality 配置，公网端口直接接收客户端连接。
    local fullConfig=$(jq -n \
        --arg uuid "$UUID" \
        --arg flow "$FLOW" \
        --argjson port "$realityPort" \
        --argjson realitySettings "$realitySettings" \
        '{
            "log": {
                "access": "/var/log/xray/access.log",
                "error": "/var/log/xray/error.log",
                "loglevel": "info"
            },
            "dns": {
                "servers": ["1.1.1.1", "1.0.0.1"],
                "queryStrategy": "UseIPv4"
            },
            "inbounds": [
                {
                    "tag": "vless-reality",
                    "listen": "0.0.0.0",
                    "port": $port,
                    "protocol": "vless",
                    "settings": {
                        "clients": [
                            ({
                                "id": $uuid
                            } + (if $flow == "" then {} else {"flow": $flow} end))
                        ],
                        "decryption": "none",
                        "fallbacks": []
                    },
                    "streamSettings": {
                        "network": "tcp",
                        "security": "reality",
                        "realitySettings": $realitySettings
                    },
                    "sniffing": {
                        "enabled": true,
                        "destOverride": ["http", "tls", "quic"],
                        "routeOnly": true
                    }
                }
            ],
            "outbounds": [
                {
                    "protocol": "freedom"
                }
            ]
        }')
    
    # 写入配置文件
    if ! echo "$fullConfig" | jq '.' > "$configPath" 2>/dev/null; then
        echo -e "${RED}❌ 错误：无法写入配置文件 $configPath${NC}" >&2
        return 1
    fi
    
    # 验证文件是否存在
    if [ ! -f "$configPath" ]; then
        echo -e "${RED}❌ 错误：配置文件创建后丢失 $configPath${NC}" >&2
        return 1
    fi
    
    # 验证JSON合法性
    if ! jq empty "$configPath" 2>/dev/null; then
        echo -e "${RED}❌ 错误：生成的JSON格式无效${NC}" >&2
        return 1
    fi
    
    return 0
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
    
    # 参数验证
    if [ -z "$uuid" ] || [ -z "$server_ip" ] || [ -z "$port" ] || [ -z "$public_key" ] || [ -z "$server_name" ] || [ -z "$short_id" ]; then
        echo -e "${RED}❌ 错误：generate_vless_reality_client_config 参数不完整${NC}" >&2
        echo -e "${YELLOW}  uuid: ${uuid:-'(空)'}${NC}" >&2
        echo -e "${YELLOW}  server_ip: ${server_ip:-'(空)'}${NC}" >&2
        echo -e "${YELLOW}  port: ${port:-'(空)'}${NC}" >&2
        echo -e "${YELLOW}  public_key: ${public_key:-'(空)'}${NC}" >&2
        echo -e "${YELLOW}  server_name: ${server_name:-'(空)'}${NC}" >&2
        echo -e "${YELLOW}  short_id: ${short_id:-'(空)'}${NC}" >&2
        return 1
    fi
    
    # VLESS Reality 标准格式 (不包含 over-tls)
    # reality 安全协议应该直接在 security 字段中指定
    local flow_param=""
    if [ -n "${FLOW:-}" ]; then
        flow_param="&flow=${FLOW}"
    fi
    local client_url="vless://${uuid}@${server_ip}:${port}?type=tcp&security=reality${flow_param}&pbk=${public_key}&fp=chrome&sni=${server_name}&sid=${short_id}&spx=%2F#${node_name}"
    
    # 验证生成的URL
    if [[ ! "$client_url" =~ ^vless:// ]]; then
        echo -e "${RED}❌ 错误：生成的URL格式错误${NC}" >&2
        return 1
    fi
    
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
    if ! [[ "$PORT" =~ ^[0-9]+$ ]] || [ "$PORT" -lt 1 ] || [ "$PORT" -gt 65535 ]; then
        echo -e "${RED}❌ 错误：端口必须是 1-65535 的数字${NC}"
        return 1
    fi

    echo -e "${YELLOW}👉 客户端兼容模式 [默认: 1]: ${NC}"
    echo -e "  ${CYAN}1${NC}. Loon/通用兼容模式（不启用 xtls-rprx-vision，握手兼容性更好）"
    echo -e "  ${CYAN}2${NC}. Xray/Vision 模式（启用 xtls-rprx-vision，客户端必须支持）"
    read -p "请选择 [默认: 1]: " FLOW_CHOICE
    FLOW_CHOICE=${FLOW_CHOICE:-1}
    if [[ "$FLOW_CHOICE" == "2" ]]; then
        FLOW="xtls-rprx-vision"
    else
        FLOW=""
    fi
    echo -e "${GREEN}✅ 已选择模式: $([ -n "$FLOW" ] && echo "Vision (${FLOW})" || echo "Loon/通用兼容模式")${NC}"
    
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
    
    if [[ "$DOMAIN_CHOICE" == "-1" ]]; then
        echo -e "${YELLOW}请输入自定义 SNI 域名 [默认: www.sony.com]: ${NC}"
        read -p "" CUSTOM_SNI
        CUSTOM_SNI=${CUSTOM_SNI:-www.sony.com}
    elif [[ "$DOMAIN_CHOICE" == "0" ]]; then
        echo "完整域名列表:"
        for i in "${!RECOMMENDED_DOMAINS[@]}"; do
            echo -e "  ${CYAN}$((i+1))${NC}. ${RECOMMENDED_DOMAINS[$i]}"
        done
        read -p "请选择序号 [默认: 1]: " DOMAIN_CHOICE
        DOMAIN_CHOICE=${DOMAIN_CHOICE:-1}
        if [[ "$DOMAIN_CHOICE" =~ ^[0-9]+$ ]] && [ "$DOMAIN_CHOICE" -ge 1 ] && [ "$DOMAIN_CHOICE" -le ${#RECOMMENDED_DOMAINS[@]} ]; then
            CUSTOM_SNI=${RECOMMENDED_DOMAINS[$((DOMAIN_CHOICE-1))]}
        else
            CUSTOM_SNI="${RECOMMENDED_DOMAINS[0]}"
        fi
    elif [[ "$DOMAIN_CHOICE" =~ ^[0-9]+$ ]] && [ "$DOMAIN_CHOICE" -ge 1 ] && [ "$DOMAIN_CHOICE" -le 5 ]; then
        case $DOMAIN_CHOICE in
            1) CUSTOM_SNI="${RECOMMENDED_DOMAINS[0]}" ;;
            2) CUSTOM_SNI="${RECOMMENDED_DOMAINS[1]}" ;;
            3) CUSTOM_SNI="${RECOMMENDED_DOMAINS[2]}" ;;
            4) CUSTOM_SNI="${RECOMMENDED_DOMAINS[21]}" ;;
            5) CUSTOM_SNI="${RECOMMENDED_DOMAINS[37]}" ;;
        esac
    elif [[ "$DOMAIN_CHOICE" =~ ^[A-Za-z0-9._-]+$ ]]; then
        CUSTOM_SNI="$DOMAIN_CHOICE"
    else
        CUSTOM_SNI="${RECOMMENDED_DOMAINS[0]}"
    fi
    
    echo -e "${GREEN}✅ 已选择 SNI 域名: ${CUSTOM_SNI}${NC}"
    
    # 1. 生成 Reality 密钥
    echo -e "${YELLOW}生成 Reality 密钥...${NC}"
    REALITY_KEYS=$(generate_reality_key_xray "")
    
    # 调试输出
    echo -e "${YELLOW}调试：Xray x25519 完整输出：${NC}"
    echo "$REALITY_KEYS"
    echo ""
    
    # 使用更健壮的方法提取密钥
    PRIVATE_KEY=$(echo "$REALITY_KEYS" | grep -Ei "PrivateKey|Private key|Private" | tail -1 | awk '{print $NF}' | sed 's/[^A-Za-z0-9_-]//g')
    PUBLIC_KEY=$(echo "$REALITY_KEYS" | grep -Ei "PublicKey|Public key|Password" | tail -1 | awk '{print $NF}' | sed 's/[^A-Za-z0-9_-]//g')
    
    # 验证密钥格式
    if ! [[ "$PRIVATE_KEY" =~ ^[A-Za-z0-9_-]{43,44}$ ]]; then
        echo -e "${RED}❌ 错误：Private Key 格式无效或未正确提取${NC}"
        echo -e "${YELLOW}  提取到的值: '${PRIVATE_KEY}'${NC}"
        echo -e "${YELLOW}  完整输出：${NC}"
        echo "$REALITY_KEYS"
        return 1
    fi
    
    if ! [[ "$PUBLIC_KEY" =~ ^[A-Za-z0-9_-]{43,44}$ ]]; then
        echo -e "${RED}❌ 错误：Public Key 格式无效或未正确提取${NC}"
        echo -e "${YELLOW}  提取到的值: '${PUBLIC_KEY}'${NC}"
        echo -e "${YELLOW}  完整输出：${NC}"
        echo "$REALITY_KEYS"
        return 1
    fi
    
    echo -e "${GREEN}✅ Reality 密钥生成成功${NC}"
    echo -e "${CYAN}  Private Key: 已生成并写入服务端配置（不在屏幕显示）${NC}"
    echo -e "${CYAN}  Public Key:  ${PUBLIC_KEY}${NC}"
    echo ""
    
    # 2. 生成 mldsa65 密钥 (可选)
    echo -e "${YELLOW}检查目标域名是否支持 ML-DSA-65...${NC}"
    MLDSA65_SEED=""
    MLDSA65_VERIFY=""
    
    # 暂时跳过 ML-DSA-65，专注于基础 Reality 握手
    echo -e "${YELLOW}⚠️ 暂时跳过 ML-DSA-65 检测，采用标准 Reality 配置${NC}"
    
    # 3. 生成 UUID 和 Short ID
    echo -e "${YELLOW}生成 UUID 和 Short ID...${NC}"
    UUID=$(/usr/local/bin/xray uuid)
    SHORT_ID=$(openssl rand -hex 8)
    
    # 验证 UUID 格式
    if ! [[ "$UUID" =~ ^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$ ]]; then
        echo -e "${RED}❌ 错误：UUID 格式无效${NC}"
        echo -e "${YELLOW}  生成的值: ${UUID}${NC}"
        return 1
    fi
    
    # 验证 Short ID 格式（应该是16字符的16进制）
    if ! [[ "$SHORT_ID" =~ ^[0-9a-f]{16}$ ]]; then
        echo -e "${RED}❌ 错误：Short ID 格式无效${NC}"
        echo -e "${YELLOW}  生成的值: ${SHORT_ID}${NC}"
        return 1
    fi
    
    echo -e "${GREEN}✅ UUID 和 Short ID 生成成功${NC}"
    echo -e "${CYAN}  UUID:     ${UUID}${NC}"
    echo -e "${CYAN}  Short ID: ${SHORT_ID}${NC}"
    echo ""
    
    # 4. 生成配置文件
    echo -e "${YELLOW}════════════════════════════════════════════════${NC}"
    echo -e "${YELLOW}参数验证总结${NC}"
    echo -e "${YELLOW}════════════════════════════════════════════════${NC}"
    echo -e "${CYAN}  节点名称:  ${NODE_NAME}${NC}"
    echo -e "${CYAN}  监听端口:  ${PORT}${NC}"
    echo -e "${CYAN}  Flow:      ${FLOW:-'(空，兼容模式)'}${NC}"
    echo -e "${CYAN}  SNI域名:   ${CUSTOM_SNI}${NC}"
    echo -e "${CYAN}  UUID:      ${UUID}${NC}"
    echo -e "${CYAN}  Public Key:${PUBLIC_KEY}${NC}"
    echo -e "${CYAN}  Short ID:  ${SHORT_ID}${NC}"
    echo -e "${YELLOW}════════════════════════════════════════════════${NC}"
    echo ""
    
    echo -e "${YELLOW}生成 Xray 配置文件...${NC}"
    if ! generate_xray_vless_vision_reality \
        "$XRAY_CONFIG" \
        "${PORT}" \
        "${CUSTOM_SNI}" \
        "443" \
        "${PRIVATE_KEY}" \
        "${PUBLIC_KEY}" \
        "${MLDSA65_SEED}" \
        "${MLDSA65_VERIFY}"; then
        echo -e "${RED}❌ 错误：配置文件生成失败${NC}"
        return 1
    fi
    
    # 验证文件是否真的存在
    if [ ! -f "$XRAY_CONFIG" ]; then
        echo -e "${RED}❌ 错误：配置文件验证失败，文件不存在：$XRAY_CONFIG${NC}"
        return 1
    fi
    echo -e "${GREEN}✅ 配置文件生成成功${NC}"
    
    # 5. 设置配置文件权限
    mkdir -p /var/log/xray
    touch /var/log/xray/access.log /var/log/xray/error.log
    if id nobody >/dev/null 2>&1; then
        chown -R nobody:nogroup /var/log/xray 2>/dev/null || chown -R nobody:nobody /var/log/xray 2>/dev/null || true
    fi
    chmod 755 /var/log/xray
    if ! chmod 600 "$XRAY_CONFIG"; then
        echo -e "${RED}❌ 错误：无法设置配置文件权限${NC}"
        echo -e "${YELLOW}调试信息：${NC}"
        ls -la "$XRAY_CONFIG" || echo "文件不存在"
        return 1
    fi
    echo -e "${GREEN}✅ 配置文件权限已设置${NC}"
    
    # 6. 验证配置
    echo -e "${YELLOW}验证配置...${NC}"
    echo -e "${YELLOW}调试：配置文件路径: $XRAY_CONFIG${NC}"
    echo -e "${YELLOW}调试：文件存在: $([ -f "$XRAY_CONFIG" ] && echo "是" || echo "否")${NC}"
    
    if [ ! -f "$XRAY_CONFIG" ]; then
        echo -e "${RED}❌ 错误：配置文件不存在！${NC}"
        echo -e "${YELLOW}目录内容：${NC}"
        ls -la "$(dirname "$XRAY_CONFIG")" || echo "目录不存在"
        return 1
    fi
    
    if /usr/local/bin/xray -test -config "$XRAY_CONFIG" > /dev/null 2>&1; then
        echo -e "${GREEN}✅ 配置校验通过${NC}"
    else
        echo -e "${RED}❌ 配置校验失败！${NC}"
        echo -e "${YELLOW}详细错误：${NC}"
        /usr/local/bin/xray -test -config "$XRAY_CONFIG" || true
        return 1
    fi
    
    # 7. 配置防火墙
    echo -e "${YELLOW}配置 UFW 防火墙...${NC}"
    if command -v ufw >/dev/null; then
        SSH_PORT=$(ss -tlnp 2>/dev/null | awk '/sshd/ {print $4}' | awk -F':' '{print $NF}' | head -n 1)
        SSH_PORT=${SSH_PORT:-22}
        ufw allow "${SSH_PORT}"/tcp > /dev/null 2>&1 || echo -e "${YELLOW}提示: SSH 端口(${SSH_PORT})放行失败，请手动检查 UFW${NC}"
        ufw allow "${PORT}"/tcp > /dev/null 2>&1 || echo -e "${YELLOW}提示: 节点端口(${PORT})放行失败，请手动检查 UFW${NC}"
        if ufw status 2>/dev/null | grep -q "Status: active"; then
            echo -e "${GREEN}UFW 已放行 SSH 端口(${SSH_PORT}) 及 节点端口(${PORT})${NC}"
        else
            read -p "UFW 当前未启用，是否启用？可能影响现有连接。(y/n，默认 n): " ENABLE_UFW
            if [[ "$ENABLE_UFW" == "y" || "$ENABLE_UFW" == "Y" ]]; then
                ufw --force enable
                echo -e "${GREEN}UFW 已启用，并放行 SSH 端口(${SSH_PORT}) 及 节点端口(${PORT})${NC}"
            else
                echo -e "${YELLOW}已跳过启用 UFW，仅尝试写入放行规则。${NC}"
            fi
        fi
    fi
    
    # 8. 配置 BBR
    echo -e "${YELLOW}配置 BBR 加速...${NC}"
    grep -q "^net.core.default_qdisc=fq" /etc/sysctl.conf || cat >> /etc/sysctl.conf <<EOF
net.core.default_qdisc=fq
EOF
    grep -q "^net.ipv4.tcp_congestion_control=bbr" /etc/sysctl.conf || cat >> /etc/sysctl.conf <<EOF
net.ipv4.tcp_congestion_control=bbr
EOF
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
    echo -e "${YELLOW}获取服务器公网IP...${NC}"
    SERVER_IP=$(curl -4s -m 5 https://api.ipify.org 2>/dev/null || \
                curl -4s -m 5 https://ipv4.icanhazip.com 2>/dev/null || \
                curl -4s -m 5 https://ifconfig.me 2>/dev/null || \
                hostname -I | awk '{print $1}' 2>/dev/null || echo "")
    
    if [ -z "$SERVER_IP" ] || [ "$SERVER_IP" = "IP获取失败" ]; then
        echo -e "${YELLOW}⚠️ 无法获取公网IP，请手动输入${NC}"
        read -p "请输入服务器公网 IP: " SERVER_IP
        if [ -z "$SERVER_IP" ]; then
            echo -e "${RED}❌ 错误：必须提供服务器 IP${NC}"
            return 1
        fi
    fi
    echo -e "${GREEN}✅ 服务器IP: ${SERVER_IP}${NC}"
    
    # 12. 生成客户端配置
    echo -e "${YELLOW}生成客户端配置...${NC}"
    
    # 参数验证
    if [ -z "$UUID" ] || [ -z "$PUBLIC_KEY" ] || [ -z "$SHORT_ID" ] || [ -z "$CUSTOM_SNI" ]; then
        echo -e "${RED}❌ 错误：客户端配置参数不完整${NC}"
        echo -e "${YELLOW}  UUID: ${UUID:-'(空)'}${NC}"
        echo -e "${YELLOW}  Public Key: ${PUBLIC_KEY:-'(空)'}${NC}"
        echo -e "${YELLOW}  Short ID: ${SHORT_ID:-'(空)'}${NC}"
        echo -e "${YELLOW}  SNI: ${CUSTOM_SNI:-'(空)'}${NC}"
        return 1
    fi
    
    # 调试：显示所有参数
    echo -e "${YELLOW}调试 - 生成链接前的参数：${NC}"
    echo -e "${CYAN}  UUID        : ${UUID}${NC}"
    echo -e "${CYAN}  SERVER_IP   : ${SERVER_IP}${NC}"
    echo -e "${CYAN}  PORT        : ${PORT}${NC}"
    echo -e "${CYAN}  PUBLIC_KEY  : ${PUBLIC_KEY}${NC}"
    echo -e "${CYAN}  CUSTOM_SNI  : ${CUSTOM_SNI}${NC}"
    echo -e "${CYAN}  SHORT_ID    : ${SHORT_ID}${NC}"
    echo -e "${CYAN}  NODE_NAME   : ${NODE_NAME}${NC}"
    echo ""
    
    ENCODED_NODE_NAME=$(echo -n "${NODE_NAME}" | jq -sRr @uri)
    VLESS_LINK=$(generate_vless_reality_client_config \
        "${UUID}" \
        "${SERVER_IP}" \
        "${PORT}" \
        "${PUBLIC_KEY}" \
        "${CUSTOM_SNI}" \
        "${SHORT_ID}" \
        "${ENCODED_NODE_NAME}")
    
    # 调试：显示生成的链接
    echo -e "${YELLOW}调试 - 生成的链接：${NC}"
    echo -e "${CYAN}${VLESS_LINK}${NC}"
    echo ""
    
    # 验证生成的链接
    if [[ ! "$VLESS_LINK" =~ ^vless:// ]]; then
        echo -e "${RED}❌ 错误：客户端配置链接生成失败${NC}"
        echo -e "${YELLOW}生成的链接: ${VLESS_LINK}${NC}"
        return 1
    fi
    
    # 检查是否包含占位符
    if [[ "$VLESS_LINK" =~ \(PublicKey\) ]]; then
        echo -e "${RED}❌ 错误：生成的链接仍包含占位符 (PublicKey)${NC}"
        echo -e "${YELLOW}这表示 PUBLIC_KEY 变量为空或格式错误${NC}"
        echo -e "${YELLOW}实际接收到的公钥值: '${PUBLIC_KEY}'${NC}"
        return 1
    fi
    
    echo -e "${GREEN}✅ 客户端配置链接生成成功${NC}"
    
    # 13. 输出结果
    clear
    echo -e "${PURPLE}════════════════════════════════════════════════════${NC}"
    echo -e "${GREEN}          🎉 安装完成！VLESS-Reality 节点就绪${NC}"
    echo -e "${PURPLE}════════════════════════════════════════════════════${NC}"
    echo ""
    
    echo -e "${YELLOW}📋 节点信息总览${NC}"
    echo -e "${CYAN}┌────────────────────────────────────────────┐${NC}"
    echo -e "${CYAN}│ 节点名称        : ${NODE_NAME}${NC}"
    echo -e "${CYAN}│ 服务器地址      : ${SERVER_IP}:${PORT}${NC}"
    echo -e "${CYAN}│ 伪装域名 (SNI)  : ${CUSTOM_SNI}${NC}"
    echo -e "${CYAN}│ UUID            : ${UUID}${NC}"
    echo -e "${CYAN}│ 公钥 (pbk)      : ${PUBLIC_KEY}${NC}"
    echo -e "${CYAN}│ Short ID        : ${SHORT_ID}${NC}"
    echo -e "${CYAN}└────────────────────────────────────────────┘${NC}"
    echo ""
    
    echo -e "${YELLOW}📝 客户端导入链接${NC}"
    echo -e "${PURPLE}═══════════════════════════════════════════════════${NC}"
    echo -e "${CYAN}${VLESS_LINK}${NC}"
    echo -e "${PURPLE}═══════════════════════════════════════════════════${NC}"
    echo ""

    echo -e "${YELLOW}📱 Loon 手动填写参数${NC}"
    echo -e "  类型/协议          : ${CYAN}VLESS${NC}"
    echo -e "  服务器             : ${CYAN}${SERVER_IP}${NC}"
    echo -e "  端口               : ${CYAN}${PORT}${NC}"
    echo -e "  UUID               : ${CYAN}${UUID}${NC}"
    echo -e "  Transport/Network  : ${CYAN}TCP${NC}"
    echo -e "  TLS/Reality        : ${CYAN}Reality${NC}"
    echo -e "  SNI/Server Name    : ${CYAN}${CUSTOM_SNI}${NC}"
    echo -e "  Public Key / pbk   : ${CYAN}${PUBLIC_KEY}${NC}"
    echo -e "  Short ID / sid     : ${CYAN}${SHORT_ID}${NC}"
    echo -e "  Fingerprint / fp   : ${CYAN}chrome${NC}"
    echo -e "  SpiderX / spx      : ${CYAN}/${NC}"
    echo -e "  Flow               : ${CYAN}${FLOW:-留空}${NC}"
    echo -e "  UDP                : ${CYAN}建议先关闭；TCP 握手成功后再测试 UDP${NC}"
    echo ""
    
    # 链接有效性验证
    echo -e "${YELLOW}📊 链接验证${NC}"
    if [[ "$VLESS_LINK" =~ pbk=([a-zA-Z0-9_-]+) ]]; then
        echo -e "${GREEN}✅ Public Key 正确: ${BASH_REMATCH[1]}${NC}"
    else
        echo -e "${RED}⚠️ Public Key 可能异常，请检查${NC}"
    fi
    
    if [[ "$VLESS_LINK" =~ sid=([a-f0-9]{16}) ]]; then
        echo -e "${GREEN}✅ Short ID 正确: ${BASH_REMATCH[1]}${NC}"
    else
        echo -e "${RED}⚠️ Short ID 格式异常${NC}"
    fi
    
    if [[ "$VLESS_LINK" =~ sni=([a-zA-Z0-9._-]+) ]]; then
        echo -e "${GREEN}✅ SNI 正确: ${BASH_REMATCH[1]}${NC}"
    else
        echo -e "${RED}⚠️ SNI 格式异常${NC}"
    fi
    echo ""
    
    # 生成二维码
    if command -v qrencode >/dev/null; then
        echo -e "${YELLOW}📱 二维码${NC}"
        qrencode -t UTF8 -m 2 "${VLESS_LINK}"
        echo ""
    else
        echo -e "${YELLOW}⚠️ qrencode 未安装，跳过二维码生成${NC}"
        echo ""
    fi
    
    echo -e "${YELLOW}📌 常用命令${NC}"
    echo -e "  • 查看实时日志      : ${CYAN}journalctl -u xray -f${NC}"
    echo -e "  • 查看配置文件      : ${CYAN}cat $XRAY_CONFIG | jq .${NC}"
    echo -e "  • 重启服务          : ${CYAN}systemctl restart xray${NC}"
    echo -e "  • 查看服务状态      : ${CYAN}systemctl status xray${NC}"
    echo -e "  • 验证配置语法      : ${CYAN}/usr/local/bin/xray -test -config $XRAY_CONFIG${NC}"
    echo -e "  • 查看配置备份      : ${CYAN}ls -la ${XRAY_CONFIG}.bak.*${NC}"
    echo ""
    
    echo -e "${YELLOW}🔐 VLESS-Reality 配置要点${NC}"
    echo -e "  • security: reality         (安全协议，不需要 over-tls)"
    echo -e "  • flow: ${FLOW:-留空}   (Loon/通用兼容模式建议留空)"
    echo -e "  • pbk: ${PUBLIC_KEY}     (客户端必须匹配)"
    echo -e "  • sni: ${CUSTOM_SNI}     (伪装域名，需真实HTTPS)"
    echo -e "  • sid: ${SHORT_ID}              (短ID标识符)"
    echo ""
    
    echo -e "${YELLOW}🚀 客户端配置方式${NC}"
    echo -e "  1. 复制上面的链接，粘贴到支持VLESS的客户端"
    echo -e "  2. 或使用二维码扫描导入"
    echo -e "  3. 支持客户端: Clash, v2rayN, SingBox 等"
    echo ""
    
    echo -e "${PURPLE}════════════════════════════════════════════════════${NC}"
    echo -e "${GREEN}✨ 更多资源${NC}"
    echo -e "  • Xray 官方文档     : https://xtls.github.io/"
    echo -e "  • Reality 说明      : https://github.com/XTLS/Xray-core"
    echo -e "  • GitHub 讨论区     : https://github.com/XTLS/Xray-core/discussions"
    echo -e "${PURPLE}════════════════════════════════════════════════════${NC}"
}

# ==================== 主程序开始 ====================
# 安装依赖 (加入 uuid-runtime，确保依赖完整)
echo -e "${YELLOW}正在安装系统依赖...${NC}"
apt-get update -qq && apt-get install -y curl jq qrencode ufw uuid-runtime openssl iproute2 ca-certificates

# 安装 Xray 核心（带重试机制）
echo -e "${YELLOW}正在安装 Xray 核心...${NC}"
XRAY_INSTALL_SCRIPT=$(mktemp)
XRAY_INSTALL_LOG=$(mktemp)
INSTALL_RETRIES=3
INSTALL_COUNT=0

while [ $INSTALL_COUNT -lt $INSTALL_RETRIES ]; do
    if curl -Ls --connect-timeout 10 --max-time 60 https://github.com/XTLS/Xray-install/raw/main/install-release.sh -o "$XRAY_INSTALL_SCRIPT" 2>/dev/null && [ -s "$XRAY_INSTALL_SCRIPT" ]; then
        if bash "$XRAY_INSTALL_SCRIPT" install >"$XRAY_INSTALL_LOG" 2>&1; then
            break
        fi
    else
        echo "下载 Xray 安装脚本失败或文件为空" > "$XRAY_INSTALL_LOG"
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
    echo -e "${YELLOW}最近一次安装日志：${NC}"
    tail -n 30 "$XRAY_INSTALL_LOG" 2>/dev/null || true
    rm -f "$XRAY_INSTALL_LOG"
    exit 1
fi
rm -f "$XRAY_INSTALL_LOG"

# 验证 Xray 可用性
if ! /usr/local/bin/xray -version >/dev/null 2>&1; then
    echo -e "${RED}❌ 错误：Xray 核心不可用！${NC}"
    exit 1
fi

XRAY_VERSION=$(/usr/local/bin/xray -version 2>/dev/null | head -n 1)
echo -e "${GREEN}✅ Xray 核心安装成功：$XRAY_VERSION${NC}"

# 调用搭建函数
setup_vless_vision_reality || exit 1
