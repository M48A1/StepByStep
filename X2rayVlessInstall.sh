#!/usr/bin/env bash

VERSION="2.0.1"
BUILD_DATE="2026-07-02"

set -Eeuo pipefail

GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[0;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
NC='\033[0m'

XRAY_BIN="/usr/local/bin/xray"
XRAY_CONFIG="/usr/local/etc/xray/config.json"
XRAY_META="/usr/local/etc/xray/reality-client.txt"
XRAY_LOG_DIR="/var/log/xray"
FLOW="xtls-rprx-vision"
INSTALL_URL="https://github.com/XTLS/Xray-install/raw/main/install-release.sh"

trap 'echo -e "${RED}错误：第 ${LINENO} 行执行失败：${BASH_COMMAND}${NC}" >&2' ERR

info() { echo -e "${CYAN}$*${NC}"; }
ok() { echo -e "${GREEN}$*${NC}"; }
warn() { echo -e "${YELLOW}$*${NC}"; }
die() { echo -e "${RED}$*${NC}" >&2; exit 1; }

print_banner() {
    clear || true
    echo -e "${BLUE}==================================================${NC}"
    echo -e "${GREEN} VLESS Vision REALITY Installer${NC}"
    echo -e "${CYAN} Version: ${VERSION} | Build: ${BUILD_DATE}${NC}"
    echo -e "${BLUE}==================================================${NC}"
}

require_root() {
    [ "${EUID}" -eq 0 ] || die "请使用 root 权限运行。"
}

check_system() {
    command -v apt-get >/dev/null 2>&1 || die "当前脚本仅支持 Debian/Ubuntu 系统。"
    command -v systemctl >/dev/null 2>&1 || die "当前系统缺少 systemctl，无法管理 Xray 服务。"
}

install_dependencies() {
    info "正在安装系统依赖..."
    apt-get update -qq
    apt-get install -y curl jq openssl iproute2 ca-certificates qrencode ufw
}

remove_old_config() {
    if [ ! -f "$XRAY_CONFIG" ]; then
        return
    fi

    warn "检测到旧 Xray 配置：$XRAY_CONFIG"
    read -r -p "是否删除旧配置并继续？(y/n): " answer
    case "$answer" in
        y|Y)
            rm -f "$XRAY_CONFIG"
            ok "旧配置已删除。"
            ;;
        *)
            warn "已取消安装，旧配置未修改。"
            exit 0
            ;;
    esac
}

install_xray() {
    if [ -x "$XRAY_BIN" ]; then
        ok "检测到已安装 Xray：$($XRAY_BIN version 2>/dev/null | head -n 1 || true)"
        return
    fi

    info "正在安装 Xray 核心..."
    local script_path log_path retries count
    script_path=$(mktemp)
    log_path=$(mktemp)
    retries=3
    count=1

    while [ "$count" -le "$retries" ]; do
        if curl -fsSL --connect-timeout 20 --max-time 180 "$INSTALL_URL" -o "$script_path"; then
            if bash "$script_path" install >"$log_path" 2>&1; then
                rm -f "$script_path" "$log_path"
                ok "Xray 核心安装成功。"
                return
            fi
        else
            echo "下载 Xray 安装脚本失败：$INSTALL_URL" >"$log_path"
        fi

        warn "Xray 安装失败，准备重试 (${count}/${retries})..."
        count=$((count + 1))
        sleep 5
    done

    warn "最近一次安装日志："
    tail -n 40 "$log_path" 2>/dev/null || true
    rm -f "$script_path" "$log_path"
    die "Xray 核心安装失败。"
}

prompt_settings() {
    read -r -p "节点名称 [默认: My_Reality]: " NODE_NAME
    NODE_NAME=${NODE_NAME:-My_Reality}

    read -r -p "监听端口 [默认: 443]: " PORT
    PORT=${PORT:-443}
    [[ "$PORT" =~ ^[0-9]+$ ]] || die "端口必须是数字。"
    [ "$PORT" -ge 1 ] && [ "$PORT" -le 65535 ] || die "端口必须在 1-65535 之间。"

    echo "选择 REALITY 目标 SNI："
    echo "  1. download-installer.cdn.mozilla.net"
    echo "  2. addons.mozilla.org"
    echo "  3. www.microsoft.com"
    echo "  4. www.apple.com"
    echo "  5. www.cloudflare.com"
    echo "  6. 自定义"
    while true; do
        read -r -p "请选择 1-6: " sni_choice
        case "$sni_choice" in
            1) SNI="download-installer.cdn.mozilla.net"; break ;;
            2) SNI="addons.mozilla.org"; break ;;
            3) SNI="www.microsoft.com"; break ;;
            4) SNI="www.apple.com"; break ;;
            5) SNI="www.cloudflare.com"; break ;;
            6)
                read -r -p "请输入自定义 SNI: " SNI
                [[ "$SNI" =~ ^[A-Za-z0-9._-]+$ ]] || die "SNI 格式不正确。"
                break
                ;;
            *) warn "请输入 1-6。" ;;
        esac
    done

    ok "已启用 Vision flow：$FLOW"
    ok "已选择 SNI：$SNI"
}

generate_keys() {
    info "正在生成 UUID、REALITY 密钥和 Short ID..."
    UUID=$("$XRAY_BIN" uuid)
    SHORT_ID=$(openssl rand -hex 8)

    local keys
    keys=$("$XRAY_BIN" x25519)
    PRIVATE_KEY=$(echo "$keys" | awk -F': ' '/Private key|PrivateKey|Private/ {print $2}' | tail -n 1 | tr -d '[:space:]')
    PUBLIC_KEY=$(echo "$keys" | awk -F': ' '/Public key|PublicKey|Password/ {print $2}' | tail -n 1 | tr -d '[:space:]')

    [[ "$UUID" =~ ^[0-9a-fA-F-]{36}$ ]] || die "UUID 生成失败：$UUID"
    [[ "$SHORT_ID" =~ ^[0-9a-f]{16}$ ]] || die "Short ID 生成失败：$SHORT_ID"
    [[ "$PRIVATE_KEY" =~ ^[A-Za-z0-9_-]{43,44}$ ]] || die "Private Key 解析失败。Xray 输出：$keys"
    [[ "$PUBLIC_KEY" =~ ^[A-Za-z0-9_-]{43,44}$ ]] || die "Public Key 解析失败。Xray 输出：$keys"

    ok "密钥生成完成。"
}

get_xray_user() {
    local user
    user=$(systemctl cat xray 2>/dev/null | awk -F= '/^[[:space:]]*User=/ {print $2; exit}' | tr -d '[:space:]')
    echo "${user:-root}"
}

write_xray_config() {
    info "正在写入 Xray 配置..."
    mkdir -p "$(dirname "$XRAY_CONFIG")" "$XRAY_LOG_DIR"
    touch "$XRAY_LOG_DIR/access.log" "$XRAY_LOG_DIR/error.log"

    jq -n \
        --arg uuid "$UUID" \
        --arg flow "$FLOW" \
        --arg email "${NODE_NAME}-vless-reality-vision" \
        --arg sni "$SNI" \
        --arg privateKey "$PRIVATE_KEY" \
        --arg shortId "$SHORT_ID" \
        --argjson port "$PORT" \
        '{
            log: {
                access: "/var/log/xray/access.log",
                error: "/var/log/xray/error.log",
                loglevel: "info"
            },
            inbounds: [
                {
                    tag: "vless-reality-vision",
                    listen: "0.0.0.0",
                    port: $port,
                    protocol: "vless",
                    settings: {
                        clients: [
                            {
                                id: $uuid,
                                flow: $flow,
                                email: $email
                            }
                        ],
                        decryption: "none"
                    },
                    streamSettings: {
                        network: "tcp",
                        security: "reality",
                        realitySettings: {
                            show: false,
                            target: ($sni + ":443"),
                            xver: 0,
                            serverNames: [$sni],
                            privateKey: $privateKey,
                            shortIds: [$shortId],
                            maxTimeDiff: 60000
                        }
                    },
                    sniffing: {
                        enabled: true,
                        destOverride: ["http", "tls", "quic"],
                        routeOnly: true
                    }
                }
            ],
            outbounds: [
                {
                    tag: "direct",
                    protocol: "freedom"
                },
                {
                    tag: "block",
                    protocol: "blackhole"
                }
            ]
        }' >"$XRAY_CONFIG"

    local xray_user
    xray_user=$(get_xray_user)
    chmod 640 "$XRAY_CONFIG"
    chmod 755 "$XRAY_LOG_DIR"
    chmod 644 "$XRAY_LOG_DIR/access.log" "$XRAY_LOG_DIR/error.log"
    if [ "$xray_user" != "root" ] && id "$xray_user" >/dev/null 2>&1; then
        chown "$xray_user":"$xray_user" "$XRAY_CONFIG" 2>/dev/null || chown "$xray_user":nogroup "$XRAY_CONFIG" 2>/dev/null || true
        chown -R "$xray_user":"$xray_user" "$XRAY_LOG_DIR" 2>/dev/null || chown -R "$xray_user":nogroup "$XRAY_LOG_DIR" 2>/dev/null || true
    fi

    "$XRAY_BIN" -test -config "$XRAY_CONFIG" >/dev/null
    ok "Xray 配置校验通过。"
}

configure_firewall() {
    if ! command -v ufw >/dev/null 2>&1; then
        return
    fi

    info "正在配置 UFW 放行规则..."
    local ssh_port
    ssh_port=$(ss -tlnp 2>/dev/null | awk '/sshd/ {print $4}' | awk -F':' '{print $NF}' | head -n 1)
    ssh_port=${ssh_port:-22}

    ufw allow "${ssh_port}/tcp" >/dev/null 2>&1 || true
    ufw allow "${PORT}/tcp" >/dev/null 2>&1 || true

    if ufw status 2>/dev/null | grep -q "Status: active"; then
        ok "UFW 已放行 SSH(${ssh_port}/tcp) 和节点端口(${PORT}/tcp)。"
        return
    fi

    read -r -p "UFW 当前未启用，是否启用？可能影响现有 SSH 连接。(y/n): " enable_ufw
    if [[ "$enable_ufw" == "y" || "$enable_ufw" == "Y" ]]; then
        ufw --force enable
        ok "UFW 已启用。"
    else
        warn "已跳过启用 UFW。请确认 VPS 商家安全组已放行 ${PORT}/tcp。"
    fi
}

enable_bbr() {
    info "正在配置 BBR..."
    grep -q '^net.core.default_qdisc=fq' /etc/sysctl.conf || echo 'net.core.default_qdisc=fq' >>/etc/sysctl.conf
    grep -q '^net.ipv4.tcp_congestion_control=bbr' /etc/sysctl.conf || echo 'net.ipv4.tcp_congestion_control=bbr' >>/etc/sysctl.conf
    sysctl -p >/dev/null 2>&1 || warn "当前环境无法立即应用 sysctl，重启后通常会生效。"
}

restart_xray() {
    info "正在启动 Xray 服务..."
    systemctl daemon-reload
    systemctl enable --now xray
    sleep 2
    systemctl is-active --quiet xray || {
        systemctl status xray --no-pager || true
        die "Xray 服务启动失败。"
    }

    if ss -tln 2>/dev/null | awk '{print $4}' | grep -Eq "(:|\\])${PORT}$"; then
        ok "Xray 已启动并监听 ${PORT}/tcp。"
    else
        ss -tlnp 2>/dev/null || true
        die "Xray 已启动，但没有监听 ${PORT}/tcp。"
    fi
}

get_public_ip() {
    SERVER_IP=$(curl -4fsS --max-time 8 https://api.ipify.org 2>/dev/null || \
        curl -4fsS --max-time 8 https://ipv4.icanhazip.com 2>/dev/null || \
        hostname -I 2>/dev/null | awk '{print $1}' || true)

    if [ -z "${SERVER_IP:-}" ]; then
        read -r -p "无法自动获取公网 IP，请手动输入: " SERVER_IP
    fi
    [ -n "$SERVER_IP" ] || die "服务器公网 IP 不能为空。"
}

write_client_info() {
    local encoded_name vless_link
    encoded_name=$(printf '%s' "$NODE_NAME" | jq -sRr @uri)
    vless_link="vless://${UUID}@${SERVER_IP}:${PORT}?type=tcp&security=reality&flow=${FLOW}&pbk=${PUBLIC_KEY}&fp=chrome&sni=${SNI}&sid=${SHORT_ID}&spx=%2F#${encoded_name}"

    cat >"$XRAY_META" <<EOF
节点名称: ${NODE_NAME}
服务器: ${SERVER_IP}
端口: ${PORT}
协议: VLESS
传输: TCP
安全: REALITY
Flow: ${FLOW}
UUID: ${UUID}
SNI: ${SNI}
Public Key: ${PUBLIC_KEY}
Short ID: ${SHORT_ID}
Fingerprint: chrome
SpiderX: /

VLESS 链接:
${vless_link}
EOF
    chmod 600 "$XRAY_META"

    clear || true
    ok "安装完成。"
    echo
    echo -e "${BLUE}================ 节点信息 ================${NC}"
    echo -e "节点名称          : ${CYAN}${NODE_NAME}${NC}"
    echo -e "服务器            : ${CYAN}${SERVER_IP}${NC}"
    echo -e "端口              : ${CYAN}${PORT}${NC}"
    echo -e "协议              : ${CYAN}VLESS${NC}"
    echo -e "传输              : ${CYAN}TCP${NC}"
    echo -e "安全              : ${CYAN}REALITY${NC}"
    echo -e "Flow              : ${CYAN}${FLOW}${NC}"
    echo -e "UUID              : ${CYAN}${UUID}${NC}"
    echo -e "SNI               : ${CYAN}${SNI}${NC}"
    echo -e "Public Key        : ${CYAN}${PUBLIC_KEY}${NC}"
    echo -e "Short ID          : ${CYAN}${SHORT_ID}${NC}"
    echo -e "Fingerprint       : ${CYAN}chrome${NC}"
    echo -e "SpiderX           : ${CYAN}/${NC}"
    echo
    echo -e "${BLUE}================ VLESS 链接 ================${NC}"
    echo -e "${CYAN}${vless_link}${NC}"
    echo

    if command -v qrencode >/dev/null 2>&1; then
        echo -e "${BLUE}================ 二维码 ================${NC}"
        qrencode -t UTF8 -m 2 "$vless_link"
        echo
    fi

    echo -e "${BLUE}================ 排查命令 ================${NC}"
    echo "systemctl status xray --no-pager"
    echo "journalctl -u xray -n 80 --no-pager"
    echo "tail -f /var/log/xray/access.log /var/log/xray/error.log"
    echo "ss -tlnp | grep ':${PORT}'"
    echo "ufw status verbose"
    echo
    warn "如果 Loon 测试时日志没有新增，请优先检查 VPS 商家安全组是否放行 ${PORT}/tcp。"
    warn "客户端先测试 TCP 访问，UDP 测试请等 TCP 可用后再开。"
    echo
    echo "节点信息已保存到：$XRAY_META"
}

main() {
    print_banner
    require_root
    check_system
    remove_old_config
    install_dependencies
    install_xray
    prompt_settings
    generate_keys
    write_xray_config
    configure_firewall
    enable_bbr
    restart_xray
    get_public_ip
    write_client_info
}

main "$@"
