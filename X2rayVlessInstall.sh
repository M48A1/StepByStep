#!/usr/bin/env bash

VERSION="2.2.1"
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
XRAY_INFO="/usr/local/etc/xray/vless-reality-info.txt"
XRAY_QR="/usr/local/etc/xray/vless-reality-qr.png"
XRAY_LOG_DIR="/var/log/xray"
XRAY_INSTALL_URL="https://github.com/XTLS/Xray-install/raw/main/install-release.sh"

FLOW="xtls-rprx-vision"
FINGERPRINT="chrome"
SPIDER_X="/"

trap 'echo -e "${RED}错误：第 ${LINENO} 行失败：${BASH_COMMAND}${NC}" >&2' ERR

info() { echo -e "${CYAN}$*${NC}"; }
ok() { echo -e "${GREEN}$*${NC}"; }
warn() { echo -e "${YELLOW}$*${NC}"; }
die() { echo -e "${RED}$*${NC}" >&2; exit 1; }

banner() {
    clear || true
    echo -e "${BLUE}==================================================${NC}"
    echo -e "${GREEN} VLESS REALITY 一键安装脚本${NC}"
    echo -e "${CYAN} Version: ${VERSION} | Build: ${BUILD_DATE}${NC}"
    echo -e "${BLUE}==================================================${NC}"
}

require_root() {
    [ "${EUID}" -eq 0 ] || die "请使用 root 权限运行。"
}

check_os() {
    command -v apt-get >/dev/null 2>&1 || die "仅支持 Debian/Ubuntu。"
    command -v systemctl >/dev/null 2>&1 || die "当前系统缺少 systemctl。"
}

install_dependencies() {
    info "安装依赖..."
    apt-get update -qq
    apt-get install -y curl jq openssl iproute2 ca-certificates qrencode ufw
}

remove_old_config() {
    if [ ! -f "$XRAY_CONFIG" ]; then
        return
    fi

    warn "检测到旧配置：$XRAY_CONFIG"
    read -r -p "是否删除旧配置并继续？(y/n): " answer
    case "$answer" in
        y|Y)
            rm -f "$XRAY_CONFIG"
            ok "旧配置已删除。"
            ;;
        *)
            warn "已取消，旧配置未修改。"
            exit 0
            ;;
    esac
}

install_xray() {
    info "安装/更新 Xray..."
    local script_path log_path
    script_path=$(mktemp)
    log_path=$(mktemp)

    if ! curl -fsSL --connect-timeout 20 --max-time 180 "$XRAY_INSTALL_URL" -o "$script_path"; then
        rm -f "$script_path" "$log_path"
        die "下载 Xray 安装脚本失败。"
    fi

    if ! bash "$script_path" install >"$log_path" 2>&1; then
        warn "Xray 安装失败，最近日志："
        tail -n 60 "$log_path" || true
        rm -f "$script_path" "$log_path"
        die "Xray 安装失败。"
    fi

    rm -f "$script_path" "$log_path"
    [ -x "$XRAY_BIN" ] || die "未找到 Xray 可执行文件：$XRAY_BIN"
    ok "$($XRAY_BIN version | head -n 1)"
}

ask_settings() {
    read -r -p "节点名称 [默认: My_VLESS]: " NODE_NAME
    NODE_NAME=${NODE_NAME:-My_VLESS}

    read -r -p "监听端口 [默认: 443]: " PORT
    PORT=${PORT:-443}
    [[ "$PORT" =~ ^[0-9]+$ ]] || die "端口必须是数字。"
    [ "$PORT" -ge 1 ] && [ "$PORT" -le 65535 ] || die "端口必须在 1-65535 之间。"

    echo "选择 REALITY 伪装目标 SNI："
    echo "  1. download-installer.cdn.mozilla.net"
    echo "  2. addons.mozilla.org"
    echo "  3. www.microsoft.com"
    echo "  4. www.apple.com"
    echo "  5. www.cloudflare.com"
    echo "  6. 自定义"
    while true; do
        read -r -p "请选择 1-6: " choice
        case "$choice" in
            1) SNI="download-installer.cdn.mozilla.net"; break ;;
            2) SNI="addons.mozilla.org"; break ;;
            3) SNI="www.microsoft.com"; break ;;
            4) SNI="www.apple.com"; break ;;
            5) SNI="www.cloudflare.com"; break ;;
            6)
                read -r -p "请输入 SNI 域名: " SNI
                [[ "$SNI" =~ ^[A-Za-z0-9._-]+$ ]] || die "SNI 格式不正确。"
                break
                ;;
            *) warn "请输入 1-6。" ;;
        esac
    done

    ok "节点名称: $NODE_NAME"
    ok "监听端口: $PORT"
    ok "SNI: $SNI"
}

generate_values() {
    info "生成 UUID、REALITY 密钥、Short ID..."
    UUID=$("$XRAY_BIN" uuid)
    SHORT_ID=$(openssl rand -hex 8)

    local key_output
    key_output=$("$XRAY_BIN" x25519)
    PRIVATE_KEY=$(echo "$key_output" | awk -F': ' '/Private key|PrivateKey|Private/ {print $2}' | tail -n 1 | tr -d '[:space:]')
    PUBLIC_KEY=$(echo "$key_output" | awk -F': ' '/Public key|PublicKey|Password/ {print $2}' | tail -n 1 | tr -d '[:space:]')

    [[ "$UUID" =~ ^[0-9a-fA-F-]{36}$ ]] || die "UUID 生成失败：$UUID"
    [[ "$SHORT_ID" =~ ^[0-9a-f]{16}$ ]] || die "Short ID 生成失败：$SHORT_ID"
    [[ "$PRIVATE_KEY" =~ ^[A-Za-z0-9_-]{43,44}$ ]] || die "Private Key 解析失败。Xray 输出：$key_output"
    [[ "$PUBLIC_KEY" =~ ^[A-Za-z0-9_-]{43,44}$ ]] || die "Public Key 解析失败。Xray 输出：$key_output"
}

xray_user() {
    local user
    user=$(systemctl show -p User --value xray 2>/dev/null | tr -d '[:space:]')
    echo "${user:-root}"
}

write_config() {
    info "写入 Xray 配置..."
    mkdir -p "$(dirname "$XRAY_CONFIG")" "$XRAY_LOG_DIR"
    touch "$XRAY_LOG_DIR/access.log" "$XRAY_LOG_DIR/error.log"

    jq -n \
        --arg uuid "$UUID" \
        --arg flow "$FLOW" \
        --arg email "${NODE_NAME}@vless-reality" \
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
                    tag: "vless-reality",
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
                            shortIds: [$shortId]
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

    local user
    user=$(xray_user)
    chmod 644 "$XRAY_CONFIG"
    chmod 755 "$XRAY_LOG_DIR"
    chmod 644 "$XRAY_LOG_DIR/access.log" "$XRAY_LOG_DIR/error.log"
    if [ "$user" != "root" ] && id "$user" >/dev/null 2>&1; then
        chown -R "$user":"$user" "$XRAY_LOG_DIR" 2>/dev/null || chown -R "$user":nogroup "$XRAY_LOG_DIR" 2>/dev/null || true
    fi

    "$XRAY_BIN" -test -config "$XRAY_CONFIG" >/dev/null
    ok "配置校验通过。"
}

open_firewall() {
    info "放行防火墙端口..."
    if command -v ufw >/dev/null 2>&1; then
        local ssh_port
        ssh_port=$(ss -tlnp 2>/dev/null | awk '/sshd/ {print $4}' | awk -F':' '{print $NF}' | head -n 1)
        ssh_port=${ssh_port:-22}
        ufw allow "${ssh_port}/tcp" >/dev/null 2>&1 || true
        ufw allow "${PORT}/tcp" >/dev/null 2>&1 || true
        if ufw status 2>/dev/null | grep -q "Status: active"; then
            ok "UFW 已放行 ${PORT}/tcp。"
        else
            warn "UFW 未启用。请确认 VPS 服务商安全组已放行 ${PORT}/tcp。"
        fi
    fi
}

enable_bbr() {
    info "配置 BBR..."
    grep -q '^net.core.default_qdisc=fq' /etc/sysctl.conf || echo 'net.core.default_qdisc=fq' >>/etc/sysctl.conf
    grep -q '^net.ipv4.tcp_congestion_control=bbr' /etc/sysctl.conf || echo 'net.ipv4.tcp_congestion_control=bbr' >>/etc/sysctl.conf
    sysctl -p >/dev/null 2>&1 || warn "sysctl 暂时未生效，重启后通常会生效。"
}

restart_xray() {
    info "启动 Xray..."
    systemctl daemon-reload
    systemctl enable --now xray
    systemctl restart xray
    sleep 2

    systemctl is-active --quiet xray || {
        systemctl status xray --no-pager || true
        die "Xray 启动失败。"
    }

    ss -tln 2>/dev/null | awk '{print $4}' | grep -Eq "(:|\\])${PORT}$" || {
        ss -tlnp || true
        die "Xray 未监听 ${PORT}/tcp。"
    }
    ok "Xray 正在监听 ${PORT}/tcp。"
}

get_public_ip() {
    SERVER_IP=$(curl -4fsS --max-time 8 https://api.ipify.org 2>/dev/null || \
        curl -4fsS --max-time 8 https://ipv4.icanhazip.com 2>/dev/null || \
        hostname -I 2>/dev/null | awk '{print $1}' || true)
    if [ -z "${SERVER_IP:-}" ]; then
        read -r -p "无法自动获取公网 IP，请输入服务器 IP: " SERVER_IP
    fi
    [ -n "$SERVER_IP" ] || die "服务器 IP 不能为空。"
}

print_result() {
    local encoded_name link
    encoded_name=$(printf '%s' "$NODE_NAME" | jq -sRr @uri)
    link="vless://${UUID}@${SERVER_IP}:${PORT}?type=tcp&security=reality&flow=${FLOW}&pbk=${PUBLIC_KEY}&fp=${FINGERPRINT}&sni=${SNI}&sid=${SHORT_ID}&spx=%2F#${encoded_name}"

    cat >"$XRAY_INFO" <<EOF
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
Fingerprint: ${FINGERPRINT}
SpiderX: ${SPIDER_X}

VLESS 链接:
${link}
EOF
    chmod 600 "$XRAY_INFO"

    clear || true
    ok "VLESS REALITY 节点搭建完成。"
    echo
    echo -e "${BLUE}============== 客户端参数 ==============${NC}"
    echo -e "服务器        : ${CYAN}${SERVER_IP}${NC}"
    echo -e "端口          : ${CYAN}${PORT}${NC}"
    echo -e "协议          : ${CYAN}VLESS${NC}"
    echo -e "传输          : ${CYAN}TCP${NC}"
    echo -e "安全          : ${CYAN}REALITY${NC}"
    echo -e "Flow          : ${CYAN}${FLOW}${NC}"
    echo -e "UUID          : ${CYAN}${UUID}${NC}"
    echo -e "SNI           : ${CYAN}${SNI}${NC}"
    echo -e "Public Key    : ${CYAN}${PUBLIC_KEY}${NC}"
    echo -e "Short ID      : ${CYAN}${SHORT_ID}${NC}"
    echo -e "Fingerprint   : ${CYAN}${FINGERPRINT}${NC}"
    echo -e "SpiderX       : ${CYAN}${SPIDER_X}${NC}"
    echo
    echo -e "${BLUE}============== 导入链接 ==============${NC}"
    echo -e "${CYAN}${link}${NC}"
    echo
    echo -e "${BLUE}============== 二维码 ==============${NC}"
    if command -v qrencode >/dev/null 2>&1; then
        qrencode -t UTF8 -m 2 "$link"
        qrencode -o "$XRAY_QR" "$link"
        chmod 600 "$XRAY_QR"
        echo
        echo "二维码图片已保存到：$XRAY_QR"
        echo
    else
        warn "未检测到 qrencode，无法输出二维码。"
    fi
    echo -e "${BLUE}============== 排查命令 ==============${NC}"
    echo "systemctl status xray --no-pager"
    echo "journalctl -u xray -n 80 --no-pager"
    echo "tail -f /var/log/xray/access.log /var/log/xray/error.log"
    echo "ss -tlnp | grep ':${PORT}'"
    echo
    warn "节点信息已保存到：$XRAY_INFO"
}

main() {
    banner
    require_root
    check_os
    remove_old_config
    install_dependencies
    install_xray
    ask_settings
    generate_values
    write_config
    open_firewall
    enable_bbr
    restart_xray
    get_public_ip
    print_result
}

main "$@"
