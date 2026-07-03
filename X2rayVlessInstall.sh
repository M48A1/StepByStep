#!/usr/bin/env bash

VERSION="2.4.4"
BUILD_DATE="2026-07-03"

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
MANAGER_BIN="/usr/local/bin/vless"
XRAY_INSTALL_URL="https://github.com/XTLS/Xray-install/raw/main/install-release.sh"

FLOW="xtls-rprx-vision"
FINGERPRINT="chrome"
SPIDER_X="/"
DEFAULT_DNS_1="1.1.1.1"
DEFAULT_DNS_2="1.0.0.1"

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
            rm -f "$XRAY_INFO" "$XRAY_QR"
            ok "旧配置、节点信息和二维码已删除。"
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
    local xray_version
    xray_version=$("$XRAY_BIN" version 2>/dev/null || "$XRAY_BIN" -version 2>/dev/null || true)
    xray_version=${xray_version%%$'\n'*}
    ok "${xray_version:-Xray 已安装。}"
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
    echo "  3. shopee.sg"
    echo "  4. www.dell.com"
    echo "  5. www.cloudflare.com"
    echo "  6. 自定义"
    while true; do
        read -r -p "请选择 1-6: " choice
        case "$choice" in
            1) SNI="download-installer.cdn.mozilla.net"; break ;;
            2) SNI="addons.mozilla.org"; break ;;
            3) SNI="shopee.sg"; break ;;
            4) SNI="www.dell.com"; break ;;
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
        --arg dns1 "$DEFAULT_DNS_1" \
        --arg dns2 "$DEFAULT_DNS_2" \
        --argjson port "$PORT" \
        '{
            log: {
                access: "/var/log/xray/access.log",
                error: "/var/log/xray/error.log",
                loglevel: "info"
            },
            dns: {
                servers: [$dns1, $dns2],
                queryStrategy: "UseIPv4"
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
            routing: {
                domainStrategy: "IPIfNonMatch",
                rules: [
                    {
                        type: "field",
                        ip: ["::/0"],
                        outboundTag: "block"
                    }
                ]
            },
            outbounds: [
                {
                    tag: "direct",
                    protocol: "freedom",
                    settings: {
                        domainStrategy: "UseIPv4"
                    }
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
    install_manager_command
}

install_manager_command() {
    local source_path
    source_path=${BASH_SOURCE[0]:-$0}
    source_path=$(readlink -f "$source_path" 2>/dev/null || printf '%s' "$source_path")

    if [ -f "$source_path" ] && grep -q "VLESS REALITY 一键安装脚本" "$source_path" 2>/dev/null; then
        install -m 755 "$source_path" "$MANAGER_BIN"
        rm -f /usr/local/bin/vless-reality 2>/dev/null || true
        ok "管理命令已安装：vless"
        echo "以后可直接运行：vless"
    else
        write_manager_stub
    fi
}

write_manager_stub() {
    cat >"$MANAGER_BIN" <<'VLESS_MANAGER_EOF'
#!/usr/bin/env bash
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
FINGERPRINT="chrome"
SPIDER_X="/"
FLOW="xtls-rprx-vision"
DEFAULT_DNS_1="1.1.1.1"
DEFAULT_DNS_2="1.0.0.1"

ok() { echo -e "${GREEN}$*${NC}"; }
warn() { echo -e "${YELLOW}$*${NC}"; }
die() { echo -e "${RED}$*${NC}" >&2; exit 1; }

require_root() {
    [ "${EUID}" -eq 0 ] || die "请使用 root 权限运行。"
}

validate_sni() {
    local value=$1
    [[ "$value" =~ ^[A-Za-z0-9._-]+$ ]]
}

validate_port() {
    local value=$1
    [[ "$value" =~ ^[0-9]+$ ]] && [ "$value" -ge 1 ] && [ "$value" -le 65535 ]
}

validate_dns_value() {
    local value=$1
    [[ "$value" =~ ^[A-Za-z0-9:._-]+$ ]]
}

public_key_from_private() {
    local private_key=$1
    local key_output
    key_output=$("$XRAY_BIN" x25519 -i "$private_key")
    printf '%s\n' "$key_output" | awk -F': ' '
        /Public key|PublicKey|Password/ { value = $2 }
        END { gsub(/[[:space:]]/, "", value); print value }
    '
}

saved_server_ip() {
    if [ -f "$XRAY_INFO" ]; then
        awk -F': ' '/^服务器:/ { print $2; exit }' "$XRAY_INFO"
    fi
}

regenerate_client_info_from_config() {
    [ -f "$XRAY_CONFIG" ] || die "未找到配置文件：$XRAY_CONFIG"

    local uuid port flow email node_name sni short_id private_key public_key server_ip encoded_name link
    uuid=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .settings.clients[0].id][0] // empty' "$XRAY_CONFIG")
    port=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .port][0] // empty' "$XRAY_CONFIG")
    flow=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .settings.clients[0].flow][0] // empty' "$XRAY_CONFIG")
    email=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .settings.clients[0].email][0] // empty' "$XRAY_CONFIG")
    sni=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.serverNames[0]][0] // empty' "$XRAY_CONFIG")
    short_id=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.shortIds[0]][0] // empty' "$XRAY_CONFIG")
    private_key=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.privateKey][0] // empty' "$XRAY_CONFIG")

    [ -n "$uuid" ] || die "无法从配置读取 UUID。"
    [ -n "$port" ] || die "无法从配置读取端口。"
    [ -n "$flow" ] || flow="$FLOW"
    [ -n "$sni" ] || die "无法从配置读取 SNI。"
    [ -n "$short_id" ] || die "无法从配置读取 Short ID。"
    [ -n "$private_key" ] || die "无法从配置读取 Private Key。"

    public_key=$(public_key_from_private "$private_key")
    [ -n "$public_key" ] || die "无法从 Private Key 推导 Public Key。"

    node_name=${email%@vless-reality}
    [ -n "$node_name" ] || node_name="My_VLESS"

    server_ip=$(saved_server_ip)
    if [ -z "${server_ip:-}" ]; then
        server_ip=$(curl -4fsS --max-time 8 https://api.ipify.org 2>/dev/null || \
            curl -4fsS --max-time 8 https://ipv4.icanhazip.com 2>/dev/null || \
            hostname -I 2>/dev/null | awk '{print $1}' || true)
    fi
    [ -n "${server_ip:-}" ] || read -r -p "请输入服务器公网 IP: " server_ip
    [ -n "$server_ip" ] || die "服务器 IP 不能为空。"

    encoded_name=$(printf '%s' "$node_name" | jq -sRr @uri)
    link="vless://${uuid}@${server_ip}:${port}?type=tcp&security=reality&flow=${flow}&pbk=${public_key}&fp=${FINGERPRINT}&sni=${sni}&sid=${short_id}&spx=%2F#${encoded_name}"

    cat >"$XRAY_INFO" <<EOF
节点名称: ${node_name}
服务器: ${server_ip}
端口: ${port}
协议: VLESS
传输: TCP
安全: REALITY
Flow: ${flow}
UUID: ${uuid}
SNI: ${sni}
Public Key: ${public_key}
Short ID: ${short_id}
Fingerprint: ${FINGERPRINT}
SpiderX: ${SPIDER_X}

VLESS 链接:
${link}
EOF
    chmod 600 "$XRAY_INFO"

    if command -v qrencode >/dev/null 2>&1; then
        qrencode -o "$XRAY_QR" "$link"
        chmod 600 "$XRAY_QR"
    fi
}

show_info() {
    if [ ! -f "$XRAY_INFO" ]; then
        warn "未找到节点信息文件，尝试从当前配置重新生成。"
        regenerate_client_info_from_config
    fi

    cat "$XRAY_INFO"
    if [ -f "$XRAY_CONFIG" ]; then
        echo
        echo "当前 Xray DNS:"
        jq -r '.dns.servers // [] | .[]' "$XRAY_CONFIG" 2>/dev/null || true
    fi
    echo
    [ -f "$XRAY_QR" ] && echo "二维码图片：$XRAY_QR"
}

show_qr() {
    if [ ! -f "$XRAY_INFO" ]; then
        regenerate_client_info_from_config
    fi

    local link
    link=$(awk '/^VLESS 链接:$/ { getline; print; exit }' "$XRAY_INFO")
    [ -n "$link" ] || die "无法读取 VLESS 链接。"

    command -v qrencode >/dev/null 2>&1 || die "未安装 qrencode。"
    qrencode -t UTF8 -m 2 "$link"
    qrencode -o "$XRAY_QR" "$link"
    chmod 600 "$XRAY_QR"
    echo
    echo "二维码图片已保存到：$XRAY_QR"
}

change_sni() {
    require_root
    [ -f "$XRAY_CONFIG" ] || die "未找到配置文件：$XRAY_CONFIG"

    local new_sni tmp_config
    new_sni=${1:-}
    if [ -z "$new_sni" ]; then
        read -r -p "请输入新的 SNI 域名: " new_sni
    fi
    validate_sni "$new_sni" || die "SNI 格式不正确。"

    tmp_config=$(mktemp)
    jq --arg sni "$new_sni" '
        (.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.target) = ($sni + ":443")
        | (.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.serverNames) = [$sni]
    ' "$XRAY_CONFIG" >"$tmp_config"

    "$XRAY_BIN" -test -config "$tmp_config" >/dev/null
    mv "$tmp_config" "$XRAY_CONFIG"
    systemctl restart xray
    regenerate_client_info_from_config
    ok "SNI 已更新为：$new_sni"
    ok "Xray 已重启，客户端信息和二维码已重新生成。"
}

change_port() {
    require_root
    [ -f "$XRAY_CONFIG" ] || die "未找到配置文件：$XRAY_CONFIG"

    local new_port old_port tmp_config
    new_port=${1:-}
    if [ -z "$new_port" ]; then
        read -r -p "请输入新的监听端口: " new_port
    fi
    validate_port "$new_port" || die "端口必须是 1-65535 之间的数字。"

    old_port=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .port][0] // empty' "$XRAY_CONFIG")
    if [ "$new_port" = "$old_port" ]; then
        warn "新端口与当前端口相同，无需修改。"
        return
    fi

    tmp_config=$(mktemp)
    jq --argjson port "$new_port" '
        (.inbounds[] | select(.protocol == "vless") | .port) = $port
    ' "$XRAY_CONFIG" >"$tmp_config"

    "$XRAY_BIN" -test -config "$tmp_config" >/dev/null
    mv "$tmp_config" "$XRAY_CONFIG"
    if command -v ufw >/dev/null 2>&1; then
        ufw allow "${new_port}/tcp" >/dev/null 2>&1 || true
    fi
    systemctl restart xray
    regenerate_client_info_from_config
    ok "端口已更新为：$new_port"
    ok "Xray 已重启，客户端信息和二维码已重新生成。"
    warn "请确认 VPS 服务商安全组已放行 ${new_port}/tcp。"
}

force_ipv4_outbound() {
    require_root
    [ -f "$XRAY_CONFIG" ] || die "未找到配置文件：$XRAY_CONFIG"

    local tmp_config
    tmp_config=$(mktemp)
    jq '
        (.outbounds[] | select(.tag == "direct" and .protocol == "freedom") | .settings.domainStrategy) = "UseIPv4"
        | .routing = (.routing // {})
        | .routing.domainStrategy = "IPIfNonMatch"
        | .routing.rules = (
            [.routing.rules[]? | select(.outboundTag != "block" or (((.ip // []) | index("::/0")) | not))]
            + [{type: "field", ip: ["::/0"], outboundTag: "block"}]
        )
    ' "$XRAY_CONFIG" >"$tmp_config"

    "$XRAY_BIN" -test -config "$tmp_config" >/dev/null
    mv "$tmp_config" "$XRAY_CONFIG"
    systemctl restart xray
    ok "已强制 Xray 使用 IPv4 出站，并阻止 IPv6 目标。"
}

change_dns() {
    require_root
    [ -f "$XRAY_CONFIG" ] || die "未找到配置文件：$XRAY_CONFIG"

    local dns1 dns2 tmp_config
    dns1=${1:-}
    dns2=${2:-}
    if [ -z "$dns1" ]; then
        read -r -p "请输入主 DNS [默认: ${DEFAULT_DNS_1}]: " dns1
        dns1=${dns1:-$DEFAULT_DNS_1}
    fi
    if [ -z "$dns2" ]; then
        read -r -p "请输入备用 DNS [默认: ${DEFAULT_DNS_2}]: " dns2
        dns2=${dns2:-$DEFAULT_DNS_2}
    fi

    validate_dns_value "$dns1" || die "主 DNS 格式不正确。"
    validate_dns_value "$dns2" || die "备用 DNS 格式不正确。"

    tmp_config=$(mktemp)
    jq --arg dns1 "$dns1" --arg dns2 "$dns2" '
        .dns = {
            servers: [$dns1, $dns2],
            queryStrategy: "UseIPv4"
        }
    ' "$XRAY_CONFIG" >"$tmp_config"

    "$XRAY_BIN" -test -config "$tmp_config" >/dev/null
    mv "$tmp_config" "$XRAY_CONFIG"
    systemctl restart xray
    ok "DNS 已更新为：$dns1, $dns2"
    ok "Xray 已重启。"
}

restart_service() {
    require_root
    systemctl restart xray
    ok "Xray 已重启。"
}

status_service() {
    systemctl status xray --no-pager
}

manager_menu() {
    while true; do
        echo
        echo -e "${BLUE}============== vless 管理菜单 ==============${NC}"
        echo "1. 查看节点信息"
        echo "2. 输出二维码"
        echo "3. 修改 SNI"
        echo "4. 修改端口"
        echo "5. 强制 IPv4 出站"
        echo "6. 修改 DNS"
        echo "7. 重启 Xray"
        echo "8. 查看 Xray 状态"
        echo "0. 退出"
        read -r -p "请选择: " choice
        case "$choice" in
            1) show_info ;;
            2) show_qr ;;
            3) change_sni ;;
            4) change_port ;;
            5) force_ipv4_outbound ;;
            6) change_dns ;;
            7) restart_service ;;
            8) status_service ;;
            0) exit 0 ;;
            *) warn "请输入 0-8。" ;;
        esac
    done
}

usage() {
    cat <<EOF
用法:
  vless                打开管理菜单
  vless show           查看节点信息
  vless qr             输出二维码
  vless sni <domain>   修改 SNI
  vless port <端口>     修改 VLESS 监听端口
  vless ipv4           强制 IPv4 出站并阻止 IPv6 目标
  vless dns <主DNS> <备用DNS>
                         修改 Xray DNS，例如：vless dns 1.1.1.1 8.8.8.8
  vless restart        重启 Xray
  vless status         查看状态
EOF
}

dispatch() {
    local command_name
    command_name=${1:-}
    case "$command_name" in
        ""|menu) manager_menu ;;
        show) show_info ;;
        qr) show_qr ;;
        sni|change-sni) shift; change_sni "${1:-}" ;;
        port|change-port) shift; change_port "${1:-}" ;;
        ipv4|force-ipv4) force_ipv4_outbound ;;
        dns|change-dns) shift; change_dns "${1:-}" "${2:-}" ;;
        restart) restart_service ;;
        status) status_service ;;
        help|-h|--help) usage ;;
        *) usage; exit 1 ;;
    esac
}

dispatch "$@"
VLESS_MANAGER_EOF

    chmod 755 "$MANAGER_BIN"
    rm -f /usr/local/bin/vless-reality 2>/dev/null || true
    ok "管理命令已安装：vless"
    echo "以后可直接运行：vless"
}

validate_sni() {
    local value=$1
    [[ "$value" =~ ^[A-Za-z0-9._-]+$ ]]
}

validate_port() {
    local value=$1
    [[ "$value" =~ ^[0-9]+$ ]] && [ "$value" -ge 1 ] && [ "$value" -le 65535 ]
}

public_key_from_private() {
    local private_key=$1
    local key_output
    key_output=$("$XRAY_BIN" x25519 -i "$private_key")
    printf '%s\n' "$key_output" | awk -F': ' '
        /Public key|PublicKey|Password/ { value = $2 }
        END { gsub(/[[:space:]]/, "", value); print value }
    '
}

saved_server_ip() {
    if [ -f "$XRAY_INFO" ]; then
        awk -F': ' '/^服务器:/ { print $2; exit }' "$XRAY_INFO"
    fi
}

regenerate_client_info_from_config() {
    [ -f "$XRAY_CONFIG" ] || die "未找到配置文件：$XRAY_CONFIG"

    local uuid port flow email node_name sni short_id private_key public_key server_ip encoded_name link
    uuid=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .settings.clients[0].id][0] // empty' "$XRAY_CONFIG")
    port=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .port][0] // empty' "$XRAY_CONFIG")
    flow=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .settings.clients[0].flow][0] // empty' "$XRAY_CONFIG")
    email=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .settings.clients[0].email][0] // empty' "$XRAY_CONFIG")
    sni=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.serverNames[0]][0] // empty' "$XRAY_CONFIG")
    short_id=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.shortIds[0]][0] // empty' "$XRAY_CONFIG")
    private_key=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.privateKey][0] // empty' "$XRAY_CONFIG")

    [ -n "$uuid" ] || die "无法从配置读取 UUID。"
    [ -n "$port" ] || die "无法从配置读取端口。"
    [ -n "$flow" ] || flow="$FLOW"
    [ -n "$sni" ] || die "无法从配置读取 SNI。"
    [ -n "$short_id" ] || die "无法从配置读取 Short ID。"
    [ -n "$private_key" ] || die "无法从配置读取 Private Key。"

    public_key=$(public_key_from_private "$private_key")
    [ -n "$public_key" ] || die "无法从 Private Key 推导 Public Key。"

    node_name=${email%@vless-reality}
    [ -n "$node_name" ] || node_name="My_VLESS"

    server_ip=$(saved_server_ip)
    if [ -z "${server_ip:-}" ]; then
        server_ip=$(curl -4fsS --max-time 8 https://api.ipify.org 2>/dev/null || \
            curl -4fsS --max-time 8 https://ipv4.icanhazip.com 2>/dev/null || \
            hostname -I 2>/dev/null | awk '{print $1}' || true)
    fi
    [ -n "${server_ip:-}" ] || read -r -p "请输入服务器公网 IP: " server_ip
    [ -n "$server_ip" ] || die "服务器 IP 不能为空。"

    encoded_name=$(printf '%s' "$node_name" | jq -sRr @uri)
    link="vless://${uuid}@${server_ip}:${port}?type=tcp&security=reality&flow=${flow}&pbk=${public_key}&fp=${FINGERPRINT}&sni=${sni}&sid=${short_id}&spx=%2F#${encoded_name}"

    cat >"$XRAY_INFO" <<EOF
节点名称: ${node_name}
服务器: ${server_ip}
端口: ${port}
协议: VLESS
传输: TCP
安全: REALITY
Flow: ${flow}
UUID: ${uuid}
SNI: ${sni}
Public Key: ${public_key}
Short ID: ${short_id}
Fingerprint: ${FINGERPRINT}
SpiderX: ${SPIDER_X}

VLESS 链接:
${link}
EOF
    chmod 600 "$XRAY_INFO"

    if command -v qrencode >/dev/null 2>&1; then
        qrencode -o "$XRAY_QR" "$link"
        chmod 600 "$XRAY_QR"
    fi
}

show_info() {
    if [ ! -f "$XRAY_INFO" ]; then
        warn "未找到节点信息文件，尝试从当前配置重新生成。"
        regenerate_client_info_from_config
    fi

    cat "$XRAY_INFO"
    if [ -f "$XRAY_CONFIG" ]; then
        echo
        echo "当前 Xray DNS:"
        jq -r '.dns.servers // [] | .[]' "$XRAY_CONFIG" 2>/dev/null || true
    fi
    echo
    if [ -f "$XRAY_QR" ]; then
        echo "二维码图片：$XRAY_QR"
    fi
}

show_qr() {
    if [ ! -f "$XRAY_INFO" ]; then
        regenerate_client_info_from_config
    fi

    local link
    link=$(awk '/^VLESS 链接:$/ { getline; print; exit }' "$XRAY_INFO")
    [ -n "$link" ] || die "无法读取 VLESS 链接。"

    if command -v qrencode >/dev/null 2>&1; then
        qrencode -t UTF8 -m 2 "$link"
        qrencode -o "$XRAY_QR" "$link"
        chmod 600 "$XRAY_QR"
        echo
        echo "二维码图片已保存到：$XRAY_QR"
    else
        die "未安装 qrencode。"
    fi
}

change_sni() {
    require_root
    [ -f "$XRAY_CONFIG" ] || die "未找到配置文件：$XRAY_CONFIG"

    local new_sni tmp_config
    new_sni=${1:-}
    if [ -z "$new_sni" ]; then
        read -r -p "请输入新的 SNI 域名: " new_sni
    fi
    validate_sni "$new_sni" || die "SNI 格式不正确。"

    tmp_config=$(mktemp)
    jq --arg sni "$new_sni" '
        (.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.target) = ($sni + ":443")
        | (.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.serverNames) = [$sni]
    ' "$XRAY_CONFIG" >"$tmp_config"

    "$XRAY_BIN" -test -config "$tmp_config" >/dev/null
    mv "$tmp_config" "$XRAY_CONFIG"
    systemctl restart xray
    regenerate_client_info_from_config
    ok "SNI 已更新为：$new_sni"
    ok "Xray 已重启，客户端信息和二维码已重新生成。"
}

change_port() {
    require_root
    [ -f "$XRAY_CONFIG" ] || die "未找到配置文件：$XRAY_CONFIG"

    local new_port old_port tmp_config
    new_port=${1:-}
    if [ -z "$new_port" ]; then
        read -r -p "请输入新的监听端口: " new_port
    fi
    validate_port "$new_port" || die "端口必须是 1-65535 之间的数字。"

    old_port=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .port][0] // empty' "$XRAY_CONFIG")
    if [ "$new_port" = "$old_port" ]; then
        warn "新端口与当前端口相同，无需修改。"
        return
    fi

    tmp_config=$(mktemp)
    jq --argjson port "$new_port" '
        (.inbounds[] | select(.protocol == "vless") | .port) = $port
    ' "$XRAY_CONFIG" >"$tmp_config"

    "$XRAY_BIN" -test -config "$tmp_config" >/dev/null
    mv "$tmp_config" "$XRAY_CONFIG"
    if command -v ufw >/dev/null 2>&1; then
        ufw allow "${new_port}/tcp" >/dev/null 2>&1 || true
    fi
    systemctl restart xray
    regenerate_client_info_from_config
    ok "端口已更新为：$new_port"
    ok "Xray 已重启，客户端信息和二维码已重新生成。"
    warn "请确认 VPS 服务商安全组已放行 ${new_port}/tcp。"
}

force_ipv4_outbound() {
    require_root
    [ -f "$XRAY_CONFIG" ] || die "未找到配置文件：$XRAY_CONFIG"

    local tmp_config
    tmp_config=$(mktemp)
    jq '
        (.outbounds[] | select(.tag == "direct" and .protocol == "freedom") | .settings.domainStrategy) = "UseIPv4"
        | .routing = (.routing // {})
        | .routing.domainStrategy = "IPIfNonMatch"
        | .routing.rules = (
            [.routing.rules[]? | select(.outboundTag != "block" or (((.ip // []) | index("::/0")) | not))]
            + [{type: "field", ip: ["::/0"], outboundTag: "block"}]
        )
    ' "$XRAY_CONFIG" >"$tmp_config"

    "$XRAY_BIN" -test -config "$tmp_config" >/dev/null
    mv "$tmp_config" "$XRAY_CONFIG"
    systemctl restart xray
    ok "已强制 Xray 使用 IPv4 出站，并阻止 IPv6 目标。"
}

validate_dns_value() {
    local value=$1
    [[ "$value" =~ ^[A-Za-z0-9:._-]+$ ]]
}

change_dns() {
    require_root
    [ -f "$XRAY_CONFIG" ] || die "未找到配置文件：$XRAY_CONFIG"

    local dns1 dns2 tmp_config
    dns1=${1:-}
    dns2=${2:-}

    if [ -z "$dns1" ]; then
        read -r -p "请输入主 DNS [默认: ${DEFAULT_DNS_1}]: " dns1
        dns1=${dns1:-$DEFAULT_DNS_1}
    fi
    if [ -z "$dns2" ]; then
        read -r -p "请输入备用 DNS [默认: ${DEFAULT_DNS_2}]: " dns2
        dns2=${dns2:-$DEFAULT_DNS_2}
    fi

    validate_dns_value "$dns1" || die "主 DNS 格式不正确。"
    validate_dns_value "$dns2" || die "备用 DNS 格式不正确。"

    tmp_config=$(mktemp)
    jq --arg dns1 "$dns1" --arg dns2 "$dns2" '
        .dns = {
            servers: [$dns1, $dns2],
            queryStrategy: "UseIPv4"
        }
    ' "$XRAY_CONFIG" >"$tmp_config"

    "$XRAY_BIN" -test -config "$tmp_config" >/dev/null
    mv "$tmp_config" "$XRAY_CONFIG"
    systemctl restart xray
    ok "DNS 已更新为：$dns1, $dns2"
    ok "Xray 已重启。"
}

restart_service() {
    require_root
    systemctl restart xray
    ok "Xray 已重启。"
}

status_service() {
    systemctl status xray --no-pager
}

manager_menu() {
    while true; do
        echo
        echo -e "${BLUE}============== vless 管理菜单 ==============${NC}"
        echo "1. 查看节点信息"
        echo "2. 输出二维码"
        echo "3. 修改 SNI"
        echo "4. 修改端口"
        echo "5. 强制 IPv4 出站"
        echo "6. 修改 DNS"
        echo "7. 重启 Xray"
        echo "8. 查看 Xray 状态"
        echo "0. 退出"
        read -r -p "请选择: " choice
        case "$choice" in
            1) show_info ;;
            2) show_qr ;;
            3) change_sni ;;
            4) change_port ;;
            5) force_ipv4_outbound ;;
            6) change_dns ;;
            7) restart_service ;;
            8) status_service ;;
            0) exit 0 ;;
            *) warn "请输入 0-8。" ;;
        esac
    done
}

usage() {
    cat <<EOF
用法:
  bash install.sh              首次安装/重装
  vless                打开管理菜单
  vless show           查看节点信息
  vless qr             输出二维码
  vless sni <domain>   修改 SNI
  vless port <端口>     修改 VLESS 监听端口
  vless ipv4           强制 IPv4 出站并阻止 IPv6 目标
  vless dns <主DNS> <备用DNS>
                         修改 Xray DNS，例如：vless dns 1.1.1.1 8.8.8.8
  vless restart        重启 Xray
  vless status         查看状态
  vless install        重新安装

说明:
  安装完成后直接输入 vless 打开菜单。
  修改 SNI、端口、IPv4 出站或 DNS 会先校验配置，通过后自动重启 Xray。
EOF
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

dispatch() {
    local command_name
    command_name=${1:-}

    case "$command_name" in
        "" )
            if [ "$(basename "$0")" = "vless" ]; then
                manager_menu
            else
                main
            fi
            ;;
        install) main ;;
        menu) manager_menu ;;
        show) show_info ;;
        qr) show_qr ;;
        sni|change-sni) shift; change_sni "${1:-}" ;;
        port|change-port) shift; change_port "${1:-}" ;;
        ipv4|force-ipv4) force_ipv4_outbound ;;
        dns|change-dns) shift; change_dns "${1:-}" "${2:-}" ;;
        restart) restart_service ;;
        status) status_service ;;
        help|-h|--help) usage ;;
        *) usage; exit 1 ;;
    esac
}

dispatch "$@"
