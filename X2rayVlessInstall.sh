#!/usr/bin/env bash

VERSION="2.4.14"
BUILD_DATE="2026-09-22"

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
    apt-get install -y curl jq openssl iproute2 ca-certificates qrencode dnsutils
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
    echo "  1. www.dell.com"
    echo "  2. shopee.sg"
    echo "  3. aws.amazon.com"
    echo "  4. www.lovelive-anime.jp"
    echo "  5. www.sjsu.edu"
    echo "  6. 自定义"
    while true; do
        read -r -p "请选择 1-6: " choice
        case "$choice" in
            1) SNI="www.dell.com" ;;
            2) SNI="shopee.sg" ;;
            3) SNI="aws.amazon.com" ;;
            4) SNI="www.lovelive-anime.jp" ;;
            5) SNI="www.sjsu.edu" ;;
            6)
                read -r -p "请输入 SNI 域名: " SNI
                validate_sni "$SNI" || { warn "SNI 格式不正确，请重新选择。"; continue; }
                ;;
            *) warn "请输入 1-6。"; continue ;;
        esac
        if confirm_sni_cdn "$SNI"; then
            break
        fi
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

test_xray_config() {
    local config_path=$1 log_path
    log_path=$(mktemp)

    if ! "$XRAY_BIN" -test -config "$config_path" >"$log_path" 2>&1; then
        warn "Xray 配置校验失败，原始输出："
        cat "$log_path" >&2 || true
        rm -f "$log_path"
        return 1
    fi

    rm -f "$log_path"
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
                            dest: ($sni + ":443"),
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

    local gate_config gate_port
    gate_port=$(choose_guard_port)
    gate_config=$(mktemp "${XRAY_CONFIG}.guard.XXXXXX")
    guard_config "$XRAY_CONFIG" "$gate_config" "$gate_port" || die "生成黑洞防护失败。"
    mv "$gate_config" "$XRAY_CONFIG"

    local user
    user=$(xray_user)
    chmod 644 "$XRAY_CONFIG"
    chmod 755 "$XRAY_LOG_DIR"
    chmod 644 "$XRAY_LOG_DIR/access.log" "$XRAY_LOG_DIR/error.log"
    if [ "$user" != "root" ] && id "$user" >/dev/null 2>&1; then
        chown -R "$user":"$user" "$XRAY_LOG_DIR" 2>/dev/null || chown -R "$user":nogroup "$XRAY_LOG_DIR" 2>/dev/null || true
    fi

    test_xray_config "$XRAY_CONFIG" || die "配置校验失败。"
    ok "配置校验通过。"
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
    link="vless://${UUID}@${SERVER_IP}:${PORT}?type=tcp&security=reality&encryption=none&flow=${FLOW}&pbk=${PUBLIC_KEY}&fp=${FINGERPRINT}&sni=${SNI}&sid=${SHORT_ID}&spx=%2F#${encoded_name}"

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

# These checks describe the current server's IPv4 DNS view, matching UseIPv4.
# CNAME references: AWS CloudFront CNAMEs, Akamai edge hostnames, Fastly routing docs.
cdn_provider_for_name() {
    local name
    name=$(printf '%s' "${1%.}" | tr '[:upper:]' '[:lower:]')
    case "$name" in
        *.cloudfront.net) printf 'Amazon CloudFront\n' ;;
        *.edgekey.net|*.edgesuite.net|*.akamaiedge.net|*.akamaized.net) printf 'Akamai\n' ;;
        *.fastly.net) printf 'Fastly\n' ;;
    esac
}

# Print matching IP/CIDR pairs. Reject a malformed list instead of treating it
# as an empty (successful) lookup; awk uses exact integers for 32-bit IPv4.
cdn_match_ipv4_ranges() {
    local addresses=${1//$'\n'/ }
    awk -v addresses="$addresses" '
        function ipnum(ip, octets, n, i, result) {
            n = split(ip, octets, ".")
            if (n != 4) return -1
            result = 0
            for (i = 1; i <= 4; i++) {
                if (octets[i] !~ /^[0-9]+$/ || octets[i] + 0 > 255) return -1
                result = result * 256 + octets[i]
            }
            return result
        }
        NF {
            if (NF != 1 || split($1, parts, "/") != 2 ||
                ipnum(parts[1]) < 0 || parts[2] !~ /^[0-9]+$/ || parts[2] + 0 > 32) {
                invalid = 1; next
            }
            cidrs[++count] = $1
            sizes[count] = 2 ^ (32 - parts[2])
            networks[count] = int(ipnum(parts[1]) / sizes[count])
        }
        END {
            if (invalid || !count) exit 2
            total = split(addresses, ips, /[[:space:]]+/)
            for (i = 1; i <= total; i++) {
                value = ipnum(ips[i])
                if (value < 0) continue
                for (j = 1; j <= count; j++) {
                    if (int(value / sizes[j]) == networks[j]) {
                        print ips[i] " ∈ " cidrs[j]
                        break
                    }
                }
            }
        }
    '
}

check_sni_cdn() {
    local domain answer aliases addresses alias provider matches ranges headers first_ip
    local detected=0 suspected=0 incomplete=0
    CDN_STATUS=unknown
    domain=$(printf '%s' "${1%.}" | tr '[:upper:]' '[:lower:]')
    printf '\n正在检测 SNI 是否使用 CDN：%s（当前服务器 IPv4 视角）\n' "$domain"
    if ! validate_sni "$domain"; then
        warn "[检测失败] SNI 域名格式不正确。"
        return 0
    fi
    if ! command -v dig >/dev/null 2>&1 || ! command -v curl >/dev/null 2>&1; then
        warn "[检测失败] 缺少 dig 或 curl；Debian/Ubuntu 可安装 dnsutils curl。"
        return 0
    fi
    if ! answer=$(dig +time=2 +tries=1 +noall +answer +comments "$domain" A 2>/dev/null || exit $?) ||
       [[ "$answer" != *"status: NOERROR,"* ]]; then
        warn "[检测失败] DNS 查询失败或域名不存在，无法判断是否使用 CDN。"
        return 0
    fi
    aliases=$(printf '%s\n' "$answer" | awk '$4 == "CNAME" {print $5}')
    addresses=$(printf '%s\n' "$answer" | awk '$4 == "A" {print $5}' | sort -u)
    # Include the input name in case the user directly selects a CDN hostname.
    while IFS= read -r alias; do
        [ -n "$alias" ] || continue
        provider=$(cdn_provider_for_name "$alias")
        if [ -n "$provider" ]; then
            detected=1
            printf '  CDN 域名特征：%s → %s\n' "$alias" "$provider"
        fi
    done <<<"$(printf '%s\n%s' "$domain" "$aliases")"
    if [ -z "$addresses" ]; then
        incomplete=1
        warn "  未解析到 IPv4 地址；当前脚本的 UseIPv4 配置可能无法连接此目标。"
    else
        printf '  IPv4 地址：%s\n' "$(printf '%s' "$addresses" | tr '\n' ' ')"
        # RFC 2544 benchmarking addresses are also commonly used by Fake-IP DNS.
        if printf "%s\n" "$addresses" | grep -Eq "^198\.(18|19)\."; then
            incomplete=1
            warn "  解析到测试网段（可能为 Fake-IP），无法据此判断真实 IP 的 CDN 归属。"
        fi
        # Official proxy ranges only. DNS hosting/ASN alone is not evidence of CDN.
        # Cache successful downloads in this process; never execute downloaded data.
        ranges=${CDN_CF_RANGES:-}
        if [ -z "$ranges" ]; then
            if ! ranges=$(curl -q --noproxy '*' --proto '=https' -fsS \
                --connect-timeout 3 --max-time 6 --max-filesize 65536 \
                https://www.cloudflare.com/ips-v4 2>/dev/null || exit $?); then
                ranges=''
            fi
        fi
        if matches=$(printf '%s\n' "$ranges" | cdn_match_ipv4_ranges "$addresses" || exit $?); then
            CDN_CF_RANGES=$ranges
            if [ -n "$matches" ]; then
                detected=1
                printf '  Cloudflare 官方代理网段命中：\n%s\n' "$matches"
            fi
        else
            incomplete=1
            warn "  Cloudflare 官方 IP 网段获取或校验失败，IP 检测未完成。"
        fi
        if [ "$detected" -eq 0 ]; then
            first_ip=${addresses%%$'\n'*}
            # Do not follow redirects: headers must belong to this SNI, not a
            # different site's redirect destination. Bypass proxy environment vars.
            if headers=$(curl -q --noproxy '*' --proto '=https' -4 -sS -I \
                --connect-timeout 3 --max-time 6 --max-filesize 65536 \
                --resolve "${domain}:443:${first_ip}" "https://${domain}/" 2>/dev/null || exit $?); then
                headers=$(printf '%s\n' "$headers" | tr -d '\r' | tr '[:upper:]' '[:lower:]')
                if printf '%s\n' "$headers" | grep -Eq '^(cf-ray:|cf-cache-status:|server:[[:space:]]*cloudflare([[:space:]]|$))'; then
                    suspected=1
                    printf '  HTTPS 响应头：发现 Cloudflare 特征（辅助证据）。\n'
                fi
                if printf '%s\n' "$headers" | grep -Eq '^x-amz-cf-(id|pop):'; then
                    suspected=1
                    printf '  HTTPS 响应头：发现 CloudFront 特征（辅助证据）。\n'
                fi
            else
                incomplete=1
                warn "  HTTPS 响应头检测失败（连接、证书或超时），不能据此认定没有 CDN。"
            fi
        fi
    fi
    if [ "$detected" -eq 1 ]; then
        CDN_STATUS=detected
        warn "[检测到 CDN] 建议考虑其他目标；允许的 SNI 仍可能产生转发流量。"
    elif [ "$suspected" -eq 1 ]; then
        CDN_STATUS=suspected
        warn "[疑似 CDN] 响应头仅为辅助证据，可能被修改或伪造。"
    elif [ "$incomplete" -eq 1 ]; then
        warn "[检测未完成] 无法确认是否使用 CDN。"
    else
        CDN_STATUS=not_detected
        ok "[未发现 CDN 特征] 检测范围有限，不代表确定没有 CDN。"
    fi
    return 0
}

confirm_sni_cdn() {
    local answer
    check_sni_cdn "$1"
    if [ -t 0 ] && { [ "$CDN_STATUS" = detected ] || [ "$CDN_STATUS" = suspected ]; }; then
        read -r -p "继续使用此 SNI？[Y/n]: " answer || return 1
        case "$answer" in n|N|no|NO) return 1 ;; esac
    fi
    return 0
}

validate_sni() {
    local value=${1%.} label
    local -a labels
    [ -n "$value" ] && [ "${#value}" -le 253 ] || return 1
    [[ "$value" =~ ^[A-Za-z0-9.-]+$ ]] || return 1
    [[ "$value" != .* && "$value" != *. && "$value" != *..* ]] || return 1
    IFS=. read -r -a labels <<<"$value"
    for label in "${labels[@]}"; do
        [ "${#label}" -le 63 ] || return 1
        [[ "$label" != -* && "$label" != *- ]] || return 1
    done
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

test_xray_config() {
    local config_path=$1 log_path
    log_path=$(mktemp)

    if ! "$XRAY_BIN" -test -config "$config_path" >"$log_path" 2>&1; then
        warn "Xray 配置校验失败，原始输出："
        cat "$log_path" >&2 || true
        rm -f "$log_path"
        return 1
    fi

    rm -f "$log_path"
}

repair_reality_config() {
    require_root
    [ -f "$XRAY_CONFIG" ] || die "未找到配置文件：$XRAY_CONFIG"

    local old_target tmp_config
    old_target=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.target][0] // empty' "$XRAY_CONFIG")
    if [ -z "$old_target" ]; then
        test_xray_config "$XRAY_CONFIG" || die "当前配置校验失败。"
        regenerate_client_info_from_config
        ok "未发现需要转换的 target 字段，已重新生成客户端信息和二维码。"
        return
    fi

    tmp_config=$(mktemp --suffix=.json)
    jq --arg oldTarget "$old_target" '
        (.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.dest) = $oldTarget
        | del(.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.target)
    ' "$XRAY_CONFIG" >"$tmp_config"

    if ! sync_guard_config "$tmp_config" || ! test_xray_config "$tmp_config"; then
        rm -f "$tmp_config"
        die "修复后的配置校验失败，未覆盖原配置。"
    fi
    chown --reference="$XRAY_CONFIG" "$tmp_config"
    chmod --reference="$XRAY_CONFIG" "$tmp_config"
    mv "$tmp_config" "$XRAY_CONFIG"
    systemctl restart xray
    regenerate_client_info_from_config
    ok "Reality 配置已修复：target -> dest"
    ok "Xray 已重启，客户端信息和二维码已重新生成。"
}

regenerate_client_info_from_config() {
    [ -f "$XRAY_CONFIG" ] || die "未找到配置文件：$XRAY_CONFIG"

    local uuid port flow email node_name sni short_id private_key public_key server_ip encoded_name link
    uuid=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .settings.clients[0].id][0] // empty' "$XRAY_CONFIG")
    port=$(jq -r '([.inbounds[] | select(.tag == "reality-sni-gate") | .port][0] // [.inbounds[] | select(.protocol == "vless") | .port][0]) // empty' "$XRAY_CONFIG")
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
    link="vless://${uuid}@${server_ip}:${port}?type=tcp&security=reality&encryption=none&flow=${flow}&pbk=${public_key}&fp=${FINGERPRINT}&sni=${sni}&sid=${short_id}&spx=%2F#${encoded_name}"

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
    confirm_sni_cdn "$new_sni" || { warn "已取消修改 SNI，原配置未修改。"; return 0; }

    tmp_config=$(mktemp --suffix=.json)
    jq --arg sni "$new_sni" '
        (.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.dest) = ($sni + ":443")
        | del(.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.target)
        | (.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.serverNames) = [$sni]
    ' "$XRAY_CONFIG" >"$tmp_config"

    if ! sync_guard_config "$tmp_config" || ! test_xray_config "$tmp_config"; then
        rm -f "$tmp_config"
        die "新 SNI 配置校验失败，未覆盖原配置。"
    fi
    chown --reference="$XRAY_CONFIG" "$tmp_config"
    chmod --reference="$XRAY_CONFIG" "$tmp_config"
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

    old_port=$(jq -r '([.inbounds[] | select(.tag == "reality-sni-gate") | .port][0] // [.inbounds[] | select(.protocol == "vless") | .port][0]) // empty' "$XRAY_CONFIG")
    if [ "$new_port" = "$old_port" ]; then
        warn "新端口与当前端口相同，无需修改。"
        return
    fi

    tmp_config=$(mktemp --suffix=.json)
    jq --argjson port "$new_port" '
        if any(.inbounds[]; .tag == "reality-sni-gate") then
            (.inbounds[] | select(.tag == "reality-sni-gate") | .port) = $port
        else (.inbounds[] | select(.protocol == "vless") | .port) = $port end
    ' "$XRAY_CONFIG" >"$tmp_config"

    if ! sync_guard_config "$tmp_config" || ! test_xray_config "$tmp_config"; then
        rm -f "$tmp_config"
        die "新端口配置校验失败，未覆盖原配置。"
    fi
    chown --reference="$XRAY_CONFIG" "$tmp_config"
    chmod --reference="$XRAY_CONFIG" "$tmp_config"
    mv "$tmp_config" "$XRAY_CONFIG"
    systemctl restart xray
    regenerate_client_info_from_config
    ok "端口已更新为：$new_port"
    ok "Xray 已重启，客户端信息和二维码已重新生成。"
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

    tmp_config=$(mktemp --suffix=.json)
    jq --arg dns1 "$dns1" --arg dns2 "$dns2" '
        .dns = {
            servers: [$dns1, $dns2],
            queryStrategy: "UseIPv4"
        }
    ' "$XRAY_CONFIG" >"$tmp_config"

    if ! sync_guard_config "$tmp_config" || ! test_xray_config "$tmp_config"; then
        rm -f "$tmp_config"
        die "新 DNS 配置校验失败，未覆盖原配置。"
    fi
    chown --reference="$XRAY_CONFIG" "$tmp_config"
    chmod --reference="$XRAY_CONFIG" "$tmp_config"
    mv "$tmp_config" "$XRAY_CONFIG"
    systemctl restart xray
    ok "DNS 已更新为：$dns1, $dns2"
    ok "Xray 已重启。"
}

# Filter the REALITY target connection; VLESS itself keeps the public listener.
guard_config() {
    local input=$1 output=$2 internal_port=${3:-45987}
    jq --argjson internal "$internal_port" '
        [.inbounds[] | select(.protocol == "vless")] as $v
        | if ($v | length) != 1 or $v[0].streamSettings.security != "reality"
          then error("仅支持单个 VLESS REALITY 入站") else . end
        | $v[0] as $v
        | ([.inbounds[] | select(.tag == "reality-sni-gate")][0] // null) as $old
        | ([.inbounds[] | select(.tag == "reality-target-gate")][0] // null) as $gate
        | ($v.streamSettings.realitySettings.serverNames | map(select(length > 0) | "full:" + .)) as $names
        | if ($names | length) == 0 then error("缺少非空 SNI") else . end
        | ($old.port // $v.port) as $public
        | ($gate.port // (if $old != null then $v.port else $internal end)) as $private
        | if $public == $private or any(.inbounds[];
            .protocol != "vless" and .tag != "reality-sni-gate" and .tag != "reality-target-gate" and .port == $private)
          then error("内部端口与其他入站端口冲突") else . end
        | ($v.streamSettings.realitySettings.target // $v.streamSettings.realitySettings.dest) as $target
        | (if $gate != null and $target == ("127.0.0.1:" + ($gate.port | tostring))
           then $gate.settings
           else ($target | capture("^(?<address>.+):(?<port>[0-9]+)$")
                 | .port |= tonumber | .address |= ltrimstr("[") | .address |= rtrimstr("]")) end) as $remote
        | if $remote == null then error("不支持的 REALITY 目标格式") else . end
        | .inbounds = ([{
            tag: "reality-target-gate", listen: "127.0.0.1",
            port: $private, protocol: "dokodemo-door",
            settings: {address: $remote.address, port: $remote.port, network: "tcp"},
            sniffing: {enabled: true, destOverride: ["tls"], routeOnly: true}
          }] + [.inbounds[] | select(.tag != "reality-sni-gate" and .tag != "reality-target-gate")
            | if .protocol == "vless" then
                .listen = ($old.listen // $v.listen // "0.0.0.0") | .port = $public
                | .streamSettings.realitySettings.dest = ("127.0.0.1:" + ($private | tostring))
                | del(.streamSettings.realitySettings.target)
              else . end])
        | .outbounds = ([.outbounds[] | select(.tag != "reality-gate-direct" and .tag != "reality-gate-block")]
            + [{tag: "reality-gate-direct", protocol: "freedom", settings: {domainStrategy: "UseIPv4"}},
               {tag: "reality-gate-block", protocol: "blackhole"}])
        | .routing.rules = ([
            {type: "field", inboundTag: ["reality-target-gate"], domain: $names, outboundTag: "reality-gate-direct"},
            {type: "field", inboundTag: ["reality-target-gate"], outboundTag: "reality-gate-block"}
          ] + [(.routing.rules // [])[] | select(
            ((.inboundTag // []) | index("reality-sni-gate")) == null and
            ((.inboundTag // []) | index("reality-target-gate")) == null)])
    ' "$input" >"$output" || return 1
    [ -s "$output" ]
}

choose_guard_port() {
    local candidate listeners
    listeners=$(ss -H -ltn) || return 1
    for candidate in {45987..46087}; do
        if ! jq -e --argjson p "$candidate" 'any(.inbounds[]; .port == $p)' "$XRAY_CONFIG" >/dev/null &&
           ! printf '%s\n' "$listeners" | awk '{print $4}' | grep -Eq ":${candidate}$"; then
            printf '%s\n' "$candidate"
            return
        fi
    done
    die "未找到可用的内部端口。"
}

sync_guard_config() {
    local path=$1 tmp
    if jq -e 'any(.inbounds[]; .tag == "reality-sni-gate" or .tag == "reality-target-gate")' "$path" >/dev/null; then
        tmp=$(mktemp "${path}.guard.XXXXXX")
        if ! guard_config "$path" "$tmp"; then
            rm -f "$tmp"
            return 1
        fi
        cat "$tmp" >"$path"
        rm -f "$tmp"
    fi
}

enable_blackhole() {
    require_root
    [ -f "$XRAY_CONFIG" ] || die "未找到配置文件：$XRAY_CONFIG"
    local tmp backup internal_port
    internal_port=$(choose_guard_port)
    tmp=$(mktemp "${XRAY_CONFIG}.guard.XXXXXX")
    if ! guard_config "$XRAY_CONFIG" "$tmp" "$internal_port" || ! test_xray_config "$tmp"; then
        rm -f "$tmp"
        die "黑洞配置校验失败，原配置未修改。"
    fi
    backup=$(mktemp "${XRAY_CONFIG}.backup.XXXXXX")
    cp -p "$XRAY_CONFIG" "$backup"
    # Preserve the existing service-readable ownership and mode.
    chown --reference="$XRAY_CONFIG" "$tmp"
    chmod --reference="$XRAY_CONFIG" "$tmp"
    mv "$tmp" "$XRAY_CONFIG"
    if ! systemctl restart xray || ! systemctl is-active --quiet xray; then
        mv "$backup" "$XRAY_CONFIG"
        systemctl restart xray || true
        die "启动失败，已恢复原配置。"
    fi
    rm -f "$backup"
    ok "SNI 黑洞防护已启用，客户端链接不变。"
    warn "目标转发仅放行配置中的 SNI；访问允许的目标仍会消耗带宽。"
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
        echo "5. 修改 DNS"
        echo "6. 重启 Xray"
        echo "7. 查看 Xray 状态"
        echo "8. 修复 Reality 配置"
        echo "9. 启用 SNI 黑洞防护"
        echo "0. 退出"
        read -r -p "请选择: " choice
        case "$choice" in
            1) show_info ;;
            2) show_qr ;;
            3) change_sni ;;
            4) change_port ;;
            5) change_dns ;;
            6) restart_service ;;
            7) status_service ;;
            8) repair_reality_config ;;
            9) enable_blackhole ;;
            0) exit 0 ;;
            *) warn "请输入 0-9。" ;;
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
  vless dns <主DNS> <备用DNS>
                         修改 Xray DNS，例如：vless dns 1.1.1.1 8.8.8.8
  vless restart        重启 Xray
  vless status         查看状态
  vless repair         转换 Reality target 字段为兼容字段 dest
  vless blackhole      启用 SNI 黑洞防护（保留客户端链接）
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
        dns|change-dns) shift; change_dns "${1:-}" "${2:-}" ;;
        restart) restart_service ;;
        status) status_service ;;
        repair) repair_reality_config ;;
        blackhole) enable_blackhole ;;
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

# These checks describe the current server's IPv4 DNS view, matching UseIPv4.
# CNAME references: AWS CloudFront CNAMEs, Akamai edge hostnames, Fastly routing docs.
cdn_provider_for_name() {
    local name
    name=$(printf '%s' "${1%.}" | tr '[:upper:]' '[:lower:]')
    case "$name" in
        *.cloudfront.net) printf 'Amazon CloudFront\n' ;;
        *.edgekey.net|*.edgesuite.net|*.akamaiedge.net|*.akamaized.net) printf 'Akamai\n' ;;
        *.fastly.net) printf 'Fastly\n' ;;
    esac
}

# Print matching IP/CIDR pairs. Reject a malformed list instead of treating it
# as an empty (successful) lookup; awk uses exact integers for 32-bit IPv4.
cdn_match_ipv4_ranges() {
    local addresses=${1//$'\n'/ }
    awk -v addresses="$addresses" '
        function ipnum(ip, octets, n, i, result) {
            n = split(ip, octets, ".")
            if (n != 4) return -1
            result = 0
            for (i = 1; i <= 4; i++) {
                if (octets[i] !~ /^[0-9]+$/ || octets[i] + 0 > 255) return -1
                result = result * 256 + octets[i]
            }
            return result
        }
        NF {
            if (NF != 1 || split($1, parts, "/") != 2 ||
                ipnum(parts[1]) < 0 || parts[2] !~ /^[0-9]+$/ || parts[2] + 0 > 32) {
                invalid = 1; next
            }
            cidrs[++count] = $1
            sizes[count] = 2 ^ (32 - parts[2])
            networks[count] = int(ipnum(parts[1]) / sizes[count])
        }
        END {
            if (invalid || !count) exit 2
            total = split(addresses, ips, /[[:space:]]+/)
            for (i = 1; i <= total; i++) {
                value = ipnum(ips[i])
                if (value < 0) continue
                for (j = 1; j <= count; j++) {
                    if (int(value / sizes[j]) == networks[j]) {
                        print ips[i] " ∈ " cidrs[j]
                        break
                    }
                }
            }
        }
    '
}

check_sni_cdn() {
    local domain answer aliases addresses alias provider matches ranges headers first_ip
    local detected=0 suspected=0 incomplete=0
    CDN_STATUS=unknown
    domain=$(printf '%s' "${1%.}" | tr '[:upper:]' '[:lower:]')
    printf '\n正在检测 SNI 是否使用 CDN：%s（当前服务器 IPv4 视角）\n' "$domain"
    if ! validate_sni "$domain"; then
        warn "[检测失败] SNI 域名格式不正确。"
        return 0
    fi
    if ! command -v dig >/dev/null 2>&1 || ! command -v curl >/dev/null 2>&1; then
        warn "[检测失败] 缺少 dig 或 curl；Debian/Ubuntu 可安装 dnsutils curl。"
        return 0
    fi
    if ! answer=$(dig +time=2 +tries=1 +noall +answer +comments "$domain" A 2>/dev/null || exit $?) ||
       [[ "$answer" != *"status: NOERROR,"* ]]; then
        warn "[检测失败] DNS 查询失败或域名不存在，无法判断是否使用 CDN。"
        return 0
    fi
    aliases=$(printf '%s\n' "$answer" | awk '$4 == "CNAME" {print $5}')
    addresses=$(printf '%s\n' "$answer" | awk '$4 == "A" {print $5}' | sort -u)
    # Include the input name in case the user directly selects a CDN hostname.
    while IFS= read -r alias; do
        [ -n "$alias" ] || continue
        provider=$(cdn_provider_for_name "$alias")
        if [ -n "$provider" ]; then
            detected=1
            printf '  CDN 域名特征：%s → %s\n' "$alias" "$provider"
        fi
    done <<<"$(printf '%s\n%s' "$domain" "$aliases")"
    if [ -z "$addresses" ]; then
        incomplete=1
        warn "  未解析到 IPv4 地址；当前脚本的 UseIPv4 配置可能无法连接此目标。"
    else
        printf '  IPv4 地址：%s\n' "$(printf '%s' "$addresses" | tr '\n' ' ')"
        # RFC 2544 benchmarking addresses are also commonly used by Fake-IP DNS.
        if printf "%s\n" "$addresses" | grep -Eq "^198\.(18|19)\."; then
            incomplete=1
            warn "  解析到测试网段（可能为 Fake-IP），无法据此判断真实 IP 的 CDN 归属。"
        fi
        # Official proxy ranges only. DNS hosting/ASN alone is not evidence of CDN.
        # Cache successful downloads in this process; never execute downloaded data.
        ranges=${CDN_CF_RANGES:-}
        if [ -z "$ranges" ]; then
            if ! ranges=$(curl -q --noproxy '*' --proto '=https' -fsS \
                --connect-timeout 3 --max-time 6 --max-filesize 65536 \
                https://www.cloudflare.com/ips-v4 2>/dev/null || exit $?); then
                ranges=''
            fi
        fi
        if matches=$(printf '%s\n' "$ranges" | cdn_match_ipv4_ranges "$addresses" || exit $?); then
            CDN_CF_RANGES=$ranges
            if [ -n "$matches" ]; then
                detected=1
                printf '  Cloudflare 官方代理网段命中：\n%s\n' "$matches"
            fi
        else
            incomplete=1
            warn "  Cloudflare 官方 IP 网段获取或校验失败，IP 检测未完成。"
        fi
        if [ "$detected" -eq 0 ]; then
            first_ip=${addresses%%$'\n'*}
            # Do not follow redirects: headers must belong to this SNI, not a
            # different site's redirect destination. Bypass proxy environment vars.
            if headers=$(curl -q --noproxy '*' --proto '=https' -4 -sS -I \
                --connect-timeout 3 --max-time 6 --max-filesize 65536 \
                --resolve "${domain}:443:${first_ip}" "https://${domain}/" 2>/dev/null || exit $?); then
                headers=$(printf '%s\n' "$headers" | tr -d '\r' | tr '[:upper:]' '[:lower:]')
                if printf '%s\n' "$headers" | grep -Eq '^(cf-ray:|cf-cache-status:|server:[[:space:]]*cloudflare([[:space:]]|$))'; then
                    suspected=1
                    printf '  HTTPS 响应头：发现 Cloudflare 特征（辅助证据）。\n'
                fi
                if printf '%s\n' "$headers" | grep -Eq '^x-amz-cf-(id|pop):'; then
                    suspected=1
                    printf '  HTTPS 响应头：发现 CloudFront 特征（辅助证据）。\n'
                fi
            else
                incomplete=1
                warn "  HTTPS 响应头检测失败（连接、证书或超时），不能据此认定没有 CDN。"
            fi
        fi
    fi
    if [ "$detected" -eq 1 ]; then
        CDN_STATUS=detected
        warn "[检测到 CDN] 建议考虑其他目标；允许的 SNI 仍可能产生转发流量。"
    elif [ "$suspected" -eq 1 ]; then
        CDN_STATUS=suspected
        warn "[疑似 CDN] 响应头仅为辅助证据，可能被修改或伪造。"
    elif [ "$incomplete" -eq 1 ]; then
        warn "[检测未完成] 无法确认是否使用 CDN。"
    else
        CDN_STATUS=not_detected
        ok "[未发现 CDN 特征] 检测范围有限，不代表确定没有 CDN。"
    fi
    return 0
}

confirm_sni_cdn() {
    local answer
    check_sni_cdn "$1"
    if [ -t 0 ] && { [ "$CDN_STATUS" = detected ] || [ "$CDN_STATUS" = suspected ]; }; then
        read -r -p "继续使用此 SNI？[Y/n]: " answer || return 1
        case "$answer" in n|N|no|NO) return 1 ;; esac
    fi
    return 0
}

validate_sni() {
    local value=${1%.} label
    local -a labels
    [ -n "$value" ] && [ "${#value}" -le 253 ] || return 1
    [[ "$value" =~ ^[A-Za-z0-9.-]+$ ]] || return 1
    [[ "$value" != .* && "$value" != *. && "$value" != *..* ]] || return 1
    IFS=. read -r -a labels <<<"$value"
    for label in "${labels[@]}"; do
        [ "${#label}" -le 63 ] || return 1
        [[ "$label" != -* && "$label" != *- ]] || return 1
    done
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

repair_reality_config() {
    require_root
    [ -f "$XRAY_CONFIG" ] || die "未找到配置文件：$XRAY_CONFIG"

    local old_target tmp_config
    old_target=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.target][0] // empty' "$XRAY_CONFIG")
    if [ -z "$old_target" ]; then
        test_xray_config "$XRAY_CONFIG" || die "当前配置校验失败。"
        regenerate_client_info_from_config
        ok "未发现需要转换的 target 字段，已重新生成客户端信息和二维码。"
        return
    fi

    tmp_config=$(mktemp --suffix=.json)
    jq --arg oldTarget "$old_target" '
        (.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.dest) = $oldTarget
        | del(.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.target)
    ' "$XRAY_CONFIG" >"$tmp_config"

    if ! sync_guard_config "$tmp_config" || ! test_xray_config "$tmp_config"; then
        rm -f "$tmp_config"
        die "修复后的配置校验失败，未覆盖原配置。"
    fi
    chown --reference="$XRAY_CONFIG" "$tmp_config"
    chmod --reference="$XRAY_CONFIG" "$tmp_config"
    mv "$tmp_config" "$XRAY_CONFIG"
    systemctl restart xray
    regenerate_client_info_from_config
    ok "Reality 配置已修复：target -> dest"
    ok "Xray 已重启，客户端信息和二维码已重新生成。"
}

regenerate_client_info_from_config() {
    [ -f "$XRAY_CONFIG" ] || die "未找到配置文件：$XRAY_CONFIG"

    local uuid port flow email node_name sni short_id private_key public_key server_ip encoded_name link
    uuid=$(jq -r '[.inbounds[] | select(.protocol == "vless") | .settings.clients[0].id][0] // empty' "$XRAY_CONFIG")
    port=$(jq -r '([.inbounds[] | select(.tag == "reality-sni-gate") | .port][0] // [.inbounds[] | select(.protocol == "vless") | .port][0]) // empty' "$XRAY_CONFIG")
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
    link="vless://${uuid}@${server_ip}:${port}?type=tcp&security=reality&encryption=none&flow=${flow}&pbk=${public_key}&fp=${FINGERPRINT}&sni=${sni}&sid=${short_id}&spx=%2F#${encoded_name}"

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
    confirm_sni_cdn "$new_sni" || { warn "已取消修改 SNI，原配置未修改。"; return 0; }

    tmp_config=$(mktemp --suffix=.json)
    jq --arg sni "$new_sni" '
        (.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.dest) = ($sni + ":443")
        | del(.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.target)
        | (.inbounds[] | select(.protocol == "vless") | .streamSettings.realitySettings.serverNames) = [$sni]
    ' "$XRAY_CONFIG" >"$tmp_config"

    if ! sync_guard_config "$tmp_config" || ! test_xray_config "$tmp_config"; then
        rm -f "$tmp_config"
        die "新 SNI 配置校验失败，未覆盖原配置。"
    fi
    chown --reference="$XRAY_CONFIG" "$tmp_config"
    chmod --reference="$XRAY_CONFIG" "$tmp_config"
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

    old_port=$(jq -r '([.inbounds[] | select(.tag == "reality-sni-gate") | .port][0] // [.inbounds[] | select(.protocol == "vless") | .port][0]) // empty' "$XRAY_CONFIG")
    if [ "$new_port" = "$old_port" ]; then
        warn "新端口与当前端口相同，无需修改。"
        return
    fi

    tmp_config=$(mktemp --suffix=.json)
    jq --argjson port "$new_port" '
        if any(.inbounds[]; .tag == "reality-sni-gate") then
            (.inbounds[] | select(.tag == "reality-sni-gate") | .port) = $port
        else (.inbounds[] | select(.protocol == "vless") | .port) = $port end
    ' "$XRAY_CONFIG" >"$tmp_config"

    if ! sync_guard_config "$tmp_config" || ! test_xray_config "$tmp_config"; then
        rm -f "$tmp_config"
        die "新端口配置校验失败，未覆盖原配置。"
    fi
    chown --reference="$XRAY_CONFIG" "$tmp_config"
    chmod --reference="$XRAY_CONFIG" "$tmp_config"
    mv "$tmp_config" "$XRAY_CONFIG"
    systemctl restart xray
    regenerate_client_info_from_config
    ok "端口已更新为：$new_port"
    ok "Xray 已重启，客户端信息和二维码已重新生成。"
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

    tmp_config=$(mktemp --suffix=.json)
    jq --arg dns1 "$dns1" --arg dns2 "$dns2" '
        .dns = {
            servers: [$dns1, $dns2],
            queryStrategy: "UseIPv4"
        }
    ' "$XRAY_CONFIG" >"$tmp_config"

    if ! sync_guard_config "$tmp_config" || ! test_xray_config "$tmp_config"; then
        rm -f "$tmp_config"
        die "新 DNS 配置校验失败，未覆盖原配置。"
    fi
    chown --reference="$XRAY_CONFIG" "$tmp_config"
    chmod --reference="$XRAY_CONFIG" "$tmp_config"
    mv "$tmp_config" "$XRAY_CONFIG"
    systemctl restart xray
    ok "DNS 已更新为：$dns1, $dns2"
    ok "Xray 已重启。"
}

# Filter the REALITY target connection; VLESS itself keeps the public listener.
guard_config() {
    local input=$1 output=$2 internal_port=${3:-45987}
    jq --argjson internal "$internal_port" '
        [.inbounds[] | select(.protocol == "vless")] as $v
        | if ($v | length) != 1 or $v[0].streamSettings.security != "reality"
          then error("仅支持单个 VLESS REALITY 入站") else . end
        | $v[0] as $v
        | ([.inbounds[] | select(.tag == "reality-sni-gate")][0] // null) as $old
        | ([.inbounds[] | select(.tag == "reality-target-gate")][0] // null) as $gate
        | ($v.streamSettings.realitySettings.serverNames | map(select(length > 0) | "full:" + .)) as $names
        | if ($names | length) == 0 then error("缺少非空 SNI") else . end
        | ($old.port // $v.port) as $public
        | ($gate.port // (if $old != null then $v.port else $internal end)) as $private
        | if $public == $private or any(.inbounds[];
            .protocol != "vless" and .tag != "reality-sni-gate" and .tag != "reality-target-gate" and .port == $private)
          then error("内部端口与其他入站端口冲突") else . end
        | ($v.streamSettings.realitySettings.target // $v.streamSettings.realitySettings.dest) as $target
        | (if $gate != null and $target == ("127.0.0.1:" + ($gate.port | tostring))
           then $gate.settings
           else ($target | capture("^(?<address>.+):(?<port>[0-9]+)$")
                 | .port |= tonumber | .address |= ltrimstr("[") | .address |= rtrimstr("]")) end) as $remote
        | if $remote == null then error("不支持的 REALITY 目标格式") else . end
        | .inbounds = ([{
            tag: "reality-target-gate", listen: "127.0.0.1",
            port: $private, protocol: "dokodemo-door",
            settings: {address: $remote.address, port: $remote.port, network: "tcp"},
            sniffing: {enabled: true, destOverride: ["tls"], routeOnly: true}
          }] + [.inbounds[] | select(.tag != "reality-sni-gate" and .tag != "reality-target-gate")
            | if .protocol == "vless" then
                .listen = ($old.listen // $v.listen // "0.0.0.0") | .port = $public
                | .streamSettings.realitySettings.dest = ("127.0.0.1:" + ($private | tostring))
                | del(.streamSettings.realitySettings.target)
              else . end])
        | .outbounds = ([.outbounds[] | select(.tag != "reality-gate-direct" and .tag != "reality-gate-block")]
            + [{tag: "reality-gate-direct", protocol: "freedom", settings: {domainStrategy: "UseIPv4"}},
               {tag: "reality-gate-block", protocol: "blackhole"}])
        | .routing.rules = ([
            {type: "field", inboundTag: ["reality-target-gate"], domain: $names, outboundTag: "reality-gate-direct"},
            {type: "field", inboundTag: ["reality-target-gate"], outboundTag: "reality-gate-block"}
          ] + [(.routing.rules // [])[] | select(
            ((.inboundTag // []) | index("reality-sni-gate")) == null and
            ((.inboundTag // []) | index("reality-target-gate")) == null)])
    ' "$input" >"$output" || return 1
    [ -s "$output" ]
}

choose_guard_port() {
    local candidate listeners
    listeners=$(ss -H -ltn) || return 1
    for candidate in {45987..46087}; do
        if ! jq -e --argjson p "$candidate" 'any(.inbounds[]; .port == $p)' "$XRAY_CONFIG" >/dev/null &&
           ! printf '%s\n' "$listeners" | awk '{print $4}' | grep -Eq ":${candidate}$"; then
            printf '%s\n' "$candidate"
            return
        fi
    done
    die "未找到可用的内部端口。"
}

sync_guard_config() {
    local path=$1 tmp
    if jq -e 'any(.inbounds[]; .tag == "reality-sni-gate" or .tag == "reality-target-gate")' "$path" >/dev/null; then
        tmp=$(mktemp "${path}.guard.XXXXXX")
        if ! guard_config "$path" "$tmp"; then
            rm -f "$tmp"
            return 1
        fi
        cat "$tmp" >"$path"
        rm -f "$tmp"
    fi
}

enable_blackhole() {
    require_root
    [ -f "$XRAY_CONFIG" ] || die "未找到配置文件：$XRAY_CONFIG"
    local tmp backup internal_port
    internal_port=$(choose_guard_port)
    tmp=$(mktemp "${XRAY_CONFIG}.guard.XXXXXX")
    if ! guard_config "$XRAY_CONFIG" "$tmp" "$internal_port" || ! test_xray_config "$tmp"; then
        rm -f "$tmp"
        die "黑洞配置校验失败，原配置未修改。"
    fi
    backup=$(mktemp "${XRAY_CONFIG}.backup.XXXXXX")
    cp -p "$XRAY_CONFIG" "$backup"
    # Preserve the existing service-readable ownership and mode.
    chown --reference="$XRAY_CONFIG" "$tmp"
    chmod --reference="$XRAY_CONFIG" "$tmp"
    mv "$tmp" "$XRAY_CONFIG"
    if ! systemctl restart xray || ! systemctl is-active --quiet xray; then
        mv "$backup" "$XRAY_CONFIG"
        systemctl restart xray || true
        die "启动失败，已恢复原配置。"
    fi
    rm -f "$backup"
    ok "SNI 黑洞防护已启用，客户端链接不变。"
    warn "目标转发仅放行配置中的 SNI；访问允许的目标仍会消耗带宽。"
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
        echo "5. 修改 DNS"
        echo "6. 重启 Xray"
        echo "7. 查看 Xray 状态"
        echo "8. 修复 Reality 配置"
        echo "9. 启用 SNI 黑洞防护"
        echo "0. 退出"
        read -r -p "请选择: " choice
        case "$choice" in
            1) show_info ;;
            2) show_qr ;;
            3) change_sni ;;
            4) change_port ;;
            5) change_dns ;;
            6) restart_service ;;
            7) status_service ;;
            8) repair_reality_config ;;
            9) enable_blackhole ;;
            0) exit 0 ;;
            *) warn "请输入 0-9。" ;;
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
  vless dns <主DNS> <备用DNS>
                         修改 Xray DNS，例如：vless dns 1.1.1.1 8.8.8.8
  vless restart        重启 Xray
  vless status         查看状态
  vless repair         转换 Reality target 字段为兼容字段 dest
  vless blackhole      启用 SNI 黑洞防护（保留客户端链接）
  vless update-manager 更新 vless 管理命令，不重装节点
  vless install        重新安装

说明:
  安装完成后直接输入 vless 打开菜单。
  修改 SNI、端口或 DNS 会先校验配置，通过后自动重启 Xray。
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
        dns|change-dns) shift; change_dns "${1:-}" "${2:-}" ;;
        restart) restart_service ;;
        status) status_service ;;
        repair) repair_reality_config ;;
        blackhole) enable_blackhole ;;
        update-manager|update) require_root; install_manager_command ;;
        help|-h|--help) usage ;;
        *) usage; exit 1 ;;
    esac
}

dispatch "$@"
