#!/usr/bin/env bash
set -Eeuo pipefail

# One-click Shadowsocks 2022 installer for Linux.
# Project: https://github.com/shadowsocks/shadowsocks-rust

readonly APP="ss2022"
readonly CONF_DIR="/etc/shadowsocks-rust"
readonly CONF_FILE="${CONF_DIR}/config.json"
readonly SERVICE_FILE="/etc/systemd/system/${APP}.service"
readonly BIN="/usr/local/bin/ssserver"
readonly MANAGER="/usr/local/bin/ss2022"
readonly DEFAULT_PORT="8388"
readonly DEFAULT_METHOD="2022-blake3-aes-256-gcm"

log() { printf '[%s] %s\n' "$APP" "$*"; }
die() { printf '[%s] ERROR: %s\n' "$APP" "$*" >&2; exit 1; }
trap 'die "安装失败，出错行：${LINENO}"' ERR

choose_bind() {
  printf '\n请选择监听地址：\n  1) IPv4\n  2) IPv6\n  3) IPv4 + IPv6（双栈）\n\n'
  read -r -p '请输入选项 [1-3，默认 1]：' choice
  case "${choice:-1}" in
    1) bind='0.0.0.0' ;;
    2) bind='::' ;;
    3) bind='dual' ;;
    *) die '无效选项' ;;
  esac
}

write_config() {
  local bind_value="$1" password
  password="${SS_PASSWORD:-$(openssl rand -base64 32 | tr -d '\n')}"
  printf '%s' "$password" | grep -Eq '^[A-Za-z0-9+/=]+$' || die '密码包含非法字符'
  if [[ "$bind_value" == dual ]]; then
    cat > "$CONF_FILE" <<EOF
{
  "servers": [
    {"server": "0.0.0.0", "server_port": ${port}, "method": "${method}", "password": "${password}", "mode": "tcp_and_udp"},
    {"server": "::", "server_port": ${port}, "method": "${method}", "password": "${password}", "mode": "tcp_and_udp"}
  ]
}
EOF
  else
    cat > "$CONF_FILE" <<EOF
{
  "server": "${bind_value}",
  "server_port": ${port},
  "method": "${method}",
  "password": "${password}",
  "mode": "tcp_and_udp",
  "fast_open": false
}
EOF
  fi
  chmod 600 "$CONF_FILE"
}

menu() {
  while true; do
    printf '\n==== Shadowsocks 2022 管理菜单 ====\n'
    printf '1) 安装 / 重新安装\n2) 查看状态\n3) 重启服务\n4) 查看日志\n5) 卸载\n0) 退出\n\n'
    read -r -p '请选择 [0-5]：' action
    case "${action:-0}" in
      1) install_server ;;
      2) systemctl status "${APP}.service" --no-pager || true ;;
      3) systemctl restart "${APP}.service" && log '服务已重启' || true ;;
      4) journalctl -u "${APP}.service" -n 80 --no-pager || true ;;
      5) uninstall_server; exit 0 ;;
      0) exit 0 ;;
      *) printf '无效选项\n' ;;
    esac
  done
}

uninstall_server() {
  systemctl disable --now "${APP}.service" 2>/dev/null || true
  rm -f "$SERVICE_FILE" "$BIN" "$MANAGER"
  rm -rf "$CONF_DIR"
  systemctl daemon-reload
  log "已卸载 ${APP}（不会修改防火墙规则）"
}

[[ "${EUID}" -eq 0 ]] || die "请使用 root 运行：sudo bash ss2022.sh"
command -v systemctl >/dev/null || die "此脚本需要 systemd"

if [[ "${1:-}" == "uninstall" ]]; then uninstall_server; exit 0; fi

detect_arch() {
  case "$(uname -m)" in
    x86_64|amd64) echo "x86_64-unknown-linux-gnu" ;;
    aarch64|arm64) echo "aarch64-unknown-linux-gnu" ;;
    *) die "不支持的架构：$(uname -m)，目前支持 x86_64 和 arm64" ;;
  esac
}

install_tools() {
  local id=""
  [[ -r /etc/os-release ]] && . /etc/os-release && id="${ID:-}"
  case "$id" in
    debian|ubuntu|linuxmint)
      apt-get update
      DEBIAN_FRONTEND=noninteractive apt-get install -y curl ca-certificates openssl
      ;;
    rocky|almalinux|centos|rhel|fedora)
      (command -v dnf >/dev/null && dnf install -y curl ca-certificates openssl) || yum install -y curl ca-certificates openssl
      ;;
    *)
      command -v curl >/dev/null || die "请先安装 curl、ca-certificates 和 openssl"
      command -v openssl >/dev/null || die "请先安装 openssl"
      ;;
  esac
}

version="${SS_VERSION:-1.25.0}"
port="${SS_PORT:-$DEFAULT_PORT}"
method="${SS_METHOD:-$DEFAULT_METHOD}"
[[ "$port" =~ ^[0-9]+$ && "$port" -ge 1 && "$port" -le 65535 ]] || die "SS_PORT 必须是 1-65535 的端口"
[[ "$method" == "$DEFAULT_METHOD" ]] || die "目前只允许使用 $DEFAULT_METHOD"

install_server() {
install_tools
choose_bind
target="$(detect_arch)"
tmp="$(mktemp -d)"
trap 'rm -rf "$tmp"' EXIT
archive="shadowsocks-v${version}.${target}.tar.xz"
url="https://github.com/shadowsocks/shadowsocks-rust/releases/download/v${version}/${archive}"

log "下载 shadowsocks-rust v${version}（${target}）"
curl --fail --location --retry 3 --proto '=https' --tlsv1.2 -o "${tmp}/${archive}" "$url"
tar -xJf "${tmp}/${archive}" -C "$tmp"
found="$(find "$tmp" -type f -name ssserver -perm -u+x -print -quit)"
[[ -n "$found" ]] || die "压缩包中找不到 ssserver"
install -m 0755 "$found" "$BIN"

mkdir -p "$CONF_DIR"
if [[ -s "$CONF_FILE" && "${SS_FORCE:-0}" != 1 ]]; then
  log "保留已有配置：$CONF_FILE（如需重建请设置 SS_FORCE=1）"
else
  write_config "$bind"
fi

cat > "$SERVICE_FILE" <<EOF
[Unit]
Description=Shadowsocks Rust SS2022 Server
After=network-online.target
Wants=network-online.target

[Service]
ExecStart=${BIN} -c ${CONF_FILE}
Restart=on-failure
RestartSec=3
User=nobody
NoNewPrivileges=true
PrivateTmp=true
ProtectSystem=strict
ProtectHome=true
ReadWritePaths=${CONF_DIR}
LimitNOFILE=1048576

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable --now "${APP}.service"
systemctl is-active --quiet "${APP}.service" || { systemctl status "${APP}.service" --no-pager; exit 1; }

install -m 0755 "$0" "$MANAGER"

ip="$(curl -4fsS --max-time 5 https://api.ipify.org 2>/dev/null || true)"
log "安装完成，监听：${port} / ${method}"
log "配置文件：${CONF_FILE}"
log "服务管理：systemctl status ${APP}; journalctl -u ${APP} -e"
printf '\n客户端参数（请妥善保管）：\n  server: %s\n  port: %s\n  method: %s\n  password: %s\n' "${ip:-你的服务器IP}" "$port" "$method" "$(sed -n 's/.*"password": "\([^"]*\)".*/\1/p' "$CONF_FILE")"
}

if [[ "$(basename "$0")" == "ss2022" ]]; then
  menu
else
  install_server
fi
