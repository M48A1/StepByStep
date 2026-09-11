#!/usr/bin/env bash
# Version: 1.0.2 | Date: 2026-09-11
set -Eeuo pipefail

# One-click Shadowsocks 2022 installer for Linux.
# Project: https://github.com/shadowsocks/shadowsocks-rust

readonly APP="ss2022"
readonly SCRIPT_VERSION="1.0.2"
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
    printf '\n==== Shadowsocks 2022 管理菜单 v%s ====\n' "$SCRIPT_VERSION"
    printf '1) 安装 / 重新安装\n2) 查看状态\n3) 重启服务\n4) 查看日志\n5) 显示节点配置 / 二维码\n6) 检查配置文件 / 删除配置\n7) 卸载\n0) 退出\n\n'
    read -r -p '请选择 [0-7]：' action
    case "${action:-0}" in
      1) install_server ;;
      2) systemctl status "${APP}.service" --no-pager || true ;;
      3) systemctl restart "${APP}.service" && log '服务已重启' || true ;;
      4) journalctl -u "${APP}.service" -n 80 --no-pager || true ;;
      5) show_node ;;
      6) inspect_config ;;
      7) uninstall_server; exit 0 ;;
      0) exit 0 ;;
      *) printf '无效选项\n' ;;
    esac
  done
}

inspect_config() {
  if [[ ! -f "$CONF_FILE" ]]; then
    log "配置文件不存在：$CONF_FILE"
    return 0
  fi
  printf '\n===== 当前配置文件（密码已隐藏） =====\n'
  sed 's/"password"[[:space:]]*:[[:space:]]*"[^"]*"/"password": "********"/g' "$CONF_FILE"
  printf '\n配置文件路径：%s\n' "$CONF_FILE"
  read -r -p '是否删除当前配置文件？输入 DELETE 确认：' confirm
  if [[ "$confirm" == "DELETE" ]]; then
    systemctl disable --now "${APP}.service" 2>/dev/null || true
    rm -f "$CONF_FILE"
    log "配置文件已删除，服务已停止"
  else
    log "已取消删除"
  fi
}

preflight_check() {
  if [[ -e "$CONF_FILE" ]]; then
    printf '\n检测到已有配置文件：%s\n' "$CONF_FILE"
    printf '当前配置（密码已隐藏）：\n'
    sed 's/"password"[[:space:]]*:[[:space:]]*"[^"]*"/"password": "********"/g' "$CONF_FILE"
    read -r -p '是否删除整个旧安装后重新安装？[y/N]：' delete_old
    if [[ "$delete_old" =~ ^[Yy]$ ]]; then
      systemctl disable --now "${APP}.service" 2>/dev/null || true
      rm -f "$SERVICE_FILE" "$BIN" "$MANAGER"
      rm -rf "$CONF_DIR"
      systemctl daemon-reload
      log '旧服务、程序和配置已全部删除，将重新安装'
    else
      log '保留旧配置，安装时不会覆盖密码和端口'
    fi
  elif [[ -e "$SERVICE_FILE" || -e "$BIN" ]]; then
    printf '\n检测到已有 SS2022 程序或服务，将继续安装并启动服务。\n'
  fi
  return 0
}

show_node() {
  [[ -s "$CONF_FILE" ]] || { log '尚未安装或配置文件不存在'; return 0; }
  local port_value method_value password_value host_value encoded uri
  port_value="$(sed -n 's/.*"server_port":[[:space:]]*\([0-9]*\).*/\1/p' "$CONF_FILE" | head -1)"
  method_value="$(sed -n 's/.*"method":[[:space:]]*"\([^"]*\)".*/\1/p' "$CONF_FILE" | head -1)"
  password_value="$(sed -n 's/.*"password":[[:space:]]*"\([^"]*\)".*/\1/p' "$CONF_FILE" | head -1)"
  host_value="$(curl -4fsS --max-time 5 https://api.ipify.org 2>/dev/null || true)"
  [[ -n "$host_value" ]] || host_value='请替换为服务器 IP 或域名'
  encoded="$(printf '%s' "${method_value}:${password_value}" | base64 | tr '+/' '-_' | tr -d '=\n')"
  uri="ss://${encoded}@${host_value}:${port_value}"
  printf '\n===== SS2022 节点配置 =====\n'
  printf '服务器：%s\n端口：%s\n加密：%s\n密码：%s\n\n节点链接：\n%s\n' "$host_value" "$port_value" "$method_value" "$password_value" "$uri"
  printf '\n原始配置文件：\n'
  sed 's/"password"[[:space:]]*:[[:space:]]*"[^"]*"/"password": "********"/g' "$CONF_FILE"
  if command -v qrencode >/dev/null 2>&1; then
    printf '\n二维码：\n'
    qrencode -t ANSIUTF8 "$uri"
  else
    printf '\n未检测到 qrencode，无法显示终端二维码。Debian/Ubuntu 可执行：\n  sudo apt-get install -y qrencode\n'
  fi
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
    # 使用静态链接的 musl 构建，避免旧版 Debian/Ubuntu 的 glibc
    # 无法运行上游 GNU 构建。
    x86_64|amd64) echo "x86_64-unknown-linux-musl" ;;
    aarch64|arm64) echo "aarch64-unknown-linux-musl" ;;
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
tmp="$(mktemp -d)"
trap 'rm -rf "$tmp"' EXIT
# 删除旧安装前保存当前安装器；脚本可能正从 $MANAGER 运行。
installer_source="${tmp}/ss2022"
install -m 0755 "$0" "$installer_source"
preflight_check || return 0
install_tools
choose_bind
target="$(detect_arch)"
archive="shadowsocks-v${version}.${target}.tar.xz"
url="https://github.com/shadowsocks/shadowsocks-rust/releases/download/v${version}/${archive}"

log "下载 shadowsocks-rust v${version}（${target}）"
curl --fail --location --retry 3 --proto '=https' --tlsv1.2 -o "${tmp}/${archive}" "$url"
tar -xJf "${tmp}/${archive}" -C "$tmp"
found="$(find "$tmp" -type f -name ssserver -perm -u+x -print -quit)"
[[ -n "$found" ]] || die "压缩包中找不到 ssserver"
"$found" --version >/dev/null || die "下载的 ssserver 无法在当前系统运行"
install -m 0755 "$found" "$BIN"
# 提前安装管理命令，即使服务启动失败，也能通过 ss2022 查看状态和日志。
install -m 0755 "$installer_source" "$MANAGER"

mkdir -p "$CONF_DIR"
if [[ -s "$CONF_FILE" && "${SS_FORCE:-0}" != 1 ]]; then
  log "保留已有配置：$CONF_FILE（如需重建请设置 SS_FORCE=1）"
else
  write_config "$bind"
fi
# 安装器和服务均以 root 运行，配置仅允许 root 读取。
chown root:root "$CONF_FILE"
chmod 600 "$CONF_FILE"

cat > "$SERVICE_FILE" <<EOF
[Unit]
Description=Shadowsocks Rust SS2022 Server
After=network-online.target
Wants=network-online.target

[Service]
ExecStart=${BIN} -c ${CONF_FILE}
Restart=on-failure
RestartSec=3
User=root
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
