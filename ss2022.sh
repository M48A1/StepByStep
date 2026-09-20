#!/usr/bin/env bash
# Version: 1.3.3 | Date: 2026-09-20
set -Eeuo pipefail

# One-click Shadowsocks 2022 installer for Linux.
# Project: https://github.com/shadowsocks/shadowsocks-rust

readonly APP="ss2022"
readonly SCRIPT_VERSION="1.3.3"
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

# ssserver 的 black_list 按客户端来源地址过滤 TCP/UDP。
# 地址库：https://github.com/gaoyifan/china-operator-ip
# 不修改系统防火墙；经境外中转的连接只能识别到中转 IP。
update_cn_acl() (
  set -Eeuo pipefail
  [[ -s "$CONF_FILE" ]] || die '请先安装 SS2022'
  command -v python3 >/dev/null || die '请先安装 python3，或重新运行安装器'
  local acl_dir work base
  acl_dir="$(dirname "$CONF_FILE")"
  work="$(mktemp -d "${acl_dir}/.cn-update.XXXXXX")"
  trap 'rm -rf "$work"' EXIT
  base='https://raw.githubusercontent.com/gaoyifan/china-operator-ip/ip-lists'
  log '下载中国大陆 IPv4 / IPv6 网段，失败时保留原有规则'
  curl -fsSL --proto '=https' --tlsv1.2 --connect-timeout 15 --max-time 120 --retry 3 "${base}/china.txt" -o "${work}/cn4"
  curl -fsSL --proto '=https' --tlsv1.2 --connect-timeout 15 --max-time 120 --retry 3 "${base}/china6.txt" -o "${work}/cn6"
  python3 - "$CONF_FILE" "$work" <<'SS2022_CN_PY'
import ipaddress
import json
import os
import pathlib
import sys

conf = pathlib.Path(sys.argv[1])
work = pathlib.Path(sys.argv[2])
acl = conf.parent / 'cn-block.acl'
config = json.loads(conf.read_text())
# 避免悄悄覆盖用户自己设置的其他 ACL。
for item in [config] + config.get('servers', []):
    if item.get('acl') and item['acl'] != str(acl):
        sys.exit('检测到自定义 ACL，请先手动合并规则：' + item['acl'])
networks = []
for version in (4, 6):
    entries = set()
    for line in (work / ('cn' + str(version))).read_text().splitlines():
        line = line.split('#', 1)[0].strip()
        if not line:
            continue
        net = ipaddress.ip_network(line, strict=True)
        if net.version != version or net.prefixlen == 0:
            sys.exit('网段类型错误或包含默认路由，拒绝更新')
        entries.add(net)
    if len(entries) < 100:
        sys.exit('网段列表过少，拒绝更新，保留旧规则')
    networks.extend(sorted(entries, key=lambda n: (int(n.network_address), n.prefixlen)))
    print('IPv%d：%d 个网段' % (version, len(entries)))
config['acl'] = str(acl)
for server in config.get('servers', []):
    server['acl'] = str(acl)
(work / 'acl').write_text('[accept_all]\n[black_list]\n' + '\n'.join(map(str, networks)) + '\n')
(work / 'config').write_text(json.dumps(config, indent=2) + '\n')
os.chmod(work / 'acl', 0o600)
os.chmod(work / 'config', 0o600)
# 同一文件系统原子替换，不让服务读到半份列表。
os.replace(work / 'acl', acl)
os.replace(work / 'config', conf)
SS2022_CN_PY
  if [[ "${1:-}" != no-restart ]]; then
    systemctl restart ss2022.service
    systemctl is-active --quiet ss2022.service || die '规则已写入，但服务启动失败，请检查 ss2022 logs'
    log '中国大陆来源 IP 屏蔽已生效（IPv4 / IPv6，TCP / UDP）'
  else
    log '中国大陆来源 IP 规则已写入，服务启动后生效'
  fi
)

get_node_name() {
  local name_file="${CONF_FILE%/*}/node-name"
  if [[ -s "$name_file" ]]; then
    cat "$name_file"
  else
    printf '%s' 'SS2022'
  fi
}

encode_node_name() {
  # 按 UTF-8 字节编码，兼容中文、空格、#、& 等名称字符。
  local LC_ALL=C value="$1" char i byte
  for ((i=0; i<${#value}; i++)); do
    char="${value:i:1}"
    case "$char" in
      [a-zA-Z0-9.~_-]) printf '%s' "$char" ;;
      *) printf -v byte '%d' "'$char"; printf '%%%02X' "$((byte & 255))" ;;
    esac
  done
}

install_manager_command() {
  {
  printf '#!/usr/bin/env bash\n'
  declare -f update_cn_acl
  declare -f get_node_name encode_node_name show_node
  cat <<'SS2022_MANAGER_EOF'
#!/usr/bin/env bash
set -Eeuo pipefail
readonly VERSION="1.3.3"
readonly CONF_FILE="/etc/shadowsocks-rust/config.json"
log() { printf '[ss2022] %s\n' "$*"; }
die() { printf '[ss2022] ERROR: %s\n' "$*" >&2; exit 1; }


show_config() {
  [[ -f "$CONF_FILE" ]] || die "配置文件不存在：$CONF_FILE"
  sed 's/"password"[[:space:]]*:[[:space:]]*"[^"]*"/"password": "********"/g' "$CONF_FILE"
}

uninstall_server() {
  read -r -p '输入 DELETE 确认完整卸载：' confirm
  [[ "$confirm" == DELETE ]] || { log '已取消'; return; }
  systemctl disable --now ss2022.service 2>/dev/null || true
  rm -f /etc/systemd/system/ss2022.service /usr/local/bin/ssserver
  rm -rf /etc/shadowsocks-rust
  systemctl daemon-reload
  log '服务、程序和配置已删除'
}

menu() {
  while true; do
    printf '\n==== Shadowsocks 2022 管理菜单 v%s ====\n' "$VERSION"
    printf '1) 查看节点配置 / 二维码\n2) 查看服务状态\n3) 重启服务\n4) 查看日志\n5) 查看配置文件（隐藏密码）\n6) 卸载\n7) 启用 / 更新中国大陆来源 IP 屏蔽（重启服务）\n0) 退出\n\n'
    read -r -p '请选择 [0-7]：' choice
    case "${choice:-0}" in
      1) show_node ;; 2) systemctl status ss2022.service --no-pager || true ;;
      3) systemctl restart ss2022.service && log '服务已重启' ;;
      4) journalctl -u ss2022.service -n 80 --no-pager ;;
      5) show_config ;; 6) uninstall_server ;; 7) update_cn_acl ;; 0) exit 0 ;; *) log '无效选项' ;;
    esac
  done
}

if [[ "${EUID}" -ne 0 ]]; then
  command -v sudo >/dev/null 2>&1 || die '请使用 root 运行'
  exec sudo -- "$0" "$@"
fi
case "${1:-}" in
  ""|menu) menu ;; show|qr) show_node ;; config) show_config ;;
  status) systemctl status ss2022.service --no-pager ;;
  restart) systemctl restart ss2022.service && log '服务已重启' ;;
  logs) journalctl -u ss2022.service -n 80 --no-pager ;;
  block-cn|update-cn) update_cn_acl ;;
  uninstall) uninstall_server ;; version|-v|--version) printf 'ss2022 %s\n' "$VERSION" ;;
  *) printf '用法：ss2022 [menu|show|config|status|restart|logs|block-cn|update-cn|uninstall|version]\n'; exit 1 ;;
esac
SS2022_MANAGER_EOF
  } > "$MANAGER"
  chmod 0755 "$MANAGER"
  [[ -x "$MANAGER" ]] || die "管理命令安装失败：$MANAGER"
  log "管理命令已安装：ss2022"
}

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

choose_node_name() {
  local default_name
  default_name="$(get_node_name)"
  IFS= read -r -p "请输入节点名称 [默认 ${default_name}]：" node_name
  node_name="${node_name:-$default_name}"
}

write_config() {
  local bind_value="$1" password
  password="${SS_PASSWORD:-$(openssl rand -base64 32 | tr -d '\n')}"
  printf '%s' "$password" | grep -Eq '^[A-Za-z0-9+/=]+$' || die '密码包含非法字符'
  if [[ "$bind_value" == dual ]]; then
    cat > "$CONF_FILE" <<EOF
{
  "ipv6_only": true,
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
  "fast_open": false,
  "ipv6_only": true
}
EOF
  fi
  chmod 600 "$CONF_FILE"
}

menu() {
  while true; do
    printf '\n==== Shadowsocks 2022 管理菜单 v%s ====\n' "$SCRIPT_VERSION"
    printf '1) 安装 / 重新安装\n2) 查看状态\n3) 重启服务\n4) 查看日志\n5) 显示节点配置 / 二维码\n6) 检查配置文件 / 删除配置\n7) 卸载\n8) 启用 / 更新中国大陆来源 IP 屏蔽（重启服务）\n0) 退出\n\n'
    read -r -p '请选择 [0-8]：' action
    case "${action:-0}" in
      1) install_server ;;
      2) systemctl status "${APP}.service" --no-pager || true ;;
      3) systemctl restart "${APP}.service" && log '服务已重启' || true ;;
      4) journalctl -u "${APP}.service" -n 80 --no-pager || true ;;
      5) show_node ;;
      6) inspect_config ;;
      7) uninstall_server; exit 0 ;;
      8) install_tools; update_cn_acl ;;
      0) exit 0 ;;
      *) printf '无效选项\n' ;;
    esac
  done
}

usage() {
  cat <<EOF
用法：
  ss2022                 打开管理菜单
  ss2022 show            显示节点配置和二维码
  ss2022 status          查看服务状态
  ss2022 restart         重启服务
  ss2022 logs            查看最近日志
  ss2022 config          检查或删除配置
  ss2022 install         安装或重新安装
  ss2022 uninstall       完整卸载
  ss2022 block-cn        启用中国大陆来源 IP 屏蔽并重启服务
  ss2022 update-cn       更新 IPv4 / IPv6 网段并重启服务
  bash ss2022.sh update-manager  仅更新管理命令，不重装或重启服务
  ss2022 version         显示脚本版本

也可以使用 bash ss2022.sh 进行首次安装，默认屏蔽中国大陆来源 IP。
已有安装可执行 bash ss2022.sh block-cn，无需重装或更换密码。
网段不自动更新，请定期执行 ss2022 update-cn。
EOF
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
      # 立即恢复新版管理命令，后续下载或启动失败时仍可进入菜单排障。
      install_manager_command "$installer_source"
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
  [[ -s "$CONF_FILE" ]] || die '尚未安装或配置文件不存在'
  command -v python3 >/dev/null || die '读取节点配置需要 python3'
  local fields bind_value port_value method_value password_value host_value uri_host userinfo uri node_name
  # JSON 解析兼容单行、缩进、单栈和双栈配置，不用 sed 猜字段。
  fields="$(python3 - "$CONF_FILE" <<'SS2022_NODE_PY'
import json
import sys
config = json.load(open(sys.argv[1]))
server = config['servers'][0] if config.get('servers') else config
for key in ('server', 'server_port', 'method', 'password'):
    value = str(server[key])
    if not value or '\n' in value or '\r' in value:
        sys.exit('配置字段为空或包含换行：' + key)
    print(value)
SS2022_NODE_PY
)" || die '无法解析节点配置'
  {
    IFS= read -r bind_value
    IFS= read -r port_value
    IFS= read -r method_value
    IFS= read -r password_value
  } <<< "$fields"
  host_value="${SS_HOST:-}"
  if [[ -z "$host_value" ]]; then
    case "$bind_value" in
      0.0.0.0) host_value="$(curl -4fsS --max-time 5 https://api.ipify.org 2>/dev/null || true)" ;;
      ::) host_value="$(curl -6fsS --max-time 5 https://api6.ipify.org 2>/dev/null || true)" ;;
      *) host_value="$bind_value" ;;
    esac
  fi
  [[ -n "$host_value" ]] || die '无法获取公网地址，请指定：sudo env SS_HOST=服务器IP或域名 bash ss2022.sh show'
  uri_host="$host_value"
  if [[ "$uri_host" == *:* && "$uri_host" != \[*\] ]]; then
    uri_host="[${uri_host}]"
  fi
  case "$method_value" in
    2022-*) userinfo="$(encode_node_name "$method_value"):$(encode_node_name "$password_value")" ;;
    *) userinfo="$(printf '%s' "${method_value}:${password_value}" | base64 | tr '+/' '-_' | tr -d '=\n')" ;;
  esac
  node_name="$(get_node_name)"
  uri="ss://${userinfo}@${uri_host}:${port_value}#$(encode_node_name "$node_name")"
  printf '\n===== SS2022 节点配置 =====\n'
  printf '节点名称：%s\n' "$node_name"
  printf '服务器：%s\n端口：%s\n加密：%s\n密码：%s\n\n节点链接：\n%s\n' "$host_value" "$port_value" "$method_value" "$password_value" "$uri"
  if command -v qrencode >/dev/null 2>&1; then
    printf '\n二维码：\n'
    qrencode -t ANSIUTF8 "$uri"
  else
    log '未安装 qrencode，暂时无法输出二维码'
  fi
}

uninstall_server() {
  systemctl disable --now "${APP}.service" 2>/dev/null || true
  rm -f "$SERVICE_FILE" "$BIN" "$MANAGER"
  rm -rf "$CONF_DIR"
  systemctl daemon-reload
  log "已卸载 ${APP}（不会修改防火墙规则）"
}

if [[ "${EUID}" -ne 0 ]]; then
  if [[ "$(basename "$0")" == "ss2022" ]]; then
    command -v sudo >/dev/null 2>&1 || die "需要 root 权限，且系统未安装 sudo"
    exec sudo -- "$0" "$@"
  fi
  die "首次安装请使用 root 权限运行"
fi
command -v systemctl >/dev/null || die "此脚本需要 systemd"

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
      DEBIAN_FRONTEND=noninteractive apt-get install -y curl ca-certificates openssl python3 qrencode tar xz-utils
      ;;
    rocky|almalinux|centos|rhel|fedora)
      (command -v dnf >/dev/null && dnf install -y curl ca-certificates openssl python3 tar xz) || yum install -y curl ca-certificates openssl python3 tar xz
      if command -v dnf >/dev/null; then
        dnf install -y qrencode || log 'qrencode 安装失败，将只输出节点链接'
      else
        yum install -y qrencode || log 'qrencode 安装失败，将只输出节点链接'
      fi
      ;;
    *)
      command -v curl >/dev/null || die "请先安装 curl、ca-certificates 和 openssl"
      command -v openssl >/dev/null || die "请先安装 openssl"
      command -v python3 >/dev/null || die "请先安装 python3"
      ;;
  esac
  command -v tar >/dev/null || die '缺少 tar，请安装后重试'
  command -v xz >/dev/null || die '缺少 xz：Debian/Ubuntu 请安装 xz-utils，RHEL/Fedora 请安装 xz'
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
install -m 0755 "${BASH_SOURCE[0]:-$0}" "$installer_source"
# 安装流程一开始就落地管理命令，后续任一步失败仍可使用 ss2022 排障。
install_manager_command "$installer_source"
preflight_check || return 0
install_tools
choose_bind
choose_node_name
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
install_manager_command "$installer_source"

mkdir -p "$CONF_DIR"
if [[ -s "$CONF_FILE" && "${SS_FORCE:-0}" != 1 ]]; then
  log "保留已有配置：$CONF_FILE（如需重建请设置 SS_FORCE=1）"
else
  write_config "$bind"
fi
printf '%s\n' "$node_name" > "${CONF_DIR}/node-name"
chmod 600 "${CONF_DIR}/node-name"
# 安装器和服务均以 root 运行，配置仅允许 root 读取。
chown root:root "$CONF_FILE"
chmod 600 "$CONF_FILE"
update_cn_acl no-restart

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
systemctl enable "${APP}.service"
systemctl restart "${APP}.service"
systemctl is-active --quiet "${APP}.service" || { systemctl status "${APP}.service" --no-pager; exit 1; }

log "安装完成，配置文件：${CONF_FILE}"
log "服务管理：systemctl status ${APP}; journalctl -u ${APP} -e"
log "以后直接输入 ss2022 打开管理菜单"

show_node
}

dispatch() {
  local command_name="${1:-}"
  case "$command_name" in
    "")
      if [[ "$(basename "$0")" == "ss2022" ]]; then
        menu
      else
        install_server
      fi
      ;;
    menu) menu ;;
    show|qr) show_node ;;
    status) systemctl status "${APP}.service" --no-pager ;;
    restart) systemctl restart "${APP}.service" && log '服务已重启' ;;
    logs) journalctl -u "${APP}.service" -n 80 --no-pager ;;
    config) inspect_config ;;
    install) install_server ;;
    update-manager) install_manager_command ;;
    block-cn|update-cn)
      install_tools
      update_cn_acl
      install_manager_command
      ;;
    uninstall) uninstall_server ;;
    version|-v|--version) printf '%s %s\n' "$APP" "$SCRIPT_VERSION" ;;
    help|-h|--help) usage ;;
    *) usage; return 1 ;;
  esac
}

dispatch "$@"
