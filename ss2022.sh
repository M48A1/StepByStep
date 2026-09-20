#!/usr/bin/env bash
# Version: 1.4.5 | Date: 2026-09-20
set -Eeuo pipefail

# One-click Shadowsocks 2022 installer for Linux.
# Project: https://github.com/shadowsocks/shadowsocks-rust

readonly APP="ss2022"
readonly SCRIPT_VERSION="1.4.5"
readonly CONF_DIR="/etc/shadowsocks-rust"
readonly CONF_FILE="${CONF_DIR}/config.json"
readonly SERVICE_FILE="/etc/systemd/system/${APP}.service"
readonly BIN="/usr/local/bin/ssserver"
readonly MANAGER="/usr/local/bin/ss2022"
readonly DEFAULT_PORT="8388"
readonly DEFAULT_METHOD="2022-blake3-aes-256-gcm"

log() { printf '[%s] %s\n' "$APP" "$*"; }
die() { printf '[%s] ERROR: %s\n' "$APP" "$*" >&2; exit 1; }

# ssserver 的 black_list 按客户端来源地址过滤 TCP/UDP。
# 地址库：https://github.com/gaoyifan/china-operator-ip
# 不修改系统防火墙；经境外中转的连接只能识别到中转 IP。
prepare_cn_acl() {
  local input="$1" prepared="$2" base
  base='https://raw.githubusercontent.com/gaoyifan/china-operator-ip/ip-lists'
  log '下载中国大陆 IPv4 / IPv6 网段，失败时保留原有规则'
  curl -fsSL --proto '=https' --tlsv1.2 --connect-timeout 15 --max-time 120 --retry 3 "${base}/china.txt" -o "${prepared}/cn4"
  curl -fsSL --proto '=https' --tlsv1.2 --connect-timeout 15 --max-time 120 --retry 3 "${base}/china6.txt" -o "${prepared}/cn6"
  python3 - "$input" "$prepared" "${CONF_DIR}/cn-block.acl" <<'SS2022_CN_PY'
import ipaddress
import json
import os
import pathlib
import sys

conf = pathlib.Path(sys.argv[1])
work = pathlib.Path(sys.argv[2])
acl = pathlib.Path(sys.argv[3])
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
servers = config.get('servers', [])
v4_ports = {x.get('server_port') for x in servers if x.get('server') == '0.0.0.0'}
v6_ports = {x.get('server_port') for x in servers if x.get('server') == '::'}
if v4_ports & v6_ports:
    config['ipv6_only'] = True
config['acl'] = str(acl)
for server in config.get('servers', []):
    server['acl'] = str(acl)
(work / 'acl').write_text('[accept_all]\n[black_list]\n' + '\n'.join(map(str, networks)) + '\n')
(work / 'config').write_text(json.dumps(config, indent=2) + '\n')
os.chmod(work / 'acl', 0o600)
os.chmod(work / 'config', 0o600)
SS2022_CN_PY
}

update_cn_acl() (
  set -Eeuo pipefail
  [[ -s "$CONF_FILE" ]] || die '请先安装 SS2022'
  command -v python3 >/dev/null || die '请先安装 python3'
  work='' changed=0 committed=0 was_active=0 was_enabled=0
  change_paths=()
  work="$(mktemp -d)"
  trap 'finish_changes "$?"' EXIT
  trap 'exit 130' INT
  trap 'exit 143' TERM
  mkdir "${work}/prepared"
  validate_config "$CONF_FILE"
  prepare_cn_acl "$CONF_FILE" "${work}/prepared"
  backup_changes "$CONF_FILE" "${CONF_DIR}/cn-block.acl"
  changed=1
  atomic_install "${work}/prepared/acl" "${CONF_DIR}/cn-block.acl" 0600
  atomic_install "${work}/prepared/config" "$CONF_FILE" 0600
  restart_and_check
  committed=1
  log '中国大陆来源 IP 屏蔽已生效（IPv4 / IPv6，TCP / UDP）'
)

unblock_cn() (
  set -Eeuo pipefail
  [[ -s "$CONF_FILE" ]] || die '请先安装 SS2022'
  command -v python3 >/dev/null || die '请先安装 python3'
  work='' changed=0 committed=0 was_active=0 was_enabled=0
  change_paths=()
  work="$(mktemp -d)"
  trap 'finish_changes "$?"' EXIT
  trap 'exit 130' INT
  trap 'exit 143' TERM
  validate_config "$CONF_FILE"
  python3 - "$CONF_FILE" "${work}/config" "${CONF_DIR}/cn-block.acl" <<'SS2022_UNBLOCK_PY'
import json
import os
import pathlib
import sys

config = json.loads(pathlib.Path(sys.argv[1]).read_text())
for item in [config] + config.get('servers', []):
    if item.get('acl') == sys.argv[3]:
        del item['acl']
    elif item.get('acl'):
        print('保留自定义 ACL：' + item['acl'])
output = pathlib.Path(sys.argv[2])
output.write_text(json.dumps(config, indent=2) + '\n')
os.chmod(output, 0o600)
SS2022_UNBLOCK_PY
  backup_changes "$CONF_FILE" "${CONF_DIR}/cn-block.acl"
  changed=1
  atomic_install "${work}/config" "$CONF_FILE" 0600
  rm -f "${CONF_DIR}/cn-block.acl"
  restart_and_check
  committed=1
  log '已解除本脚本的中国大陆来源 IP 屏蔽；节点地址、端口和密码保持不变'
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

install_manager_command() (
  staged='' key=''
  staged="$(mktemp "${MANAGER}.XXXXXX")"
  trap 'rm -f "$staged"' EXIT
  {
    printf '#!/usr/bin/env bash\nset -Eeuo pipefail\n'
    for key in APP SCRIPT_VERSION CONF_DIR CONF_FILE SERVICE_FILE BIN MANAGER DEFAULT_PORT DEFAULT_METHOD; do
      printf 'readonly %s=%q\n' "$key" "${!key}"
    done
    # 从内存里的函数生成完整管理命令，支持 bash <(curl ...) 和 curl | bash。
    declare -f log die prepare_cn_acl update_cn_acl unblock_cn get_node_name encode_node_name install_manager_command prompt choose_node_name write_config menu usage inspect_config preflight_check show_config show_node uninstall_server detect_arch install_tools install_server dispatch validate_config atomic_install backup_changes finish_changes restart_and_check main
    printf '\nif [[ "${BASH_SOURCE[0]:-$0}" == "$0" ]]; then main "$@"; fi\n'
  } > "$staged"
  bash -n "$staged"
  chmod 0755 "$staged"
  mv -f "$staged" "$MANAGER"
  log "管理命令已更新：ss2022 v${SCRIPT_VERSION}"
)

prompt() {
  # 管道安装时 stdin 是脚本源码，交互必须从终端读取。
  if [[ -t 0 ]]; then
    IFS= read -r -p "$1" "$2" || die '输入已结束，操作取消'
  elif [[ -r /dev/tty ]] && { true </dev/tty; } 2>/dev/null; then
    IFS= read -r -p "$1" "$2" </dev/tty || die '输入已结束，操作取消'
  else
    die '此操作需要交互终端。请先将脚本保存为 ss2022.sh，再运行 sudo bash ss2022.sh'
  fi
}

choose_node_name() {
  local default_name
  default_name="$(get_node_name)"
  prompt "请输入节点名称 [默认 ${default_name}]：" node_name
  node_name="${node_name:-$default_name}"
}

write_config() {
  local output_file="${1:-$CONF_FILE}" password
  password="${SS_PASSWORD:-$(openssl rand -base64 32 | tr -d '\n')}"
  printf '%s' "$password" | grep -Eq '^[A-Za-z0-9+/=]+$' || die '密码包含非法字符'
  cat > "$output_file" <<EOF
{
  "server": "0.0.0.0",
  "server_port": ${port},
  "method": "${method}",
  "password": "${password}",
  "mode": "tcp_and_udp",
  "fast_open": true
}
EOF
  chmod 600 "$output_file"
}

menu() {
  while true; do
    printf '\n==== Shadowsocks 2022 管理菜单 v%s ====\n' "$SCRIPT_VERSION"
    printf '1) 安装 / 重新安装\n2) 查看状态\n3) 重启服务\n4) 查看日志\n5) 显示节点配置 / 二维码\n6) 检查配置文件 / 删除配置\n7) 卸载\n8) 启用 / 更新中国大陆来源 IP 屏蔽（默认关闭，重启服务）\n9) 解除中国大陆来源 IP 屏蔽（重启服务）\n0) 退出\n\n'
    prompt '请选择 [0-9]：' action
    case "${action:-0}" in
      1) install_server ;;
      2) systemctl status "${APP}.service" --no-pager || true ;;
      3) systemctl restart "${APP}.service" && log '服务已重启' || true ;;
      4) journalctl -u "${APP}.service" -n 80 --no-pager || true ;;
      5) show_node ;;
      6) inspect_config ;;
      7) uninstall_server; exit 0 ;;
      8) install_tools; update_cn_acl ;;
      9) unblock_cn ;;
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
  ss2022 unblock-cn      解除本脚本的大陆来源 IP 屏蔽并重启服务
  bash ss2022.sh update-manager  仅更新管理命令，不重装或重启服务
  ss2022 version         显示脚本版本

每次执行 bash ss2022.sh 或 ss2022 install，都会自动清理旧安装并重新生成配置。
安装默认不屏蔽中国大陆来源 IP。
旧版安装可执行 bash ss2022.sh unblock-cn 解除屏蔽，无需重装或更换密码。
仅在需要屏蔽时执行 ss2022 block-cn；启用后网段需手动执行 ss2022 update-cn 更新。
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
  prompt '是否删除当前配置文件？输入 DELETE 确认：' confirm
  if [[ "$confirm" == "DELETE" ]]; then
    systemctl disable --now "${APP}.service" 2>/dev/null || true
    rm -f "$CONF_FILE"
    log "配置文件已删除，服务已停止"
  else
    log "已取消删除"
  fi
}

preflight_check() {
  if [[ -e "$CONF_DIR" || -e "$BIN" || -e "$SERVICE_FILE" || -e "$MANAGER" ]]; then
    log '检测到旧安装，将自动停止服务并清理全部旧配置、程序和服务文件'
    systemctl disable --now "${APP}.service" 2>/dev/null || true
    rm -f "$SERVICE_FILE" "$BIN" "$MANAGER"
    rm -rf "$CONF_DIR"
    systemctl daemon-reload
    log '旧安装已清理，将重新生成节点配置'
  fi
}

show_config() {
  [[ -f "$CONF_FILE" ]] || die "配置文件不存在：$CONF_FILE"
  sed 's/"password"[[:space:]]*:[[:space:]]*"[^"\\]*\(\\.[^"\\]*\)*"/"password": "********"/g' "$CONF_FILE"
}

show_node() {
  [[ -s "$CONF_FILE" ]] || die '尚未安装或配置文件不存在'
  command -v python3 >/dev/null || die '读取节点配置需要 python3'
  local fields bind_value port_value method_value password_value host_value uri_host userinfo uri node_name qx_name
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
  # 与 jinqians/ss-2022.sh 的 b64_url / View 使用相同的分享编码：
  # Base64URL(method:原始密码)，去掉换行和 padding。
  # 这是客户端兼容格式，不是 SIP002 为 AEAD-2022 指定的百分号编码格式。
  userinfo="$(printf '%s' "${method_value}:${password_value}" | base64 | tr '+/' '-_' | tr -d '=\r\n')"
  node_name="$(get_node_name)"
  uri="ss://${userinfo}@${uri_host}:${port_value}#$(encode_node_name "$node_name")"
  printf '\n===== SS2022 节点配置 =====\n'
  printf '节点名称：%s\n' "$node_name"
  printf '服务器：%s\n端口：%s\n加密：%s\n密码：%s\n\n节点导入链接（复制下一整行到“输入 SS URI”）：\n%s\n' "$host_value" "$port_value" "$method_value" "$password_value" "$uri"
  # Quantumult X 原生配置使用原始 Base64 密钥，不使用 URI 的百分号编码。
  # 逗号是字段分隔符；节点名称中的逗号和换行不能原样写入配置行。
  qx_name="${node_name//,/，}"
  qx_name="${qx_name//$'\r'/ }"
  qx_name="${qx_name//$'\n'/ }"
  printf '\nQuantumult X 配置（复制下一整行到配置文件的 [server_local] 下）：\n'
  printf 'shadowsocks=%s:%s, method=%s, password=%s, fast-open=true, udp-relay=true, tag=%s\n' "$uri_host" "$port_value" "$method_value" "$password_value" "${qx_name:-SS2022}"
  if command -v qrencode >/dev/null 2>&1; then
    printf '\n节点导入二维码（与上方链接相同）：\n'
    qrencode -t ANSIUTF8 "$uri"
  else
    log '未安装 qrencode，暂时无法输出二维码'
  fi
}

uninstall_server() {
  local confirm
  prompt '输入 DELETE 确认完整卸载：' confirm
  [[ "$confirm" == DELETE ]] || { log '已取消'; return; }
  read -r -p '输入 DELETE 确认完整卸载：' confirm
  [[ "$confirm" == DELETE ]] || { log '已取消'; return; }
  systemctl disable --now ss2022.service 2>/dev/null || true
  rm -f /etc/systemd/system/ss2022.service /usr/local/bin/ssserver
  rm -rf /etc/shadowsocks-rust
  systemctl daemon-reload
  log '服务、程序和配置已删除'
}

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

install_server() (
  set -Eeuo pipefail
  local version="${SS_VERSION:-1.25.0}" port="${SS_PORT:-$DEFAULT_PORT}" method="${SS_METHOD:-$DEFAULT_METHOD}"
  work='' changed=0 committed=0 was_active=0 was_enabled=0
  local node_name target archive url found
  change_paths=()
  log "SS2022 安装脚本版本：v${SCRIPT_VERSION}"
  log "即将安装 shadowsocks-rust：v${version}"
  [[ "$version" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]] || die 'SS_VERSION 格式应为 1.25.0'
  preflight_check
  install_manager_command
  install_tools
  [[ "$port" =~ ^[0-9]{1,5}$ ]] || die 'SS_PORT 必须是 1-65535 的端口'
  port="$((10#$port))"
  [[ "$port" -ge 1 && "$port" -le 65535 ]] || die 'SS_PORT 必须是 1-65535 的端口'
  [[ "$method" == "$DEFAULT_METHOD" ]] || die "目前只允许使用 $DEFAULT_METHOD"
  log '监听 IPv4：0.0.0.0'
  choose_node_name
  target="$(detect_arch)"
  work="$(mktemp -d)"
  trap 'finish_changes "$?"' EXIT
  trap 'exit 130' INT
  trap 'exit 143' TERM
  archive="shadowsocks-v${version}.${target}.tar.xz"
  url="https://github.com/shadowsocks/shadowsocks-rust/releases/download/v${version}/${archive}"
  log "下载 shadowsocks-rust v${version}（${target}）"
  curl --fail --location --retry 3 --connect-timeout 15 --max-time 300 --proto '=https' --tlsv1.2 -o "${work}/${archive}" "$url"
  mkdir "${work}/unpack"
  tar -xJf "${work}/${archive}" -C "${work}/unpack"
  found="$(find "${work}/unpack" -type f -name ssserver -perm -u+x -print -quit)"
  [[ -n "$found" ]] || die '压缩包中找不到 ssserver'
  "$found" --version >/dev/null || die '下载的 ssserver 无法在当前系统运行'
  write_config "${work}/input.json"
  validate_config "${work}/input.json"
  printf '%s\n' "$node_name" > "${work}/node-name"
  cat > "${work}/service" <<EOF
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
  backup_changes "$CONF_DIR" "$BIN" "$SERVICE_FILE"
  changed=1
  systemctl stop "${APP}.service" 2>/dev/null || true
  mkdir -p "$CONF_DIR"
  chmod 0700 "$CONF_DIR"
  atomic_install "${work}/input.json" "$CONF_FILE" 0600
  atomic_install "${work}/node-name" "${CONF_DIR}/node-name" 0600
  atomic_install "$found" "$BIN" 0755
  atomic_install "${work}/service" "$SERVICE_FILE" 0644
  systemctl daemon-reload
  restart_and_check
  systemctl enable "${APP}.service"
  committed=1
  log "安装完成：脚本 v${SCRIPT_VERSION} / shadowsocks-rust v${version}"
  log '未启用中国大陆来源 IP 屏蔽，允许大陆客户端连接'
  log '防火墙和云安全组需要放行实际 SS 端口的 TCP / UDP'
  log '以后直接输入 ss2022 打开管理菜单'
  # 节点导出失败不应撤销已经正常运行的服务。
  (show_node) || log '服务已启动，但节点导出失败；可设置 SS_HOST 后执行 ss2022 show'
)

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
    unblock-cn)
      unblock_cn
      install_manager_command
      ;;
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


validate_config() {
  python3 - "$1" <<'SS2022_VALIDATE_PY'
import base64
import json
import sys
try:
    config = json.load(open(sys.argv[1]))
    servers = config.get('servers') or [config]
    for server in servers:
        if server['method'] != '2022-blake3-aes-256-gcm':
            raise ValueError('仅支持 2022-blake3-aes-256-gcm')
        if len(base64.b64decode(server['password'], validate=True)) != 32:
            raise ValueError('SS2022 密钥必须是 32 字节的 Base64 编码')
        port = server['server_port']
        if type(port) is not int or not 1 <= port <= 65535:
            raise ValueError('端口必须是 1-65535 的整数')
        if not isinstance(server['server'], str) or not server['server'].strip():
            raise ValueError('监听地址为空')
except (ValueError, TypeError, KeyError, AttributeError):
    sys.exit('配置无效：请检查 JSON、监听地址、端口、加密方式和 32 字节 Base64 密钥')
SS2022_VALIDATE_PY
}

atomic_install() (
  source="$1" destination="$2" mode="$3" staged=''
  staged="$(mktemp "${destination}.XXXXXX")"
  trap 'rm -f "$staged"' EXIT
  install -m "$mode" "$source" "$staged"
  mv -f "$staged" "$destination"
)

backup_changes() {
  local i
  change_paths=("$@")
  mkdir "${work}/backup"
  for ((i=0; i<${#change_paths[@]}; i++)); do
    if [[ -e "${change_paths[i]}" ]]; then
      cp -a "${change_paths[i]}" "${work}/backup/${i}"
    fi
  done
  # 全新安装或清理旧安装后，服务文件尚未写入，无需查询旧状态。
  if [[ -f "$SERVICE_FILE" ]]; then
    if systemctl is-active --quiet "${APP}.service" 2>/dev/null; then was_active=1; fi
    if systemctl is-enabled --quiet "${APP}.service" 2>/dev/null; then was_enabled=1; fi
  fi
}

finish_changes() {
  local result="$1" i restore_failed=0
  trap - ERR
  set +e
  if [[ "$changed" == 1 && "$committed" == 0 ]]; then
    log '操作失败，正在撤销本次未完成的更改'
    systemctl stop "${APP}.service" 2>/dev/null
    if [[ "$was_enabled" == 0 ]]; then systemctl disable "${APP}.service" 2>/dev/null; fi
    for ((i=0; i<${#change_paths[@]}; i++)); do
      rm -rf "${change_paths[i]}" || restore_failed=1
      if [[ -e "${work}/backup/${i}" ]]; then
        cp -a "${work}/backup/${i}" "${change_paths[i]}" || restore_failed=1
      fi
    done
    systemctl daemon-reload || restore_failed=1
    if [[ "$was_active" == 1 ]]; then
      systemctl restart "${APP}.service" || restore_failed=1
    fi
    if [[ "$restore_failed" == 1 ]]; then
      log "恢复未完成，备份保留在：${work}/backup"
      return 1
    fi
    log '已撤销本次更改；安装前已清理的旧配置不会恢复'
  fi
  rm -rf "$work"
  return "$result"
}

restart_and_check() {
  systemctl restart "${APP}.service"
  # 等待启动后再检查，避免将启动即崩溃误报为安装成功。
  sleep 2
  systemctl is-active --quiet "${APP}.service" || {
    systemctl status "${APP}.service" --no-pager || true
    die '服务启动失败，请执行 ss2022 logs 查看日志'
  }
}

main() {
  case "${1:-}" in
    help|-h|--help) usage; return ;;
    version|-v|--version) printf '%s %s\n' "$APP" "$SCRIPT_VERSION"; return ;;
  esac
  [[ "$EUID" -eq 0 ]] || die '请使用 root 权限运行，例如 sudo bash ss2022.sh'
  command -v systemctl >/dev/null || die '此脚本需要 Linux / systemd'
  umask 077
  trap 'printf "[ss2022] ERROR: 操作失败，出错行：%s\n" "$LINENO" >&2' ERR
  dispatch "$@"
}

if [[ "${BASH_SOURCE[0]:-$0}" == "$0" ]]; then main "$@"; fi
