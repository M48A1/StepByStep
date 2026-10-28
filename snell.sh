#!/usr/bin/env bash
# snell-manager-managed
# Snell 官方服务端一键安装（Linux + systemd 247+，glibc）。
# 官方说明：https://kb.nssurge.com/surge-knowledge-base/release-notes/snell
# 用法：sudo bash snell.sh
# 运行后先选择版本，再输入端口；也可用环境变量预先指定。
# 自动安装：sudo env SNELL_VERSION=5.0.1 SNELL_PORT=23456 bash snell.sh
# 将完整入口定义为函数，流式运行时也能用 declare -f 保存管理命令。
# Bash 会先读取完整函数体，再执行入口，避免重新读取已消费的 /dev/fd 管道。
snell_entry() {
set -Eeuo pipefail
umask 077

readonly BIN=/usr/local/bin/snell-server
readonly CONF_DIR=/etc/snell
readonly CONF=/etc/snell/snell-server.conf
readonly UNIT=/etc/systemd/system/snell-server.service
readonly META=/etc/snell/installer-version
readonly CLIENT_CONF=/etc/snell/client-options.conf
readonly TFO_SYSCTL=/etc/sysctl.d/99-snell-tfo.conf
readonly SERVICE=snell-server.service
readonly MANAGER=/usr/local/bin/snell
readonly BACKUP_ROOT=/var/backups/snell
CHANGE_PATHS=("$BIN" "$CONF_DIR" "$UNIT" "$MANAGER" "$TFO_SYSCTL")
readonly CHANGE_PATHS
WORK=''
INSTALLING=0
TFO_CHANGED=0
TFO_OLD=''
BACKUP_DIR=''
WAS_ACTIVE=0
WAS_ENABLED=0

log() { printf '[snell] %s\n' "$*"; }
die() { printf '[snell] 错误：%s\n' "$*" >&2; exit 1; }

usage() {
  cat <<'EOF'
用法：snell [install|reinstall|info|config|status|start|stop|restart|logs]

首次运行安装；安装后直接输入 snell 打开管理菜单，非 root 自动调用 sudo。
reinstall 重新选择参数并重装，备份保存在 /var/backups/snell，失败时恢复。
支持 Debian 11+、Ubuntu 22.04+、Rocky/AlmaLinux 9+、Fedora 等系统。
要求 systemd >= 247 及 glibc；不支持 macOS、Alpine/OpenRC 或无 systemd 的容器。

首次安装会提示选择版本并输入端口，不设置默认版本和端口。
随后询问 TFO、连接复用、密钥、节点地址，以及 v6 传输模式。
也可设置环境变量跳过对应输入：
  SNELL_VERSION  5.0.1 或 6.0.0（官方文档仍标为测试版）
  SNELL_PORT     监听端口，1024–65535
  SNELL_PSK      16–128 位字母/数字/_/-，默认生成 32 位随机密钥
  SNELL_HOST     输出节点使用的公网 IPv4/IPv6 或域名，默认自动检测公网 IP
  SNELL_TFO      true/false，默认 false；开启时持久化内核 TFO 支持
  SNELL_REUSE    true/false，Surge 连接复用，默认 false
  SNELL_MODE     v6 使用 default/unshaped，默认 default
  SNELL_NETWORK ipv4/ipv6/dual，默认 ipv4，仅影响 Snell 的监听与出站
                ipv6（关闭 IPv4）要求 Snell v6 及可用的 IPv6 网络

无交互终端时必须指定版本和端口，其他选项使用上述默认值。

示例：
  sudo bash snell.sh
  sudo env SNELL_PORT=34567 SNELL_HOST=proxy.example.com bash snell.sh
  sudo env SNELL_VERSION=6.0.0 SNELL_PORT=34567 bash snell.sh
  sudo bash snell.sh info

可选仅 IPv4（关闭 IPv6）、仅 IPv6（关闭 IPv4）、双栈。
不会禁用主机网卡或改变 SSH 网络。请在云安全组及主机防火墙放行对应端口：
v5 使用 TCP 和 UDP；v6 使用 TCP。脚本会打印对应命令。
修改 /etc/snell/snell-server.conf 后运行 restart 生效。
EOF
}

check_system() {
  [[ $(uname -s) == Linux ]] || die '请在 Linux 服务器上运行此脚本。'
  [[ $EUID -eq 0 ]] || die '请使用 sudo bash snell.sh 或以 root 运行。'
  command -v systemctl >/dev/null && [[ -d /run/systemd/system ]] ||
    die '需要正在运行的 systemd。'
  local systemd_version
  systemd_version=$(systemctl --version | awk 'NR == 1 {print $2}')
  [[ $systemd_version =~ ^[0-9]+$ ]] && (( systemd_version >= 247 )) ||
    die '需要 systemd 247 或更新版本（例如 Debian 11+ / Ubuntu 22.04+）。'
  getconf GNU_LIBC_VERSION >/dev/null 2>&1 || die '官方二进制需要 glibc，不支持 Alpine/musl。'
}

valid_port() {
  [[ $1 =~ ^[0-9]{1,5}$ ]] && (( 10#$1 >= 1024 && 10#$1 <= 65535 ))
}

valid_host() {
  [[ $1 == YOUR_SERVER_IP || $1 =~ ^[a-zA-Z0-9][a-zA-Z0-9.-]*$ ||
     ( $1 == *:* && $1 =~ ^[0-9a-fA-F:]+$ ) ||
     $1 =~ ^\[[0-9a-fA-F:]+\]$ ]]
}

valid_detected_ip() {
  local value=$1 family=$2 part count=0 rest first second
  local parts=()
  if [[ $family == 4 ]]; then
    [[ $value =~ ^[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}$ ]] || return 1
    IFS=. read -r -a parts <<< "$value"
    for part in "${parts[@]}"; do (( 10#$part <= 255 )) || return 1; done
    first=$((10#${parts[0]})); second=$((10#${parts[1]}))
    (( first > 0 && first < 224 && first != 10 && first != 127 )) || return 1
    (( !(first == 169 && second == 254) && !(first == 172 && second >= 16 && second <= 31) &&
       !(first == 192 && second == 168) && !(first == 100 && second >= 64 && second <= 127) ))
  else
    # 只接收全球单播 IPv6，拒绝 HTML、内网地址及格式不完整的响应。
    [[ $value =~ ^[23][0-9a-fA-F]{3}:[0-9a-fA-F:]+$ && $value != *:::* ]] || return 1
    [[ $value != *: || $value == *:: ]] || return 1
    rest=${value#*::}
    [[ $rest != *::* ]] || return 1
    IFS=: read -r -a parts <<< "$value"
    for part in "${parts[@]}"; do
      [[ -n $part ]] || continue
      (( ${#part} <= 4 )) || return 1
      count=$((count + 1))
    done
    if [[ $value == *::* ]]; then
      (( count < 8 ))
    else
      [[ $value != *: ]] && (( count == 8 ))
    fi
  fi
}

detect_public_ip() {
  command -v curl >/dev/null || return 1
  local family url result
  local families=()
  case "${SNELL_NETWORK:-ipv4}" in
    ipv4) families=(4) ;;
    ipv6) families=(6) ;;
    dual) families=(4 6) ;;
    *) return 1 ;;
  esac
  for family in "${families[@]}"; do
    for url in https://api64.ipify.org https://ifconfig.co/ip; do
      if result=$(curl -q -fsS "-$family" --noproxy '*' --proto '=https' \
          --connect-timeout 2 --max-time 4 --max-filesize 128 "$url" 2>/dev/null); then
        result=${result//$'\r'/}
        if valid_detected_ip "$result" "$family"; then
          printf '%s\n' "$result"
          return 0
        fi
      fi
    done
  done
  return 1
}

choose_host() {
  local interactive=${1:-false} detected='' choice
  [[ -z ${SNELL_HOST:-} ]] || return 0
  log '正在自动检测服务器公网 IP…' >&2
  detected=$(detect_public_ip) || detected=''
  if [[ -n $detected ]]; then
    log "检测到公网 IP：${detected}" >&2
  else
    log '未能检测到公网 IP，可手动填写或稍后替换节点中的 YOUR_SERVER_IP。' >&2
  fi
  if [[ $interactive == true ]]; then
    while true; do
      if [[ -n $detected ]]; then
        printf '服务器 IP 或域名 [回车使用 %s，可输入其他地址覆盖]：' "$detected" >&2
      else
        printf '请输入服务器公网 IP 或域名 [回车稍后手动填写]：' >&2
      fi
      IFS= read -r -u 8 choice || die '输入已结束，安装已取消。'
      if [[ -z $choice ]] || valid_host "$choice"; then
        SNELL_HOST=${choice:-${detected:-YOUR_SERVER_IP}}
        return
      fi
      printf '请输入 IP 或域名，不要包含协议或端口。\n' >&2
    done
  fi
  SNELL_HOST=${detected:-YOUR_SERVER_IP}
}

ask_bool() {
  local name=$1 prompt=$2 choice
  [[ -z ${!name} ]] || return 0
  while true; do
    printf '%s [y/N]：' "$prompt" >&2
    IFS= read -r -u 8 choice || die '输入已结束，安装已取消。'
    case "$choice" in
      y|Y|yes|YES) printf -v "$name" '%s' true; return ;;
      ''|n|N|no|NO) printf -v "$name" '%s' false; return ;;
      *) printf '请输入 y 或 n。\n' >&2 ;;
    esac
  done
}

collect_options() {
  # 从终端读取，避免其他标准输入内容被当成安装选项。
  if [[ -t 0 ]]; then
    exec 8<&0
  elif ! { exec 8</dev/tty; } 2>/dev/null; then
    if [[ -n ${SNELL_VERSION:-} && -n ${SNELL_PORT:-} ]]; then
      choose_host false
      return 0
    fi
    die '无法读取交互终端，请同时指定 SNELL_VERSION 和 SNELL_PORT。'
  fi
  local choice
  if [[ -z ${SNELL_VERSION:-} ]]; then
    printf '\n请选择 Snell 版本：\n  1) v5.0.1\n  2) v6.0.0（官方文档仍标为测试版，客户端须支持 v6）\n' >&2
    while true; do
      printf '请输入选项 [1/2，无默认值]：' >&2
      IFS= read -r -u 8 choice || die '输入已结束，安装已取消。'
      case "$choice" in
        1|5.0.1) SNELL_VERSION=5.0.1; break ;;
        2|6.0.0)
          if [[ $(uname -m) == armv7l ]]; then
            printf '官方 v6 未提供 armv7l，请选择 v5.0.1。\n' >&2
            continue
          fi
          SNELL_VERSION=6.0.0; break ;;
        *) printf '请输入 1 或 2，不能留空。\n' >&2 ;;
      esac
    done
  fi
  if [[ -z ${SNELL_PORT:-} ]]; then
    while true; do
      printf '请输入自定义监听端口 [1024–65535，无默认值]：' >&2
      IFS= read -r -u 8 choice || die '输入已结束，安装已取消。'
      if valid_port "$choice"; then
        SNELL_PORT=$((10#$choice))
        break
      fi
      printf '端口无效，请输入 1024–65535 的整数。\n' >&2
    done
  fi
  if [[ -z ${SNELL_NETWORK:-} ]]; then
    printf '\nSnell 网络模式（只影响 Snell）：\n  1) 仅 IPv4，关闭 IPv6\n  2) 仅 IPv6，关闭 IPv4（需 Snell v6）\n  3) IPv4 + IPv6 双栈\n' >&2
    while true; do
      printf '请选择 [1/2/3，默认 1]：' >&2
      IFS= read -r -u 8 choice || die '输入已结束，安装已取消。'
      case "$choice" in
        ''|1|ipv4) SNELL_NETWORK=ipv4; break ;;
        2|ipv6)
          if [[ $SNELL_VERSION != 6.0.0 ]]; then
            printf '仅 IPv6 出站控制需要 v6，请选其他模式或重新运行选择 v6。\n' >&2
            continue
          fi
          SNELL_NETWORK=ipv6; break ;;
        3|dual) SNELL_NETWORK=dual; break ;;
        *) printf '请输入 1、2 或 3。\n' >&2 ;;
      esac
    done
  fi
  SNELL_TFO=${SNELL_TFO:-}
  SNELL_REUSE=${SNELL_REUSE:-}
  ask_bool SNELL_TFO '是否启用 TCP Fast Open（同时开启并持久化服务器内核 TFO 支持）'
  ask_bool SNELL_REUSE '是否启用 Surge 连接复用 reuse'
  if [[ ! ${SNELL_PSK+x} ]]; then
    while true; do
      printf '请输入自定义 PSK [回车生成随机密钥，输入不回显]：' >&2
      IFS= read -r -s -u 8 choice || die '输入已结束，安装已取消。'
      printf '\n' >&2
      if [[ -z $choice || $choice =~ ^[a-zA-Z0-9_-]{16,128}$ ]]; then
        SNELL_PSK=$choice
        break
      fi
      printf 'PSK 必须为 16–128 位字母、数字、下划线或连字符。\n' >&2
    done
  fi
  choose_host true
  if [[ $SNELL_VERSION == 6.0.0 && -z ${SNELL_MODE:-} ]]; then
    printf '\nv6 传输模式：\n  1) default：加密并进行流量整形\n  2) unshaped：保留加密，关闭流量整形\n' >&2
    while true; do
      printf '请选择 [1/2，默认 1]：' >&2
      IFS= read -r -u 8 choice || die '输入已结束，安装已取消。'
      case "$choice" in
        ''|1|default) SNELL_MODE=default; break ;;
        2|unshaped) SNELL_MODE=unshaped; break ;;
        *) printf '请输入 1 或 2。\n' >&2 ;;
      esac
    done
  fi
  exec 8<&-
}

validate_options() {
  VERSION=${SNELL_VERSION:-}
  PORT=${SNELL_PORT:-}
  PSK=${SNELL_PSK:-}
  HOST=${SNELL_HOST:-YOUR_SERVER_IP}
  TFO=${SNELL_TFO:-false}
  REUSE=${SNELL_REUSE:-false}
  MODE=${SNELL_MODE:-default}
  NETWORK=${SNELL_NETWORK:-ipv4}
  case "$NETWORK" in ipv4|ipv6|dual) ;; *) die 'SNELL_NETWORK 必须为 ipv4、ipv6 或 dual。' ;; esac
  [[ $NETWORK != ipv6 || $VERSION == 6.0.0 ]] || die '仅 IPv6 模式需要 Snell v6。'
  case "$TFO" in true|false) ;; *) die 'SNELL_TFO 必须为 true 或 false。' ;; esac
  case "$REUSE" in true|false) ;; *) die 'SNELL_REUSE 必须为 true 或 false。' ;; esac
  case "$MODE" in default|unshaped) ;; *) die 'SNELL_MODE 仅支持 default 或 unshaped。' ;; esac
  [[ $VERSION == 6.0.0 || $MODE == default ]] || die 'unshaped 模式仅适用于 v6。'
  case "$VERSION" in 5.0.1|6.0.0) ;; *) die 'SNELL_VERSION 仅支持 5.0.1 或 6.0.0。' ;; esac
  valid_port "$PORT" || die '端口必须是 1024–65535 的整数。'
  PORT=$((10#$PORT))
  [[ -z $PSK || $PSK =~ ^[a-zA-Z0-9_-]{16,128}$ ]] || die 'PSK 必须为 16–128 位字母、数字、下划线或连字符。'
  valid_host "$HOST" || die 'SNELL_HOST 请填写公网 IP 或域名，不要包含协议或端口。'
  case "$(uname -m)" in
    x86_64) ARCH=amd64 ;;
    i386|i686) ARCH=i386 ;;
    aarch64|arm64) ARCH=aarch64 ;;
    armv7l)
      [[ $VERSION == 5.0.1 ]] || die '官方 v6 下载列表未提供 armv7l。'
      ARCH=armv7l ;;
    *) die "不支持的架构：$(uname -m)" ;;
  esac
}

install_dependencies() {
  if command -v curl >/dev/null && command -v unzip >/dev/null &&
     command -v openssl >/dev/null && command -v ss >/dev/null &&
     command -v sysctl >/dev/null &&
     { [[ -s /etc/ssl/certs/ca-certificates.crt ]] || [[ -s /etc/pki/tls/certs/ca-bundle.crt ]]; }; then
    return
  fi
  log '安装 curl、unzip、openssl、CA 证书和端口检测工具…'
  if command -v apt-get >/dev/null; then
    apt-get update
    DEBIAN_FRONTEND=noninteractive apt-get install -y ca-certificates curl unzip openssl iproute2 procps
  elif command -v dnf >/dev/null; then
    dnf install -y ca-certificates curl unzip openssl iproute procps-ng
  elif command -v yum >/dev/null; then
    yum install -y ca-certificates curl unzip openssl iproute procps-ng
  else
    die '未找到 apt-get/dnf/yum，请先自行安装 curl、unzip、openssl、CA 证书、ss 和 sysctl。'
  fi
}

snapshot_installation() {
  local reinstall=$1 path index=0
  if [[ $reinstall == true ]]; then
    install -d -m 700 "$BACKUP_ROOT"
    BACKUP_DIR=$(mktemp -d "$BACKUP_ROOT/$(date +%Y%m%d-%H%M%S)-XXXXXX")
    log "重装备份：$BACKUP_DIR"
  else
    BACKUP_DIR="$WORK/backup"
    mkdir "$BACKUP_DIR"
  fi
  for path in "${CHANGE_PATHS[@]}"; do
    if [[ -e $path || -L $path ]]; then
      cp -a -- "$path" "$BACKUP_DIR/$index"
    fi
    printf '%s\t%s\n' "$index" "$path" >> "$BACKUP_DIR/paths.txt"
    index=$((index + 1))
  done
  systemctl is-active --quiet "$SERVICE" && WAS_ACTIVE=1
  systemctl is-enabled --quiet "$SERVICE" && WAS_ENABLED=1
  printf 'active=%s\nenabled=%s\n' "$WAS_ACTIVE" "$WAS_ENABLED" > "$BACKUP_DIR/service-state"
}

cleanup() {
  local rc=$? index=0 path failed=0
  trap - EXIT
  if (( INSTALLING )); then
    log '安装未完成，恢复变更前的文件和服务状态。' >&2
    systemctl stop "$SERVICE" >/dev/null 2>&1 || true
    systemctl disable "$SERVICE" >/dev/null 2>&1 || true
    for path in "${CHANGE_PATHS[@]}"; do
      rm -rf -- "$path" || failed=1
      if [[ -e $BACKUP_DIR/$index || -L $BACKUP_DIR/$index ]]; then
        cp -a -- "$BACKUP_DIR/$index" "$path" || failed=1
      fi
      index=$((index + 1))
    done
    if (( TFO_CHANGED )); then
      sysctl -q -w "net.ipv4.tcp_fastopen=$TFO_OLD" || failed=1
    fi
    systemctl daemon-reload >/dev/null 2>&1 || failed=1
    if (( WAS_ENABLED )); then systemctl enable "$SERVICE" >/dev/null 2>&1 || failed=1; fi
    if (( WAS_ACTIVE )); then systemctl start "$SERVICE" || failed=1; fi
    if (( failed )); then
      log "自动恢复未完全成功，请检查备份：$BACKUP_DIR" >&2
      # 初次安装的备份在 WORK 内；恢复失败时保留现场。
      exit 1
    fi
  fi
  [[ -z $WORK ]] || rm -rf -- "$WORK"
  exit "$rc"
}

atomic_install() {
  local source=$1 dest=$2 mode=$3 staging
  staging=$(mktemp "${dest}.tmp.XXXXXX")
  if ! install -m "$mode" "$source" "$staging" || ! mv -f -- "$staging" "$dest"; then
    rm -f -- "$staging"
    return 1
  fi
}

prepare_network() {
  DENY_FAMILY=''
  case "$NETWORK" in
    ipv4) LISTEN="0.0.0.0:$PORT"; FAMILIES='AF_INET AF_UNIX AF_NETLINK'; DENY_FAMILY=ipv6 ;;
    ipv6)
      FAMILIES='AF_INET6 AF_UNIX AF_NETLINK'
      DENY_FAMILY=ipv4
      LISTEN=''
      local address index prefix scope flags interface formatted offset
      [[ -r /proc/net/if_inet6 ]] || die '当前系统未启用 IPv6。'
      while read -r address index prefix scope flags interface; do
        [[ $scope == 00 && $interface != lo ]] || continue
        (( (16#$flags & 72) == 0 )) || continue
        formatted=${address:0:4}
        for offset in 4 8 12 16 20 24 28; do formatted+=":${address:offset:4}"; done
        LISTEN+="${LISTEN:+,}[$formatted]:$PORT"
      done < /proc/net/if_inet6
      [[ -n $LISTEN ]] || die '仅 IPv6 模式需要可用的全局 IPv6 地址。'
      # 绑定具体 IPv6 地址，避免 [::] 接受 IPv4-mapped 入站连接。
      ;;
    dual)
      FAMILIES='AF_INET AF_INET6 AF_UNIX AF_NETLINK'
      if [[ $VERSION == 6.* ]]; then
        LISTEN="0.0.0.0:$PORT,[::]:$PORT"
      else
        [[ $(sysctl -n net.ipv6.bindv6only) == 0 ]] || die 'v5 双栈要求 bindv6only=0；请使用 v6 或仅 IPv4 模式。'
        LISTEN="[::]:$PORT"
      fi ;;
  esac
  if [[ $NETWORK != ipv4 ]]; then
    [[ -r /proc/net/if_inet6 ]] && awk '$6 != "lo" {found=1} END {exit !found}' /proc/net/if_inet6 ||
      die '未检测到非回环 IPv6 地址，请先为服务器配置 IPv6。'
  fi
}

write_server_config() {
  printf '[snell-server]\nlisten = %s\npsk = %s\n' "$LISTEN" "$PSK"
  if [[ $VERSION == 6.* ]]; then
    printf 'mode = %s\n' "$MODE"
    case "$NETWORK" in
      ipv4) printf 'dns-ip-preference = ipv4-only\n' ;;
      ipv6)
        printf 'dns-ip-preference = ipv6-only\n'
        # 纯 IPv6 服务不能访问仅有 IPv4 地址的系统 DNS。
        printf 'dns = 2606:4700:4700::1111,2001:4860:4860::8888\n' ;;
      dual) printf 'dns-ip-preference = default\n' ;;
    esac
  elif [[ $NETWORK == ipv4 ]]; then
    printf 'ipv6 = false\n'
  else
    printf 'ipv6 = true\n'
  fi
}

config_value() {
  awk -v key="$1" '
    /^[[:space:]]*[#;]/ {next}
    index($0, "=") {
      k = substr($0, 1, index($0, "=") - 1)
      gsub(/^[[:space:]]+|[[:space:]]+$/, "", k)
      if (k == key) {
        v = substr($0, index($0, "=") + 1)
        gsub(/^[[:space:]]+|[[:space:]]+$/, "", v)
        print v; exit
      }
    }' "${2:-$CONF}"
}

enable_kernel_tfo() {
  [[ $TFO == true ]] || return 0
  command -v sysctl >/dev/null || die '开启 TFO 需要 sysctl，请安装 procps/procps-ng。'
  if [[ -e $TFO_SYSCTL ]]; then
    [[ $(head -n 1 "$TFO_SYSCTL") == '# Created by snell.sh' ]] || die "已有非本脚本创建的 ${TFO_SYSCTL}，不会覆盖。"
  fi
  TFO_OLD=$(sysctl -n net.ipv4.tcp_fastopen) || die '当前内核不支持 TFO。'
  [[ $TFO_OLD =~ ^[0-9]{1,6}$ ]] || die '无法识别当前内核 TFO 设置。'
  printf '%s\n' "$TFO_OLD" > "$BACKUP_DIR/kernel-tfo"
  local value=$((10#$TFO_OLD | 3))
  # 保留其他位，仅增加客户端与服务端支持；不强制所有监听端口启用 TFO。
  TFO_CHANGED=1
  sysctl -q -w "net.ipv4.tcp_fastopen=$value" || die '无法启用内核 TFO（容器可能限制此操作）。'
  install -d -m 755 /etc/sysctl.d
  printf '# Created by snell.sh\nnet.ipv4.tcp_fastopen = %s\n' "$value" > "$TFO_SYSCTL"
  chmod 644 "$TFO_SYSCTL"
}

show_info() {
  [[ -s $CONF && -s $META ]] || die '未找到本脚本安装的 Snell。'
  local version protocol listen port psk host mode network=legacy tfo=false reuse=false
  version=$(cat "$META")
  protocol=${version%%.*}
  listen=$(config_value listen)
  port=${listen##*:}
  psk=$(config_value psk)
  host=YOUR_SERVER_IP
  if [[ -s $CLIENT_CONF ]]; then
    host=$(config_value host "$CLIENT_CONF")
    tfo=$(config_value tfo "$CLIENT_CONF")
    reuse=$(config_value reuse "$CLIENT_CONF")
    network=$(config_value network "$CLIENT_CONF")
  fi
  host=${SNELL_HOST:-${host:-YOUR_SERVER_IP}}
  valid_host "$host" || die 'SNELL_HOST 请填写公网 IP 或域名。'
  if [[ $host == *:* && $host != \[*\] ]]; then host="[$host]"; fi
  printf '\nSnell %s\n监听地址：%s\n配置文件：%s\n\n' "$version" "$listen" "$CONF"
  printf '网络模式：%s\n' "${network:-legacy（请查看原始配置）}"
  log '以下为 Surge [Proxy] 配置片段（包含密钥，请妥善保管）：'
  printf '\n[Proxy]\n'
  printf 'Snell = snell, %s, %s, psk=%s, version=%s' "$host" "$port" "$psk" "$protocol"
  printf ', tfo=%s, reuse=%s' "${tfo:-false}" "${reuse:-false}"
  case "$network" in
    ipv4) printf ', ip-version=v4-only' ;;
    ipv6) printf ', ip-version=v6-only' ;;
  esac
  if [[ $protocol == 6 ]]; then
    mode=$(config_value mode)
    printf ', mode=%s' "${mode:-default}"
  fi
  printf '\n\n'
  log '已有 [Proxy] 区域时，只复制上面的 Snell = ... 这一整行到该区域，无需重复添加 [Proxy]。'
  log '这是一段节点配置；节点名称 Snell 可自行修改。以后运行 snell info 可再次查看。'
  [[ $host != YOUR_SERVER_IP ]] || log '请将 YOUR_SERVER_IP 替换成服务器公网 IP 或域名。'
  log "请在云安全组和主机防火墙放行 ${port}/tcp。"
  printf '  UFW:       ufw allow %s/tcp\n' "$port"
  printf '  firewalld: firewall-cmd --permanent --add-port=%s/tcp\n' "$port"
  if [[ $protocol == 5 ]]; then
    log "v5 QUIC Proxy Mode 还需要放行 ${port}/udp。"
    printf '  UFW:       ufw allow %s/udp\n' "$port"
    printf '  firewalld: firewall-cmd --permanent --add-port=%s/udp\n' "$port"
    log '旧版 Surge 可尝试将节点 version=5 改为 version=4（不使用 QUIC Proxy Mode）。'
  fi
  log '按实际使用的防火墙执行对应命令；firewalld 修改后执行 firewall-cmd --reload。'
  log '查看状态：systemctl status snell-server；查看日志：journalctl -u snell-server -e'
}

is_installed() {
  [[ -s $META && -s $CONF && -x $BIN && -f $UNIT ]]
}

write_manager() {
  {
    printf '#!/usr/bin/env bash\n# snell-manager-managed\n'
    declare -f snell_entry || return 1
    printf '\nsnell_entry "$@"\n'
  } > "$1"
}

install_manager() {
  local staging
  staging=$(mktemp "${MANAGER}.tmp.XXXXXX")
  if write_manager "$staging" && chmod 755 "$staging" && mv -f -- "$staging" "$MANAGER"; then
    return 0
  fi
  rm -f -- "$staging"
  return 1
}

check_manager_target() {
  local marker
  if [[ -e $MANAGER || -L $MANAGER ]]; then
    [[ ! -L $MANAGER && -f $MANAGER ]] &&
      marker=$(head -n 2 "$MANAGER") && [[ $marker == *'# snell-manager-managed'* ]] ||
      die "已有其他程序占用 ${MANAGER}，不会覆盖。"
  fi
}

install_snell() {
  local reinstall=${1:-false} path
  check_manager_target
  if is_installed && [[ $reinstall != true ]]; then
    install_manager
    log '保留现有安装，已安装/更新 snell 管理命令。修改参数请使用 snell reinstall。'
    show_info
    return
  fi
  if [[ $reinstall == true ]]; then
    is_installed || die '未找到完整的现有安装，请先运行 snell install。'
  else
    for path in "$BIN" "$CONF_DIR" "$UNIT"; do
      [[ ! -e $path && ! -L $path ]] || die "检测到已有文件：${path}，请先检查现有安装；不会覆盖。"
    done
    if systemctl cat "$SERVICE" >/dev/null 2>&1; then
      die '检测到同名 systemd 服务，请先检查现有安装。'
    fi
  fi
  for path in "${CHANGE_PATHS[@]}"; do
    [[ ! -L $path ]] || die "不支持符号链接安装路径：$path"
  done
  install_dependencies
  collect_options
  validate_options
  prepare_network
  [[ -n $PSK ]] || PSK=$(openssl rand -hex 16)
  WORK=$(mktemp -d)
  trap cleanup EXIT
  trap 'exit 130' INT
  trap 'exit 143' TERM
  local url="https://dl.nssurge.com/snell/snell-server-v${VERSION}-linux-${ARCH}.zip"
  log "从官方源下载 Snell ${VERSION} (${ARCH})…"
  curl --fail --show-error --location --proto '=https' --proto-redir '=https' \
    --tlsv1.2 --connect-timeout 15 --max-time 180 --retry 3 \
    "$url" -o "$WORK/snell.zip"
  unzip -tq "$WORK/snell.zip" >/dev/null
  # 只读取指定文件，不将归档中的其他路径解压到系统。
  unzip -p "$WORK/snell.zip" snell-server > "$WORK/snell-server"
  [[ -s $WORK/snell-server ]] || die '下载的归档中没有 snell-server。'
  chmod 700 "$WORK/snell-server"
  "$WORK/snell-server" --help > "$WORK/help.txt" 2>&1 ||
    die '二进制无法运行，请检查 CPU 架构或升级系统 glibc。'
  write_server_config > "$WORK/snell-server.conf"
  write_manager "$WORK/manager"
  cat > "$WORK/snell-server.service" <<EOF
[Unit]
Description=Snell Proxy Server
Wants=network-online.target
After=network-online.target

[Service]
Type=simple
DynamicUser=yes
LoadCredential=snell.conf:${CONF}
ExecStart=${BIN} -c %d/snell.conf
Restart=on-failure
RestartSec=3
LimitNOFILE=65535
NoNewPrivileges=yes
PrivateTmp=yes
ProtectSystem=strict
ProtectHome=yes
RestrictAddressFamilies=${FAMILIES}
IPAddressDeny=${DENY_FAMILY}

[Install]
WantedBy=multi-user.target
EOF
  snapshot_installation "$reinstall"
  INSTALLING=1
  if [[ $reinstall == true ]]; then systemctl stop "$SERVICE"; fi
  local sockets
  sockets=$(ss -H -lntu)
  if awk -v port="$PORT" '$5 ~ (":" port "$") {found=1} END {exit !found}' <<< "$sockets"; then
    die "端口 ${PORT} 已被占用，请重新选择。"
  fi
  install -d -m 700 "$CONF_DIR"
  install -d -m 755 /usr/local/bin
  atomic_install "$WORK/snell-server.conf" "$CONF" 600
  atomic_install "$WORK/snell-server" "$BIN" 755
  atomic_install "$WORK/snell-server.service" "$UNIT" 644
  atomic_install "$WORK/manager" "$MANAGER" 755
  printf '%s\n' "$VERSION" > "$WORK/meta"
  printf 'host = %s\ntfo = %s\nreuse = %s\nnetwork = %s\n' "$HOST" "$TFO" "$REUSE" "$NETWORK" > "$WORK/client"
  atomic_install "$WORK/meta" "$META" 600
  atomic_install "$WORK/client" "$CLIENT_CONF" 600
  enable_kernel_tfo
  systemctl daemon-reload
  systemctl enable --now "$SERVICE"
  sleep 2
  if ! systemctl is-active --quiet "$SERVICE"; then
    die '服务启动失败。请用 journalctl -u snell-server -n 50 --no-pager 查看原因。'
  fi
  INSTALLING=0
  log '安装完成，服务已启动并设置开机自启。以后输入 snell 打开管理菜单。'
  show_info
}

show_config() {
  is_installed || die '尚未安装 Snell。'
  show_info
  printf '\n--- 服务端配置（包含密钥）---\n'
  cat "$CONF"
  if [[ -f $CLIENT_CONF ]]; then
    printf '\n--- 节点选项 ---\n'
    cat "$CLIENT_CONF"
  fi
  printf '\n--- systemd 服务配置 ---\n'
  cat "$UNIT"
}

select_action() {
  if [[ -t 0 ]]; then
    exec 8<&0
  elif ! { exec 8</dev/tty; } 2>/dev/null; then
    die '没有交互终端，请指定 snell info、snell config 或 snell reinstall 等命令。'
  fi
  printf '\nSnell 管理菜单\n  1) 查看现有配置和节点\n  2) 查看服务状态\n  3) 重装 / 重新选择网络与其他参数\n  4) 启动\n  5) 停止\n  6) 重启\n  7) 查看日志\n  0) 退出\n' >&2
  local choice
  while true; do
    printf '请选择：' >&2
    IFS= read -r -u 8 choice || die '输入已结束。'
    case "$choice" in
      1) ACTION=config; break ;;
      2) ACTION=status; break ;;
      3) ACTION=reinstall; break ;;
      4) ACTION=start; break ;;
      5) ACTION=stop; break ;;
      6) ACTION=restart; break ;;
      7) ACTION=logs; break ;;
      0) ACTION=quit; break ;;
      *) printf '请输入 0–7。\n' >&2 ;;
    esac
  done
  exec 8<&-
}

main() {
  [[ $# -le 1 ]] || { usage >&2; return 1; }
  ACTION=${1:-auto}
  case "$ACTION" in
    help|-h|--help) usage; return ;;
    auto|menu|install|reinstall|info|config|status|start|stop|restart|logs) ;;
    *) usage >&2; return 1 ;;
  esac
  if [[ $EUID -ne 0 && $(uname -s) == Linux ]]; then
    command -v sudo >/dev/null || die '请以 root 运行。'
    local elevated_dir elevated_rc
    elevated_dir=$(mktemp -d)
    # sudo 通常关闭 /dev/fd 描述符，先将完整入口保存到私有临时目录。
    trap 'rm -rf -- "$elevated_dir"' EXIT
    trap 'exit 130' INT
    trap 'exit 143' TERM
    write_manager "$elevated_dir/snell.sh"
    if sudo bash "$elevated_dir/snell.sh" "$@"; then elevated_rc=0; else elevated_rc=$?; fi
    rm -rf -- "$elevated_dir"
    trap - EXIT INT TERM
    return "$elevated_rc"
  fi
  check_system
  command -v flock >/dev/null || die '缺少 flock，请安装 util-linux。'
  exec 9>/run/lock/snell-installer.lock
  flock -n 9 || die '另一个 Snell 管理进程正在运行。'
  if [[ $ACTION == auto ]]; then
    if is_installed; then ACTION=menu; else ACTION=install; fi
  fi
  if [[ $ACTION == menu ]]; then
    check_manager_target
    install_manager
    select_action
  fi
  case "$ACTION" in
    install) install_snell ;;
    reinstall) install_snell true ;;
    info) show_info ;;
    config) show_config ;;
    status) systemctl status "$SERVICE" --no-pager ;;
    start|stop|restart) systemctl "$ACTION" "$SERVICE" ;;
    logs) flock -u 9; journalctl -u "$SERVICE" -n 100 -f ;;
    quit) return ;;
  esac
}

if [[ -z ${BASH_SOURCE[0]:-} || ${BASH_SOURCE[0]} == "$0" ]]; then
  main "$@"
fi
}

snell_entry "$@"
