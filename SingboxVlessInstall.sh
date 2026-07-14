#!/usr/bin/env bash
set -Eeuo pipefail

# sing-box VLESS Reality IPv4-only 一键安装脚本
readonly VERSION="1.0.0"
readonly CONF_DIR="/etc/sing-box"
readonly CONF_FILE="${CONF_DIR}/config.json"
readonly INFO_FILE="${CONF_DIR}/vless-info.env"
readonly LINK_FILE="${CONF_DIR}/vless-link.txt"
readonly SYSCTL_FILE="/etc/sysctl.d/99-sing-box-ipv4-only.conf"
readonly MANAGER="/usr/local/bin/vless"

PORT=443
PORT_EXPLICIT=0
SNI="www.dell.com"
SNI_EXPLICIT=0

say() { printf '\033[1;34m[信息]\033[0m %s\n' "$*"; }
ok() { printf '\033[1;32m[完成]\033[0m %s\n' "$*"; }
warn() { printf '\033[1;33m[注意]\033[0m %s\n' "$*"; }
die() { printf '\033[1;31m[错误]\033[0m %s\n' "$*" >&2; exit 1; }

need_root() { [[ $EUID -eq 0 ]] || die "请用 root 运行：sudo bash $0 $*"; }
valid_port() { [[ "$1" =~ ^[0-9]+$ ]] && ((10#$1 >= 1 && 10#$1 <= 65535)) || die "端口必须为 1-65535"; }
valid_sni() { [[ "$1" =~ ^([A-Za-z0-9]([A-Za-z0-9-]*[A-Za-z0-9])?\.)+[A-Za-z]{2,63}$ ]] || die "SNI 格式不正确：$1"; }

choose_port() {
  ((PORT_EXPLICIT == 0)) || return
  if [[ ! -t 0 ]]; then
    say "非交互运行，默认使用端口：$PORT"
    return
  fi
  printf '\n请选择 VLESS 监听端口：\n'
  printf '  1) 443（默认）\n'
  printf '  2) 8443\n'
  printf '  3) 自定义端口\n'
  read -r -p '请输入选项 [1-3，默认 1]：' choice
  case "${choice:-1}" in
    1) PORT=443 ;;
    2) PORT=8443 ;;
    3)
      read -r -p '请输入端口 [1-65535]：' PORT
      valid_port "$PORT"
      ;;
    *) die "无效的端口选项：$choice" ;;
  esac
  say "已选择端口：$PORT"
}

choose_sni() {
  ((SNI_EXPLICIT == 0)) || return
  if [[ ! -t 0 ]]; then
    say "非交互运行，默认选择 Dell：$SNI"
    return
  fi
  printf '\n请选择 Reality SNI：\n'
  printf '  1) Dell（www.dell.com）\n'
  printf '  2) 新加坡虾皮（shopee.sg）\n'
  printf '  3) 自定义域名\n'
  read -r -p '请输入选项 [1-3，默认 1]：' choice
  case "${choice:-1}" in
    1) SNI="www.dell.com" ;;
    2) SNI="shopee.sg" ;;
    3)
      read -r -p '请输入自定义 SNI：' SNI
      valid_sni "$SNI"
      ;;
    *) die "无效的 SNI 选项：$choice" ;;
  esac
  say "已选择 SNI：$SNI"
}

install_deps() {
  if command -v apt-get >/dev/null; then
    apt-get update
    DEBIAN_FRONTEND=noninteractive apt-get install -y curl ca-certificates jq openssl qrencode iproute2
  elif command -v dnf >/dev/null; then
    dnf install -y curl ca-certificates jq openssl qrencode iproute
  elif command -v yum >/dev/null; then
    yum install -y curl ca-certificates jq openssl qrencode iproute
  else
    die "仅支持 Debian/Ubuntu/RHEL/Rocky/Alma Linux"
  fi
}

install_singbox() {
  if command -v sing-box >/dev/null; then
    say "sing-box 已安装：$(sing-box version | sed -n '1p')"
  else
    say "使用官方脚本安装 sing-box"
    curl -fsSL https://sing-box.app/install.sh | sh
    command -v sing-box >/dev/null || die "sing-box 安装失败"
  fi
}

disable_ipv6() {
  say "关闭系统 IPv6（立即生效并永久保存）"
  cat >"$SYSCTL_FILE" <<'EOF'
net.ipv6.conf.all.disable_ipv6 = 1
net.ipv6.conf.default.disable_ipv6 = 1
net.ipv6.conf.lo.disable_ipv6 = 1
EOF
  sysctl -p "$SYSCTL_FILE" >/dev/null
}

public_ipv4() {
  local addr
  addr="$(curl -4fsS --max-time 8 https://api.ipify.org 2>/dev/null || true)"
  [[ "$addr" =~ ^([0-9]{1,3}\.){3}[0-9]{1,3}$ ]] || addr="$(ip -4 route get 1.1.1.1 2>/dev/null | awk '{for(i=1;i<=NF;i++)if($i=="src"){print $(i+1);exit}}')"
  [[ "$addr" =~ ^([0-9]{1,3}\.){3}[0-9]{1,3}$ ]] || die "无法取得服务器 IPv4"
  printf '%s' "$addr"
}

generate_keys() {
  UUID="$(sing-box generate uuid)"
  local pair
  pair="$(sing-box generate reality-keypair)"
  PRIVATE_KEY="$(awk -F': *' 'tolower($1) ~ /privatekey|private key/{print $2;exit}' <<<"$pair")"
  PUBLIC_KEY="$(awk -F': *' 'tolower($1) ~ /publickey|public key/{print $2;exit}' <<<"$pair")"
  SHORT_ID="$(openssl rand -hex 8)"
  [[ -n "$UUID" && -n "$PRIVATE_KEY" && -n "$PUBLIC_KEY" ]] || die "Reality 密钥解析失败：$pair"
}

write_config() {
  mkdir -p "$CONF_DIR"
  umask 077
  jq -n --argjson port "$PORT" --arg uuid "$UUID" --arg sni "$SNI" \
    --arg private "$PRIVATE_KEY" --arg sid "$SHORT_ID" '{
      log:{level:"info",timestamp:true},
      dns:{
        servers:[{type:"udp",tag:"cloudflare",server:"1.1.1.1",server_port:53}],
        final:"cloudflare",
        strategy:"ipv4_only"
      },
      inbounds:[{
        type:"vless",tag:"vless-reality-in",listen:"0.0.0.0",listen_port:$port,
        users:[{name:"vless",uuid:$uuid,flow:"xtls-rprx-vision"}],
        tls:{enabled:true,server_name:$sni,reality:{
          enabled:true,
          handshake:{server:$sni,server_port:443,domain_resolver:{server:"cloudflare",strategy:"ipv4_only"}},
          private_key:$private,short_id:[$sid]
        }}
      }],
      outbounds:[{type:"direct",tag:"direct",domain_resolver:{server:"cloudflare",strategy:"ipv4_only"}}],
      route:{default_domain_resolver:{server:"cloudflare",strategy:"ipv4_only"},final:"direct"}
    }' >"$CONF_FILE"
  chmod 600 "$CONF_FILE"
}

check_config() {
  local output
  output="$(sing-box check -c "$CONF_FILE" 2>&1)" || { printf '%s\n' "$output" >&2; die "配置校验失败"; }
}

save_node() {
  SERVER_IPV4="$(public_ipv4)"
  local link
  link="vless://${UUID}@${SERVER_IPV4}:${PORT}?encryption=none&flow=xtls-rprx-vision&security=reality&sni=${SNI}&fp=chrome&pbk=${PUBLIC_KEY}&sid=${SHORT_ID}&type=tcp#sing-box-${SERVER_IPV4}"
  {
    printf 'SERVER_IPV4=%q\n' "$SERVER_IPV4"
    printf 'PORT=%q\n' "$PORT"
    printf 'UUID=%q\n' "$UUID"
    printf 'SNI=%q\n' "$SNI"
    printf 'PUBLIC_KEY=%q\n' "$PUBLIC_KEY"
    printf 'SHORT_ID=%q\n' "$SHORT_ID"
  } >"$INFO_FILE"
  printf '%s\n' "$link" >"$LINK_FILE"
  chmod 600 "$INFO_FILE" "$LINK_FILE"
}

install_manager() {
  local temp_manager
  temp_manager="$(mktemp /tmp/vless-manager.XXXXXX)"
  {
    printf '%s\n' '#!/usr/bin/env bash' 'set -Eeuo pipefail'
    printf 'readonly VERSION=%q\n' "$VERSION"
    printf 'readonly CONF_DIR=%q\n' "$CONF_DIR"
    printf 'readonly CONF_FILE=%q\n' "$CONF_FILE"
    printf 'readonly INFO_FILE=%q\n' "$INFO_FILE"
    printf 'readonly LINK_FILE=%q\n' "$LINK_FILE"
    printf 'readonly SYSCTL_FILE=%q\n' "$SYSCTL_FILE"
    printf 'readonly MANAGER=%q\n' "$MANAGER"
    printf '%s\n' 'PORT=443' 'PORT_EXPLICIT=0' 'SNI="www.dell.com"' 'SNI_EXPLICIT=0'
    declare -f say ok warn die need_root valid_port valid_sni \
      choose_port choose_sni install_deps install_singbox disable_ipv6 \
      public_ipv4 generate_keys write_config check_config save_node \
      show_node install_manager install_node uninstall_node status_node \
      usage parse_args main_menu main
    printf '%s\n' 'main "$@"'
  } >"$temp_manager"
  install -m 755 "$temp_manager" "$MANAGER"
  rm -f "$temp_manager"
}

show_node() {
  [[ -s "$INFO_FILE" && -s "$LINK_FILE" ]] || die "节点尚未安装"
  # shellcheck disable=SC1090
  source "$INFO_FILE"
  printf '\n服务器 IPv4 : %s\n端口          : %s\nUUID          : %s\nSNI           : %s\n公钥          : %s\nShort ID      : %s\n\n节点链接：\n' \
    "$SERVER_IPV4" "$PORT" "$UUID" "$SNI" "$PUBLIC_KEY" "$SHORT_ID"
  cat "$LINK_FILE"
  if [[ -t 1 ]] && command -v qrencode >/dev/null; then printf '\n'; qrencode -t ANSIUTF8 <"$LINK_FILE"; fi
  printf '\n管理命令：vless\n配置文件：%s\n' "$CONF_FILE"
}

install_node() {
  need_root install
  command -v systemctl >/dev/null || die "系统不支持 systemd"
  choose_port
  choose_sni
  valid_port "$PORT"; valid_sni "$SNI"
  install_deps; install_singbox
  if ss -H -ltn "sport = :$PORT" 2>/dev/null | grep -q .; then die "TCP 端口 $PORT 已被占用"; fi
  disable_ipv6; generate_keys; write_config; check_config; save_node
  install_manager
  systemctl daemon-reload
  systemctl enable sing-box >/dev/null
  systemctl restart sing-box
  sleep 1
  if ! systemctl is-active --quiet sing-box; then
    journalctl -u sing-box --no-pager -n 50 >&2 || true
    die "sing-box 启动失败"
  fi
  ok "VLESS Reality 已安装；监听、DNS、出站和系统均限制为 IPv4"
  warn "请在云服务器安全组/防火墙放行 TCP $PORT，本脚本不自动修改防火墙"
  show_node
}

uninstall_node() {
  need_root uninstall
  if [[ "${YES:-0}" != 1 ]]; then
    read -r -p "确认删除 sing-box 节点和配置？[y/N] " answer
    [[ "$answer" =~ ^[Yy]$ ]] || exit 0
  fi
  systemctl disable --now sing-box >/dev/null 2>&1 || true
  rm -rf "$CONF_DIR"; rm -f "$MANAGER" "$SYSCTL_FILE"
  warn "已卸载；为避免网络中断，本次没有自动重新开启 IPv6"
}

status_node() {
  systemctl status sing-box --no-pager || true
  printf '\nIPv6：all=%s default=%s\n' \
    "$(sysctl -n net.ipv6.conf.all.disable_ipv6 2>/dev/null || printf '?')" \
    "$(sysctl -n net.ipv6.conf.default.disable_ipv6 2>/dev/null || printf '?')"
}

usage() {
  cat <<EOF
sing-box VLESS Reality 纯 IPv4 一键脚本 v$VERSION

安装：bash install.sh install [--port 443] [--sni 域名]

也可以直接运行 bash install.sh 打开交互菜单。

交互安装时可选择：
  1) Dell：www.dell.com
  2) 新加坡虾皮：shopee.sg
  3) 自定义 SNI

使用 --sni 可跳过交互选择，例如：
  bash install.sh install --sni shopee.sg

管理：
  vless            打开交互管理菜单
  vless show       显示链接和二维码
  vless status     查看服务和 IPv6 状态
  vless restart    重启服务
  vless logs       查看实时日志
  vless uninstall  卸载
  vless help       帮助

脚本不会自动修改服务器防火墙。
DNS 固定使用 Cloudflare IPv4：1.1.1.1。
EOF
}

parse_args() {
  while (($#)); do
    case "$1" in
      --port) [[ $# -ge 2 ]] || die "--port 缺少值"; PORT="$2"; PORT_EXPLICIT=1; shift 2 ;;
      --sni) [[ $# -ge 2 ]] || die "--sni 缺少值"; SNI="$2"; SNI_EXPLICIT=1; shift 2 ;;
      *) die "未知参数：$1" ;;
    esac
  done
}

main_menu() {
  [[ -t 0 ]] || { usage; return; }
  while true; do
    printf '\n========================================\n'
    printf ' sing-box VLESS Reality 管理菜单\n'
    printf '========================================\n'
    printf '  1) 安装 VLESS Reality\n'
    printf '  2) 显示节点信息\n'
    printf '  3) 查看运行状态\n'
    printf '  4) 重启 sing-box\n'
    printf '  5) 查看实时日志\n'
    printf '  6) 卸载节点\n'
    printf '  0) 退出\n'
    read -r -p '请选择 [0-6]：' choice
    case "$choice" in
      1) install_node; return ;;
      2) show_node ;;
      3) status_node ;;
      4) need_root restart; systemctl restart sing-box; ok "sing-box 已重启" ;;
      5) printf '按 Ctrl+C 退出日志。\n'; journalctl -u sing-box --output cat -f ;;
      6) uninstall_node; return ;;
      0) return ;;
      *) warn "无效选项，请重新选择" ;;
    esac
  done
}

main() {
  local cmd="${1:-menu}"
  (($# == 0)) || shift
  case "$cmd" in
    menu) main_menu ;;
    install) parse_args "$@"; install_node ;;
    show) show_node ;;
    status) status_node ;;
    restart) need_root restart; systemctl restart sing-box; ok "sing-box 已重启" ;;
    logs) journalctl -u sing-box --output cat -f ;;
    uninstall) [[ "${1:-}" != "--yes" ]] || YES=1; uninstall_node ;;
    help|-h|--help) usage ;;
    *) usage; die "未知命令：$cmd" ;;
  esac
}

main "$@"
