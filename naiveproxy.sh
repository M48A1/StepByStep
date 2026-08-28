#!/usr/bin/env bash
# NaiveProxy OneClick v1.0.0 — Debian/Ubuntu x86_64 + systemd
# Public installer only: no account password, SSH key or fixed server domain.
# Downloaded official binaries are verified against GitHub API SHA-256 digests.
set -Eeuo pipefail

case "${1:-}" in
  --help|-h)
    cat <<'NAIVE_HELP'
NaiveProxy 一键安装 v1.0.0
支持：Debian/Ubuntu x86_64、root、systemd、公网 IPv4，80/443 端口空闲。
首次运行交互安装；已安装时进入 naive-manager，不覆盖现有配置。

参数：
  --domain example.com   已直接解析到本机的域名
  --username naive       代理用户名（默认 naive）
  --password-file PATH   从 root 私有文件读取密码；不支持明文命令行密码
  --yes                  非交互确认，须指定域名；默认随机密码
  --check                只检查，不安装依赖或修改服务
  --version              显示安装脚本版本

不停止已有网站，不清空防火墙，不修改 SSH，不自动升级已有部署。
安装后：naive-manager；节点链接：naive-manager link。
NAIVE_HELP
    exit 0
    ;;
  --version) echo '1.0.0'; exit 0 ;;
esac

[[ "$(uname -s)" == Linux && "$(uname -m)" == x86_64 ]] || { echo '仅支持 Linux x86_64。' >&2; exit 1; }
[[ "${EUID}" -eq 0 ]] || { echo '请先以 root 登录 VPS。' >&2; exit 1; }
[[ -r /etc/os-release && -d /run/systemd/system ]] || { echo '需要 Debian/Ubuntu 和 systemd。' >&2; exit 1; }
. /etc/os-release
case "${ID:-}" in debian|ubuntu) ;; *) echo '仅支持 Debian/Ubuntu。' >&2; exit 1 ;; esac
command -v sha256sum >/dev/null || { echo '缺少 sha256sum，无法验证脚本内嵌内容。' >&2; exit 1; }

umask 077
naive_stage_dir=$(mktemp -d /var/tmp/naive-bootstrap.XXXXXX)
trap 'rm -rf -- "$naive_stage_dir"' EXIT
cat > "$naive_stage_dir/naive-manager.py" <<'NAIVE_MANAGER_PAYLOAD'
#!/usr/bin/python3
"""本 VPS 的 NaiveProxy 单站点/单用户管理工具，仅使用 Python 标准库。"""

import argparse
import base64
from datetime import datetime, timezone
import fcntl
import getpass
import grp
import hashlib
import json
import os
from pathlib import Path
import re
import secrets
import shlex
import shutil
import subprocess
import sys
import tempfile
from urllib.parse import quote, urlencode

VERSION = "1.1.0"
AUTH_LINE = re.compile(r"(?m)^(?P<indent>[ \t]*)basic_auth[ \t]+(?P<args>[^\n]+)$")
SITE_LINE = re.compile(r"(?m)^[ \t]*:(\d+)[ \t]*,[ \t]*([A-Za-z0-9.-]+)(?::(\d+))?[ \t]*\{[ \t]*$")
MANAGED_FILES = ("Caddyfile", "client.json", "shadowrocket-http2.txt")


class ManagerError(Exception):
    pass


def credentials_valid(username, password):
    if not re.fullmatch(r"[A-Za-z0-9._-]{1,64}", username):
        raise ManagerError("用户名只允许 1–64 位英文字母、数字、点、下划线和短横线。")
    if not re.fullmatch(r"[A-Za-z0-9_-]{16,128}", password):
        raise ManagerError("密码须为 16–128 位英文字母、数字、下划线或短横线；可直接回车随机生成。")


def parse_config(text):
    auths = list(AUTH_LINE.finditer(text))
    sites = list(SITE_LINE.finditer(text))
    if len(auths) != 1 or len(sites) != 1:
        raise ManagerError("仅支持一个站点和一组 basic_auth；请保留当前 ':443, 域名 {' 结构。未修改配置。")
    try:
        auth = shlex.split(auths[0].group("args"), comments=True)
    except ValueError as exc:
        raise ManagerError("basic_auth 引号不完整。") from exc
    if len(auth) != 2:
        raise ManagerError("basic_auth 应为：basic_auth 用户名 密码。")
    credentials_valid(*auth)
    site = sites[0]
    port = int(site.group(1))
    domain = site.group(2).lower()
    named_port = int(site.group(3) or 443)
    if not 1 <= port <= 65535 or named_port != port:
        raise ManagerError("两个站点地址必须使用同一端口；非 443 示例：:8443, 域名:8443 {")
    labels = domain.split(".")
    if len(domain) > 253 or len(labels) < 2 or any(
        not re.fullmatch(r"[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?", label) for label in labels
    ):
        raise ManagerError("域名格式不正确，请使用普通域名或 punycode。")
    return {"domain": domain, "port": port, "username": auth[0], "password": auth[1]}


def client_config(info):
    user = quote(info["username"], safe="")
    password = quote(info["password"], safe="")
    return {
        "listen": "socks://127.0.0.1:1080",
        "proxy": f'https://{user}:{password}@{info["domain"]}:{info["port"]}',
    }


def shadowrocket_link(info):
    authority = f'{info["username"]}:{info["password"]}@{info["domain"]}:{info["port"]}'
    encoded = base64.b64encode(authority.encode()).decode()
    query = urlencode({"peer": info["domain"], "alpn": "h2", "padding": "1", "allowInsecure": "0"})
    return f"http2://{encoded}?{query}#Naive-HTTP2"


def replace_credentials(text, username, password):
    parse_config(text)
    credentials_valid(username, password)
    return AUTH_LINE.sub(lambda m: f'{m.group("indent")}basic_auth {username} {password}', text)


def validate_adapted(document, info):
    """Use Caddy's own parsed representation to catch extra/imported sites."""
    apps = document.get("apps", {})
    servers = list(apps.get("http", {}).get("servers", {}).values())
    if len(servers) != 1 or servers[0].get("listen") != [f':{info["port"]}']:
        raise ManagerError("解析结果包含额外监听地址或站点；仅支持当前单站点结构。")
    domains = set(apps.get("tls", {}).get("certificates", {}).get("automate", []))
    forwarders = []

    def walk(value):
        if isinstance(value, dict):
            if value.get("handler") == "forward_proxy":
                forwarders.append(value)
            for matcher in value.get("match", []):
                domains.update(matcher.get("host", []))
            for child in value.values():
                walk(child)
        elif isinstance(value, list):
            for child in value:
                walk(child)

    walk(servers)
    if domains != {info["domain"]} or len(forwarders) != 1:
        raise ManagerError("解析结果包含额外域名/代理，或缺少域名自动证书；未应用修改。")
    auth = forwarders[0].get("auth_credentials", [])
    expected = base64.b64encode(base64.b64encode(f'{info["username"]}:{info["password"]}'.encode())).decode()
    if auth != [expected]:
        raise ManagerError("Caddy 解析出的认证信息与导出节点不一致；未应用修改。")


def confirm(message):
    return input(message + " [输入 y 确认，其余取消]：").strip().lower() == "y"


class Manager:
    def __init__(self, config_dir="/etc/naiveproxy", backup_dir="/var/backups/naiveproxy", service_gid=None):
        self.config_dir = Path(config_dir)
        self.backup_dir = Path(backup_dir)
        self.config = self.config_dir / "Caddyfile"
        self.binary = "/opt/naiveproxy/caddy" if Path("/opt/naiveproxy/caddy").is_file() else "/opt/naiveproxy/releases/v2.11.2-naive/caddy"
        self.service = "naiveproxy.service"
        self.uid = os.geteuid()
        self.gid = grp.getgrnam("naiveproxy").gr_gid if service_gid is None else service_gid

    def read(self):
        return self.config.read_text(encoding="utf-8")

    def command(self, args, capture=False):
        return subprocess.run(args, check=True, text=True, stdout=subprocess.PIPE if capture else None)

    def atomic_write(self, path, data, mode, uid=None, gid=None):
        path = Path(path)
        fd, temporary = tempfile.mkstemp(prefix=".naive-manager-", dir=path.parent)
        try:
            with os.fdopen(fd, "wb") as stream:
                os.fchmod(stream.fileno(), mode)
                os.fchown(stream.fileno(), self.uid if uid is None else uid, self.gid if gid is None else gid)
                stream.write(data)
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(temporary, path)
        finally:
            if os.path.exists(temporary):
                os.unlink(temporary)

    def snapshot(self):
        saved = {}
        for name in MANAGED_FILES:
            path = self.config_dir / name
            if path.is_symlink():
                raise ManagerError(f"拒绝处理符号链接：{path}")
            if path.exists():
                stat = path.stat()
                saved[name] = (path.read_bytes(), stat.st_mode & 0o777, stat.st_uid, stat.st_gid)
            else:
                saved[name] = None
        return saved

    def restore_snapshot(self, saved):
        for name, record in saved.items():
            path = self.config_dir / name
            if record is None:
                path.unlink(missing_ok=True)
            else:
                data, mode, uid, gid = record
                self.atomic_write(path, data, mode, uid, gid)

    def backup(self, saved=None):
        saved = self.snapshot() if saved is None else saved
        if not saved.get("Caddyfile"):
            raise ManagerError("找不到主配置，无法备份。")
        if self.backup_dir.is_symlink():
            raise ManagerError("备份目录不能是符号链接。")
        self.backup_dir.mkdir(parents=True, exist_ok=True, mode=0o700)
        self.backup_dir.chmod(0o700)
        stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S.%fZ") + "-" + secrets.token_hex(3)
        folder = self.backup_dir / stamp
        folder.mkdir(mode=0o700)
        manifest = {"version": VERSION, "created_utc": stamp, "files": {}}
        for name, record in saved.items():
            if record is None:
                continue
            data, mode, uid, gid = record
            self.atomic_write(folder / name, data, 0o600, self.uid, os.getegid())
            manifest["files"][name] = {"sha256": hashlib.sha256(data).hexdigest(), "mode": mode, "uid": uid, "gid": gid}
        self.atomic_write(folder / "manifest.json", (json.dumps(manifest, indent=2) + "\n").encode(), 0o600, self.uid, os.getegid())
        print(f"备份已保存：{folder}")
        return folder

    def stage(self, text):
        directory = Path(tempfile.mkdtemp(prefix=".manager-stage-", dir=self.config_dir))
        directory.chmod(0o750)
        os.chown(directory, self.uid, self.gid)
        path = directory / "Caddyfile"
        self.atomic_write(path, text.encode(), 0o640)
        return directory, path

    def validate(self, path):
        prefix = [
            "runuser", "-u", "naiveproxy", "--", "env",
            "XDG_DATA_HOME=/var/lib/naiveproxy/data", "XDG_CONFIG_HOME=/var/lib/naiveproxy/config",
            self.binary,
        ]
        flags = ["--config", str(path), "--adapter", "caddyfile"]
        adapted = self.command(prefix + ["adapt"] + flags, capture=True)
        validate_adapted(json.loads(adapted.stdout), parse_config(path.read_text(encoding="utf-8")))
        self.command(prefix + ["validate"] + flags)

    def active(self):
        return subprocess.run(["systemctl", "is-active", "--quiet", self.service]).returncode == 0

    def reload_service(self):
        self.command(["systemctl", "reload", self.service])
        if not self.active():
            raise ManagerError("重载后服务未处于 active 状态。")

    def sync_client_files(self, info):
        self.atomic_write(self.config_dir / "client.json", (json.dumps(client_config(info), indent=2) + "\n").encode(), 0o600, self.uid, os.getegid())
        self.atomic_write(self.config_dir / "shadowrocket-http2.txt", (shadowrocket_link(info) + "\n").encode(), 0o600, self.uid, os.getegid())

    def apply(self, candidate, expected_original=None):
        if not self.active():
            raise ManagerError("服务当前未运行。请先启动服务再修改，以便验证重载是否成功。")
        if expected_original is not None and self.read() != expected_original:
            raise ManagerError("编辑期间原配置已被其他程序修改；为避免覆盖，已取消。")
        info = parse_config(candidate)
        directory, staged = self.stage(candidate)
        try:
            self.validate(staged)
        finally:
            shutil.rmtree(directory)
        if expected_original is not None and self.read() != expected_original:
            raise ManagerError("检查期间原配置发生变化；已取消，请重新编辑。")
        saved = self.snapshot()
        backup = self.backup(saved)
        try:
            self.atomic_write(self.config, candidate.encode(), 0o640)
            self.sync_client_files(info)
            self.reload_service()
        except BaseException as exc:
            print("应用失败，正在恢复修改前的配置和客户端文件。")
            try:
                self.restore_snapshot(saved)
                self.reload_service()
            except BaseException as rollback_error:
                raise ManagerError(f"自动恢复未完成，请检查服务；备份：{backup}；恢复错误：{rollback_error}") from exc
            raise ManagerError("应用失败，旧文件已恢复并重载。") from exc
        print("配置检查通过并已重载。客户端 JSON 和小火箭链接已同步。")
        print("若更换了账号、域名或端口，请重新导入节点；已发出的旧链接不会自动更新。")
        return backup

    def show(self):
        print(AUTH_LINE.sub(lambda m: m.group("indent") + "basic_auth <用户名已隐藏> <密码已隐藏>", self.read()))
        print("账号密码和节点链接请使用 info 或 link 命令查看。")

    def info(self, link_only=False):
        info = parse_config(self.read())
        if not link_only:
            print("以下信息包含密码，请勿公开截图或转发。")
            for label, key in [("域名", "domain"), ("端口", "port"), ("用户名", "username"), ("密码", "password")]:
                print(f"{label}：{info[key]}")
            print("Shadowrocket：HTTP2；Padding 开；允许不安全关；UoT 关。\n")
        print(shadowrocket_link(info))

    def check(self):
        parse_config(self.read())
        self.validate(self.config)
        print("检查通过；未修改文件、未重载服务。")

    def edit(self):
        original = self.read()
        directory, staged = self.stage(original)
        try:
            print("编辑的是临时副本；保存退出后才会检查并应用。")
            self.command(["nano", str(staged)])
            candidate = staged.read_text(encoding="utf-8")
            if candidate == original:
                print("没有修改。")
            elif confirm("保存并应用修改？"):
                self.apply(candidate, expected_original=original)
        finally:
            shutil.rmtree(directory)

    def change_credentials(self):
        original = self.read()
        info = parse_config(original)
        username = input(f'用户名 [回车保留 {info["username"]}]：').strip() or info["username"]
        password = getpass.getpass("新密码 [回车生成随机密码，输入不显示]：")
        if password:
            if password != getpass.getpass("再输入一次新密码："):
                raise ManagerError("两次密码不一致，未修改。")
        else:
            password = secrets.token_urlsafe(24)
        candidate = replace_credentials(original, username, password)
        if confirm("确认修改账号密码？旧小火箭节点将需要更新"):
            self.apply(candidate, expected_original=original)
            self.info()

    def backups(self):
        if not self.backup_dir.exists():
            return []
        return sorted([p for p in self.backup_dir.iterdir() if p.is_dir() and not p.is_symlink() and (p / "manifest.json").is_file()], reverse=True)

    def restore(self, name=None):
        options = self.backups()
        if not options:
            raise ManagerError("还没有备份。")
        if name is None:
            for index, folder in enumerate(options, 1):
                print(f"{index}. {folder.name}")
            selection = input("输入要恢复的编号 [回车取消]：").strip()
            if not selection:
                return
            if not selection.isdigit() or not 1 <= int(selection) <= len(options):
                raise ManagerError("编号无效。")
            folder = options[int(selection) - 1]
        else:
            matches = [p for p in options if p.name == name]
            if len(matches) != 1:
                raise ManagerError("备份名称无效，只接受 backups 命令列出的名称。")
            folder = matches[0]
        path = folder / "Caddyfile"
        if path.is_symlink():
            raise ManagerError("备份主配置不能是符号链接。")
        data = path.read_bytes()
        manifest = json.loads((folder / "manifest.json").read_text())
        if hashlib.sha256(data).hexdigest() != manifest["files"]["Caddyfile"]["sha256"]:
            raise ManagerError("备份校验失败，拒绝恢复。")
        if confirm(f"恢复备份 {folder.name}？恢复前会另存当前配置"):
            self.apply(data.decode(), expected_original=self.read())

    def dispatch(self, action, argument=None):
        if action == "status":
            self.command(["systemctl", "status", self.service, "--no-pager", "-l"])
        elif action == "show":
            self.show()
        elif action == "info":
            self.info()
        elif action == "link":
            self.info(link_only=True)
        elif action == "check":
            self.check()
        elif action == "backup":
            self.backup()
        elif action == "backups":
            for folder in self.backups():
                print(folder.name)
        elif action == "credentials":
            self.change_credentials()
        elif action == "edit":
            self.edit()
        elif action == "restore":
            self.restore(argument)
        elif action == "reload":
            if confirm("检查并重载磁盘上的当前配置？"):
                current = self.read()
                self.apply(current, expected_original=current)
        elif action == "logs":
            self.command(["journalctl", "-u", self.service, "-n", "50", "--no-pager"])
        elif action in ("start", "stop", "restart"):
            labels = {"start": "启动", "stop": "停止", "restart": "重启"}
            if confirm(f"确认{labels[action]}服务？停止/重启会影响当前代理连接"):
                if action != "stop":
                    self.check()
                self.command(["systemctl", action, self.service])
                print("命令执行成功。")

    def menu(self):
        actions = {
            "1": ("查看服务状态", "status"), "2": ("查看配置（隐藏账号密码）", "show"),
            "3": ("查看账号密码 / 小火箭链接", "info"), "4": ("修改账号密码", "credentials"),
            "5": ("编辑完整配置（临时副本）", "edit"), "6": ("检查配置", "check"),
            "7": ("检查并重载当前配置", "reload"), "8": ("立即备份", "backup"),
            "9": ("选择备份恢复", "restore"), "10": ("查看最近日志", "logs"),
            "11": ("启动服务", "start"), "12": ("停止服务", "stop"), "13": ("重启服务", "restart"),
        }
        while True:
            print(f"\n=== NaiveProxy 管理 v{VERSION} ===")
            for number, (label, _) in actions.items():
                print(f"{number:>2}. {label}")
            print(" 0. 退出")
            choice = input("选择：").strip()
            if choice == "0":
                return
            if choice not in actions:
                print("请输入菜单中的编号。")
                continue
            try:
                self.dispatch(actions[choice][1])
            except (ManagerError, OSError, ValueError, subprocess.CalledProcessError) as exc:
                print(f"错误：{exc}")
            input("按回车返回菜单…")


def main():
    parser = argparse.ArgumentParser(description="本机 NaiveProxy 中文管理菜单；不带参数进入菜单。")
    parser.add_argument("action", nargs="?", choices=["menu", "status", "show", "info", "link", "check", "backup", "backups", "credentials", "edit", "restore", "reload", "logs", "start", "stop", "restart"], default="menu")
    parser.add_argument("backup_name", nargs="?", help="仅 restore 可选：backups 列出的备份目录名称")
    parser.add_argument("--version", action="version", version=VERSION)
    args = parser.parse_args()
    if args.backup_name and args.action != "restore":
        parser.error("备份名称只适用于 restore。")
    if os.geteuid() != 0:
        parser.error("请使用 root 登录，或执行 sudo naive-manager。")
    os.umask(0o077)
    try:
        with open("/run/lock/naive-manager.lock", "a") as lock:
            try:
                fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
            except BlockingIOError:
                raise ManagerError("另一个管理脚本正在运行，请先退出它。")
            manager = Manager()
            if args.action == "menu":
                manager.menu()
            else:
                manager.dispatch(args.action, args.backup_name)
    except (KeyboardInterrupt, EOFError):
        print("\n已取消。")
        return 130
    except (ManagerError, OSError, ValueError, KeyError, subprocess.CalledProcessError) as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
NAIVE_MANAGER_PAYLOAD
cat > "$naive_stage_dir/installer.py" <<'NAIVE_INSTALLER_PAYLOAD'
#!/usr/bin/python3
"""Single-file installer payload. No secrets or machine-specific domain embedded."""
import argparse
import base64
import errno
import getpass
import grp
import hashlib
import ipaddress
import json
import os
from pathlib import Path
import platform
import pwd
import re
import runpy
import secrets
import shutil
import socket
import ssl
import struct
import subprocess
import sys
import tarfile
import tempfile
import time
import urllib.request

VERSION = "1.0.0"
MANAGER_SOURCE = Path(__file__).with_name("naive-manager.py")
RELEASE_API = "https://api.github.com/repos/klzgrad/forwardproxy/releases/latest"


class InstallError(Exception):
    pass


def run(args, **kwargs):
    return subprocess.run(args, check=True, **kwargs)


def fetch(url, limit=128 * 1024 * 1024):
    if not url.startswith("https://"):
        raise InstallError("仅允许 HTTPS 下载。")
    request = urllib.request.Request(url, headers={"User-Agent": "NaiveProxy-OneClick/" + VERSION})
    with urllib.request.urlopen(request, timeout=30) as response:
        if not response.url.startswith("https://"):
            raise InstallError("拒绝重定向到非 HTTPS 地址。")
        data = response.read(limit + 1)
    if len(data) > limit:
        raise InstallError("下载文件超过大小限制。")
    return data


def valid_domain(domain):
    domain = domain.strip().lower()
    labels = domain.split(".")
    if len(labels) < 2 or len(domain) > 253 or any(
        not re.fullmatch(r"[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?", label) for label in labels
    ):
        raise InstallError("请输入域名，不要带 https://、端口或路径；中文域名请使用 punycode。")
    try:
        ipaddress.ip_address(domain)
    except ValueError:
        return domain
    raise InstallError("这里需要域名，不能填写 IP 地址。")


def dns_records(domain):
    answers = {}
    for kind in ("A", "AAAA"):
        data = json.loads(fetch(f"https://dns.google/resolve?name={domain}&type={kind}", 1024 * 1024))
        if data.get("Status") != 0:
            raise InstallError(f"公共 DNS 查询失败：{domain}（{kind}）。")
        number = 1 if kind == "A" else 28
        answers[kind] = sorted({row["data"] for row in data.get("Answer", []) if row.get("type") == number})
    return answers


def public_address(ipv6=False):
    host = "api6.ipify.org" if ipv6 else "api.ipify.org"
    address = str(ipaddress.ip_address(fetch(f"https://{host}", 128).decode().strip()))
    if ipaddress.ip_address(address).version != (6 if ipv6 else 4):
        raise InstallError("公网 IP 检测结果类型不正确。")
    return address


def check_dns(domain):
    records = dns_records(domain)
    ipv4 = public_address()
    print(f"本机公网 IPv4：{ipv4}\n域名 A 记录：{', '.join(records['A']) or '无'}")
    if set(records["A"]) != {ipv4}:
        raise InstallError("域名 A 记录必须仅指向本机公网 IPv4。若使用 Cloudflare，请关闭代理（仅 DNS）。")
    if records["AAAA"]:
        try:
            ipv6 = public_address(ipv6=True)
        except Exception as exc:
            raise InstallError("存在 AAAA 记录，但无法核对本机公网 IPv6。请先修正/删除不适用的 AAAA 记录。") from exc
        if {str(ipaddress.ip_address(a)) for a in records["AAAA"]} != {ipv6}:
            raise InstallError("AAAA 记录与本机公网 IPv6 不一致，请先修正。")
    return ipv4


def check_ports():
    for family in (socket.AF_INET, socket.AF_INET6):
        for kind, port in ((socket.SOCK_STREAM, 80), (socket.SOCK_STREAM, 443), (socket.SOCK_DGRAM, 443)):
            try:
                with socket.socket(family, kind) as sock:
                    if family == socket.AF_INET6:
                        sock.setsockopt(socket.IPPROTO_IPV6, socket.IPV6_V6ONLY, 1)
                    sock.bind(("0.0.0.0" if family == socket.AF_INET else "::", port))
            except OSError as exc:
                if family == socket.AF_INET6 and exc.errno in (errno.EAFNOSUPPORT, errno.EPROTONOSUPPORT, errno.EADDRNOTAVAIL):
                    continue
                transport = "TCP" if kind == socket.SOCK_STREAM else "UDP"
                raise InstallError(f"{transport} {port} 端口不可绑定，可能已有服务占用。不会停止现有服务。") from exc


def choose_release(document):
    tag = document.get("tag_name", "")
    if document.get("draft") or document.get("prerelease") or not re.fullmatch(r"v[0-9]+\.[0-9]+\.[0-9]+-naive", tag):
        raise InstallError("官方版本标签格式发生变化，请人工检查后更新安装脚本。")
    matches = [a for a in document.get("assets", []) if a.get("name") == "caddy-forwardproxy-naive.tar.xz"]
    if len(matches) != 1:
        raise InstallError("没有找到预期的官方 Linux 服务端发布包。")
    asset = matches[0]
    if not re.fullmatch(r"sha256:[0-9a-f]{64}", asset.get("digest") or ""):
        raise InstallError("官方 API 没有提供 SHA-256，拒绝跳过校验。")
    expected = f"https://github.com/klzgrad/forwardproxy/releases/download/{tag}/caddy-forwardproxy-naive.tar.xz"
    if asset.get("browser_download_url") != expected:
        raise InstallError("发布包下载地址与预期官方仓库不一致。")
    return tag, asset


def unpack_verified(archive, asset):
    if hashlib.sha256(archive.read_bytes()).hexdigest() != asset["digest"].split(":", 1)[1]:
        raise InstallError("发布包 SHA-256 校验失败，已停止。")
    files = {}
    with tarfile.open(archive, "r:xz") as tar:
        for name in ("caddy", "LICENSE", "README.md"):
            members = [m for m in tar.getmembers() if m.name == "caddy-forwardproxy-naive/" + name]
            if len(members) != 1 or not members[0].isfile() or members[0].size > 128 * 1024 * 1024:
                raise InstallError("压缩包内容不符合预期，拒绝解包。")
            files[name] = tar.extractfile(members[0]).read()
    data = files["caddy"]
    if len(data) < 64 or data[:4] != b"\x7fELF" or data[4:6] != b"\x02\x01" or struct.unpack_from("<H", data, 18)[0] != 62:
        raise InstallError("服务端程序不是 Linux x86_64 ELF，已停止。")
    return files


def make_caddyfile(info):
    template = '''{
    order forward_proxy before file_server
    admin 127.0.0.1:2019
    log {
        exclude http.log.error
    }
}

:443, DOMAIN {
    tls {
        issuer acme {
            dir https://acme-v02.api.letsencrypt.org/directory
        }
    }
    encode zstd gzip
    forward_proxy {
        basic_auth USERNAME PASSWORD
        hide_ip
        hide_via
        probe_resistance
        acl {
            deny 0.0.0.0/8 10.0.0.0/8 100.64.0.0/10 127.0.0.0/8 169.254.0.0/16 172.16.0.0/12 192.168.0.0/16 224.0.0.0/4 240.0.0.0/4
            deny ::/128 ::1/128 fc00::/7 fe80::/10 ff00::/8
            allow all
        }
    }
    file_server {
        root /var/www/naiveproxy
    }
}
'''
    values = {"DOMAIN": info["domain"], "USERNAME": info["username"], "PASSWORD": info["password"]}
    return re.sub(r"\b(DOMAIN|USERNAME|PASSWORD)\b", lambda match: values[match.group()], template)


UNIT = '''[Unit]
Description=NaiveProxy HTTPS forward proxy
Documentation=https://github.com/klzgrad/naiveproxy
Wants=network-online.target
After=network-online.target

[Service]
Type=notify
User=naiveproxy
Group=naiveproxy
Environment=XDG_DATA_HOME=/var/lib/naiveproxy/data
Environment=XDG_CONFIG_HOME=/var/lib/naiveproxy/config
StateDirectory=naiveproxy
StateDirectoryMode=0750
WorkingDirectory=/var/lib/naiveproxy
ExecStart=/opt/naiveproxy/caddy run --config /etc/naiveproxy/Caddyfile --adapter caddyfile
ExecReload=/opt/naiveproxy/caddy reload --config /etc/naiveproxy/Caddyfile --adapter caddyfile
Restart=on-failure
RestartSec=5s
TimeoutStopSec=15s
LimitNOFILE=1048576
UMask=0027
AmbientCapabilities=CAP_NET_BIND_SERVICE
CapabilityBoundingSet=CAP_NET_BIND_SERVICE
NoNewPrivileges=true
PrivateTmp=true
PrivateDevices=true
ProtectSystem=strict
ProtectHome=true
ProtectKernelTunables=true
ProtectKernelModules=true
ProtectControlGroups=true
RestrictSUIDSGID=true
RestrictRealtime=true
LockPersonality=true
RestrictAddressFamilies=AF_INET AF_INET6 AF_UNIX AF_NETLINK

[Install]
WantedBy=multi-user.target
'''


class Installer:
    def __init__(self, root=Path("/"), manager_source=MANAGER_SOURCE):
        self.root = Path(root)
        self.manager_source = Path(manager_source)
        self.api = runpy.run_path(str(self.manager_source), run_name="embedded_manager")
        self.created = []
        self.created_dirs = []
        self.unit_written = False

    def path(self, absolute):
        return self.root / absolute.lstrip("/")

    def layout_state(self):
        required = ["/etc/naiveproxy/Caddyfile", "/etc/systemd/system/naiveproxy.service", "/usr/local/sbin/naive-manager"]
        exists = [self.path(p).exists() for p in required]
        binary = self.path("/opt/naiveproxy/caddy").is_file() or self.path("/opt/naiveproxy/releases/v2.11.2-naive/caddy").is_file()
        if all(exists) and binary:
            return "installed"
        collisions = required + ["/etc/naiveproxy/client.json", "/etc/naiveproxy/shadowrocket-http2.txt", "/etc/naiveproxy/install-manifest.json", "/opt/naiveproxy/caddy", "/var/www/naiveproxy/index.html"]
        if any(self.path(p).exists() or self.path(p).is_symlink() for p in collisions):
            return "partial"
        return "fresh"

    def check_account(self):
        try:
            user = pwd.getpwnam("naiveproxy")
        except KeyError:
            return
        try:
            group = grp.getgrnam("naiveproxy")
        except KeyError as exc:
            raise InstallError("存在同名用户但没有对应组，请人工检查。") from exc
        if user.pw_uid == 0 or user.pw_dir != "/var/lib/naiveproxy" or not user.pw_shell.endswith("/nologin") or user.pw_gid != group.gr_gid:
            raise InstallError("已有 naiveproxy 用户不符合独立服务账户条件，不会改动它。")

    def ensure_account(self):
        self.check_account()
        try:
            group = grp.getgrnam("naiveproxy")
        except KeyError:
            run(["groupadd", "--system", "naiveproxy"])
            group = grp.getgrnam("naiveproxy")
        try:
            user = pwd.getpwnam("naiveproxy")
        except KeyError:
            run(["useradd", "--system", "--gid", "naiveproxy", "--home-dir", "/var/lib/naiveproxy", "--shell", "/usr/sbin/nologin", "--comment", "NaiveProxy service", "naiveproxy"])
            user = pwd.getpwnam("naiveproxy")
        return user.pw_uid, group.gr_gid

    def directory(self, path, mode, uid=0, gid=0):
        if path.is_symlink():
            raise InstallError(f"目标目录不能是符号链接：{path}")
        if not path.exists():
            path.mkdir(parents=True, mode=mode)
            self.created_dirs.append(path)
        elif path == self.path("/usr/local/sbin"):
            return
        path.chmod(mode)
        os.chown(path, uid, gid)

    def write_new(self, path, data, mode, uid=0, gid=0):
        if path.exists() or path.is_symlink():
            raise InstallError(f"拒绝覆盖已有文件：{path}")
        with path.open("xb") as stream:
            self.created.append(path)
            os.fchmod(stream.fileno(), mode)
            os.fchown(stream.fileno(), uid, gid)
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())

    def rollback(self):
        if self.unit_written:
            stopped = subprocess.run(["systemctl", "disable", "--now", "naiveproxy.service"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            if stopped.returncode and subprocess.run(["systemctl", "is-active", "--quiet", "naiveproxy.service"]).returncode == 0:
                print("无法停止本次创建的服务，保留所有文件以便人工处理。请检查 systemctl status naiveproxy。")
                return
        for path in reversed(self.created):
            path.unlink(missing_ok=True)
        if self.unit_written:
            subprocess.run(["systemctl", "daemon-reload"], check=False)
        for path in reversed(self.created_dirs):
            try:
                path.rmdir()
            except OSError:
                pass
        print("本次创建的程序和配置已撤回。已安装的系统依赖、服务账户以及可能生成的证书状态会保留。")

    def install(self, info, tag, asset, files, archive, work):
        candidate = work / "Caddyfile"
        candidate.write_text(make_caddyfile(info))
        candidate.chmod(0o600)
        binary = work / "caddy"
        binary.write_bytes(files["caddy"])
        binary.chmod(0o755)
        env = dict(os.environ, XDG_DATA_HOME=str(work / "data"), XDG_CONFIG_HOME=str(work / "config"))
        adapted = run([str(binary), "adapt", "--config", str(candidate), "--adapter", "caddyfile"], text=True, capture_output=True, env=env)
        self.api["validate_adapted"](json.loads(adapted.stdout), info)
        run([str(binary), "validate", "--config", str(candidate), "--adapter", "caddyfile"], env=env)
        try:
            uid, gid = self.ensure_account()
            for path, mode, owner, group in [
                ("/opt/naiveproxy", 0o755, 0, 0), ("/opt/naiveproxy/releases", 0o755, 0, 0),
                (f"/opt/naiveproxy/releases/{tag}", 0o755, 0, 0), ("/opt/naiveproxy/downloads", 0o755, 0, 0),
                ("/etc/naiveproxy", 0o750, 0, gid), ("/var/lib/naiveproxy", 0o750, uid, gid),
                ("/var/www/naiveproxy", 0o755, 0, 0), ("/usr/local/sbin", 0o755, 0, 0),
            ]:
                self.directory(self.path(path), mode, owner, group)
            for name, data in files.items():
                self.write_new(self.path(f"/opt/naiveproxy/releases/{tag}/{name}"), data, 0o755 if name == "caddy" else 0o644)
            self.write_new(self.path(f"/opt/naiveproxy/downloads/{tag}.tar.xz"), archive.read_bytes(), 0o644)
            stable = self.path("/opt/naiveproxy/caddy")
            stable.symlink_to(f"releases/{tag}/caddy")
            self.created.append(stable)
            self.write_new(self.path("/etc/naiveproxy/Caddyfile"), candidate.read_bytes(), 0o640, 0, gid)
            self.write_new(self.path("/etc/naiveproxy/client.json"), (json.dumps(self.api["client_config"](info), indent=2) + "\n").encode(), 0o600)
            self.write_new(self.path("/etc/naiveproxy/shadowrocket-http2.txt"), (self.api["shadowrocket_link"](info) + "\n").encode(), 0o600)
            self.write_new(self.path("/usr/local/sbin/naive-manager"), self.manager_source.read_bytes(), 0o755)
            page = b'<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Welcome</title><h1>Welcome</h1><p>This site is being prepared. Please check back later.</p></html>\n'
            self.write_new(self.path("/var/www/naiveproxy/index.html"), page, 0o644)
            self.write_new(self.path("/etc/systemd/system/naiveproxy.service"), UNIT.encode(), 0o644)
            self.unit_written = True
            manifest = {"installer_version": VERSION, "server_release": tag, "server_archive_sha256": asset["digest"], "domain": info["domain"], "created_utc": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())}
            self.write_new(self.path("/etc/naiveproxy/install-manifest.json"), (json.dumps(manifest, indent=2) + "\n").encode(), 0o600)
            run(["systemd-analyze", "verify", "/etc/systemd/system/naiveproxy.service"])
            run(["/usr/local/sbin/naive-manager", "check"])
            run(["systemctl", "daemon-reload"])
            run(["systemctl", "enable", "--now", "naiveproxy.service"])
            run(["systemctl", "is-active", "--quiet", "naiveproxy.service"])
        except BaseException:
            self.rollback()
            raise


def probe_proxy(info, expected_ip):
    context = ssl.create_default_context()
    context.set_alpn_protocols(["http/1.1"])
    with socket.create_connection(("127.0.0.1", 443), timeout=6) as sock:
        with context.wrap_socket(sock, server_hostname=info["domain"]) as tls:
            tls.settimeout(8)
            auth = base64.b64encode(f'{info["username"]}:{info["password"]}'.encode()).decode()
            tls.sendall(f"CONNECT api.ipify.org:80 HTTP/1.1\r\nHost: api.ipify.org:80\r\nProxy-Authorization: Basic {auth}\r\n\r\n".encode())
            headers = b""
            while b"\r\n\r\n" not in headers:
                part = tls.recv(4096)
                if not part or len(headers) > 16384:
                    raise InstallError("代理握手响应不完整。")
                headers += part
            if not headers.startswith(b"HTTP/1.1 200 "):
                raise InstallError("代理认证或 CONNECT 失败。")
            tls.sendall(b"GET / HTTP/1.1\r\nHost: api.ipify.org\r\nConnection: close\r\n\r\n")
            body = b""
            while len(body) < 65536:
                part = tls.recv(4096)
                if not part:
                    break
                body += part
            if not body.startswith(b"HTTP/1.1 200 ") or expected_ip.encode() not in body.split(b"\r\n\r\n", 1)[-1]:
                raise InstallError("代理出口测试没有返回预期公网 IP。")


def wait_ready(info, expected_ip, timeout=90):
    deadline = time.monotonic() + timeout
    last_error = ""
    while time.monotonic() < deadline:
        try:
            probe_proxy(info, expected_ip)
            return True
        except Exception as exc:
            last_error = str(exc)
            print("等待证书签发/代理就绪…", flush=True)
            time.sleep(min(3, max(0, deadline - time.monotonic())))
    print("安装已完成，但尚未通过证书或代理测试：" + last_error)
    print("不要关闭证书验证；检查 DNS 和云防火墙 TCP 80/443，再运行 naive-manager logs。")
    return False


def parser():
    result = argparse.ArgumentParser(description="NaiveProxy 一键安装：Debian/Ubuntu x86_64，固定 443 端口。重复运行不重装。")
    result.add_argument("--domain", help="已直接解析到本机的域名")
    result.add_argument("--username", default="naive", help="默认 naive")
    result.add_argument("--password-file", type=Path, help="从 root 私有文件读取密码，不把密码放入命令行")
    result.add_argument("--yes", action="store_true", help="无需最终确认；首次安装须同时指定 --domain")
    result.add_argument("--check", action="store_true", help="只检查，不安装系统依赖或修改服务")
    return result


def main(argv=None):
    args = parser().parse_args(argv)
    if os.geteuid() != 0 or platform.system() != "Linux" or platform.machine() != "x86_64":
        raise InstallError("仅支持 root 下的 Debian/Ubuntu Linux x86_64。")
    if not Path("/run/systemd/system").is_dir():
        raise InstallError("需要以 systemd 启动的系统，不适用于普通容器。")
    installer = Installer()
    state = installer.layout_state()
    if state == "partial":
        raise InstallError("检测到已有或不完整的同名部署，不会覆盖。请先人工检查 /etc/naiveproxy 和 naiveproxy.service。")
    if state == "installed":
        manager_stat = installer.path("/usr/local/sbin/naive-manager").stat()
        if manager_stat.st_uid != 0 or manager_stat.st_mode & 0o022:
            raise InstallError("已有管理脚本不是 root 独占写入，不会自动执行。")
        print("检测到已安装 NaiveProxy；保留现有配置、账号、证书和程序版本。")
        if args.domain:
            actual = installer.api["parse_config"](installer.path("/etc/naiveproxy/Caddyfile").read_text())
            if valid_domain(args.domain) != actual["domain"]:
                raise InstallError("已有节点域名不同，不会自动修改；请使用 naive-manager edit。")
        if args.check:
            run(["/usr/local/sbin/naive-manager", "check"])
        elif sys.stdin.isatty() and not args.yes:
            run(["/usr/local/sbin/naive-manager"])
        else:
            print("管理命令：naive-manager；查看节点链接：naive-manager link")
        return 0
    check_ports()
    installer.check_account()
    load_state = subprocess.run(["systemctl", "show", "naiveproxy.service", "-p", "LoadState", "--value"], text=True, capture_output=True).stdout.strip()
    if load_state != "not-found":
        raise InstallError("systemd 已有同名服务，但不属于可识别的部署，不会覆盖。")
    if shutil.disk_usage("/").free < 300 * 1024 * 1024:
        raise InstallError("磁盘剩余空间不足 300 MB。")
    if args.check and not args.domain:
        print("系统、目录和 80/443 端口检查通过。可用 --check --domain 你的域名 继续检查 DNS。")
        return 0
    if not args.domain and (args.yes or not sys.stdin.isatty()):
        raise InstallError("非交互安装必须指定 --domain；没有修改服务。")
    domain = valid_domain(args.domain or input("请输入已解析到本机的域名："))
    expected_ip = check_dns(domain)
    if args.check:
        print("DNS 和系统检查通过；没有安装、修改或重载服务。")
        return 0
    password = ""
    if args.password_file:
        stat = args.password_file.stat()
        if stat.st_uid != 0 or stat.st_mode & 0o077 or args.password_file.is_symlink():
            raise InstallError("密码文件必须由 root 拥有、权限 0600/0400，且不能是符号链接。")
        password = args.password_file.read_text().rstrip("\r\n")
        if not password:
            raise InstallError("密码文件为空。")
    elif sys.stdin.isatty() and not args.yes:
        password = getpass.getpass("代理密码 [回车自动生成，输入不显示]：")
        if password and password != getpass.getpass("再次输入代理密码："):
            raise InstallError("两次密码不一致。")
    password = password or secrets.token_urlsafe(24)
    installer.api["credentials_valid"](args.username, password)
    info = {"domain": domain, "port": 443, "username": args.username, "password": password}
    print(f"将安装到 {domain}:443；使用独立服务账户；不会更改 SSH 或防火墙。")
    if not args.yes and input("确认安装？[输入 y 继续]：").strip().lower() != "y":
        print("已取消；未安装代理服务。")
        return 0
    with tempfile.TemporaryDirectory(prefix="naive-install-", dir="/var/tmp") as temporary:
        work = Path(temporary)
        tag, asset = choose_release(json.loads(fetch(RELEASE_API, 4 * 1024 * 1024)))
        print("下载并校验官方服务端：" + tag, flush=True)
        archive = work / "caddy.tar.xz"
        archive.write_bytes(fetch(asset["browser_download_url"]))
        files = unpack_verified(archive, asset)
        installer.install(info, tag, asset, files, archive, work)
    ready = wait_ready(info, expected_ip)
    print("\n管理命令：naive-manager")
    print("小火箭 HTTP2 节点（包含密码，请勿公开）：")
    print(installer.api["shadowrocket_link"](info))
    print("节点已保存到 /etc/naiveproxy/shadowrocket-http2.txt")
    print("Padding 开启；允许不安全关闭；无需安装 MITM 证书。")
    if ready:
        print("证书验证和认证代理出口测试通过。请再用小火箭测试你自己的网络。")
    return 0 if ready else 2


if __name__ == "__main__":
    os.umask(0o077)
    try:
        raise SystemExit(main())
    except (KeyboardInterrupt, EOFError):
        print("\n已取消。", file=sys.stderr)
        raise SystemExit(130)
    except Exception as exc:
        print(f"错误：{exc}", file=sys.stderr)
        raise SystemExit(1)
NAIVE_INSTALLER_PAYLOAD

# Both payloads must be complete and match their build-time digest before any install.
(cd "$naive_stage_dir" && sha256sum --check --status <<'NAIVE_PAYLOAD_HASHES'
9df3f514ed063db1f59aa1f377b44d798a508e27340f6fcfa26b00f83522ead7  naive-manager.py
9205230f00f262301e51f0b28c78018a5ffd34351052bbaf114e6bce7325438b  installer.py
NAIVE_PAYLOAD_HASHES
) || { echo '脚本内容不完整或校验失败，没有开始安装。' >&2; exit 1; }

naive_check_only=0
for naive_arg in "$@"; do
  [[ "$naive_arg" != --check ]] || naive_check_only=1
done
naive_missing=()
for naive_command in python3 nano; do
  command -v "$naive_command" >/dev/null || naive_missing+=("$naive_command")
done
dpkg-query -W -f='${Status}' ca-certificates 2>/dev/null | grep -q 'install ok installed' || naive_missing+=(ca-certificates)
if [[ "${#naive_missing[@]}" -gt 0 ]]; then
  if [[ "$naive_check_only" -eq 1 ]]; then
    echo "缺少依赖：${naive_missing[*]}。检查模式未安装任何软件。" >&2
    exit 1
  fi
  echo "安装所需系统依赖：${naive_missing[*]}"
  apt-get update
  DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends "${naive_missing[@]}"
fi

# Prevent two instances from installing at once; existing manager has its own lock.
exec 9>/run/lock/naive-installer.lock
flock -n 9 || { echo '另一个安装脚本正在运行。' >&2; exit 1; }
python3 "$naive_stage_dir/installer.py" "$@"
