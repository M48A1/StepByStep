# X2rayVlessInstall

```bash
bash <(curl -Ls https://raw.githubusercontent.com/M48A1/StepByStep/refs/heads/main/X2rayVlessInstall.sh)
```





# naiveproxy


```bash
bash <(curl -Ls https://raw.githubusercontent.com/M48A1/StepByStep/refs/heads/main/naiveproxy.sh)
```


# ss2022

```bash
bash <(curl -Ls https://raw.githubusercontent.com/M48A1/StepByStep/refs/heads/main/ss2022.sh)
```

# Snell

在 Linux 服务器上运行（Debian 11+ / Ubuntu 22.04+ / Rocky、AlmaLinux 9+，需要 systemd 247+）：

```bash
sudo bash snell.sh
```

首次运行会先提示选择版本（v5.0.1 / v6.0.0），再输入自定义端口（1024–65535），两项均无默认值。随后从官方源安装，配置开机自启并输出 Surge 节点。

安装后直接输入 `snell` 打开管理菜单，可查看配置、查看状态、重装、启动、停止、重启或查看日志。非 root 用户会自动调用 sudo。已有旧版脚本安装的用户，运行一次新版 `sudo bash snell.sh` 即可添加管理命令，保留原配置。

交互安装还会询问：

- 网络模式：仅 IPv4（关闭 IPv6）、仅 IPv6（关闭 IPv4）、双栈。只影响 Snell 的监听和出站，不修改主机网卡、系统 IPv6 开关或 SSH。仅 IPv6 要求 Snell v6，并绑定安装时检测到的全局 IPv6 地址，使用 Cloudflare/Google 的 IPv6 DNS；地址变化后需调整配置。v5 双栈要求系统 `net.ipv6.bindv6only=0`。
- 是否开启 TCP Fast Open（TFO），默认关闭客户端 TFO。选择开启时，会保留现有内核标志并添加 TFO 客户端/服务端支持，写入 `/etc/sysctl.d/99-snell-tfo.conf` 持久化；安装失败会尝试恢复。选择关闭时不会禁用系统全局 TFO。
- 是否开启 Surge 连接复用 `reuse`，默认关闭。
- 自定义 PSK（输入隐藏，回车生成随机密钥）和服务器公网 IPv4/IPv6/域名（可留空稍后填写）。
- v6 传输模式：`default`（默认，加密与流量整形）或 `unshaped`（加密，关闭流量整形）。服务端与输出节点保持一致。

节点地址、网络模式、TFO 和 reuse 保存在 `/etc/snell/client-options.conf`，`snell info` 会读取这些设置。TFO 和 reuse 是客户端节点参数；脚本不会写入未经确认的服务端 `tfo` 开关。TFO 能否协商成功还取决于客户端和网络路径。

```bash
# 首次安装时自定义端口和节点地址
sudo env SNELL_PORT=34567 SNELL_HOST=proxy.example.com bash snell.sh

# 无交互终端时，必须指定版本和端口（其他选项使用默认值）
sudo env SNELL_VERSION=6.0.0 SNELL_PORT=34567 bash snell.sh

# 可用 SNELL_TFO、SNELL_REUSE、SNELL_MODE 预设对应选项
sudo env SNELL_VERSION=6.0.0 SNELL_PORT=34567 SNELL_TFO=true SNELL_REUSE=true SNELL_MODE=default bash snell.sh

# 网络模式可用 SNELL_NETWORK=ipv4 / ipv6 / dual 预设
sudo env SNELL_VERSION=6.0.0 SNELL_PORT=34567 SNELL_NETWORK=ipv6 bash snell.sh

# 安装后的管理命令
snell             # 打开菜单
snell info        # 查看节点与网络模式
snell config      # 查看完整配置（含密钥）和 systemd 配置
snell status      # 查看状态
snell logs        # 跟踪日志
snell restart     # 重启
snell reinstall   # 重新选择版本、端口、网络等参数后重装
```

重装会在下载和准备完成后，备份到 `/var/backups/snell/时间戳-随机字符/`，再停服务替换文件。备份包含原配置目录、程序、管理命令、systemd 服务和 TFO 配置；`paths.txt` 记录备份编号与原路径，`service-state` 记录原服务状态。重装失败会尝试恢复原文件、内核 TFO 值和服务启用/运行状态，成功后也保留备份。重装会重新询问所有参数，PSK 留空会生成新密钥，需同步更新客户端。

请在云安全组和主机防火墙放行监听端口：v5 需要 TCP/UDP，v6 需要 TCP。脚本会显示防火墙命令；输出中的 `YOUR_SERVER_IP` 需替换为服务器公网 IP 或域名。

配置文件位于 `/etc/snell/snell-server.conf`，修改后执行 `snell restart`。切换网络模式请使用 `snell reinstall`，以便同步更新服务的地址族限制和客户端节点。请将安装脚本保存为文件再运行，以便复制到 `/usr/local/bin/snell`。
版本与兼容性说明见 [Snell 官方文档](https://kb.nssurge.com/surge-knowledge-base/release-notes/snell)。

模拟测试（不需要 root、Linux 或网络）：`PYTHONDONTWRITEBYTECODE=1 python3 -m unittest discover -s tests -v`。
