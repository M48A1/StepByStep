#!/bin/bash

# ==================== 颜色定义 ====================
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[0;33m'
BLUE='\033[0;34m'
PURPLE='\033[0;35m'
CYAN='\033[0;36m'
NC='\033[0m'

clear
echo -e "${BLUE}==================================================${NC}"
echo -e "${GREEN}   VLESS-Reality Ultimate Final 版   ${NC}"
echo -e "${BLUE}==================================================${NC}"

if [ "$EUID" -ne 0 ]; then
    echo -e "${RED}错误：请使用 root 权限运行此脚本！${NC}"
    exit 1
fi

XRAY_CONFIG="/usr/local/etc/xray/config.json"

# 检查并删除旧配置
if [ -f "$XRAY_CONFIG" ]; then
    echo -e "${RED}⚠️ 检测到旧配置${NC}"
    read -p "是否删除旧配置继续？(y/n，默认 y): " REMOVE_OLD
    [[ "$REMOVE_OLD" != "n" && "$REMOVE_OLD" != "N" ]] && rm -f "$XRAY_CONFIG"
fi

# ==================== 用户输入 ====================
echo -e "${YELLOW}👉 请输入节点名称 [默认: My_Reality]: ${NC}"
read -p "" NODE_NAME
NODE_NAME=${NODE_NAME:-My_Reality}

echo -e "${YELLOW}👉 请输入伪装 SNI 域名 [默认: www.sony.com]: ${NC}"
read -p "" CUSTOM_SNI
CUSTOM_SNI=${CUSTOM_SNI:-www.sony.com}

echo -e "${YELLOW}👉 请输入端口号 [默认: 443]: ${NC}"
read -p "" CUSTOM_PORT
PORT=${CUSTOM_PORT:-443}

echo -e "${BLUE}--------------------------------------------------${NC}"
echo -e "${GREEN}节点名称 → ${NODE_NAME}${NC}"
echo -e "${GREEN}SNI 域名 → ${CUSTOM_SNI}${NC}"
echo -e "${GREEN}端口     → ${PORT}${NC}"
echo -e "${BLUE}--------------------------------------------------${NC}"

# 安装依赖和 Xray
echo -e "${YELLOW}[1/7] 安装依赖...${NC}"
apt-get update -qq && apt-get install -y curl jq qrencode ufw

echo -e "${YELLOW}[2/7] 安装 Xray...
