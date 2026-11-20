#!/usr/bin/env bash
set -euo pipefail
# Called only after setup-runtime.ps1 verified this distro belongs to the bundle.
export DEBIAN_FRONTEND=noninteractive
if [ -f /etc/apt/sources.list.d/ubuntu.sources ]; then
  sed -i 's,http://,https://,g' /etc/apt/sources.list.d/ubuntu.sources
fi
apt-get update -qq
apt-get install -y -qq docker.io docker-compose-v2 python3 ca-certificates
systemctl start docker
docker info --format '{{.ServerVersion}}'
