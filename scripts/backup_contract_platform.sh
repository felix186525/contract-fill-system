#!/usr/bin/env bash
set -euo pipefail

APP_DIR="${APP_DIR:-/opt/contract-platform}"
DATA_DIR="${DATA_DIR:-/opt/contract-platform-data}"
TEMPLATE_DIR="${TEMPLATE_DIR:-/opt/contract-platform-templates}"
BACKUP_DIR="${BACKUP_DIR:-/opt/contract-platform-backups}"
LOG_DIR="${LOG_DIR:-${APP_DIR}/log}"
RETENTION_DAYS="${RETENTION_DAYS:-30}"

mkdir -p "$BACKUP_DIR" "$LOG_DIR"
chmod 700 "$BACKUP_DIR"

timestamp="$(date +%Y%m%d_%H%M%S)"
archive="${BACKUP_DIR}/contract-platform-full-${timestamp}.tar.gz"
log_file="${LOG_DIR}/backup-${timestamp}.log"

{
  echo "[$(date '+%F %T')] backup started"
  echo "archive=${archive}"

  tar \
    --exclude="${APP_DIR}/output" \
    --exclude="${APP_DIR}/uploads" \
    --exclude="${APP_DIR}/__pycache__" \
    --exclude="${APP_DIR}/log" \
    -czf "$archive" \
    "$APP_DIR" \
    "$DATA_DIR" \
    "$TEMPLATE_DIR" \
    /etc/systemd/system/contract-platform.service \
    /etc/nginx 2>&1

  chmod 600 "$archive"
  sha256sum "$archive" > "${archive}.sha256"
  chmod 600 "${archive}.sha256"

  find "$BACKUP_DIR" -type f \( -name 'contract-platform-full-*.tar.gz' -o -name 'contract-platform-full-*.tar.gz.sha256' \) -mtime +"$RETENTION_DAYS" -delete

  echo "size=$(du -h "$archive" | awk '{print $1}')"
  echo "sha256=$(awk '{print $1}' "${archive}.sha256")"
  echo "[$(date '+%F %T')] backup finished"
} >> "$log_file" 2>&1

