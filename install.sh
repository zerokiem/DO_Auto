#!/usr/bin/env bash
set -Eeuo pipefail

RELEASE_VERSION="v1.1.0"
INSTALL_DIR="${DOFFICE_INSTALL_DIR:-${PWD}/DO_Auto}"
SOURCE_URL="https://github.com/zerokiem/DO_Auto/archive/refs/tags/${RELEASE_VERSION}.tar.gz"

say() { printf '%s\n' "$*"; }
fail() { printf 'LOI: %s\n' "$*" >&2; exit 1; }

if [[ ! -f "${INSTALL_DIR}/docker-compose.yml" ]]; then
    if [[ -e "$INSTALL_DIR" && -n "$(find "$INSTALL_DIR" -mindepth 1 -maxdepth 1 -print -quit 2>/dev/null)" ]]; then
        fail "Thu muc da co du lieu nhung khong phai DOffice Auto: ${INSTALL_DIR}"
    fi
    command -v tar >/dev/null 2>&1 || fail "Can lenh tar."
    TEMP_DIR="$(mktemp -d "${TMPDIR:-/tmp}/doffice-install.XXXXXX")"
    trap 'case "${TEMP_DIR:-}" in "${TMPDIR:-/tmp}"/doffice-install.*) rm -rf -- "$TEMP_DIR" ;; esac' EXIT
    say "Tai DOffice Auto ${RELEASE_VERSION} ..."
    if command -v curl >/dev/null 2>&1; then
        curl -fL "$SOURCE_URL" -o "${TEMP_DIR}/source.tar.gz"
    elif command -v wget >/dev/null 2>&1; then
        wget -O "${TEMP_DIR}/source.tar.gz" "$SOURCE_URL"
    else
        fail "Can curl hoac wget."
    fi
    mkdir -p "$INSTALL_DIR"
    tar -xzf "${TEMP_DIR}/source.tar.gz" -C "$INSTALL_DIR" --strip-components=1
fi

cd "$INSTALL_DIR"
[[ -f .env ]] || cp .env.example .env

# Ghi lai duong dan du lieu tuy chon de nhung lan docker compose sau van dung.
if [[ -n "${DOFFICE_DATA_HOST:-}" ]]; then
    grep -v '^DOFFICE_DATA_HOST=' .env > .env.tmp || true
    printf 'DOFFICE_DATA_HOST=%s\n' "$DOFFICE_DATA_HOST" >> .env.tmp
    mv .env.tmp .env
fi

if docker info >/dev/null 2>&1; then
    DOCKER=(docker)
elif command -v sudo >/dev/null 2>&1 && sudo -n docker info >/dev/null 2>&1; then
    DOCKER=(sudo -n docker)
else
    fail "Docker chua cai/chua chay, hoac tai khoan chua co quyen truy cap Docker."
fi

if "${DOCKER[@]}" compose version >/dev/null 2>&1; then
    COMPOSE=("${DOCKER[@]}" compose)
elif command -v docker-compose >/dev/null 2>&1; then
    if docker-compose version >/dev/null 2>&1; then
        COMPOSE=(docker-compose)
    else
        COMPOSE=(sudo -n docker-compose)
    fi
elif [[ -x /usr/local/bin/docker-compose ]]; then
    COMPOSE=(sudo -n /usr/local/bin/docker-compose)
else
    fail "Khong tim thay Docker Compose."
fi

say "Build va khoi dong DOffice Auto ..."
"${COMPOSE[@]}" up -d --build
"${COMPOSE[@]}" ps

PORT="${DOFFICE_PORT:-8877}"
say ""
say "DA CAI XONG: ${INSTALL_DIR}"
say "Dashboard: http://IP-MAY-CHU:${PORT}"
say "May headless/NAS: tao state.json tren Windows roi chep vao:"
say "  ${INSTALL_DIR}/playwright/.auth/state.json"
