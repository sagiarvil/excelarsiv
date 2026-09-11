#!/usr/bin/env bash
# install_mac.sh — Antigravity macOS Eklenti Kurulum Betiği
# Apple Silicon (M1/M2/M3/M4) ve Intel Mac Destekli

set -e

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "🍎 ANTIGRAVITY MACOS EKLENTİ KURULUMU (Apple Silicon & Intel)"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

# 1. Node.js kontrolü
if ! command -v node &> /dev/null; then
    echo "⚠️ Node.js bulunamadı! Homebrew ile kuruluyor..."
    if ! command -v brew &> /dev/null; then
        echo "❌ Homebrew bulunamadı. Lütfen önce Homebrew kurun:"
        echo '/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"'
        exit 1
    fi
    brew install node
fi

# 2. Kurulumu çalıştır
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
node "$DIR/install_plugins.js"
