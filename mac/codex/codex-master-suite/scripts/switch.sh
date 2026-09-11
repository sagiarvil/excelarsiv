#!/usr/bin/env bash
# switch.sh — Antigravity macOS Çoklu Hesap Değiştirici Betiği
# Apple Silicon (M1/M2/M3/M4) ve Intel Mac Uyumlu

set -e

GEMINI_DIR="$HOME/.gemini"
REG_PATH="$GEMINI_DIR/accounts_registry.json"

function show_help() {
    echo "Antigravity macOS Hesap Yöneticisi"
    echo "Kullanım:"
    echo "  ./switch.sh list                 # Kayıtlı hesapları listele"
    echo "  ./switch.sh capture <isim> <mail> # Aktif oturumu kaydet"
    echo "  ./switch.sh switch <isim|mail>   # Hesaba geçiş yap"
    echo "  ./switch.sh clean                # Oturumu temizle"
}

case "$1" in
    list)
        if [ -f "$REG_PATH" ]; then
            cat "$REG_PATH"
        else
            echo "Kayıtlı hesap bulunamadı."
        fi
        ;;
    clean)
        echo "Oturum temizleniyor..."
        rm -f "$GEMINI_DIR/oauth_creds.json" "$GEMINI_DIR/google_accounts.json"
        echo "Oturum temizlendi."
        ;;
    *)
        show_help
        ;;
esac
