# 🍎 macOS Hızlı Kurulum Rehberi (Apple Silicon & Intel)

Bu eklenti paketi **macOS** işletim sistemi (M1, M2, M3, M4 ve Intel Mac) için özel olarak optimize edilmiştir. Tüm komutlar macOS Terminali (Zsh / Bash) ve Homebrew ekosistemine tam uyumludur.

---

## ⚡ 1. Tek Komutla Otomatik Kurulum (En Hızlı)

macOS Terminal uygulamasını açın ve bu klasöre gidip çalıştırın:

```bash
chmod +x install_mac.sh && ./install_mac.sh
```

Bu betik:
- Gerekli Node.js ortamını kontrol eder (yoksa Homebrew ile kurar).
- Eklentileri sisteminize (`~/.gemini/config/plugins/`) anında tescil eder.

---

## 🛠️ 2. Geliştirici Ortamı Kurulumu (Homebrew veya Laravel Herd)

Sisteminizde PHP, Python veya Node.js eksikse aşağıdaki yöntemlerden birini seçin:

### Yöntem A: Homebrew ile CLI Kurulumu (Tavsiye Edilen)
```bash
# Homebrew yoksa kurun:
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# Temel araçları kurun:
brew install node python git php composer
```

### Yöntem B: Görsel Arayüz (Laravel Herd & DBngin)
- **Laravel Herd macOS:** [herd.laravel.com](https://herd.laravel.com) (PHP 8.3/8.4, Nginx ve SSL tek tıkla)
- **DBngin macOS:** [dbngin.com](https://dbngin.com) (MySQL, PostgreSQL ve Redis tek tıkla)

---

## 🔄 3. Çoklu Hesap Değiştirici (Antigravity Switch)

Antigravity çoklu hesap oturumlarını yönetmek ve limit izlerini sıfırlamak için:

```bash
chmod +x switch.sh
./switch.sh list                 # Kayıtlı hesapları listele
./switch.sh switch <isim|mail>   # Hesaba geçiş yap
./switch.sh clean                # Oturumu temizle
```

---

## 🔍 4. Kurulumu Doğrulama

Terminalden eklentilerin aktif olduğunu teyit edin:

```bash
~/.gemini/bin/agy plugin list
```

Tüm 18 Uzman Ajan ve 23 Beceri macOS üzerinde tam kapasiteyle çalışmaya hazırdır!
