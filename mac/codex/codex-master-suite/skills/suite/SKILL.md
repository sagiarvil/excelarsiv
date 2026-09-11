---
name: suite
description: >-
  Lists the complete Antigravity & AI Suite catalog. Triggered when the user types
  '/suite' or asks to see all agents, skills, instructions, execution methods, and tool paths.
  Recognized by all AI models (Gemini 3.8/3.7, Claude Sonnet/Opus, GPT, Cursor).
---

# 🎛️ AI SUITE CATALOG & EXECUTION PROTOCOL (/suite)

> **Evrensel Tetikleyici:** Kullanıcı sohbet alanına /suite yazdığında veya 'ajanları listele', 'becerileri listele', 'suite yapısını göster' dediğinde devreye girer.  
> **Model Desteği:** Tüm modeller (Gemini, Claude, GPT, Cursor) tarafından doğrudan tanınır.  
> **Dil Standartı:** %100 Türkçe çıktı zorunludur.

---

## 🎯 AJANIN /suite YANIT TALİMATI

Kullanıcı /suite yazdığında ajan istisnasız şu 6 bölümü eksiksiz ve yapılandırılmış olarak sunar:

1. **🏛️ EVRENSEL STANDARTLAR VE MİMARİ İLKELER:**
   - %100 Türkçe iletişim kuralı (toolAction ve toolSummary dahil).
   - %100 BOM'suz saf UTF-8 standardı.
   - Karpathy Cerrahi Disiplini: Sıfır Dolgu (Zero-Preamble), Az Token · Çok İş.
   - 4 Kalite Kapısı (Syntax PASS, BOM-suz PASS, Güvenlik PASS, Kanıt PASS).
   - Taşınabilir Dinamik Yollar (os.homedir(), $HOME, $HOME).

2. **🤖 16 UZMAN AJAN KATALOĞU (AGENTS):**
   Her ajan için:
   - Adı, Rolü ve Uzmanlık Alanı
   - Çekirdek Talimatı (System Directive)
   - Nasıl Çalıştırıldığı (Subagent / Direct Role / Execution Standard)

3. **⚡ 21 ÇEKİRDEK BECERİ (SKILLS) & SLASH KOMUTLARI:**
   Her beceri için:
   - Beceri Adı & Slash Komutu (/<skill-name>)
   - Tetikleyicileri
   - Operasyonel Talimatları
   - Nasıl Çalıştığı (Tools & Runtime)

4. **🛠️ DETERMINİSTİK CLI ARAÇLARI (TOOLS HUB):**
   - Hata kontrolü (error_checker.js), BOM temizliği (bom_utf8_scan.js), SEO analiz motoru (audit_engine_v3.js), kanca muhafızları (pre_tool_guard.js, ensure_utf8_nobom.js).

5. **🔄 OTONOM GÜNCELLEME VE KENDİNİ GELİŞTİRME (SELF-EVOLUTION):**
   - **Canlı Geliştirme İlkesi:** Proje esnasında eksik araç veya kural tespit edildiğinde eklenti doğrudan güncellenir.
   - **Otomatik Versiyonlama:** Her güncellemede `plugin.json` ve `installed_version.json` sürümü otomatik artırılır (SemVer: patch, minor, major).
   - **Tarihçeli Değişiklik Günlüğü:** Tüm güncellemeler `CHANGELOG.md` dosyasına işlenir.
   - **Çift Yönlü Canlı Senkronizasyon:** Masaüstü repo ile sistem kurulu eklenti dizini anında eşitlenir.
   - **Resmi CLI Tescili:** `node scripts/self_updater.js --patch "..."` komutu çalıştırılarak `agy plugin install` ile yeni sürüm anında tescil edilir.


7. **🧬 OTONOM DİL SENTEZLEME VE AZ TOKEN MİMARİSİ (LANGUAGE SYNTHESIZER):**
   - **Tek Komutla Dil Ekleme:** `node scripts/language_synthesizer.js <dil_adi>` (örn: python, typescript, golang, rust).
   - **Otomatik 3 Katman:** `rules/<lang>-rules.md`, `skills/ag-<lang>-developer/SKILL.md`, `agents/<lang>-developer/agent.md`.
   - **Çift Dilli Mimari (Bilingual Token Optimization):** Ajana/modele sistem talimatları yüksek yoğunluklu İngilizce (minimum token, sıfır halüsinasyon), kullanıcıya çıktılar %100 Türkçe!
   - **Canlı Tescil:** `self_updater.js` ile sürüm otomatik artar, Desktop ve Global eşitlenir, `agy plugin install` ile sisteme kaydedilir.

6. **💻 TERMINAL / CLI ÇALIŞTIRMA KOMUTLARI:**
   - `node scripts/suite.js` -> Tam katalog çıktısı
   - `node scripts/self_updater.js --status` -> Güncelleme durumu ve yetenek havuzu
   - `agy plugin list` -> Kurulu eklentileri doğrulama
