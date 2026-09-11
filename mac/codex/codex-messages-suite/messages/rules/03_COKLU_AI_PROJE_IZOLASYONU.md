# 09_MULTI_AI_PROJECT_ISOLATION — ÇOKLU YAPAY ZEKA PROJE İZOLASYONU VE HARİTALANDIRMA PROTOKOLÜ

> **Evrensel Kural:** Farklı Yapay Zekaların (Gemini, Claude, Cursor vb.) Proje Dosya Biçimlerine Tam Uyum.  
> **Temel Standart:** Çakışmasız Birlikte Yaşama (Non-Destructive Coexistence) · Merkezi Hafıza Köprüsü · %100 BOM'suz Saf UTF-8

---

## 🎯 1. ÇOKLU YAPAY ZEKA PROJE KAYIT FARKLILIKLARI

Sistemde çalışan farklı yapay zekalar projeleri kendi mimarilerine göre farklı formatlarda kaydeder:

| Yapay Zeka Ortamı | Proje Kayıt Formatı | Konum |
|---|---|---|
| **Gemini / Antigravity** | UUID JSON Dosyası | `~/.gemini/config/projects/<uuid>.json` |
| **Claude Code / Desktop** | Path Slug Klasörü & `.claude.json` | `~/.claude/projects/C--<path>/` |
| **Yerel Proje Hafızası** | Proje İçi Klasör | `<proje_kökü>/.claude/messages/` veya `/.gemini/` |
| **Genel / Cursor / Generic** | Proje İçi Kurallar | `<proje_kökü>/CLAUDE.md` veya `GEMINI.md` |

---

## 🛡️ 2. İZOLASYON VE KORUMA İLKELERİ

1. **Formatlara Müdahale Yasağı:**
   - Gemini çalışırken Claude'un `.claude/projects/` veya `.claude.json` dosyalarını bozamaz veya kendi formatına zorlayamaz.
   - Claude çalışırken Gemini'nin `~/.gemini/config/projects/<uuid>.json` yapısına zarar veremez.
2. **Merkezi Hafıza Köprüsü (`ai_projects_map.json`):**
   - Her proje hem Gemini UUID'si hem Claude slug'ı ile merkezi haritada eşleştirilir.
   - Hangi yapay zeka devrede olursa olsun, ortak plan (`PLAN.md`), durum (`_OKU.txt`) ve kurallar (`PROJE_TALİMATLARI.md`) `gemini-messages-suite/messages/projects/<proje>/` üzerinden okunur.
3. **Dinamik Çözümleme (Auto-Resolution):**
   - Ajan veya kullanıcı bir projeden bahsettiğinde (`satis`, `C--Laravel Herd-www-satis` veya UUID), `ai_project_resolver.js` motoru arka planda otomatik eşleme yapar ve doğru proje bağlamını getirir.

---

## 🔄 3. ÇOKLU AI EŞLEME MOTORU KULLANIMI

macOS Terminalinden projeler arasındaki canlı köprüyü tazelemek için:

```bash
node ~/.gemini/config/plugins/gemini-messages-suite/scripts/ai_project_resolver.js
```
