# 📋 KULLANICI KİŞİSEL VE GENEL TALİMATLARI

> **Son Güncelleme:** 2026-09-11  
> **Kapsam:** Tüm Projeler, Kurulumlar ve Betikler İçin Evrensel Geçerli Kullanıcı Direktifleri

---

## 📌 GENEL VE EVRENSEL DİREKTİFLER

| Tarih | Talimat | Hangi Dosyaya Eklendi | Durum |
|---|---|---|---|
| 2026-09-11 | Kurulumlarda ve betiklerde asla sabit kullanıcı dizini (örn: C:\Users\<kullanici_adi>\) yazılmamalıdır; tüm bilgisayarların kullanıcı yapısı farklı olduğundan dinamik kullanıcı değişkenleri (os.homedir(), $HOME, $HOME, $HOME) kullanılmalıdır. | messages/rules/KULLANICI_TALİMATLARI.md | 🟢 Aktif / Zorunlu |

---

## 🎯 DİNAMİK YOL VE TAŞINABİLİRLİK (PORTABILITY) STANDARDI

1. **Node.js Betiklerinde:**
   - Asla sabit ('$HOME/...') kullanıcı yolu yazılamaz.
   - Her zaman 'os.homedir()' veya 'process.env.USERPROFILE || process.env.HOME || os.homedir()' üzerinden dinamik çözümlenir.
   - Masaüstü: path.join(os.homedir(), 'Desktop')
   - Gemini Ana Dizini: path.join(os.homedir(), '.gemini')
   - Claude Ana Dizini: path.join(os.homedir(), '.claude')

2. **macOS Terminal ve Shell Komutlarında:**
   - Sabit kullanıcı adı yerine $HOME veya $HOME kullanılır.
   - Örnek: ~/.gemini/bin/agy veya ~/.gemini/bin/agy

3. **Dokümantasyon ve Kurulum Kılavuzlarında:**
   - Örnek yollarda $HOME, $HOME veya göreceli (relative) taşınabilir yollar kullanılır.
