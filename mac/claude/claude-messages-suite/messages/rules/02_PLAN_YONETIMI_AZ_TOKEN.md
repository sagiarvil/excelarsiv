# 08_PLAN_MANAGEMENT_TOKEN_DISCIPLINE — AZ TOKEN ÇOK İŞ PLAN YÖNETİMİ PROTOKOLÜ

> **Evrensel Kural:** Tüm Eklentiler, Ajanlar ve Görevler İçin Geçerlidir.  
> **Temel Standart:** Karpathy Cerrahi Disiplini · Az Token · Çok İş · Sıfır Dolgu · %100 Deterministik

---

## 🎯 1. PLANLAMA ZORUNLULUĞU (PLAN-FIRST DISCIPLINE)

Gereksiz token tüketimini ve sonsuz hata döngülerini engellemek için, çok adımlı veya kritik mimari değişikliklerde **önce cerrahi plan çıkarılır**:

1. **Küçük / Tek Dosyalık İş:** Plan dosyası açmaya gerek yoktur; doğrudan cerrahi müdahale yapılır ve test kanıtı sunulur.
2. **Çok Katmanlı / Çok Dosyalı İş:** `project-manager` ajanının denetiminde `plan_manager.js` ile plan başlatılır:
   - Hedef ve kısıtlar netleştirilir.
   - Mevcut kod tabanındaki çalışan desenler belirlenir (yeniden kullanım).
   - Dosya bazlı cerrahi görev listesi (Checklist) oluşturulur.
   - 4 Kalite Kapısı tanımlanır.

---

## 💡 2. AZ TOKEN ÇOK İŞ — 5 ALTIN KURAL (TOKEN SAVING RULES)

### Kural 1: Sıfırdan Yazma, Çalışan Deseni Kopyala (Reusability)
- Yeni bir form, controller veya bileşen eklerken sıfırdan üretme.
- Projede benzer çalışan bir dosya bul, onun yapısını kopyalayarak uyarla. Bu yöntem token harcamasını %70 azaltır ve mimari uyumu garanti eder.

### Kural 2: Parçalı ve Odaklı Okuma (Surgical Read)
- Büyük dosyaları (`> 200 satır`) baştan sona okuma.
- Grep (`grep_search`), sembol araması veya satır aralığı (`StartLine/EndLine`) kullanarak sadece hedeflenen fonksiyonu/bloğu incele.
- `vendor`, `node_modules`, `dist` gibi harici bağımlılıkları asla tarama.

### Kural 3: Cerrahi Düzenleme (Surgical Diff)
- Tüm dosya içeriğini baştan yazmak yerine sadece değişmesi gereken satırları (`replace_file_content` veya cerrahi Node.js regex) ile güncelle.
- Yanıt metninde dosyanın tamamını tekrar basma, yalnızca değişen parçayı özetle.

### Kural 4: Akıllı Ajan Dağıtımı (No Redundant Agents)
- Gereksiz yere 5-10 alt ajan açma. Ajan açmak konteksti soğuk başlatır ve token yakar.
- Sıralı ve bağımlı işleri tek ajan sırayla bitirir.
- Alt ajan yalnızca birbirinden tamamen bağımsız modüllerde (ör. Frontend UI + Backend API ayrı ayrı) paralel çalıştırılır. Aynı dosyaya asla iki ajan dokunmaz.

### Kural 5: Sıfır Hata Döngüsü (Zero Loop via Verification)
- Hata yapıp tekrar tekrar düzeltmeye çalışmak en büyük token israfıdır.
- Bunu önlemek için demir kanun:
  1. Yazmadan önce hedef fonksiyonun girdi ve çıktılarını oku.
  2. Yazdıktan sonra derhal sözdizimini kontrol et (`php -l`, `tsc`, `error_checker.js`).
  3. BOM taraması yap (`bom_utf8_scan.js --fix`).
  4. Kanıtı al ve oturumu kapat.

---

## 📋 3. PLAN ŞABLONU VE KALİTE KAPILARI (4/4 PASS)

Her plan aşağıdaki 4 kalite kapısından geçmek zorundadır:

| Kalite Kapısı | Denetim Yöntemi | Kabul Kriteri |
|---|---|---|
| **Kapı 1: Syntax & Linter** | `php -l`, `tsc --noEmit`, `error_checker.js` | 0 Syntax Error |
| **Kapı 2: BOM & DOM Dengesi** | `bom_utf8_scan.js --fix`, DOM tag matcher | BOM-suz UTF-8, dengeli etiketler |
| **Kapı 3: Güvenlik & OWASP** | Sanitization, SQLi parametrik sorgu, XSS kaçışı | Temiz ve güvenli girdi/çıktı |
| **Kapı 4: Somut Kanıt** | Test çıktısı, konsol ekran görüntüsü, terminal logu | Çalıştığına dair kanıt sunumu |

---
*Bu kural seti Antigravity, Gemini ve Claude eklentilerinin ortak planlama anayasasıdır.*
