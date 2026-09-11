import sys
import urllib.request
import urllib.error
import re

def main():
    if len(sys.argv) < 2:
        print("Kullanım: python seo_checker.py <url>")
        sys.exit(1)

    url = sys.argv[1]
    
    # Güvenlik için basit bir User-Agent ekliyoruz
    req = urllib.request.Request(
        url, 
        data=None, 
        headers={
            'User-Agent': 'SuperSeoMotorBot/3.0'
        }
    )

    try:
        with urllib.request.urlopen(req) as response:
            data = response.read().decode('utf-8')
            byte_size = len(data.encode('utf-8'))
            
            print("\n=== SÜPER SEO MOTORU (PYTHON V3.0) ===")
            print(f"[+] Hedef URL: {url}")
            
            # 14KB AST Kontrolü (14.336 byte)
            print(f"[+] Sayfa Boyutu: {byte_size} byte")
            if byte_size > 14336:
                print("[!] UYARI (TOKEN-BLOAT-001): 14KB AST bütçesi aşıldı! AI motorları veriyi kesebilir.")
            else:
                print("[+] BAŞARILI: 14KB AST bütçesi sınırları içerisinde.")
                
            # H1 Etiketi Kontrolü
            h1_count = len(re.findall(r'(?i)<h1[\s>]', data))
            print(f"[+] H1 Etiketi Sayısı: {h1_count}")
            if h1_count != 1:
                print("[!] UYARI (TECH-H1-001): Sayfada tam olarak 1 adet H1 olmalıdır.")
                
            # Canonical Kontrolü
            has_canonical = bool(re.search(r'(?i)<link[^>]*rel=["\']canonical["\'][^>]*>', data))
            print(f"[+] Canonical Belirteci: {'VAR' if has_canonical else 'YOK'}")
            if not has_canonical:
                print("[!] UYARI (TECH-CANON-001): Canonical etiketi bulunamadı!")
                
            print("======================================\n")
            
    except urllib.error.URLError as e:
        print(f"[X] HATA: Bağlantı kurulamadı - {e.reason}")
    except Exception as e:
        print(f"[X] HATA: Beklenmeyen bir hata oluştu - {e}")

if __name__ == "__main__":
    main()
