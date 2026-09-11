import sys
import json
import urllib.request
import urllib.error

def main():
    if len(sys.argv) < 4:
        print("Kullanım: python indexnow_pusher.py <host_adresi> <indexnow_key> <url_listesi_txt_yolu>")
        sys.exit(1)

    host = sys.argv[1]
    key = sys.argv[2]
    url_list_path = sys.argv[3]
    
    try:
        with open(url_list_path, 'r', encoding='utf-8') as f:
            urls = [line.strip() for line in f if line.strip()]
    except Exception as e:
        print(f"[X] HATA: Dosya okunamadı - {e}")
        sys.exit(1)

    data = {
        "host": host,
        "key": key,
        "keyLocation": f"https://{host}/{key}.txt",
        "urlList": urls
    }
    
    payload = json.dumps(data).encode('utf-8')
    req = urllib.request.Request(
        "https://api.indexnow.org/indexnow",
        data=payload,
        headers={'Content-Type': 'application/json; charset=utf-8'}
    )

    print("\n=== SÜPER SEO MOTORU (INDEXNOW PUSHER V3.0) ===")
    try:
        with urllib.request.urlopen(req) as response:
            status = response.getcode()
            print(f"[+] API Yanıt Kodu: {status}")
            if status in [200, 202]:
                print(f"[+] BAŞARILI: {len(urls)} URL arama motorlarına iletildi.")
            else:
                print(f"[X] HATA: Gönderim başarısız.")
    except urllib.error.URLError as e:
        print(f"[X] BAĞLANTI HATASI: {e}")

if __name__ == "__main__":
    main()
