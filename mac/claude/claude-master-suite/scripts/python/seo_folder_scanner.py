import os
import sys
import re

def analyze_file(file_path):
    byte_size = os.path.getsize(file_path)
    
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
        
    # Vektör 1
    h1_count = len(re.findall(r'(?i)<h1[\s>]', content))
    has_canonical = bool(re.search(r'(?i)<link[^>]*rel=["\']canonical["\'][^>]*>', content))
    has_noindex = bool(re.search(r'(?i)<meta[^>]+name=["\']robots["\'][^>]+content=["\'][^"\']*noindex', content))
    
    # Vektör 4
    has_llms_discovery = bool(re.search(r'(?i)rel=["\']describedby["\'][^>]+href=["\']/llms\.txt["\']', content))
    has_markdown_alt = bool(re.search(r'(?i)rel=["\']alternate["\'][^>]+type=["\']text/markdown["\']', content))
    
    # Vektör 6
    has_data_chunk_id = bool(re.search(r'(?i)data-chunk-id', content))
    h2h3_count = len(re.findall(r'(?i)<(h2|h3)[\s>]', content))
    
    # Vektör 7
    has_wikidata = bool(re.search(r'(?i)wikidata\.org/wiki/Q', content))
    
    return {
        'file': file_path,
        'size': byte_size,
        'h1': h1_count,
        'canonical': has_canonical,
        'noindex': has_noindex,
        'llms_discovery': has_llms_discovery,
        'data_chunk_id': has_data_chunk_id,
        'h2h3': h2h3_count,
        'wikidata': has_wikidata
    }

def scan_directory(target_dir):
    results = []
    for root, dirs, files in os.walk(target_dir):
        for file in files:
            if file.endswith('.html') or file.endswith('.htm'):
                full_path = os.path.join(root, file)
                results.append(analyze_file(full_path))
    return results

def main():
    if len(sys.argv) < 2:
        print("Kullanım: python seo_folder_scanner.py <klasor_yolu>")
        sys.exit(1)
        
    target_dir = sys.argv[1]
    
    print("\n=== SÜPER SEO MOTORU V3.0 (KAPSAMLI KLASÖR TARAMASI) ===")
    print(f"[+] Hedef Dizin: {os.path.abspath(target_dir)}\n")
    
    try:
        report = scan_directory(target_dir)
        
        if not report:
            print("[!] Belirtilen klasörde HTML dosyası bulunamadı.")
            sys.exit(0)
            
        for res in report:
            print(f"📄 Dosya: {res['file']}")
            
            if res['size'] > 14336:
                print(f"   [X] HATA (TOKEN-BLOAT-001): Boyut {res['size']} byte (14KB aşıldı!)")
            else:
                print(f"   [V] PASS: Boyut {res['size']} byte")
                
            print(f"   [V] PASS: Tek H1 (TECH-H1-001)" if res['h1'] == 1 else f"   [X] HATA: H1 Sayısı {res['h1']}")
            print(f"   [V] PASS: Canonical var (TECH-CANON-001)" if res['canonical'] else f"   [X] HATA: Canonical YOK")
            
            if res['noindex']:
                print(f"   [X] KRİTİK HATA (TECH-NOINDEX-001): Accidental noindex tespit edildi!")
                
            print(f"   [V] PASS: llms.txt discovery" if res['llms_discovery'] else f"   [!] UYARI: describedby=llms.txt YOK")
            print(f"   [V] PASS: data-chunk-id" if res['data_chunk_id'] else f"   [!] UYARI: data-chunk-id bulunamadı")
            print(f"   [V] PASS: H2/H3 Matrisi ({res['h2h3']})" if res['h2h3'] >= 3 else f"   [!] UYARI: H2/H3 yapısı zayıf ({res['h2h3']})")
            print(f"   [V] PASS: Wikidata referansı" if res['wikidata'] else f"   [!] UYARI: Wikidata (QID) eksik")
            
            print("-" * 50)
            
    except Exception as e:
        print(f"[X] HATA: {e}")

if __name__ == "__main__":
    main()
