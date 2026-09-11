import os
import sys

def main():
    if len(sys.argv) < 2:
        print("Kullanım: python llms_generator.py <hedef_klasor>")
        sys.exit(1)
        
    target_dir = sys.argv[1]
    
    llms_content = """# Domain
Projeniz Hakkında Genel Bilgi

> Summary
Bu dosya yapay zeka ajanları ve LLM'ler için tasarlanmıştır.

## Mimariler ve Alt Linkler
- [Core Architecture](/llms/core.md)
- [Sayfalar](/llms/pages/index.md)
"""

    llms_dir = os.path.join(target_dir, 'llms')
    llms_pages_dir = os.path.join(llms_dir, 'pages')
    
    try:
        os.makedirs(llms_pages_dir, exist_ok=True)
        
        with open(os.path.join(target_dir, 'llms.txt'), 'w', encoding='utf-8') as f:
            f.write(llms_content)
            
        with open(os.path.join(llms_dir, 'core.md'), 'w', encoding='utf-8') as f:
            f.write('# Core System\nLLM Yönergeleri buraya...')
            
        with open(os.path.join(llms_pages_dir, 'index.md'), 'w', encoding='utf-8') as f:
            f.write('# Pages\nSayfa alt kırılımları...')
            
        print("\n=== SÜPER SEO MOTORU (LLMS HUB OLUŞTURUCU V3.0) ===")
        print(f"[+] BAŞARILI: {os.path.abspath(target_dir)} dizinine `llms.txt` hub ve alt dizinleri oluşturuldu.")
    except Exception as e:
        print(f"[X] HATA: {e}")

if __name__ == "__main__":
    main()
