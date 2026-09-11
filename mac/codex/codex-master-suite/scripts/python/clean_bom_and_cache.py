import os
import sys

def walk_dir(directory, callback):
    for root, dirs, files in os.walk(directory):
        # 'vendor', 'node_modules' ve '.git' dizinlerini hariç tut
        dirs[:] = [d for d in dirs if d not in ('vendor', 'node_modules', '.git')]
        for f in files:
            file_path = os.path.join(root, f)
            callback(file_path)

def main():
    root_dir = '$HOME/Sites/satis'
    count = 0
    target_extensions = ('.php', '.json', '.html', '.js', '.css')
    bom_marker = b'\xef\xbb\xbf'

    def process_file(file_path):
        nonlocal count
        if file_path.lower().endswith(target_extensions):
            try:
                with open(file_path, 'rb') as f:
                    buf = f.read()

                if buf.startswith(bom_marker):
                    clean_buf = buf[len(bom_marker):]
                    with open(file_path, 'wb') as f:
                        f.write(clean_buf)

                    rel_path = file_path.replace(root_dir, '')
                    print(f"[BOM TEMİZLENDİ]: {rel_path}")
                    count += 1
            except Exception as e:
                print(f"[HATA]: {file_path} işlenirken hata: {e}", file=sys.stderr)

    if os.path.exists(root_dir):
        walk_dir(root_dir, process_file)
        print(f"Toplam {count} dosyadaki UTF-8 BOM temizlendi.")
    else:
        print(f"[UYARI]: Kök dizin bulunamadı: {root_dir}")

    # Cache klasörünü boşalt
    cache_dir = '$HOME/Sites/satis/storage/cache'
    if os.path.exists(cache_dir):
        for f in os.listdir(cache_dir):
            fp = os.path.join(cache_dir, f)
            try:
                if os.path.isfile(fp):
                    os.unlink(fp)
            except Exception:
                pass
        print('PageCache (storage/cache) tamamen temizlendi.')

if __name__ == '__main__':
    main()
