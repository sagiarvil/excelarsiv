#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
error_checker.py — Kapsamlı Proje Hata ve Sözdizim Kontrol Aracı (Bug Hunter & QA)

Bu araç belirtilen dizin veya dosyada:
1. PHP Sözdizim Doğrulama (php -l)
2. Python Sözdizim Doğrulama (py_compile)
3. JavaScript / Node.js Sözdizim Doğrulama (node --check)
4. JSON Format Doğrulama (json.loads)
5. UTF-8 BOM Tespiti (\xef\xbb\xbf)
6. Güvenlik ve Mantık Açığı Statik Analizi (eval, XSS, raw SQL)

Kullanım:
    python error_checker.py [hedef_klasor_veya_dosya] [--fix-bom] [--json] [--strict]
"""

import sys
import os
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass
import re
import json
import subprocess
from datetime import datetime

CANONICAL_PHP = r"php"
CANONICAL_NODE = r"node"
CANONICAL_PYTHON = r"python3"

PHP_EXE = CANONICAL_PHP if os.path.exists(CANONICAL_PHP) else "php"
NODE_EXE = CANONICAL_NODE if os.path.exists(CANONICAL_NODE) else "node"
PYTHON_BIN = 'python3'

IGNORED_DIRS = {
    'node_modules', 'vendor', '.git', 'storage', 'cache',
    'dist', 'build', '.vscode', '.idea', 'tmp', 'temp'
}

BUG_PATTERNS = [
    (re.compile(r"eval\s*\(", re.I), "SECURITY", "SEC-EVAL", "eval() kullanımı tespit edildi"),
    (re.compile(r"\bextract\s*\(\s*\$_(GET|POST|REQUEST)", re.I), "SECURITY", "SEC-EXTRACT", "extract($_GET/$_POST) değişken enjeksiyon riski"),
    (re.compile(r"echo\s+\$_(GET|POST|REQUEST)\[[^\]]+\]\s*;", re.I), "SECURITY", "SEC-XSS", "Doğrudan kullanıcı girdisi echo ediliyor (XSS riski)"),
    (re.compile(r"\$_(GET|POST|REQUEST)\[[^\]]+\]\s*\.\s*(['\"][^'\"]*SELECT|INSERT|UPDATE|DELETE)", re.I), "SECURITY", "SEC-SQLI", "Ham SQL sorgusuna girdi birleştirme"),
    (re.compile(r"\b(shell_exec|exec|system|passthru|popen|proc_open)\s*\(", re.I), "SECURITY", "SEC-CMD-EXEC", "Sistem komut yürütücü çağrısı"),
    (re.compile(r"var_dump\s*\(|console\.log\s*\(", re.I), "QUALITY", "QUAL-DEBUG-LEAK", "Debug çıktısı (var_dump/console.log) kalıntısı")
]

def run_checker():
    target = "."
    fix_bom = False
    is_json = False
    is_strict = False

    args = sys.argv[1:]
    for arg in args:
        if arg == "--fix-bom":
            fix_bom = True
        elif arg == "--json":
            is_json = True
        elif arg == "--strict":
            is_strict = True
        elif not arg.startswith("--"):
            target = arg

    resolved_target = os.path.abspath(target)
    if not os.path.exists(resolved_target):
        print(f"❌ HATA: Hedef bulunamadı: {resolved_target}", file=sys.stderr)
        sys.exit(2)

    report = {
        "target": resolved_target,
        "timestamp": datetime.now().isoformat(),
        "scanned_files": 0,
        "stats": {"pass": 0, "fail": 0, "warn": 0, "info": 0},
        "errors": [],
        "warnings": []
    }

    def add_log(status, code, path_str, msg, detail=""):
        if status == "FAIL":
            report["stats"]["fail"] += 1
            report["errors"].append({"code": code, "file": path_str, "message": msg, "detail": detail})
        elif status == "WARN":
            report["stats"]["warn"] += 1
            report["warnings"].append({"code": code, "file": path_str, "message": msg, "detail": detail})
        elif status == "INFO":
            report["stats"]["info"] += 1
        elif status == "PASS":
            report["stats"]["pass"] += 1

    def scan_file(file_path):
        ext = os.path.splitext(file_path)[1].lower()
        if ext not in {'.php', '.js', '.mjs', '.cjs', '.json', '.py', '.html', '.sql'}:
            return

        report["scanned_files"] += 1
        rel_path = os.path.relpath(file_path, resolved_target)

        try:
            with open(file_path, 'rb') as f:
                raw_bytes = f.read()
        except Exception as e:
            add_log("FAIL", "IO-ERR", rel_path, str(e))
            return

        # BOM kontrolü
        if raw_bytes.startswith(b'\xef\xbb\xbf'):
            if fix_bom:
                with open(file_path, 'wb') as f:
                    f.write(raw_bytes[3:])
                raw_bytes = raw_bytes[3:]
                add_log("WARN", "BOM-FIXED", rel_path, "UTF-8 BOM silindi.")
            else:
                add_log("FAIL", "BOM-DETECTED", rel_path, "UTF-8 BOM tespit edildi.")

        try:
            text_content = raw_bytes.decode('utf-8')
        except UnicodeDecodeError:
            try:
                text_content = raw_bytes.decode('utf-8', errors='replace')
                add_log("WARN", "NON-UTF8", rel_path, "Dosya bozuk karakter içeriyor.")
            except Exception:
                text_content = ""

        # PHP Kontrolü
        if ext == '.php':
            cmd = [PHP_EXE, "-l", file_path]
            try:
                res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=10)
                if res.returncode != 0 or "No syntax errors detected" not in res.stdout:
                    err_msg = res.stdout.strip() or res.stderr.strip()
                    add_log("FAIL", "PHP-SYNTAX", rel_path, "PHP Sözdizim Hatası", err_msg)
                else:
                    report["stats"]["pass"] += 1
            except Exception as e:
                add_log("FAIL", "PHP-EXEC-ERR", rel_path, str(e))

        # Node.js Kontrolü
        elif ext in {'.js', '.mjs', '.cjs'}:
            cmd = [NODE_EXE, "--check", file_path]
            try:
                res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=10)
                if res.returncode != 0:
                    add_log("FAIL", "JS-SYNTAX", rel_path, "JS Sözdizim Hatası", res.stderr.strip())
                else:
                    report["stats"]["pass"] += 1
            except Exception as e:
                add_log("FAIL", "JS-EXEC-ERR", rel_path, str(e))

        # Python Kontrolü
        elif ext == '.py':
            cmd = [PYTHON_EXE, "-m", "py_compile", file_path]
            try:
                res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=10)
                if res.returncode != 0:
                    add_log("FAIL", "PY-SYNTAX", rel_path, "Python Sözdizim Hatası", res.stderr.strip())
                else:
                    report["stats"]["pass"] += 1
            except Exception as e:
                add_log("FAIL", "PY-EXEC-ERR", rel_path, str(e))

        # JSON Kontrolü
        elif ext == '.json':
            try:
                json.loads(text_content)
                report["stats"]["pass"] += 1
            except Exception as e:
                add_log("FAIL", "JSON-PARSE", rel_path, f"Geçersiz JSON: {e}")

        # Kalıp Taraması (Güvenlik / Kalite)
        if ext in {'.php', '.js'}:
            for pat, cat, code, msg in BUG_PATTERNS:
                if pat.search(text_content):
                    if cat == "SECURITY":
                        if is_strict:
                            add_log("FAIL", code, rel_path, msg)
                        else:
                            add_log("WARN", code, rel_path, msg)
                    else:
                        add_log("INFO", code, rel_path, msg)

    if os.path.isfile(resolved_target):
        scan_file(resolved_target)
    else:
        for root, dirs, files in os.walk(resolved_target):
            dirs[:] = [d for d in dirs if d not in IGNORED_DIRS]
            for f in files:
                scan_file(os.path.join(root, f))

    if is_json:
        print(json.dumps(report, indent=2, ensure_ascii=False))
    else:
        print("=" * 60)
        print("🔍 HATA VE SÖZDİZİM KONTROL RAPORU (Python Sürümü)")
        print(f"Hedef  : {resolved_target}")
        print(f"Taranan: {report['scanned_files']} dosya")
        print("=" * 60)

        if not report["errors"] and not report["warnings"]:
            print("✅ TEBRİKLER: Hiçbir sözdizim veya kritik hata tespit edilmedi.")
        else:
            if report["errors"]:
                print(f"\n🔴 KRİTİK HATALAR ({len(report['errors'])} adet):")
                for idx, e in enumerate(report["errors"], 1):
                    print(f"  {idx}. [{e['code']}] {e['file']} → {e['message']}")
                    if e.get("detail"):
                        print(f"     Detay: {e['detail'].splitlines()[0]}")

            if report["warnings"]:
                print(f"\n🟡 UYARILAR ({len(report['warnings'])} adet):")
                for idx, w in enumerate(report["warnings"], 1):
                    print(f"  {idx}. [{w['code']}] {w['file']} → {w['message']}")

        print("\n" + "=" * 60)
        print("SONUÇ ÖZETİ:")
        print(f"  BAŞARILI (PASS) : {report['stats']['pass']}")
        print(f"  HATALI   (FAIL) : {report['stats']['fail']}")
        print(f"  UYARI    (WARN) : {report['stats']['warn']}")
        print(f"  BİLGİ    (INFO) : {report['stats']['info']}")
        print("=" * 60)

    sys.exit(1 if report["stats"]["fail"] > 0 else 0)

if __name__ == '__main__':
    run_checker()
