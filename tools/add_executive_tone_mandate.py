with open('calisma_prorokolu/EXCELARSIV_DECISION_OS_V15_1_RUNTIME_KERNEL.py', 'r', encoding='utf-8') as f:
    content = f.read()

rule_addition = '''
    # MANDATE RULE 254: C-LEVEL CORPORATE TONE & ZERO INFORMAL PHRASING
    # Şablonlarda, sayfa başlıklarında, kılavuz kartlarında veya tablolarda asla '5 yaşında çocuk', 'çocuğa anlatır gibi',
    # 'eli5' veya laubali/gayriresmi eğitim jargonu kullanılamaz. Açıklamalar daima Siemens, McKinsey ve Fortune 500
    # kurumsal standartlarında; 'Yalın Kurumsal Açıklama', 'Yönetici Özeti', 'Teknik Olmayan Berrak İfade' seviyesinde
    # ve C-Level saygınlığında olmalıdır.
'''

if 'MANDATE RULE 254' not in content:
    target = '    # MANDATE RULE 253: INFOGRAPHIC COLOR HARMONY & VECTOR CHART STYLING'
    content = content.replace(target, rule_addition + '\n' + target)

with open('calisma_prorokolu/EXCELARSIV_DECISION_OS_V15_1_RUNTIME_KERNEL.py', 'w', encoding='utf-8') as f:
    f.write(content)

print("Rule 254 (C-Level Corporate Tone) Kernel py dosyasına başarıyla işlendi.")
