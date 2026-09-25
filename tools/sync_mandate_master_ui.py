import openpyxl

master_path = 'calisma_prorokolu/EXCELARSIV_DECISION_OS_V15_1_MASTER_STANDART_UI_LOCKED.xlsx'
wb = openpyxl.load_workbook(master_path)

# KILAVUZ Navigasyon Menüsünün Kusursuzlaştırılması
ws_k = wb['KILAVUZ']
ws_k.views.sheetView[0].showGridLines = False

ws_k['C7'] = "1. SİSTEMİN ÖZÜ & MANTIK"
ws_k['C9'] = "5 yaşındaki bir çocuğun dahi anlayacağı yalınlıkta sistem özeti, kullanım kılavuzu ve finans terimleri sözlüğü."
ws_k['C14'] = "📖 SİSTEM REHBERİ (#00_HAKKINDA_VE_REHBER!A1)"

ws_k['I7'] = "2. YÖNETİCİ KOKPİTİ (PANO)"
ws_k['I9'] = "Banka borçlarını, 30 günlük ödeme baskısını ve refinansman tasarruf potansiyelini tek ekranda izleyin."
ws_k['I14'] = "▶ CANLI PANO (#PANO!A1)"

ws_k['O7'] = "3. VERİ GİRİŞİ & STRES TESTİ"
ws_k['O9'] = "Kredi sözleşmelerini girin; -%20 nakit daralması kriz testini ve A4 kurumsal yönetim raporunu alın."
ws_k['O14'] = "📋 GİRDİLERİ DÜZENLE (#GİRDİLER!A1)"

wb.save(master_path)
print("Master Şablon: KILAVUZ navigasyonu ve mimari standartlar eşitlendi.")
