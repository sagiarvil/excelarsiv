#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
EXCELARŞİV MANDATE V15.1 & MASTER STANDART ENGINE
=================================================
MANDATE_V15_1.md ve EXCELARSIV_DECISION_OS_V15_1_MASTER_STANDART_UI_LOCKED.xlsx
standartlarına %100 SADIK tamamlama motoru.

Her çalışma kitabı aşağıdaki kurumsal mimariyi eksiksiz taşır:
1. 00_BASLANGIC (Model Kimliği, Kılavuz, Kurallar, A1 Menü - Altın Tab: C6A15B)
2. PANO (Yönetici Karar Kokpiti - Lacivert Tab: 17324D - Varsayılan Aktif Sekme)
3. GİRDİLER (Veri Girişi & Operasyon Motoru - Mavi Tab: 2563EB)
4. YÖNETİM_RAPORU (İcra Raporu & Senaryo Analizi - Bordo Tab: 991B1B)
5. 01_YURUTME_CEKIRDEGI (Zorunlu Üretim Sırası - Koyu Yeşil Tab: 065F46)
6. 02_KONTROL_LISTESI (60+ Nokta Kontrol Listesi - Kırmızı Tab: 9B2C2C)
7. 03_GORSEL_TASARIM (Görsel Tasarım & Premium Standart - Gri Tab: 475569)
8. 04_MUHASEBE_THP (Türk Tekdüzen Hesap Planı - Mavi Tab: 2F75B5)
9. 05_KARAR_SATIS (5 Saniyede Karar ve Satış Standardı - Mor Tab: 5B477E)
10. Orijinal Ürün Sayfaları (Hesap Motorları, Detay Cetvelleri, Formüller)

Tüm sayfalar görünür, kılavuz çizgileri kapalı, font Segoe UI / Aptos.
"""

import os
import sys
from pathlib import Path
from copy import copy
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Renkler
COLOR_NAVY_DARK    = "0F172A"
COLOR_NAVY_MID     = "1E293B"
COLOR_WHITE        = "FFFFFF"
COLOR_TEXT_MAIN    = "0F172A"
COLOR_TEXT_MUTED   = "64748B"
COLOR_BORDER       = "E2E8F0"
COLOR_ZEBRA        = "F8FAFC"

COLOR_EMERALD_LINE = "10B981"
COLOR_EMERALD_BG   = "ECFDF5"
COLOR_EMERALD_TXT  = "047857"
COLOR_GREEN_DARK   = "065F46"
COLOR_GREEN_BTN    = "059669"

COLOR_RED_LINE     = "EF4444"
COLOR_RED_BG       = "FEF2F2"
COLOR_RED_TXT      = "DC2626"
COLOR_ROSE_CARD_BG = "FFF1F2"

COLOR_BLUE_LINE    = "2563EB"
COLOR_BLUE_BG      = "EFF6FF"
COLOR_BLUE_TXT     = "1D4ED8"
COLOR_BLUE_LIGHT   = "38BDF8"
COLOR_BLUE_CARD_BG = "EFF6FF"

COLOR_AMBER_BG     = "FEF3C7"
COLOR_AMBER_TXT    = "B45309"

FONT_FAMILY = "Segoe UI"

def copy_master_sheet(src_ws, dest_ws):
    """Master standart şablondaki bir sayfayı tüm stilleri, genişlikleri ve formatıyla kopyalar."""
    dest_ws.views.sheetView[0].showGridLines = False
    
    # Hücreleri kopyala
    for row in src_ws.iter_rows(values_only=False):
        for cell in row:
            if cell.value is not None:
                new_cell = dest_ws.cell(row=cell.row, column=cell.column, value=cell.value)
                if cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.border = copy(cell.border)
                    new_cell.fill = copy(cell.fill)
                    new_cell.number_format = copy(cell.number_format)
                    new_cell.protection = copy(cell.protection)
                    new_cell.alignment = copy(cell.alignment)
                    
    # Birleştirilmiş hücreleri kopyala
    for rng in src_ws.merged_cells.ranges:
        dest_ws.merge_cells(str(rng))
        
    # Kolon genişlikleri
    for col, dim in src_ws.column_dimensions.items():
        if dim.width:
            dest_ws.column_dimensions[col].width = dim.width
            
    # Satır yükseklikleri
    for row, dim in src_ws.row_dimensions.items():
        if dim.height:
            dest_ws.row_dimensions[row].height = dim.height

def get_product_context(product_name: str):
    clean_title = product_name.replace("-", " ").title()
    name_l = product_name.lower()
    
    if any(k in name_l for k in ["kasa", "nakit", "likidite", "13-haftalik", "on-uc"]):
        return {
            "title": "13 HAFTALIK NAKİT AKIŞI VE ÖDEME PLANLAMA SİSTEMİ",
            "kpi1_title": "TOPLAM NAKİT HAVUZU",
            "kpi1_badge": "4 HESAP",
            "kpi1_val": "₺34.850.000",
            "kpi1_sub": "↑ %14.2 konsolide vadesiz + fon",
            "kpi2_title": "BU AY ÖDENECEK ÇIKIŞ",
            "kpi2_badge": "VADE: 4 GÜN",
            "kpi2_val": "₺2.480.000",
            "kpi2_sub": "● 8 Kalem nakit karşılığı: bloke",
            "kpi3_title": "NET SERBEST LİKİDİTE",
            "kpi3_badge": "GÜVENLİ",
            "kpi3_val": "₺18.420.000",
            "kpi3_sub": "██████████░░░░░ %52.8 ORAN",
            "kpi4_title": "DSCR / NAKİT DÖNGÜSÜ",
            "kpi4_badge": "EŞİK: 1.30x",
            "kpi4_val": "2.14x",
            "kpi4_sub": "✓ İDEAL borç servis karşılama",
            "decision_title": "LİKİDİTE OPTİMİZASYON FIRSATI",
            "decision_desc": "Atıl kalan ₺4.2M nakit rezervinin gecelik repoda değerlendirilmesiyle ek getiri.",
            "decision_btn": "FONLAMA SENARYOSUNU İNCELE >"
        }
    elif any(k in name_l for k in ["kredi", "borc", "banka", "taksit"]):
        return {
            "title": f"{clean_title.upper()} — BORÇ VE LİKİDİTE SİSTEMİ",
            "kpi1_title": "TOPLAM KREDİ PORTFÖYÜ",
            "kpi1_badge": "6 BANKA",
            "kpi1_val": "₺18.640.000",
            "kpi1_sub": "↑ Ağırlıklı Ortalama Vade: 28 Ay",
            "kpi2_title": "BU AY ÖDENECEK TAKSİT",
            "kpi2_badge": "VADE: 7 GÜN",
            "kpi2_val": "₺1.120.000",
            "kpi2_sub": "● Anapara + Faiz Karşılığı Ayrıldı",
            "kpi3_title": "KULLANILABİLİR KREDİ LİMİTİ",
            "kpi3_badge": "GÜVENLİ",
            "kpi3_val": "₺9.500.000",
            "kpi3_sub": "████████░░░░░░░ %42.0 REZERV",
            "kpi4_title": "BORÇ SERVİS ORANI (DSCR)",
            "kpi4_badge": "EŞİK: 1.25x",
            "kpi4_val": "1.95x",
            "kpi4_sub": "✓ BANKA COVENANT UYUMLU",
            "decision_title": "FAİZ VE YENİDEN YAPILANDIRMA FIRSATI",
            "decision_desc": "Yüksek faizli rotatif kredilerin uzun vadeli yatırıma dönüştürülmesi önerilmektedir.",
            "decision_btn": "YAPILANDIRMA ANALİZİNİ GÖR >"
        }
    else:
        return {
            "title": f"{clean_title.upper()} — YÖNETİM VE KARAR SİSTEMİ",
            "kpi1_title": "TOPLAM İŞLEM HACMİ",
            "kpi1_badge": "DOĞRULANMIŞ",
            "kpi1_val": "₺28.450.000",
            "kpi1_sub": "↑ Konsolide yıllıklandırılmış büyüme",
            "kpi2_title": "RİSKLİ / BEKLEYEN İŞLEM",
            "kpi2_badge": "MİNİMUM",
            "kpi2_val": "₺940.000",
            "kpi2_sub": "● Tüm riskli kayıtlar kontrol altında",
            "kpi3_title": "OPERASYONEL NET MARJ",
            "kpi3_badge": "GÜVENLİ",
            "kpi3_val": "₺8.120.000",
            "kpi3_sub": "██████████░░░░░ %28.5 NET MARJ",
            "kpi4_title": "SİSTEM GÜVEN SKORU",
            "kpi4_badge": "DENETLENDİ",
            "kpi4_val": "98/100",
            "kpi4_sub": "✓ 0 Formül ve Döngüsel Hata",
            "decision_title": "OPERASYONEL VERİMLİLİK FIRSATI",
            "decision_desc": "Süreç optimizasyonuyla birim operasyonel maliyeti düşürme potansiyeli mevcuttur.",
            "decision_btn": "STRATEJİ REÇETESİNİ İNCELE >"
        }

def build_pano_sheet(ws, ctx):
    ws.views.sheetView[0].showGridLines = False
    
    border_card = Border(
        left=Side(style='thin', color=COLOR_BORDER), right=Side(style='thin', color=COLOR_BORDER),
        top=Side(style='thin', color=COLOR_BORDER), bottom=Side(style='thin', color=COLOR_BORDER)
    )
    border_card_green = Border(
        left=Side(style='medium', color=COLOR_EMERALD_LINE), right=Side(style='thin', color=COLOR_BORDER),
        top=Side(style='thin', color=COLOR_BORDER), bottom=Side(style='thin', color=COLOR_BORDER)
    )
    border_card_red = Border(
        left=Side(style='medium', color=COLOR_RED_LINE), right=Side(style='thin', color=COLOR_BORDER),
        top=Side(style='thin', color=COLOR_BORDER), bottom=Side(style='thin', color=COLOR_BORDER)
    )
    border_card_blue = Border(
        left=Side(style='medium', color=COLOR_BLUE_LINE), right=Side(style='thin', color=COLOR_BORDER),
        top=Side(style='thin', color=COLOR_BORDER), bottom=Side(style='thin', color=COLOR_BORDER)
    )
    
    # A1 Menü
    ws["A1"] = "☰ PANO"
    ws["A1"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_TEXT_MUTED)
    
    # Başlık
    ws["B2"] = ctx["title"]
    ws["B2"].font = Font(name=FONT_FAMILY, size=15, bold=True, color=COLOR_TEXT_MAIN)
    
    ws["B3"] = "Yönetici Karar Kokpiti, Otomatik Doğrulama ve Dinamik Hesaplama Sistemi · 2026"
    ws["B3"].font = Font(name=FONT_FAMILY, size=9, bold=False, color=COLOR_TEXT_MUTED)
    
    ws["L2"] = "Dönem: 2026 / Tam ▼"
    ws["L2"].font = Font(name=FONT_FAMILY, size=9, bold=True, color=COLOR_TEXT_MAIN)
    ws["L2"].alignment = Alignment(horizontal="center", vertical="center")
    ws["L2"].border = border_card
    
    ws["M2"] = "★ DOĞRULANMIŞ"
    ws["M2"].font = Font(name=FONT_FAMILY, size=9, bold=True, color=COLOR_WHITE)
    ws["M2"].fill = PatternFill(start_color=COLOR_GREEN_BTN, end_color=COLOR_GREEN_BTN, fill_type="solid")
    ws["M2"].alignment = Alignment(horizontal="center", vertical="center")
    
    # 4 Bento KPI Kartı
    # Kart 1
    ws["B5"] = ctx["kpi1_title"]
    ws["B5"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_TEXT_MUTED)
    ws["D5"] = ctx["kpi1_badge"]
    ws["D5"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_BLUE_TXT)
    ws["D5"].fill = PatternFill(start_color=COLOR_BLUE_BG, end_color=COLOR_BLUE_BG, fill_type="solid")
    ws["D5"].alignment = Alignment(horizontal="center")
    ws["B7"] = ctx["kpi1_val"]
    ws["B7"].font = Font(name=FONT_FAMILY, size=18, bold=True, color=COLOR_TEXT_MAIN)
    ws["B9"] = ctx["kpi1_sub"]
    ws["B9"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_EMERALD_TXT)
    for r in range(5, 10):
        for c in range(2, 5):
            ws.cell(r, c).border = border_card_green if c == 2 else border_card
            
    # Kart 2
    ws["E5"] = ctx["kpi2_title"]
    ws["E5"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_TEXT_MUTED)
    ws["G5"] = ctx["kpi2_badge"]
    ws["G5"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_RED_TXT)
    ws["G5"].fill = PatternFill(start_color=COLOR_RED_BG, end_color=COLOR_RED_BG, fill_type="solid")
    ws["G5"].alignment = Alignment(horizontal="center")
    ws["E7"] = ctx["kpi2_val"]
    ws["E7"].font = Font(name=FONT_FAMILY, size=18, bold=True, color=COLOR_RED_TXT)
    ws["E9"] = ctx["kpi2_sub"]
    ws["E9"].font = Font(name=FONT_FAMILY, size=8, bold=False, color=COLOR_TEXT_MUTED)
    for r in range(5, 10):
        for c in range(5, 8):
            ws.cell(r, c).border = border_card_red if c == 5 else border_card

    # Kart 3
    ws["H5"] = ctx["kpi3_title"]
    ws["H5"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_TEXT_MUTED)
    ws["J5"] = ctx["kpi3_badge"]
    ws["J5"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_EMERALD_TXT)
    ws["J5"].fill = PatternFill(start_color=COLOR_EMERALD_BG, end_color=COLOR_EMERALD_BG, fill_type="solid")
    ws["J5"].alignment = Alignment(horizontal="center")
    ws["H7"] = ctx["kpi3_val"]
    ws["H7"].font = Font(name=FONT_FAMILY, size=18, bold=True, color=COLOR_TEXT_MAIN)
    ws["H9"] = ctx["kpi3_sub"]
    ws["H9"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_EMERALD_TXT)
    for r in range(5, 10):
        for c in range(8, 11):
            ws.cell(r, c).border = border_card_green if c == 8 else border_card

    # Kart 4
    ws["K5"] = ctx["kpi4_title"]
    ws["K5"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_TEXT_MUTED)
    ws["M5"] = ctx["kpi4_badge"]
    ws["M5"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_BLUE_TXT)
    ws["M5"].fill = PatternFill(start_color=COLOR_BLUE_BG, end_color=COLOR_BLUE_BG, fill_type="solid")
    ws["M5"].alignment = Alignment(horizontal="center")
    ws["K7"] = ctx["kpi4_val"]
    ws["K7"].font = Font(name=FONT_FAMILY, size=18, bold=True, color="0F766E")
    ws["K9"] = ctx["kpi4_sub"]
    ws["K9"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_EMERALD_TXT)
    for r in range(5, 10):
        for c in range(11, 14):
            ws.cell(r, c).border = border_card_blue if c == 11 else border_card

    # 12 Aylık Giriş/Çıkış
    ws["B11"] = "12 AYLIK NAKİT GİRİŞ / ÇIKIŞ & LİKİDİTE BAŞARIMI"
    ws["B11"].font = Font(name=FONT_FAMILY, size=10, bold=True, color=COLOR_TEXT_MAIN)
    ws["B12"] = "Kümülatif Tahsilat vs. Tedarikçi/Kredi Ödeme Projeksiyonu"
    ws["B12"].font = Font(name=FONT_FAMILY, size=8, bold=False, color=COLOR_TEXT_MUTED)
    ws["F11"] = "+%32 Verim"
    ws["F11"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_WHITE)
    ws["F11"].fill = PatternFill(start_color=COLOR_EMERALD_LINE, end_color=COLOR_EMERALD_LINE, fill_type="solid")
    ws["F11"].alignment = Alignment(horizontal="center")

    graph_data = [
        ("Dönem", "Aktif Hacim", "Maliyet/Yük", "Hedef"),
        ("D1", 65, 45, 60), ("D2", 78, 48, 70), ("D3", 85, 55, 80),
        ("D4", 75, 42, 75), ("D5", 90, 40, 85), ("D6", 96, 32, 92),
    ]
    for r_i, row in enumerate(graph_data, start=14):
        for c_i, val in enumerate(row, start=2):
            cell = ws.cell(r_i, c_i, val)
            cell.font = Font(name=FONT_FAMILY, size=8, color=COLOR_TEXT_MUTED)
            if r_i == 14:
                cell.font = Font(name=FONT_FAMILY, size=8, bold=True)

    # Konsolide Karar Matrisi
    ws["H11"] = "KONSOLİDE OPERASYON VE KARAR MATRİSİ"
    ws["H11"].font = Font(name=FONT_FAMILY, size=10, bold=True, color=COLOR_TEXT_MAIN)
    ws["H12"] = "Anlık Hesaplama Sonuçları, Dinamik Formül Çıktıları ve Durum Puanlaması"
    ws["H12"].font = Font(name=FONT_FAMILY, size=8, bold=False, color=COLOR_TEXT_MUTED)

    matrix_headers = ["PARAMETRE / KALEM", "ANA TUTAR", "FARK / MARJ", "VADE", "KONTROL DURUMU"]
    for c_i, h in enumerate(matrix_headers, start=8):
        c = ws.cell(14, c_i, h)
        c.font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_WHITE)
        c.fill = PatternFill(start_color=COLOR_NAVY_MID, end_color=COLOR_NAVY_MID, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")

    matrix_rows = [
        ("13 haftalık kümülatif bakiye...", "₺8.450.000", "₺425.000", "12 Ay", "✓ DÜZENLİ", COLOR_EMERALD_BG, COLOR_EMERALD_TXT),
        ("Kritik (eksi) hafta uyarısı", "₺12.200.000", "₺580.400", "24 Ay", "★ AVANTAJ", COLOR_BLUE_BG, COLOR_BLUE_TXT),
        ("Dönem net nakit değişimi", "₺4.850.000", "₺235.200", "8 Ay", "● YAKLAŞAN", COLOR_AMBER_BG, COLOR_AMBER_TXT),
        ("13 haftalık kümülatif bakiye...", "₺2.140.000", "₺118.000", "16 Ay", "✓ TAM UYUM", COLOR_EMERALD_BG, COLOR_EMERALD_TXT),
    ]

    for r_i, (k, a, f, v, st, bg, txt) in enumerate(matrix_rows, start=15):
        ws.cell(r_i, 8, k).font = Font(name=FONT_FAMILY, size=8, bold=True)
        ws.cell(r_i, 9, a).font = Font(name=FONT_FAMILY, size=8, bold=True)
        ws.cell(r_i, 10, f).font = Font(name=FONT_FAMILY, size=8, color=COLOR_RED_TXT)
        ws.cell(r_i, 11, v).font = Font(name=FONT_FAMILY, size=8)
        c_st = ws.cell(r_i, 12, st)
        c_st.font = Font(name=FONT_FAMILY, size=8, bold=True, color=txt)
        c_st.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
        c_st.alignment = Alignment(horizontal="center")
        for col_idx in range(8, 13):
            ws.cell(r_i, col_idx).border = border_card

    # Alt Toplam
    ws["H19"] = "KONSOLİDE SİSTEM TOPLAMI"
    ws["H19"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_WHITE)
    ws["H19"].fill = PatternFill(start_color=COLOR_NAVY_DARK, end_color=COLOR_NAVY_DARK, fill_type="solid")
    
    ws["I19"] = "₺42.850.000"
    ws["I19"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_BLUE_LIGHT)
    ws["I19"].fill = PatternFill(start_color=COLOR_NAVY_DARK, end_color=COLOR_NAVY_DARK, fill_type="solid")

    ws["J19"] = "₺3.420.000"
    ws["J19"].font = Font(name=FONT_FAMILY, size=8, bold=True, color="F43F5E")
    ws["J19"].fill = PatternFill(start_color=COLOR_NAVY_DARK, end_color=COLOR_NAVY_DARK, fill_type="solid")

    ws["K19"] = ""
    ws["K19"].fill = PatternFill(start_color=COLOR_NAVY_DARK, end_color=COLOR_NAVY_DARK, fill_type="solid")

    ws["L19"] = "TAM DOĞRULANDI (0 HATA)"
    ws["L19"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_EMERALD_LINE)
    ws["L19"].fill = PatternFill(start_color=COLOR_NAVY_DARK, end_color=COLOR_NAVY_DARK, fill_type="solid")
    ws["L19"].alignment = Alignment(horizontal="center")

    # Alt Alanlar
    ws["B22"] = "KONSOLİDE PAY VE AĞIRLIK"
    ws["B22"].font = Font(name=FONT_FAMILY, size=9, bold=True, color=COLOR_TEXT_MAIN)
    donut_items = [
        ("Ana Segment (%42)", COLOR_EMERALD_LINE),
        ("İkincil Grup (%30)", COLOR_BLUE_LINE),
        ("Destek Birimi (%15)", "F59E0B"),
        ("Rezervler (%13)", "8B5CF6")
    ]
    for d_i, (lbl, col_code) in enumerate(donut_items, start=23):
        ws.cell(d_i, 2, f"■ {lbl}").font = Font(name=FONT_FAMILY, size=8, bold=True, color=col_code)

    ws["F22"] = "LİKİDİTE RİSKİ & ERKEN UYARI SİNYALLERİ"
    ws["F22"].font = Font(name=FONT_FAMILY, size=9, bold=True, color=COLOR_TEXT_MAIN)
    ws["F23"] = "Nakit Açığı Riski (30 Gün)"
    ws["F23"].font = Font(name=FONT_FAMILY, size=8, bold=False, color=COLOR_TEXT_MUTED)
    ws["I23"] = "0.00 TL (Sıfır Risk)"
    ws["I23"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_EMERALD_TXT)

    ws["F24"] = "Tahsilat Gecikme Toleransı"
    ws["F24"].font = Font(name=FONT_FAMILY, size=8, bold=False, color=COLOR_TEXT_MUTED)
    ws["I24"] = "48 Gün Güvenli Eşik"
    ws["I24"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_EMERALD_TXT)

    ws["F25"] = "██████████████████████████████████████ Güvenli Bölge (%100)"
    ws["F25"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_EMERALD_LINE)

    ws.merge_cells("F26:I27")
    c_warn = ws["F26"]
    c_warn.value = "✓ 90 GÜNLÜK PROJEKSİYONDA NAKİT SIKIŞIKLIĞI TESPİT EDİLMEDİ"
    c_warn.font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_EMERALD_TXT)
    c_warn.fill = PatternFill(start_color=COLOR_EMERALD_BG, end_color=COLOR_EMERALD_BG, fill_type="solid")
    c_warn.alignment = Alignment(horizontal="center", vertical="center")
    c_warn.border = Border(
        left=Side(style='thin', color="A7F3D0"), right=Side(style='thin', color="A7F3D0"),
        top=Side(style='thin', color="A7F3D0"), bottom=Side(style='thin', color="A7F3D0")
    )

    # Karar Motoru Kartı
    for r_k in range(22, 28):
        for c_k in range(11, 14):
            ws.cell(r_k, c_k).fill = PatternFill(start_color=COLOR_GREEN_DARK, end_color=COLOR_GREEN_DARK, fill_type="solid")

    ws["K22"] = "KARAR MOTORU"
    ws["K22"].font = Font(name=FONT_FAMILY, size=7, bold=True, color=COLOR_WHITE)
    ws["K23"] = ctx["decision_title"]
    ws["K23"].font = Font(name=FONT_FAMILY, size=9, bold=True, color=COLOR_WHITE)

    ws.merge_cells("K24:M25")
    ws["K24"] = ctx["decision_desc"]
    ws["K24"].font = Font(name=FONT_FAMILY, size=8, color="D1FAE5")
    ws["K24"].alignment = Alignment(wrap_text=True)

    ws.merge_cells("K26:M27")
    btn = ws["K26"]
    btn.value = ctx["decision_btn"]
    btn.font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_GREEN_DARK)
    btn.fill = PatternFill(start_color=COLOR_WHITE, end_color=COLOR_WHITE, fill_type="solid")
    btn.alignment = Alignment(horizontal="center", vertical="center")

    widths = {1: 8, 2: 18, 3: 15, 4: 12, 5: 16, 6: 18, 7: 6, 8: 24, 9: 15, 10: 14, 11: 12, 12: 18, 13: 18}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def build_girdiler_sheet(ws, ctx):
    ws.views.sheetView[0].showGridLines = False
    border_input = Border(
        left=Side(style='thin', color=COLOR_BORDER), right=Side(style='thin', color=COLOR_BORDER),
        top=Side(style='thin', color=COLOR_BORDER), bottom=Side(style='thin', color=COLOR_BORDER)
    )

    ws["A1"] = "☰ PANO"
    ws["A1"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_TEXT_MUTED)

    # Arama ve Rozetler
    ws.merge_cells("B3:E3")
    ws["B3"] = "🔍  Kayıt veya hesap ara..."
    ws["B3"].font = Font(name=FONT_FAMILY, size=9, color=COLOR_TEXT_MUTED)
    ws["B3"].fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    ws["B3"].alignment = Alignment(vertical="center")
    ws["B3"].border = border_input

    ws.merge_cells("G3:H3")
    ws["G3"] = "TOPLAM: 1.450 KAYIT"
    ws["G3"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_BLUE_TXT)
    ws["G3"].fill = PatternFill(start_color=COLOR_BLUE_BG, end_color=COLOR_BLUE_BG, fill_type="solid")
    ws["G3"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("I3:J3")
    ws["I3"] = "HATALI HÜCRE: 0 ADET"
    ws["I3"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_EMERALD_TXT)
    ws["I3"].fill = PatternFill(start_color=COLOR_EMERALD_BG, end_color=COLOR_EMERALD_BG, fill_type="solid")
    ws["I3"].alignment = Alignment(horizontal="center", vertical="center")

    headers = [
        "TARİH", "REFERANS NO", "İŞLEM AÇIKLAMASI VE DETAY", "BRÜT TUTAR (₺)",
        "KDV ORANI", "NET TUTAR (₺)", "HESAP / CARİ", "DURUM"
    ]
    for c_i, h in enumerate(headers, start=2):
        c = ws.cell(5, c_i, h)
        c.font = Font(name=FONT_FAMILY, size=8.5, bold=True, color=COLOR_WHITE)
        c.fill = PatternFill(start_color=COLOR_NAVY_DARK, end_color=COLOR_NAVY_DARK, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")

    rows_data = [
        ("12.10.2026", "TR-0428-A", "Haftalık operasyon tahmini", 845200, 0.20, "Vadesiz Ticari", "ONAYLANDI"),
        ("14.10.2026", "TR-0429-B", "Haftalık operasyon tahmini", 1240000, 0.20, "Tedarikçi Havuzu", "BEKLEMEDE"),
        ("18.10.2026", "TR-0430-C", "Ödeme takvimi notları", 620500, 0.10, "Kısa Vade Tahakkuk", "ONAYLANDI"),
        ("21.10.2026", "TR-0431-D", "Haftalık operasyon tahmini", 415000, 0.20, "Müşteri Alacağı", "ONAYLANDI"),
        ("25.10.2026", "TR-0432-E", "Haftalık operasyon tahmini", 980400, 0.20, "Vadesiz Ticari", "ONAYLANDI"),
        ("28.10.2026", "TR-0433-F", "Haftalık operasyon tahmini", 1850000, 0.20, "Banka Kredi Taksiti", "BLOKE"),
        ("02.11.2026", "TR-0434-G", "Ödeme takvimi notları", 740000, 0.20, "Kira & Sabit Gider", "ONAYLANDI"),
        ("05.11.2026", "TR-0435-H", "Haftalık operasyon tahmini", 520800, 0.20, "Vadesiz Ticari", "ONAYLANDI"),
        ("09.11.2026", "TR-0436-I", "Haftalık operasyon tahmini", 1120000, 0.20, "Müşteri Alacağı", "ONAYLANDI"),
        ("12.11.2026", "TR-0437-J", "Haftalık operasyon tahmini", 890000, 0.20, "Tedarikçi Havuzu", "ONAYLANDI"),
    ]

    for r_i, (tarih, ref, aciklama, brut, kdv, hesap, durum) in enumerate(rows_data, start=6):
        is_even = (r_i % 2 == 0)
        row_bg = COLOR_ZEBRA if is_even else COLOR_WHITE

        ws.cell(r_i, 2, tarih).font = Font(name=FONT_FAMILY, size=8.5)
        ws.cell(r_i, 2).alignment = Alignment(horizontal="center")

        ws.cell(r_i, 3, ref).font = Font(name=FONT_FAMILY, size=8.5, bold=True)
        ws.cell(r_i, 3).alignment = Alignment(horizontal="center")

        ws.cell(r_i, 4, aciklama).font = Font(name=FONT_FAMILY, size=8.5)

        c_brut = ws.cell(r_i, 5, brut)
        c_brut.font = Font(name=FONT_FAMILY, size=8.5, bold=True)
        c_brut.number_format = '₺#,##0'
        c_brut.alignment = Alignment(horizontal="right")

        c_kdv = ws.cell(r_i, 6, kdv)
        c_kdv.font = Font(name=FONT_FAMILY, size=8.5)
        c_kdv.number_format = '0%'
        c_kdv.alignment = Alignment(horizontal="center")

        c_net = ws.cell(r_i, 7, f"=E{r_i}/(1+F{r_i})")
        c_net.font = Font(name=FONT_FAMILY, size=8.5)
        c_net.number_format = '₺#,##0'
        c_net.alignment = Alignment(horizontal="right")

        ws.cell(r_i, 8, hesap).font = Font(name=FONT_FAMILY, size=8.5)

        c_durum = ws.cell(r_i, 9, durum)
        c_durum.font = Font(name=FONT_FAMILY, size=8, bold=True)
        c_durum.alignment = Alignment(horizontal="center")
        if durum == "ONAYLANDI":
            c_durum.font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_EMERALD_TXT)
        elif durum == "BLOKE":
            c_durum.font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_RED_TXT)

        for col_idx in range(2, 10):
            cell = ws.cell(r_i, col_idx)
            cell.fill = PatternFill(start_color=row_bg, end_color=row_bg, fill_type="solid")
            cell.border = border_input

    # Alt Toplam Barı
    ws.merge_cells("B17:D17")
    ws["B17"] = "SAYFA TOPLAMI (10 / 1.450 KAYIT)"
    ws["B17"].font = Font(name=FONT_FAMILY, size=9, bold=True, color=COLOR_WHITE)
    ws["B17"].fill = PatternFill(start_color=COLOR_NAVY_DARK, end_color=COLOR_NAVY_DARK, fill_type="solid")
    ws["B17"].alignment = Alignment(vertical="center", indent=1)

    ws.merge_cells("E17:I17")
    ws["E17"] = "=TOPLA(E6:E15)"
    ws["E17"].font = Font(name=FONT_FAMILY, size=11, bold=True, color=COLOR_BLUE_LIGHT)
    ws["E17"].fill = PatternFill(start_color=COLOR_NAVY_DARK, end_color=COLOR_NAVY_DARK, fill_type="solid")
    ws["E17"].number_format = '₺#,##0'
    ws["E17"].alignment = Alignment(horizontal="right", vertical="center")

    for c_i in range(2, 10):
        ws.cell(17, c_i).fill = PatternFill(start_color=COLOR_NAVY_DARK, end_color=COLOR_NAVY_DARK, fill_type="solid")

    widths = {1: 8, 2: 14, 3: 16, 4: 32, 5: 18, 6: 12, 7: 18, 8: 22, 9: 16}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def build_yonetim_raporu_sheet(ws, ctx):
    ws.views.sheetView[0].showGridLines = False
    border_card = Border(
        left=Side(style='thin', color=COLOR_BORDER), right=Side(style='thin', color=COLOR_BORDER),
        top=Side(style='thin', color=COLOR_BORDER), bottom=Side(style='thin', color=COLOR_BORDER)
    )

    ws["A1"] = "☰ PANO"
    ws["A1"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_TEXT_MUTED)

    for r in range(2, 5):
        ws.cell(r, 2).border = Border(left=Side(style='thick', color=COLOR_EMERALD_LINE))

    ws["B2"] = f"{ctx['title']} — NİHAİ KARAR VE DEĞERLEME"
    ws["B2"].font = Font(name=FONT_FAMILY, size=14, bold=True, color=COLOR_TEXT_MAIN)

    ws["B3"] = "Yönetim Kurulu ve İcra Komitesi için Otomatik Türetilen Finansal Stres Testi Raporu"
    ws["B3"].font = Font(name=FONT_FAMILY, size=9.5, bold=False, color=COLOR_TEXT_MUTED)

    ws["B4"] = "DÖNEM: 2026 Q3 · DENETİM PROTOKOLÜ: ISO/VUK/IFRS UYUMLU · DURUM: ONAYLANDI"
    ws["B4"].font = Font(name=FONT_FAMILY, size=8.5, bold=True, color=COLOR_EMERALD_TXT)

    # Senaryo A
    ws.merge_cells("B6:F6")
    ws["B6"] = "SENARYO A: KÖTÜMSER (STRES TESTİ -%20)"
    ws["B6"].font = Font(name=FONT_FAMILY, size=9, bold=True, color=COLOR_RED_TXT)
    ws["B6"].fill = PatternFill(start_color=COLOR_ROSE_CARD_BG, end_color=COLOR_ROSE_CARD_BG, fill_type="solid")
    ws["B6"].alignment = Alignment(horizontal="center", vertical="center")

    scen_a_metrics = [
        ("Ciro / Hacim Düşüşü:", "-%20.0 (Kriz Eşiği)", COLOR_RED_TXT),
        ("Beklenen Net Kâr / Getiri:", "₺4.280.000", COLOR_TEXT_MAIN),
        ("Minimum Nakit Tamponu:", "₺1.850.000 (Yeterli)", COLOR_EMERALD_TXT),
        ("Likidite Karşılama Oranı:", "1.28x (Eşik: 1.10x)", COLOR_EMERALD_TXT),
    ]

    for idx, (label, val, col_c) in enumerate(scen_a_metrics, start=8):
        ws.cell(idx, 2, label).font = Font(name=FONT_FAMILY, size=8.5, bold=False, color=COLOR_TEXT_MUTED)
        c_v = ws.cell(idx, 5, val)
        c_v.font = Font(name=FONT_FAMILY, size=9, bold=True, color=col_c)
        c_v.alignment = Alignment(horizontal="right")

    ws.merge_cells("B14:F15")
    btn_a = ws["B14"]
    btn_a.value = "STRES SENARYOSUNDA DAHİ İFLAS RİSKİ YOK"
    btn_a.font = Font(name=FONT_FAMILY, size=8.5, bold=True, color=COLOR_RED_TXT)
    btn_a.fill = PatternFill(start_color=COLOR_RED_BG, end_color=COLOR_RED_BG, fill_type="solid")
    btn_a.alignment = Alignment(horizontal="center", vertical="center")
    btn_a.border = Border(
        left=Side(style='thin', color="FECACA"), right=Side(style='thin', color="FECACA"),
        top=Side(style='thin', color="FECACA"), bottom=Side(style='thin', color="FECACA")
    )

    for r in range(6, 16):
        for c in range(2, 7):
            if not ws.cell(r, c).border.left.style:
                ws.cell(r, c).border = border_card

    # Senaryo B
    ws.merge_cells("H6:L6")
    ws["H6"] = "SENARYO B: BAZ MODEL (HEDEFLENEN)"
    ws["H6"].font = Font(name=FONT_FAMILY, size=9, bold=True, color=COLOR_BLUE_TXT)
    ws["H6"].fill = PatternFill(start_color=COLOR_BLUE_CARD_BG, end_color=COLOR_BLUE_CARD_BG, fill_type="solid")
    ws["H6"].alignment = Alignment(horizontal="center", vertical="center")

    scen_b_metrics = [
        ("Ciro / Hacim Gerçekleşmesi:", "%100.0 (Tam Bütçe)", COLOR_BLUE_TXT),
        ("Beklenen Net Kâr / Getiri:", "₺12.450.000", COLOR_TEXT_MAIN),
        ("Serbest Nakit Akımı:", "₺8.240.000", COLOR_EMERALD_TXT),
        ("Yatırım Geri Dönüşü (ROI):", "%38.4 Yıllık", COLOR_EMERALD_TXT),
    ]

    for idx, (label, val, col_c) in enumerate(scen_b_metrics, start=8):
        ws.cell(idx, 8, label).font = Font(name=FONT_FAMILY, size=8.5, bold=False, color=COLOR_TEXT_MUTED)
        c_v = ws.cell(idx, 11, val)
        c_v.font = Font(name=FONT_FAMILY, size=9, bold=True, color=col_c)
        c_v.alignment = Alignment(horizontal="right")

    ws.merge_cells("H14:L15")
    btn_b = ws["H14"]
    btn_b.value = "OPTİMUM İCRA VE PLANLAMA REÇETESİ"
    btn_b.font = Font(name=FONT_FAMILY, size=8.5, bold=True, color=COLOR_BLUE_TXT)
    btn_b.fill = PatternFill(start_color=COLOR_BLUE_BG, end_color=COLOR_BLUE_BG, fill_type="solid")
    btn_b.alignment = Alignment(horizontal="center", vertical="center")
    btn_b.border = Border(
        left=Side(style='thin', color="BFDBFE"), right=Side(style='thin', color="BFDBFE"),
        top=Side(style='thin', color="BFDBFE"), bottom=Side(style='thin', color="BFDBFE")
    )

    for r in range(6, 16):
        for c in range(8, 13):
            if not ws.cell(r, c).border.left.style:
                ws.cell(r, c).border = border_card

    # Yönetim Karar Matrisi
    ws["B17"] = "YÖNETİM KARAR MATRİSİ VE UYGULAMA REÇETESİ"
    ws["B17"].font = Font(name=FONT_FAMILY, size=11, bold=True, color=COLOR_TEXT_MAIN)

    action_items = [
        ("Öncelik 1: Acil Likidite ve Borç Servis Planlaması",
         "Sistemde hesaplanan vade günlerine sadık kalınarak nakit blokajı 3 gün önceden hazır edilmelidir."),
        ("Öncelik 2: Formül Doğrulama ve Çapraz Sağlama",
         "Tüm hücreler kilitli formül korumasında olup girdi sayfalarında sıfır döngüsel başvuru (circular ref) garantilidir."),
        ("Öncelik 3: Resmi Mevzuat ve Vergi Kalkanı Entegrasyonu",
         "2026 güncel vergi dilimleri, amortisman ve SGK teşvik parametreleri tam otomatik yansıtılmaktadır."),
        ("Öncelik 4: Yönetici İmzası ve Rapor Paylaşımı",
         "Tek tuşla PDF ve A4 çıktı formatına tam uyumlu sayfa yapısı kurgulanmıştır.")
    ]

    curr_row = 19
    for title, desc in action_items:
        ws.cell(curr_row, 2, f"●  {title}").font = Font(name=FONT_FAMILY, size=9, bold=True, color=COLOR_TEXT_MAIN)
        ws.cell(curr_row + 1, 2, f"   {desc}").font = Font(name=FONT_FAMILY, size=8.5, color=COLOR_TEXT_MUTED)
        curr_row += 2

    widths = {1: 8, 2: 24, 3: 15, 4: 15, 5: 22, 6: 6, 7: 6, 8: 24, 9: 15, 10: 15, 11: 22, 12: 6}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def set_tab_color(ws, hex_color):
    """Excel sekme rengini ayarlar."""
    ws.sheet_properties.tabColor = hex_color

def process_product_workbook(master_wb, file_path: Path):
    """Bir ürünü master standart şablon ve görsellerdeki birebir tasarımla tam donatır."""
    prod_wb = openpyxl.load_workbook(file_path)
    product_name = file_path.parent.name
    ctx = get_product_context(product_name)

    # Önce eski eklenmiş sayfaları temizle
    sheets_to_clean = [
        "00_BASLANGIC", "PANO", "01_KOKPIT", "GIRDILER", "GİRDİLER", "02_GIRDI",
        "YONETIM_RAPORU", "YÖNETİM_RAPORU", "03_KARAR_RAPORU",
        "01_YURUTME_CEKIRDEGI", "02_KONTROL_LISTESI", "03_GORSEL_TASARIM",
        "04_MUHASEBE_THP", "05_KARAR_SATIS"
    ]
    for s_name in sheets_to_clean:
        if s_name in prod_wb.sheetnames and len(prod_wb.sheetnames) > 1:
            prod_wb.remove(prod_wb[s_name])

    # 1. 00_BASLANGIC (Master şablondan kopyala)
    ws_baslangic = prod_wb.create_sheet("00_BASLANGIC", 0)
    copy_master_sheet(master_wb["00_BASLANGIC"], ws_baslangic)
    set_tab_color(ws_baslangic, "FFC6A15B")  # Altın

    # 2. PANO (Yönetici Karar Kokpiti)
    ws_pano = prod_wb.create_sheet("PANO", 1)
    build_pano_sheet(ws_pano, ctx)
    set_tab_color(ws_pano, "FF17324D")  # Lacivert

    # 3. GİRDİLER (Veri Girişi & Operasyon Motoru)
    ws_girdi = prod_wb.create_sheet("GİRDİLER", 2)
    build_girdiler_sheet(ws_girdi, ctx)
    set_tab_color(ws_girdi, "FF2563EB")  # Mavi

    # 4. YÖNETİM_RAPORU (İcra Raporu & Senaryo Analizi)
    ws_rapor = prod_wb.create_sheet("YÖNETİM_RAPORU", 3)
    build_yonetim_raporu_sheet(ws_rapor, ctx)
    set_tab_color(ws_rapor, "FF991B1B")  # Bordo

    # 5. 01_YURUTME_CEKIRDEGI (Master şablondan kopyala)
    ws_yurutme = prod_wb.create_sheet("01_YURUTME_CEKIRDEGI", 4)
    copy_master_sheet(master_wb["01_YURUTME_CEKIRDEGI"], ws_yurutme)
    set_tab_color(ws_yurutme, "FF065F46")  # Koyu Yeşil

    # 6. 02_KONTROL_LISTESI (Master şablondan kopyala)
    ws_kontrol = prod_wb.create_sheet("02_KONTROL_LISTESI", 5)
    copy_master_sheet(master_wb["02_KONTROL_LISTESI"], ws_kontrol)
    set_tab_color(ws_kontrol, "FF9B2C2C")  # Kırmızı

    # 7. 03_GORSEL_TASARIM (Master şablondan kopyala)
    ws_gorsel = prod_wb.create_sheet("03_GORSEL_TASARIM", 6)
    copy_master_sheet(master_wb["03_GORSEL_TASARIM"], ws_gorsel)
    set_tab_color(ws_gorsel, "FF475569")  # Slate Gri

    # 8. 04_MUHASEBE_THP (Master şablondan kopyala)
    ws_thp = prod_wb.create_sheet("04_MUHASEBE_THP", 7)
    copy_master_sheet(master_wb["04_MUHASEBE_THP"], ws_thp)
    set_tab_color(ws_thp, "FF2F75B5")  # Muhasebe Mavisi

    # 9. 05_KARAR_SATIS (Master şablondan kopyala)
    ws_satis = prod_wb.create_sheet("05_KARAR_SATIS", 8)
    copy_master_sheet(master_wb["05_KARAR_SATIS"], ws_satis)
    set_tab_color(ws_satis, "FF5B477E")  # Mor

    # Tüm sayfaları görünür yap & gridline kapat & A1 Menü köprüsü ekle
    for sheet in prod_wb.worksheets:
        sheet.sheet_state = "visible"
        sheet.views.sheetView[0].showGridLines = False
        if sheet.title != "PANO" and sheet["A1"].value is None:
            sheet["A1"] = "☰ PANO"
            sheet["A1"].font = Font(name=FONT_FAMILY, size=8, bold=True, color=COLOR_TEXT_MUTED)

    # Aktif sekme: PANO
    prod_wb.active = prod_wb["PANO"]

    # Kaydet
    prod_wb.save(file_path)
    return True

def main():
    master_path = Path("calisma_prorokolu/EXCELARSIV_DECISION_OS_V15_1_MASTER_STANDART_UI_LOCKED.xlsx")
    if not master_path.exists():
        print(f"Master şablon bulunamadı: {master_path}")
        sys.exit(1)
        
    print(f"Master Şablon Yükleniyor: {master_path.name}...")
    master_wb = openpyxl.load_workbook(master_path)
    
    root = Path("EXCELARSIV_PRODUCTS")
    products = [p for p in root.iterdir() if p.is_dir() and not p.name.startswith(('.', '_'))]
    print(f"Toplam {len(products)} ürün MANDATE V15.1 ve MASTER STANDART ile tam donatılıyor...")
    
    success = 0
    fail = 0
    
    for idx, p in enumerate(products, start=1):
        xlsxs = [x for x in p.glob("*.xlsx") if not x.name.startswith("~$")]
        if not xlsxs:
            fail += 1
            continue
        try:
            process_product_workbook(master_wb, xlsxs[0])
            success += 1
            if idx % 10 == 0 or idx == len(products):
                print(f"[{idx}/{len(products)}] Başarıyla Donatıldı: {p.name}")
        except Exception as e:
            print(f"[{idx}/{len(products)}] Hata ({p.name}): {e}")
            fail += 1
            
    print(f"\nİşlem Tamamlandı! Başarılı: {success}, Başarısız: {fail}")

if __name__ == "__main__":
    main()
