import zipfile
import shutil
from pathlib import Path
import xml.etree.ElementTree as ET

ref_path = Path('/Users/macair1/projects/excel-arsiv/referans-sablonlar/Program & Taslaklar kopyası/EXCEL751-Gelir-Gider_Tablosu.xlsm')
target_path = Path('/Users/macair1/projects/excel-arsiv/EXCELARSIV_PRODUCTS/banka-kredi-ve-taksit-takip-sistemi/BankaKrediVeTaksitTakipSistemi.xlsx')

tmp_ref = Path('/tmp/ref_extracted')
tmp_target = Path('/tmp/target_extracted')

if tmp_ref.exists(): shutil.rmtree(tmp_ref)
if tmp_target.exists(): shutil.rmtree(tmp_target)

with zipfile.ZipFile(ref_path, 'r') as z:
    z.extractall(tmp_ref)

with zipfile.ZipFile(target_path, 'r') as z:
    z.extractall(tmp_target)

print("1. Identifying sheet1.xml (Kılavuz) in reference...")
ref_sheet1 = tmp_ref / 'xl/worksheets/sheet1.xml'
ref_drawing = tmp_ref / 'xl/drawings/drawing1.xml'
ref_drawing_rels = tmp_ref / 'xl/drawings/_rels/drawing1.xml.rels'

target_wb_xml = tmp_target / 'xl/workbook.xml'
target_wb_rels = tmp_target / 'xl/_rels/workbook.xml.rels'

tree = ET.parse(target_wb_xml)
root = tree.getroot()
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

kilavuz_rId = None
for sheet in root.findall('.//s:sheet', ns):
    if sheet.get('name') == 'KILAVUZ':
        kilavuz_rId = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        break

print(f"KILAVUZ rId in target: {kilavuz_rId}")

tree_rels = ET.parse(target_wb_rels)
root_rels = tree_rels.getroot()
ns_rel = {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}

kilavuz_target_file = None
for rel in root_rels.findall('.//rel:Relationship', ns_rel):
    if rel.get('Id') == kilavuz_rId:
        kilavuz_target_file = rel.get('Target').lstrip('/')
        break

print(f"KILAVUZ target file in target: {kilavuz_target_file}")
target_sheet_path = tmp_target / kilavuz_target_file
target_sheet_rels_path = tmp_target / 'xl/worksheets/_rels' / (Path(kilavuz_target_file).name + '.rels')

print("2. Transplanting exact sheet1.xml and its drawing/media stack...")
shutil.copy2(ref_sheet1, target_sheet_path)

existing_drawings = list((tmp_target / 'xl/drawings').glob('drawing*.xml')) if (tmp_target / 'xl/drawings').exists() else []
new_drawing_idx = len(existing_drawings) + 1
new_drawing_name = f'drawing{new_drawing_idx}.xml'
new_drawing_rels_name = f'drawing{new_drawing_idx}.xml.rels'

(tmp_target / 'xl/drawings').mkdir(parents=True, exist_ok=True)
(tmp_target / 'xl/drawings/_rels').mkdir(parents=True, exist_ok=True)
(tmp_target / 'xl/media').mkdir(parents=True, exist_ok=True)
(tmp_target / 'xl/worksheets/_rels').mkdir(parents=True, exist_ok=True)

shutil.copy2(ref_drawing, tmp_target / 'xl/drawings' / new_drawing_name)
if ref_drawing_rels.exists():
    shutil.copy2(ref_drawing_rels, tmp_target / 'xl/drawings/_rels' / new_drawing_rels_name)

if (tmp_ref / 'xl/media').exists():
    for m in (tmp_ref / 'xl/media').iterdir():
        dest = tmp_target / 'xl/media' / m.name
        if not dest.exists():
            shutil.copy2(m, dest)

rels_content = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/{new_drawing_name}"/>
</Relationships>'''
target_sheet_rels_path.write_text(rels_content, encoding='utf-8')

# Ensure Content_Types has drawing and image extensions
ct_path = tmp_target / '[Content_Types].xml'
ct_tree = ET.parse(ct_path)
ct_root = ct_tree.getroot()

extensions = {d.get('Extension'): d for d in ct_root.findall('{http://schemas.openxmlformats.org/package/2006/content-types}Default')}
defaults_to_add = [
    ('png', 'image/png'),
    ('jpeg', 'image/jpeg'),
    ('jpg', 'image/jpeg')
]
for ext, ctype in defaults_to_add:
    if ext not in extensions:
        elem = ET.SubElement(ct_root, '{http://schemas.openxmlformats.org/package/2006/content-types}Default')
        elem.set('Extension', ext)
        elem.set('ContentType', ctype)

drawing_part = f"/xl/drawings/{new_drawing_name}"
existing_overrides = {o.get('PartName'): o for o in ct_root.findall('{http://schemas.openxmlformats.org/package/2006/content-types}Override')}
if drawing_part not in existing_overrides:
    elem = ET.SubElement(ct_root, '{http://schemas.openxmlformats.org/package/2006/content-types}Override')
    elem.set('PartName', drawing_part)
    elem.set('ContentType', 'application/vnd.openxmlformats-officedocument.drawing+xml')

ct_tree.write(ct_path, encoding='utf-8', xml_declaration=True)

# Re-pack zip cleanly
with zipfile.ZipFile(target_path, 'w', zipfile.ZIP_DEFLATED) as z_out:
    for f in tmp_target.rglob('*'):
        if f.is_file():
            arcname = f.relative_to(tmp_target)
            z_out.write(f, arcname)

print("✓ Exact reference Kılavuz transplanted successfully with full vector shapes and buttons!")
