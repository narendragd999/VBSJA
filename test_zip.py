import zipfile
import re
import os

def sanitize_hindi_filename(name: str) -> str:
    # Only remove invalid Windows filename characters: < > : " / \ | ? * and control chars
    # Keep all Hindi/Devanagari letters and matras (U+0900 to U+097F), English, numbers, spaces, brackets, hyphens
    clean = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '', name)
    # Replace en-dash / em-dash with standard hyphen
    clean = clean.replace('–', '-').replace('—', '-')
    clean = re.sub(r'\s+', ' ', clean).strip()
    return clean

cats = [
    'आशीर्वाद का दीया',
    'नमो सेवा–अनकही कहानियाँ (ग्रामीण क्षेत्र)',
    'नमो सेवा–अनकही कहानियाँ (शहरी क्षेत्र)',
    'मेरे सपनों का भारत–चित्रसेवा',
    'रक्तदान शिविर',
    'सेवा की कहानी',
    'सेवा के रंग, राष्ट्र के संग',
    'सेवा सेतु अभियान (ग्रामीण क्षेत्र)'
]

zip_path = 'test_hindi_zip.zip'
with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for i, c in enumerate(cats, 1):
        clean_name = sanitize_hindi_filename(c)
        fname = f'{i:02d}_{clean_name}.pdf'
        print('Zipping filename:', fname)
        zf.writestr(fname, b'%PDF-1.4 dummy')

with zipfile.ZipFile(zip_path, 'r') as zf:
    for info in zf.infolist():
        print('In zip:', info.filename, 'Flag bits:', bin(info.flag_bits))
