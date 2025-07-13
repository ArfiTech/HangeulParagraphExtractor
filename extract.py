# -*- coding: utf-8 -*-

import sys
import os
import zipfile
import xml.etree.ElementTree as ET

if len(sys.argv) != 2:
    sys.exit(1)
HWPIX_FILE = sys.argv[1]
OUTPUT_DIR = ''  # 찾은 단어가 담긴 문단들이 포함된 hwpml 파일들이 저장될 위치
CHAR_NAMES = []  # 원하는 단어 나열

os.makedirs(OUTPUT_DIR, exist_ok=True)

def extract_paragraphs_from_hwpx(hwpx_path):
    paragraphs = []
    with zipfile.ZipFile(hwpx_path, 'r') as z:
        for entry in z.namelist():
            lower = entry.lower()
            if lower.startswith('contents/section') and lower.endswith('.xml'):
                xml_data = z.read(entry).decode('utf-8', errors='ignore')
                xml_data = xml_data.replace('xmlns="', 'xmlns_removed="')
                root = ET.fromstring(xml_data)
                for elem in root.iter():
                    tag = elem.tag.split('}')[-1]
                    if tag.upper() == 'P':
                        text = ''.join(elem.itertext()).strip()
                        if text and not text.startswith('-'):
                            paragraphs.append(text)
    return paragraphs

def group_paragraphs(paragraphs, names):
    grouped = {n: [] for n in names}
    for p in paragraphs:
        if p.startswith('-'):
            continue
        for n in names:
            if n in p:
                grouped[n].append(p)
    return grouped

def write_hwpml(name, paras):
    out_path = os.path.join(OUTPUT_DIR, f"{name}.hwpml")
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write('<HWPML>\n')
        f.write('  <HEAD/>\n')
        f.write('  <BODY>\n')
        f.write('    <SECTION>\n')
        for p in paras:
            esc = p.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            f.write(f'      <P>{esc}</P>\n')
        f.write('    </SECTION>\n')
        f.write('  </BODY>\n')
        f.write('</HWPML>\n')
    print(f"[{name}] {len(paras)} paragraphs written to {out_path}")

def main():
    paras = extract_paragraphs_from_hwpx(HWPIX_FILE)
    print(f"Total paragraphs extracted: {len(paras)}")
    grouped = group_paragraphs(paras, CHAR_NAMES)
    for name, ps in grouped.items():
        print(f"Processing {name}: {len(ps)} paragraphs")
        if ps:
            write_hwpml(name, ps)
    print("Done:", OUTPUT_DIR)

if __name__ == '__main__':
    main()
