"""Extract text and image info from the PAP report DOCX file."""
import sys
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT

doc_path = r"C:\Users\matis\Downloads\kidtokid-main\PAP-Relatorio\EPB-ModeloPapV8_VERSÃO 2030 (1).docx"
out_path = r"C:\Users\matis\Downloads\kidtokid-main\PAP-Relatorio\report_text.txt"

doc = Document(doc_path)

# Count images in the document
image_count = 0
for rel in doc.part.rels.values():
    if "image" in rel.reltype:
        image_count += 1

with open(out_path, "w", encoding="utf-8") as f:
    f.write(f"=== DOCUMENT IMAGE COUNT: {image_count} images found in document rels ===\n\n")
    
    para_num = 0
    for para in doc.paragraphs:
        para_num += 1
        style_name = para.style.name if para.style else "None"
        text = para.text.strip()
        
        # Check if paragraph contains images
        has_image = False
        img_count_in_para = 0
        for run in para.runs:
            if run._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing'):
                has_image = True
                img_count_in_para += 1
            if run._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pict'):
                has_image = True
                img_count_in_para += 1
                
        # Also check for inline shapes
        drawings = para._element.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
        if drawings:
            has_image = True
            img_count_in_para = max(img_count_in_para, len(drawings))
        
        img_marker = f" [🖼️ IMAGE x{img_count_in_para}]" if has_image else ""
        
        if text or has_image:
            f.write(f"[P{para_num}|{style_name}]{img_marker} {text}\n")
    
    # Also extract tables
    f.write("\n\n=== TABLES ===\n")
    for t_idx, table in enumerate(doc.tables):
        f.write(f"\n--- Table {t_idx + 1} ---\n")
        for row_idx, row in enumerate(table.rows):
            cells = [cell.text.strip().replace('\n', ' | ') for cell in row.cells]
            f.write(f"  Row {row_idx}: {' || '.join(cells)}\n")

print(f"Done. Extracted {para_num} paragraphs, {image_count} images total.")
