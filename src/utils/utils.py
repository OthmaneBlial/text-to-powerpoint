import subprocess
import os
from PyQt5.QtGui import QPixmap

def convert_pptx_to_image(pptx_path, output_image_path):
    try:
        # Convert PPTX to PDF using LibreOffice
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', pptx_path, '--outdir', os.path.dirname(pptx_path)], check=True)
        pdf_path = pptx_path.replace('.pptx', '.pdf')
        
        # Convert the first page of PDF to PNG using ImageMagick
        subprocess.run(['convert', '-density', '300', f"{pdf_path}[0]", output_image_path], check=True)
        
        # Clean up PDF
        os.remove(pdf_path)
        return output_image_path
    except Exception as e:
        return None

def load_image(image_path):
    if os.path.exists(image_path):
        pixmap = QPixmap(image_path)
        return pixmap
    return None
