# -*- coding: utf-8 -*-
"""
Created on Thu Sep 14 16:11:31 2023

@author: Marloes Venema
"""

from pdf2image import convert_from_path
from pptx import Presentation
from pptx.util import Inches
import os
import io
import math

# Change the pdf and ppt names accordingly
pdf_name = "test.pdf"
ppt_name = "test.pptx"

# Use half of the available cores for conversion
conversion_cores = math.ceil(os.cpu_count() / 2)

# Convert PDF with to high-res PNGs
images = convert_from_path(pdf_name, dpi=1000, fmt='png', thread_count=conversion_cores)

# Determine aspect ratio of PDF presentation
ratio = images[0].width / float(images[0].height)
width = 10
height = width / ratio

# Set up PPT presentation
prs = Presentation()
prs.slide_width = Inches(width)
prs.slide_height = Inches(height)

# Add new slide for each PNG image
for img in images:
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Convert image object to virtual file object
    tmpFile = io.BytesIO()
    img.save(tmpFile, "PNG")

    slide.shapes.add_picture(tmpFile, 0, 0, height=Inches(height))
    
prs.save(ppt_name)
