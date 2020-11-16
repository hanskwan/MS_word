#!/usr/bin/env python3
# -*- coding: utf-8 -*-

#Importing packages
import pandas as pd
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH 
from time import gmtime, strftime

# Read excel
hk_slope = pd.read_excel("HK Slope 2020.xlsx")

# Set up words
doc = docx.Document()

style = doc.styles['Normal']
font = style.font
font.name = "Calibri"
font.size = Pt(12)

section = doc.sections[0]
header = section.header
htable=header.add_table(1, 2, Cm(10))
htab_cells=htable.rows[0].cells
ht0=htab_cells[0].add_paragraph()
logo=ht0.add_run()
logo.add_picture("CEDD.png", width=docx.shared.Cm(2),height=docx.shared.Cm(2))
ht1 = htab_cells[1].add_paragraph('Slope Information System')

footer = section.footer
footer.add_paragraph("Record Retrieved From SIS on " + strftime("%Y-%m-%d %H:%M:%S", gmtime()))

# headings

# First Paragraph
doc.add_heading("Basic Information",1).italic = True
doc.add_paragraph(str(hk_slope.columns[0]) +" : "+ str(hk_slope.iloc[0,0]))
doc.add_paragraph(str(hk_slope.columns[1]) +" : "+ str(hk_slope.iloc[0,1]))
doc.add_paragraph(str(hk_slope.columns[2]) +" : "+ str(hk_slope.iloc[0,2]))
doc.add_page_break()

# Second Paragraph
doc.add_heading("Slope Part",2).italic = True
doc.add_paragraph(str(hk_slope.columns[3]) +" : "+ str(hk_slope.iloc[0,3]))
doc.add_paragraph(str(hk_slope.columns[4]) +" : "+ str(hk_slope.iloc[0,4]))
doc.add_paragraph(str(hk_slope.columns[5]) +" : "+ str(hk_slope.iloc[0,5]))
doc.add_page_break()

# Third Paragraph
doc.add_heading("Slope Change",3).italic = True

significant = "Not Significant"
doc.add_paragraph(significant)
doc.add_page_break()

# Forth Paragraph
doc.add_heading("Map",4).italic = True
doc.add_picture("hfc.png", width=docx.shared.Cm(15),height=docx.shared.Cm(20))

# Export Document
doc.save("HKSS Report Nov 2020.docx")

    
