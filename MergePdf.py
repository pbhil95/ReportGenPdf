import os
from PyPDF2 import PdfFileMerger
import re

source_dir = './Files/ResultPdf/'
source_dir1 = './Files/ResultPdf/Output/'
source_dir2 = './Files/Final/'
backpage='./Files/BackPage/backpage.pdf'
merger = PdfFileMerger()
merger1 = PdfFileMerger()

for item in os.listdir(source_dir):
    if item.endswith('pdf'):
        #print(item)
        merger = PdfFileMerger()
        #print(source_dir + item)
        merger.append(source_dir + item)
        merger.append(backpage)
        merger.write(source_dir + 'Output/'+item)
merger.close()

for item in sorted(os.listdir(source_dir1),key=len):
    if item.endswith('pdf'):
        merger1.append(source_dir1 + item)       
merger1.write(source_dir2 +'ReportCard_XiSc23'+'.pdf')
merger1.close()
                