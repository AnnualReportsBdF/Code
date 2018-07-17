# -*- coding: utf-8 -*-
"""
Created on Mon Jun 25 17:33:30 2018

@author: guill
"""
import sys
from os.path import dirname
sys.path.append('F:/site-packages')

import textract
import re
import os  
import pandas as pd
import numpy as np
from io import StringIO
import pytesseract 
from PIL import Image
#import wand.image
import cProfile
import multiprocessing 
#! pip install line_profiler
#! pip install C:/Users/guill/Downloads/PythonMagick-0.9.13-cp36-cp36m-win_amd64.whl
import pdfminer
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
import pdf2image
import PyPDF2
import shutil
import time
import random
import langdetect
pytesseract.pytesseract.tesseract_cmd = 'F:/site-packages/pytesseract/tesseract/tesseract'

def convert(file):
    try:
        text = textract.process(file)
        text=str(text,"utf-8")
        return text
    except UnicodeDecodeError:
        infile = PyPDF2.PdfFileReader(open(file, "rb"))
        #infile.decrypt('')
        text = []
        for page in range(0, infile.getNumPages()):  
            currentPage = infile.getPage(page)
            text_inpage = currentPage.extractText()
            #thispage = myText.split()
            text.append(text_inpage)
            text2 = ' '.join(text)
        return text2
    

#a=convert('C:/Users/guill/Desktop/banque_france/CAC_40_test/HERMES/document [12].pdf')


def detect_language(pdf_toread,pdf_to_out,auxiliary_path,num_pages,sample,langue_1):
    lan1 = ''
    lan2 = ''
    while (len(lan1) < sample):
        liste_random_page=random.sample(range(0, num_pages), 1)[0]
        pdf_to_out.addPage(pdf_toread.getPage(liste_random_page))
        newname = auxiliary_path +"out_" + str(liste_random_page) + ".pdf"    
        outputStream = open(newname, "wb")
        pdf_to_out.write(outputStream)
        outputStream.close()
        #print(liste_random_page)
        lan1=''
        lan1 += pytesseract.image_to_string(pdf2image.convert_from_path(auxiliary_path+"out_"+str(liste_random_page)+".pdf")[0],lang=langue_1)
    lan2 += pytesseract.image_to_string(pdf2image.convert_from_path(auxiliary_path+"out_"+str(liste_random_page)+".pdf")[0],lang='eng')
    out_lan1 = langdetect.detect_langs(lan1)[0]
    out_lan2 = langdetect.detect_langs(lan2)[0]
    lang=re.findall(r"[a-zA-Z]+",str(max(out_lan1,out_lan2)))[0]
    if lang=='en':
       lang='eng'
    else :
       lang='fra'    
    return lang



#Fonction advanced_conversion : MARCHE !
#Possible optimisation au niveau de detect_langage : on a déjà OCR un échantillon du pdf,
# on peut garder cette échantillon et boucler uniquement sur les pages non échantillonnées par la suite
def advanced_conversion(file,auxiliary_path,sample,langue_1):
    #try:
        text = convert(file)
    #scan = 0
        pdf_toread = PyPDF2.PdfFileReader(open(file, "rb"))
    #pdf_toread.decrypt('')
        num_pages = pdf_toread.getNumPages()
        if text == '\x0c'*num_pages:
            scan = 1
            text=''
        #lang='fra'
            pdf_to_out=PyPDF2.PdfFileWriter()
            
            lang = detect_language(pdf_toread=pdf_toread,pdf_to_out=pdf_to_out,auxiliary_path=auxiliary_path,
                               num_pages=num_pages,sample=sample,langue_1=langue_1)
        
        #pdf_toread = PyPDF2.PdfFileReader(file)
            for i in range(0,num_pages):
                pdf_to_out=PyPDF2.PdfFileWriter()
                pdf_to_out.addPage(pdf_toread.getPage(i))
                newname = auxiliary_path +"out_" + str(i) + ".pdf"    
                outputStream = open(newname, "wb")
                pdf_to_out.write(outputStream)
                outputStream.close()
                text += pytesseract.image_to_string(pdf2image.convert_from_path(auxiliary_path+"out_"+str(i)+".pdf")[0],lang=lang)
        else:
            scan = 0
            lang = re.findall(r"[a-zA-Z]+",str(langdetect.detect_langs(text)[0]))[0]       
            return text, scan, lang
    #except:
     #       text ='Error'
      #      scan = 'Error'
       #     lang = 'Error'
        #    return text, scan, lang

#advanced_conversion('C:/Users/guill/Desktop/banque_france/CAC_40_test/HERMES/document [11].pdf',
 #                   "C:/Users/guill/Desktop/banque_france/outfile//",300,'fra')

if __name__ == '__main__':    
    firm = []
    scan = []
    year = []
    text = []
    lang = []
    whole_list = []
    num_done = 0 
    def print_result(res):
        global num_done
        start = time.time()
        num_done += 1
        print('finish in {} for item {}'.format(time.time()-start,num_done))
    
    
    p = multiprocessing.Pool(multiprocessing.cpu_count()-1)
    start_time = time.time()

    #out_write=open('C:/Users/guill/Desktop/banque_france/test.txt','rU')
    for fn in os.listdir('C:/Users/guill/Desktop/banque_france/CAC_40_test'):
    #    print(fn)
        for fn2 in os.listdir('C:/Users/guill/Desktop/banque_france/CAC_40_test//'+fn):
     #       print(fn2)
         #   firm.append(fn)
          #  year.append(re.findall(r'\d+',fn2))
             #res=p.apply_async(advanced_conversion,args=['C:/Users/guill/Desktop/banque_france/CAC_40_test//'+fn+'//'+fn2,
              #                              "C:/Users/guill/Desktop/banque_france/outfile//",
               #                             300,'fra'],callback = whole_list.append) 
               res=p.apply_async(advanced_conversion,args=['C:/Users/guill/Desktop/banque_france/CAC_40_test//'+fn+'//'+fn2,
                                           "C:/Users/guill/Desktop/banque_france/outfile//",
                                            300,'fra'],callback = print_result) 
 
            #text.append(av[0])
            #scan.append(av[1])
            #lang.append(av[2])
    p.close()
    p.join()
    #for i in range(0,len(whole_list)):
     #   text.append(whole_list[i][0])
      #  scan.append(whole_list[i][1])
       # lang.append(whole_list[i][2])
    #out_write.write(res)
    print(time.time() - start_time)       
         #firm[firm.index(fn2)]=fn
        #text[fn2] = advanced_conversion(fn2,'C:/Users/guill/Desktop/banque_france/outfile//')



