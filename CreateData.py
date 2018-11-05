# -*- coding: utf-8 -*-
"""
Created on Fri Aug 03 17:33:45 2018

@author: R550427
"""

import pandas as pd

import sys
import os
import io
import re
import multiprocessing
import time

#Convert pdf to text
import textract
import PyPDF2
from PIL import Image
from pypdfocr.pypdfocr_gs import PyGs
import pytesseract
pytesseract.pytesseract.tesseract_cmd = 'Lib/site-packages/pytesseract/tesseract/tesseract'
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage

import langdetect

import warnings
warnings.filterwarnings("ignore")

#==============================================================================
def LanguageName(lang):
    threeLetters = ["fra", "spa", "ita", "deu", "eng", "nld", "ell", "tur", 
                    "est", "fin", "swe", "gle", "lit", "ltz", "mlt", "por", 
                    "slv", "slk"]
    twoLetters = ["fr", "es", "it", "de", "en", "nl", "el", "tr", "et", "fi", 
                  "sv", "ga", "lt", "lb", "mt", "pt", "sl", "sk"]
    #https://fr.wikipedia.org/wiki/Liste_des_codes_ISO_639-1
    if (lang in threeLetters):
        lang = twoLetters[threeLetters.index(lang)]
    elif (lang in twoLetters):
        lang = threeLetters[twoLetters.index(lang)]
    return lang
#==============================================================================
def RetrievePdfInfos(pdf):
    date_start_pos = [m.start() for m in re.compile("[ ][0-9]{4}[ ]").finditer(pdf)][0]
    country = pdf.split()[0]
    company = pdf[len(country)+1:date_start_pos]
    year = pdf[date_start_pos+1:date_start_pos+5]
    lang = pdf[date_start_pos+6:date_start_pos+9]
    if (pdf.split()[-1][:-4].isdigit()):
        pdfType = pdf[pdf.find(lang)+4:pdf.find(pdf.split()[-1])-1]
    else:
        pdfType = pdf.split()[-1][:-4]
    return country, company, year, lang, pdfType
#==============================================================================
def CountPagesNumber(path):
    pdf = PyPDF2.PdfFileReader(path)
    if pdf.isEncrypted:
        pdf.decrypt('')
    return pdf.numPages
#==============================================================================
def ConvertPdftoText(path):
    try:
        text = unicode(textract.process(path), "utf-8")
        return text
    except:
        text = ""
        pagesNumber = CountPagesNumber(path)
        pdf = PyPDF2.PdfFileReader(path)
        if pdf.isEncrypted:
            pdf.decrypt('')
        for ith in range(pagesNumber):
            page = pdf.getPage(ith)
            text += page.extractText()
        text = unicode(text.encode("utf-8"), "utf-8")
        return text
#==============================================================================
def ConvertScanToText(path, language):
    text = ""
    pagesNumber = CountPagesNumber(path)
    PyGs({}).make_img_from_pdf(path)
    for ith in range(pagesNumber):
        imagePath = path[:-4] + "_" + str(ith + 1) + ".JPG"
        image = Image.open(imagePath, mode="r")
        text += pytesseract.image_to_string(image, lang = language)
        os.remove(imagePath)
    return unicode(text.encode("utf-8"), "utf-8")
#==============================================================================
def ConvertFileToText(path, language):
    text = ConvertPdftoText(path)
    pagesNumber = CountPagesNumber(path)
    scannedFile = 0
    
    if text in ["\x0c" * pagesNumber, ""]:
        scannedFile = 1
        text = ConvertScanToText(path, language)
        
    languageEstimated = LanguageName(str(langdetect.detect_langs(text))[1:3])
    
    # If the pdf language is confusing, extract the text with a more precise tool (but less efficient)
    if ((LanguageName(str(langdetect.detect_langs(text))[1:3]) != language) & (scannedFile == 0)):
        prm = PDFResourceManager()
        iob = io.BytesIO()
        device = TextConverter(prm, iob, codec = "utf-8", laparams = LAParams())
        pdf = open(path, "rb")
        interpreter = PDFPageInterpreter(prm, device)
        for page in PDFPage.get_pages(pdf, set(), maxpages = 0, password = "", caching = True, check_extractable = True):
            interpreter.process_page(page)
        text = iob.getvalue()
        pdf.close()
        device.close()
        iob.close()
        
        languageEstimated = LanguageName(str(langdetect.detect_langs(text))[1:3])
        
    return text, scannedFile, languageEstimated
#==============================================================================
def CreateDataWithoutDuplicates(directoryPath):
    dataDic = []
    ithFile = 0
    total_nb = len([pdf for pdf in os.listdir(directoryPath) if (pdf[-4:] == ".pdf")])
    sys.stdout.write("\r")
    sys.stdout.write("CreateDataWithoutDuplicates -- [" + total_nb * " " + "] 0%")
    
    # Create the dataframe of unique pdfs
    for pdf in os.listdir(directoryPath):
        if (pdf[-4:] == ".pdf"):
            ithFile += 1
            path = directoryPath + "/" + pdf
            pdfInfos = RetrievePdfInfos(pdf)
            try:
                pagesNumber = CountPagesNumber(path)
                textInfos = ConvertFileToText(path, pdfInfos[3])
                dataDic.append({"Country": pdfInfos[0], "Company": pdfInfos[1], "Year": pdfInfos[2], 
                            "Text": textInfos[0].encode("utf-8"), "Scan": textInfos[1],
                            "Pages_Number": pagesNumber, "Language_Expected": pdfInfos[3], 
                            "Language_Estimated": textInfos[2], "Type": pdfInfos[4]})
            except:
                pass
            
            sys.stdout.write("\r")
            sys.stdout.write("CreateDataWithoutDuplicates -- [" + ithFile * "#" + (total_nb-ithFile) * " " + "] " + str(int((float(ithFile)/total_nb)*100)) + "%")

    df = pd.DataFrame(dataDic)
    df.to_csv(directoryPath[:directoryPath.rfind("/")+1] + str(directoryPath[directoryPath.rfind("/")+1:]) + "-data.csv")
    print("\nCreate Data Without Duplicates: OK")
#==============================================================================  
def CreateDataWithDuplicates(directoryPath): 
    dataDic = []
    shortNames = []

    # Retrieve all shortNames (names that should be unique)
    for duplicate in os.listdir(directoryPath):
        date_start_pos = [m.start() for m in re.compile("[ ][0-9]{4}[ ]").finditer(duplicate)][0]
        shortNames.append(duplicate[:date_start_pos+5])
    shortNames = list(set(shortNames))
    
    # Stock all duplicates together
    pdfs = []
    for shortName in shortNames:
        duplicates = []
        for duplicate in os.listdir(directoryPath):
             if (duplicate.rfind(shortName) != -1):
                 duplicates.append(duplicate)
        pdfs.append(duplicates) 
    sys.stdout.write("\r")
    sys.stdout.write("CreateDataWithDuplicates -- [" + len(pdfs) * " " + "] 0%")
    
    # Create the dataframe of duplicates
    for i in range(len(pdfs)):
        d = pdfs[i]
        pdfInfos = RetrievePdfInfos(d[0])
        pagesNumbers = []
        texts = []
        scans = []
        estimatedlanguages = []
        for j in range(len(d)):
            path = directoryPath + "/" + d[j]
            try:
                pagesNumbers.append(CountPagesNumber(path))
                textInfos = ConvertFileToText(path, pdfInfos[3])
                texts.append(textInfos[0].encode('utf-8'))
                scans.append(textInfos[1])
                estimatedlanguages.append(textInfos[2])
            except:
                pass
            
        dataDic.append({"Country": pdfInfos[0], "Company": pdfInfos[1], "Year": pdfInfos[2], "Text": texts, 
                            "Scan": scans, "Pages_Number": pagesNumbers, "Language_Expected": pdfInfos[3], 
                            "Language_Estimated": estimatedlanguages, "Type": pdfInfos[4]})
        sys.stdout.write("\r")
        sys.stdout.write("CreateDataWithDuplicates -- [" + (i+1) * "#" + (len(pdfs)-(i+1)) * " " + "] " + str(int((float(i+1)/len(pdfs))*100)) + "%")
            
    df = pd.DataFrame(dataDic)
    df.to_csv(directoryPath[:directoryPath.rfind("/")+1] + str(directoryPath[directoryPath.rfind("/")+1:]) + "-data.csv")
    print("\nCreate Data With Duplicates: OK")
#==============================================================================     

# MAIN
directoryPath = "C:/temp/Annual_Reports_Netherlands_cleaned"
if (__name__ == '__main__'):
    start = time.time()
    p1 = multiprocessing.Process(target=CreateDataWithoutDuplicates, args=[directoryPath + "/Folder1"] )
    p2 = multiprocessing.Process(target=CreateDataWithoutDuplicates, args=[directoryPath + "/Folder2"])
    p3 = multiprocessing.Process(target=CreateDataWithoutDuplicates, args=[directoryPath + "/Folder3"])
    p4 = multiprocessing.Process(target=CreateDataWithoutDuplicates, args=[directoryPath + "/Folder4"])
    p5 = multiprocessing.Process(target=CreateDataWithoutDuplicates, args=[directoryPath + "/Folder5"])
    p6 = multiprocessing.Process(target=CreateDataWithDuplicates, args=[directoryPath + "/Duplicates"])
    p1.start()
    p2.start()
    p3.start()
    p4.start()
    p5.start()
    p6.start()
    p1.join()
    p2.join()
    p3.join()
    p4.join()
    p5.join()
    p6.join()
    print("\nSpent time: " + str(int(time.time()-start)) + " seconds")
            
"""   
def CreateData(directoryPath, isMultiprocessing):
    
    if (isMultiprocessing == 0):
        start = time.time()
        for folder in os.listdir(directoryPath):
            if ((folder[-4] != '.') & (folder != 'Error')):
                print('\n' + str(folder))
                if (folder != 'Duplicates'):
                    CreateDataWithoutDuplicates(directoryPath + '/' + folder)
                else:
                    CreateDataWithDuplicates(directoryPath + '/Duplicates')
        print('\nSpent time: ' + str(int(time.time()-start)) + ' seconds')
    
    else:
        if (__name__ == '__main__'):
            start = time.time()
            p1 = multiprocessing.Process(target=CreateDataWithoutDuplicates, args=[directoryPath + '/Folder1'] )
            p2 = multiprocessing.Process(target=CreateDataWithoutDuplicates, args=[directoryPath + '/Folder2'])
            p3 = multiprocessing.Process(target=CreateDataWithoutDuplicates, args=[directoryPath + '/Folder3'])
            p4 = multiprocessing.Process(target=CreateDataWithoutDuplicates, args=[directoryPath + '/Folder4'])
            p5 = multiprocessing.Process(target=CreateDataWithoutDuplicates, args=[directoryPath + '/Folder5'])
            p6 = multiprocessing.Process(target=CreateDataWithDuplicates, args=[directoryPath + '/Duplicates'])
            p1.start()
            p2.start()
            p3.start()
            p4.start()
            p5.start()
            p6.start()
            p1.join()
            p2.join()
            p3.join()
            p4.join()
            p5.join()
            p6.join()
            print('\nSpent time: ' + str(int(time.time()-start)) + ' seconds')
            
            
CreateData(directoryPath = 'C:/temp/Annual_Reports_France_cleaned', isMultiprocessing = 1)
"""