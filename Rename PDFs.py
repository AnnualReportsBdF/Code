import re
import os
import shutil
import langdetect
import PyPDF2
import random
import textract
import pdf2image
import pytesseract
pytesseract.pytesseract.tesseract_cmd = 'Lib/site-packages/pytesseract/tesseract/tesseract'


def CountPagesNumber(filePath):
    pdfReader = PyPDF2.PdfFileReader(filePath)
    return pdfReader.numPages


def isScannedFile(filePath):
    text = textract.process(filePath)
    pagesNumber = CountPagesNumber(filePath)
    scannedFile = 0
    if text == '\x0c' * pagesNumber:
        scannedFile = 1
    return scannedFile


def FindCountryLanguage(country):
    if (country == 'France'):
        countryLanguage = 'fra'
    elif (country == 'Spain'):
        countryLanguage = 'spa'
    elif (country == 'Italy'):
        countryLanguage = 'ita'
    elif (country == 'Germany'):
        countryLanguage = 'deu'
    elif (country == 'England'):
        countryLanguage = 'eng'
    #https://github.com/tesseract-ocr/tesseract/wiki/Data-Files
    return countryLanguage


def LanguageName(language):
    if (language == 'fr'):
        language = 'fra'
    elif (language == 'es'):
        language = 'spa'
    elif (language == 'it'):
        language = 'ita'
    elif (language == 'de'):
        language = 'deu'
    elif (language == 'en'):
        language = 'eng'
    #https://fr.wikipedia.org/wiki/Liste_des_codes_ISO_639-1
    return language


def DetectLanguageForScan(filePath, countryLanguage, sample):
    pdfReader = PyPDF2.PdfFileReader(filePath)
    pdfWriter = PyPDF2.PdfFileWriter()
    pagesNumber = CountPagesNumber(filePath)
    lan1 = ""
    temporaryFolderPath = "Temporary Folder"
    if os.path.exists(temporaryFolderPath):
        shutil.rmtree(temporaryFolderPath)
    os.makedirs(temporaryFolderPath)
   
    while (len(lan1) < sample):
        randomPage = random.sample(range(0, pagesNumber), 1)[0]
        pdfWriter.addPage(pdfReader.getPage(randomPage))
        temporaryFilePath = temporaryFolderPath +"/out_" + str(randomPage) + ".pdf"
        
        stream = open(temporaryFilePath, "wb")
        pdfWriter.write(stream)
        stream.close()
        
        lan1 = "" + pytesseract.image_to_string(pdf2image.convert_from_path(temporaryFilePath)[0], lang=countryLanguage)
        
        if (len(lan1) < sample):
            os.remove(temporaryFilePath)
        
    lan2 = "" + pytesseract.image_to_string(pdf2image.convert_from_path(temporaryFilePath)[0],lang='eng')
    shutil.rmtree(temporaryFolderPath)
    out_lan1 = langdetect.detect_langs(lan1)[0]
    out_lan2 = langdetect.detect_langs(lan2)[0]
    lang=re.findall(r"[a-zA-Z]+",str(max(out_lan1,out_lan2)))[0]
    
    return lang


def DetectLanguage(filePath, countryLanguage):
    if isScannedFile(filePath) == 1:
        sample = 300
        lang = DetectLanguageForScan(filePath, countryLanguage, sample)
    else:
        lang = langdetect.detect(unicode(textract.process(filePath), 'utf-8'))
    lang = LanguageName(lang)
    return lang


def KeepOnlyPdfsNeeded(directoryPath):
    fullYearPdfName = []
    arsNonePdfName = []
    arsPdfName = []
    pdfName = ""
    FilesNumber = 0
    DeletedFilesNumber = 0
    
    for country in os.listdir(directoryPath):
        for year in os.listdir(directoryPath + '/' + country):
            for pdf in os.listdir(directoryPath + '/' + country + '/' + year):
                filePath = directoryPath + '/' + country + '/' + year + '/' + pdf
                FilesNumber += 1
        
                if pdf.find('FullYear') == -1 & pdf.find('ARS-None') == -1 & pdf.find('ARS') == -1 & pdf.find('AR S') == -1:
                    DeletedFilesNumber += 1
                    os.remove(filePath)
            
    print('Keep only PDFs needed: OK')
    print(str(FilesNumber - DeletedFilesNumber) + ' out of ' + str(FilesNumber) + ' files were kept.')

    
def RenamePdfs(directoryPath, newDirectoryPath):
    if os.path.exists(newDirectoryPath):
        shutil.rmtree(newDirectoryPath)
    os.makedirs(newDirectoryPath)
    for country in os.listdir(directoryPath):
        for year in os.listdir(directoryPath + '/' + country):
            for pdf in os.listdir(directoryPath + '/' + country + '/' + year):
                filePath = directoryPath + '/' + country + '/' + year + '/' + pdf
                countryLanguage = FindCountryLanguage(country)
                language = DetectLanguage(filePath, countryLanguage)
                pdfType = ""
                companyName = ""
        
                if pdf.find('SA') != -1:
                    companyName = pdf[0:pdf.find('SA')-1]
                elif pdf.find('SE') != -1:
                    companyName = pdf[0:pdf.find('SE')-1]
                elif pdf.find('SCA') != -1:
                    companyName = pdf[0:pdf.find('SCA')-1]
            
                if pdf.find('FullYear') != -1:
                    pdfType = 'FullYear'
                elif pdf.find('ARS-None') != -1:
                    pdfType = 'ARS-None'
                elif pdf.find('ARS') != -1:
                    pdfType = 'ARS'
                elif pdf.find('AR S') != -1:
                    pdfType = 'AR S'
            
                pdfName = country + ' ' + companyName + ' ' + str(year) + ' ' + language + ' ' + pdfType + '.pdf'

                if os.path.exists(newDirectoryPath + '/' + pdfName):
                    pdfName = pdfName[:-4] + ' 2.pdf'
                    if os.path.exists(newDirectoryPath + '/' + pdfName):
                        pdfName = pdfName[:-6] + ' 3.pdf'
                        if os.path.exists(newDirectoryPath + '/' + pdfName):
                            pdfName = pdfName[:-6] + ' 4.pdf'
                            if os.path.exists(newDirectoryPath + '/' + pdfName):
                                pdfName = pdfName[:-6] + ' 5.pdf'
                
                shutil.copyfile(filePath, newDirectoryPath + '/' + pdfName)
    print('Rename all PDFs: OK')

    
KeepOnlyPdfsNeeded(directoryPath = "France base")
RenamePdfs(directoryPath = "France base", newDirectoryPath = "Annual reports")
 
