import pandas as pd
import sys
import os
import io
import multiprocessing
import time
#Convert pdf to text
import textract
import PyPDF2
from PIL import Image
from pypdfocr.pypdfocr_gs import PyGs
import pytesseract
pytesseract.pytesseract.tesseract_cmd = 'Lib/site-packages/pytesseract/tesseract/tesseract'
import pdfminer
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
import langdetect
import csv
import warnings
warnings.filterwarnings("ignore")


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


def RetrieveInfosCompany(path):
    fileName = path[path.rfind('/')+1:-4]
    words = fileName.split(' ')
    language = words[-1]
    year = words[-2]
    country = words[0]
    companyName = ''
    for i in range(len(words)-3):
        i += 1
        companyName += words[i] + ' '
    companyName = companyName[:-1]
    return country, companyName, year, language


def CountPagesNumber(path):
    pdfReader = PyPDF2.PdfFileReader(path)
    return pdfReader.numPages


def ConvertPdftoText(path):
    import unidecode
    try:
        text = unicode(textract.process(path), 'utf-8')
        return text
    except ValueError:
        text = ""
        pagesNumber = CountPagesNumber(path)
        pdf = PyPDF2.PdfFileReader(path)
        if pdf.isEncrypted:
            pdf.decrypt('')
        for ith in range(pagesNumber):
            page = pdf.getPage(ith)
            text += page.extractText()
        text = unicode(text, 'utf-8').encode('utf-8')
        return text
    
    
def ConvertScanToText(path, language):
    text = ''
    pagesNumber = CountPagesNumber(path)
    PyGs({}).make_img_from_pdf(path)
    for ith in range(pagesNumber):
        imagePath = path[:-4] + '_' + str(ith + 1) + '.JPG'
        image = Image.open(imagePath, mode='r')
        text += pytesseract.image_to_string(image, lang = language)
        os.remove(imagePath)
    return unicode(text.encode('utf-8'), 'utf-8')


def ConvertFileToText(path, language):
    text = ConvertPdftoText(path)
    scannedFile = 0
    pagesNumber = CountPagesNumber(path)
    if text in ['\x0c' * pagesNumber, '']:
        scannedFile = 1
        text = ConvertScanToText(path, language)
    languageEstimated = str(langdetect.detect_langs(text))
    languageEstimated = str(LanguageName(languageEstimated[1:3]))
    if ((languageEstimated != language) & (scannedFile == 0)):
        prm = PDFResourceManager()
        bio = io.BytesIO()
        device = TextConverter(prm, bio, codec = 'utf-8', laparams = LAParams())
        pdf = open(path, 'rb')
        interpreter = PDFPageInterpreter(prm, device)
        for page in PDFPage.get_pages(pdf, set(), maxpages = 0, password = "", caching = True, check_extractable = True):
            interpreter.process_page(page)
        text = bio.getvalues()
        pdf.close()
        device.close()
        bio.close()
        languageEstimated = str(langdetect.detect_langs(text))
        languageEstimated = str(LanguageName(languageEstimated[1:3]))
    return text, scannedFile, languageEstimated


def CreateDataWithoutDuplicates(directoryPath):
    count = 0
    ithFile = 0
    for file in os.listdir(directoryPath):
        if (file[-4:] == '.pdf'):
            count += 1
    sys.stdout.write("[" + count * " " + "] 0%")
    dataDic = []
    for file in os.listdir(directoryPath):
        if (file[-4:] == '.pdf'):
            ithFile += 1
            path = directoryPath + '/' + file
            pagesNumber = CountPagesNumber(path)
            infosFile = RetrieveInfosCompany(path)
            infosText = ConvertFileToText(path, infosFile[3])
        
            dataDic.append({'Country': infosFile[0], 'Company': infosFile[1], 'Year': infosFile[2], 
                        'Text': infosText[0].encode('utf-8'), 
                        'Scan': infosText[1],'Pages Number': pagesNumber, 'Language Expected': infosFile[3], 
                        'Language Estimated': infosText[2]})
        
            percentage = int(ithFile * 100 / count)
            spacesNumber = count - ithFile
            sys.stdout.write("\r")
            sys.stdout.write("[" + ithFile * "#" + spacesNumber * " " + "] " + str(percentage) + "%")
    df = pd.DataFrame(dataDic)
    df.to_csv(directoryPath[:directoryPath.rfind('/')+1] + str(directoryPath[directoryPath.rfind('/')+1:]) + '-data.csv')
    print('\nCreate Data Without Duplicates: OK')
    
    
def CreateDataWithDuplicates(directoryPath):  
    pdfTypes = ['ARS-None', 'FullYear', 'ARS', 'AR S']
    shortNames = []
    pdfsToMerge = []
    dataDic = []
    for duplicate in os.listdir(directoryPath):
         if (duplicate[-4:] == '.pdf'):
            lastWord = ((duplicate[:-4]).split())[-1]
            if (lastWord.isdigit()):
                shortNames.append(duplicate[:duplicate.rfind(lastWord)-1])
            else :
                shortNames.append(duplicate[:-4])    
    count = len(set(shortNames))
    sys.stdout.write("[" + count * " " + "] 0%")
    dictShortNames = {x:shortNames.count(x) for x in shortNames}
    for i in range(len(dictShortNames)):
        shortName = dictShortNames.keys()[i]
        numberOfDuplicates = dictShortNames.values()[i]
        jthFile = 0
        text = ''
        isScans = []
        isScan = ""
        estimatedLanguages = []
        estimatedLanguage = ""
        pagesTotalNumber = 0
        while (jthFile < numberOfDuplicates):                 
            if (jthFile == 0):
                path = directoryPath + '/' + shortName + '.pdf'
            else:
                path = directoryPath + '/' + shortName + ' ' + str(jthFile+1) + '.pdf'
            
            if (os.path.isfile(path) == False):
                numberOfDuplicates += 1
            else:
                for type in pdfTypes:
                    if (shortName.rfind(type) != -1):
                        infosFile = RetrieveInfosCompany(directoryPath + '/' + shortName[:shortName.rfind(type)-1] + '.pdf')
                pagesNumber = CountPagesNumber(path)
                infosText = ConvertFileToText(path, infosFile[3])
                text += infosText[0] + ' '
                isScans.append(infosText[1])
                pagesTotalNumber += pagesNumber
                estimatedLanguages.append(infosText[2])
            jthFile += 1
        for language in set(estimatedLanguages):
            estimatedLanguage += language + " "
        for scan in set(isScans):
            isScan += str(scan) + " "
        dataDic.append({'Country': infosFile[0], 'Company': infosFile[1], 'Year': infosFile[2], 
                        'Text': text.encode('utf-8'), 
                        'Scan': isScan[:-1],'Pages Number': pagesTotalNumber, 'Language Expected': infosFile[3], 
                        'Language Estimated': estimatedLanguage[:-1]})
        percentage = int((i+1) * 100 / count)
        spacesNumber = count - (i+1)
        sys.stdout.write("\r")
        sys.stdout.write("[" + (i+1) * "#" + spacesNumber * " " + "] " + str(percentage) + "%")
    df = pd.DataFrame(dataDic)
    df.to_csv(directoryPath[:directoryPath.rfind('/')+1] + str(directoryPath[directoryPath.rfind('/')+1:]) + '-data.csv')
    print('\nCreate Data With Duplicates: OK')
    
    
def CreateData(directoryPath, isMultiprocessing):
    if (isMultiprocessing == 0):
        start = time.time()
        for folder in os.listdir(directoryPath):
            if (folder[-4:] not in ['.csv', '.JPG', '.pdf']):
                print('\n' + str(folder))
                if (folder != 'Duplicates'):
                    CreateDataWithoutDuplicates(directoryPath + '/' + folder)
                else:
                    CreateDataWithDuplicates(directoryPath + '/Duplicates')
        print('\nSpent time: ' + str(int(time.time()-start)) + ' seconds')
    else:
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
        
        
CreateData(directoryPath = 'Annual reports', isMultiprocessing = 0)