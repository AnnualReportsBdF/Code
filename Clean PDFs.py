import sys
import re
import os
import shutil
import random
import langdetect
import PyPDF2
import textract
from PIL import Image
import pdf2image
from pypdfocr.pypdfocr_gs import PyGs
import pytesseract
pytesseract.pytesseract.tesseract_cmd = 'Lib/site-packages/pytesseract/tesseract/tesseract'

import warnings
warnings.filterwarnings("ignore")


def CountPagesNumber(filePath):
    pdfReader = PyPDF2.PdfFileReader(filePath)
    return pdfReader.numPages


def ConvertPdftoText(filePath):
    try:
        text = unicode(textract.process(filePath), 'utf-8')
        return text
    except ValueError:
        text = ""
        pagesNumber = CountPagesNumber(filePath)
        pdf = PyPDF2.PdfFileReader(filePath)
        if pdf.isEncrypted:
            pdf.decrypt('')
        for ith in range(pagesNumber):
            page = pdf.getPage(ith)
            text += page.extractText()
        text = unicode(text.encode('utf-8'), 'utf-8')
        return text
    
    
def ConvertScanToText(filePath, language):
    text = ''
    pagesNumber = CountPagesNumber(filePath)
    PyGs({}).make_img_from_pdf(filePath)
    for ith in range(pagesNumber):
        imagePath = filePath[:-4] + '_' + str(ith + 1) + '.JPG'
        image = Image.open(imagePath, mode='r')
        text += pytesseract.image_to_string(image, lang = language)
        os.remove(imagePath)
    return unicode(text.encode('utf-8'), 'utf-8')


def isScannedFile(filePath):
    text = ConvertPdftoText(filePath)
    pagesNumber = CountPagesNumber(filePath)
    if text in ['\x0c' * pagesNumber, '']:
        scannedFile = 1
    else:
        scannedFile = 0
    return scannedFile


def ConvertFileToText(filePath, language):
    text = ConvertPdftoText(filePath)
    pagesNumber = CountPagesNumber(filePath)
    if text in ['\x0c' * pagesNumber, '']:
        text = ConvertScanToText(path, language)
    return text


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
        lang = langdetect.detect(ConvertPdftoText(filePath))
    lang = LanguageName(str(lang)[:2])
    return lang


def RetrieveFileInfosAfterRenamingWithPdfType(filePath, pdfTypes):
    fileName = filePath[filePath.rfind('/')+1:-4]
    for type in pdfTypes:
        if (fileName.find(type) != -1):
            pdfType = fileName[fileName.find(type):fileName.find(type)+len(type)]
    language = fileName[fileName.find(pdfType)-4:fileName.find(pdfType)-1]
    year = fileName[fileName.find(language)-5:fileName.find(language)-1]
    country = fileName.split(' ')[0]
    companyName = fileName[fileName.find(country)+len(country)+1:fileName.find(year)-1]
    return country, companyName, year, language, pdfType


def KeepOnlyPdfsNeeded(directoryPath, pdfTypes):
    count = 0
    ithFile = 0
    for country in os.listdir(directoryPath):
        if ((country[-4:] not in ['.pdf', '.csv']) & (country != 'Duplicates')):
            for year in os.listdir(directoryPath + '/' + country):
                for pdf in os.listdir(directoryPath + '/' + country + '/' + year):
                    count += 1
    sys.stdout.write("[" + count * " " + "] 0%")           
    for country in os.listdir(directoryPath):
        if ((country[-4:] not in ['.pdf', '.csv']) & (country != 'Duplicates')):
            for year in os.listdir(directoryPath + '/' + country):
                for pdf in os.listdir(directoryPath + '/' + country + '/' + year):
                    fileToDelete = ''
                    ithFile += 1
                    filePath = directoryPath + '/' + country + '/' + year + '/' + pdf
                    for type in pdfTypes:
                        if (pdf.find(type) == -1):
                            fileToDelete += 'y'
                        else:
                            fileToDelete += 'n'
                    if (fileToDelete == 'y' * len(pdfTypes)):
                        os.remove(filePath)
                    percentage = int(ithFile * 100 / count)
                    spacesNumber = count - ithFile
                    sys.stdout.write("\r")
                    sys.stdout.write("[" + ithFile * "#" + spacesNumber * " " + "] " + str(percentage) + "%")
    print('\nKeep only PDFs needed: OK\n')
    
    
def RenamePdfs(directoryPath, pdfTypes):
    count = 0
    ithFile = 0
    for country in os.listdir(directoryPath):
        if ((country[-4:] not in ['.pdf', '.csv']) & (country != 'Duplicates')):
            countryLanguage = FindCountryLanguage(country)
            for year in os.listdir(directoryPath + '/' + country):
                for pdf in os.listdir(directoryPath + '/' + country + '/' + year):
                    count += 1
    sys.stdout.write("[" + count * " " + "] 0%")
    duplicates = []
    if not (os.path.exists(directoryPath + '/Duplicates')):
        os.makedirs(directoryPath +  '/Duplicates')
    for country in os.listdir(directoryPath):
        if ((country[-4:] not in ['.pdf', '.csv']) & (country != 'Duplicates')):
            countryLanguage = FindCountryLanguage(country)
            for year in os.listdir(directoryPath + '/' + country):
                for pdf in os.listdir(directoryPath + '/' + country + '/' + year):
                    ithFile += 1
                    isDuplicate = 0
                    possibleOtherPdfNames = []
                    filePath = directoryPath + '/' + country + '/' + year + '/' + pdf
                    language = DetectLanguage(filePath, countryLanguage)
                    pdfType = ""
                    for type in pdfTypes:
                        if (pdf.find(type) != -1):
                            pdfType = type
                    companyName = pdf[:pdf.rfind(((pdf[:pdf.rfind(pdfType)-1]).split())[-1])-1]
                    pdfName = country + ' ' + companyName + ' ' + str(year) + ' ' + language + ' ' + pdfType + '.pdf'
                    # Find all possible other pdf names
                    for type in pdfTypes:
                        if (pdf.find(type) != -1):
                            if (countryLanguage != 'eng'):
                                if (language != countryLanguage) :
                                    possibleOtherPdfNames.append(country + ' ' + companyName + ' ' + str(year) + ' ' 
                                                                 + countryLanguage + ' ' + type + '.pdf')
                                else:
                                    possibleOtherPdfNames.append(country + ' ' + companyName + ' ' + str(year) + ' ' 
                                                                 + 'eng' + ' ' + type + '.pdf')  
                        else:
                            for language in [countryLanguage, 'eng']:
                                possibleOtherPdfNames.append(country + ' ' + companyName + ' ' + str(year) + ' ' 
                                                             + language + ' ' + type + '.pdf')
                    # Si le fichier existe déjà dans le dossier cible
                    if os.path.exists(directoryPath + '/' + pdfName):
                        duplicates.append(pdfName) # on garde dans une liste le 1er doublon pour le supprimer plus tard et
                                                   # continuer de le comparer avec les autres documents
                        isDuplicate = 1
                        pdfName = pdfName[:-4] + ' 2.pdf'
                        number = 3
                        while (os.path.exists(directoryPath + '/Duplicates/' + pdfName)):
                            pdfName = pdfName[:-6] + ' ' + str(number) + '.pdf'
                            number += 1 
                    else:
                        for possiblePdfName in possibleOtherPdfNames:
                            if (os.path.exists(directoryPath + '/' + possiblePdfName)):
                                isDuplicate = 1
                                duplicates.append(possiblePdfName) # Même raison
                                if (os.path.exists(directoryPath + '/Duplicates/' + pdfName)):
                                    pdfName = pdfName[:-4] + ' 2.pdf'
                                    number = 3
                                    while (os.path.exists(directoryPath + '/Duplicates/' + pdfName)):
                                        pdfName = pdfName[:-6] + ' ' + str(number) + '.pdf'
                                        number += 1 
                    if (isDuplicate == 0):
                        os.rename(filePath, directoryPath + '/' + pdfName)
                    else:
                        os.rename(filePath, directoryPath + '/Duplicates/' + pdfName)
                    percentage = int(ithFile * 100 / (count+1))
                    spacesNumber = (count+1) - ithFile
                    sys.stdout.write("\r")
                    sys.stdout.write("[" + ithFile * "#" + spacesNumber * " " + "] " + str(percentage) + "%")        
    for pdfName in set(duplicates):
        if (os.path.exists(directoryPath + '/' + pdfName)):
            os.rename(directoryPath + '/' + pdfName, directoryPath + '/Duplicates/' + pdfName)
    for pdf in os.listdir(directoryPath):
        if (pdf[-4:] == '.pdf'):
            for type in pdfTypes:
                if (pdf.find(type) != -1):
                    newPdfName = pdf[:pdf.find(type)-1] + '.pdf'
                    os.rename(directoryPath + '/' + pdf, directoryPath + '/' + newPdfName)
    for country in os.listdir(directoryPath):
        if ((country[-4:] not in ['.pdf', '.csv']) & (country != 'Duplicates')):
            countryLanguage = FindCountryLanguage(country)
            for year in os.listdir(directoryPath + '/' + country):
                if (len(os.listdir(directoryPath + '/' + country + '/' + year)) == 0):
                    shutil.rmtree(directoryPath + '/' + country + '/' + year)
            if (len(os.listdir(directoryPath + '/' + country)) == 0):
                shutil.rmtree(directoryPath + '/' + country)  
    sys.stdout.write("\r")
    sys.stdout.write("[" + (count+1) * "#" + "] " + "100%")
    print('\nRename all PDFs: OK\n')
    
    
def FindPdfsToDeleteOrMerge(duplicatesDirectoryPath, pdfTypes):
    pdfNumber = len(os.listdir(duplicatesDirectoryPath))
    pdfsToDelete = []
    pdfsPairsToMerge = []
    for element1 in os.listdir(duplicatesDirectoryPath):
        element1Infos = RetrieveFileInfosAfterRenamingWithPdfType(duplicatesDirectoryPath + '/' + element1, pdfTypes)
        shortName1 = element1Infos[0] + ' ' + element1Infos[1] + ' ' + element1Infos[2]
        for element2 in os.listdir(duplicatesDirectoryPath):
            if (os.listdir(duplicatesDirectoryPath).index(element2) >  os.listdir(duplicatesDirectoryPath).index(element1)):
                element2Infos = RetrieveFileInfosAfterRenamingWithPdfType(duplicatesDirectoryPath + '/' + element2, pdfTypes)
                shortName2 = element2Infos[0] + ' ' + element2Infos[1] + ' ' + element2Infos[2]
                if (shortName1 == shortName2):
                    if ((element1Infos[4] != element2Infos[4]) 
                        & (element1Infos[4] in pdfTypes) 
                        & (element2Infos[4] in pdfTypes)): 
                        # S'ils n'ont pas le même type de pdf, on garde ARS-None (= on supprime FullYear)
                        if (pdfTypes.index(element1Infos[4]) < pdfTypes.index(element2Infos[4])):
                            pdfsToDelete.append(element2)
                        else:
                            pdfsToDelete.append(element1)               
                    else:                       # S'ils ont le même type et une langue différente, on garde le pdf en anglais
                        if (element1Infos[3] != element2Infos[3]):
                            if ((element1Infos[3] == 'eng') | (element2Infos[3] == 'eng')):
                                if (element1Infos[3] != 'eng'):
                                    pdfsToDelete.append(element1)
                                else:
                                    pdfsToDelete.append(element2)
                            else:
                                if ((element1Infos[3] == element1CountryLanguage) 
                                    | (element2Infos[3] == element2CountryLanguage)):
                                    if (element1Infos[3] != element1CountryLanguage):
                                        pdfsToDelete.append(element1)
                                    else:
                                        pdfsToDelete.append(element2)
                                else:
                                    # Aucun des pdfs est en anglais ou dans la langue du pays => on supprime les 2?
                                    pdfsToDelete.append(element1)
                                    pdfsToDelete.append(element2)
    for element1 in os.listdir(duplicatesDirectoryPath):
        element1Infos = RetrieveFileInfosAfterRenamingWithPdfType(duplicatesDirectoryPath + '/' + element1, pdfTypes)
        shortName1 = element1Infos[0] + ' ' + element1Infos[1] + ' ' + element1Infos[2]
        for element2 in os.listdir(duplicatesDirectoryPath):
            if ((element1 not in pdfsToDelete) 
                & (element2 not in pdfsToDelete) 
                & (os.listdir(duplicatesDirectoryPath).index(element2) >  os.listdir(duplicatesDirectoryPath).index(element1))):
                element2Infos = RetrieveFileInfosAfterRenamingWithPdfType(duplicatesDirectoryPath + '/' + element2, pdfTypes)
                shortName2 = element2Infos[0] + ' ' + element2Infos[1] + ' ' + element2Infos[2]
                if (shortName1 == shortName2):
                    element1Text = re.sub(r" {2,}", " ", re.sub(r"[^A-Za-z ]+", " ", 
                                   ConvertFileToText(duplicatesDirectoryPath + '/' + element1, element1Infos[3])))
                    element2Text = re.sub(r" {2,}", " ", re.sub(r"[^A-Za-z ]+", " ", 
                                   ConvertFileToText(duplicatesDirectoryPath + '/' + element2, element2Infos[3])))
                    if (element1Text == element2Text):
                        if (len(element1) != min(len(element1), len(element2))):
                            pdfsToDelete.append(element1)
                        else:
                            pdfsToDelete.append(element2)
                    else:
                        if (len(element1Text) > len(element2Text)):
                            if (element1Text.find(element2Text) != -1): #element2Text est dans element1Text
                                pdfsToDelete.append(element2)
                            else: #pdfs différents => on merge
                                pdfsPairsToMerge.append([element1, element2])         
                        else:
                            if (element2Text.find(element1Text) != -1):
                                pdfsToDelete.append(element1)
                            else:
                                pdfsPairsToMerge.append([element1, element2])                       
    pdfsToDelete = set(pdfsToDelete)
    return pdfsToDelete, pdfsPairsToMerge


def DeleteDuplicates(duplicatesDirectoryPath, pdfTypes):
    pdfsNumber = len(os.listdir(duplicatesDirectoryPath))
    directoryPath = duplicatesDirectoryPath[0:duplicatesDirectoryPath.find('/Duplicates')]
    pdfsWhichNeedAction = FindPdfsToDeleteOrMerge(duplicatesDirectoryPath, pdfTypes)
    pdfsToDelete = pdfsWhichNeedAction[0]
    pdfsPairsToMerge = pdfsWhichNeedAction[1]
    pdfsToMerge = []
    # PDFs to merge later
    for pdf in pdfsPairsToMerge:
        pdfsToMerge.append(pdf[0])
        pdfsToMerge.append(pdf[1])
    pdfsToMerge = set(pdfsToMerge)
    # PDFs to keep in non-duplicates
    for pdf in os.listdir(duplicatesDirectoryPath):
        if ((pdf not in pdfsToDelete) & (pdf not in pdfsToMerge)):
            for type in pdfTypes:
                if (pdf.find(type) != -1):
                    newPdfName = pdf[:pdf.find(type)-1] + '.pdf'
            os.rename(duplicatesDirectoryPath + '/' + pdf, directoryPath + '/' + newPdfName)
    # PDFs to delete
    for pdf in pdfsToDelete:
        os.remove(duplicatesDirectoryPath + '/' + pdf)
    sys.stdout.write("[" + pdfsNumber * "#" + "] 100%")
    print('\nDelete duplicates: OK')
    
    
def CleanPdfs(directoryPath, pdfTypes):
    duplicatesDirectoryPath = directoryPath + '/Duplicates'
    KeepOnlyPdfsNeeded(directoryPath, pdfTypes)
    RenamePdfs(directoryPath, pdfTypes)
    DeleteDuplicates(duplicatesDirectoryPath, pdfTypes)
    
    
CleanPdfs(directoryPath = "Annual reports", pdfTypes = ['ARS-None', 'FullYear', 'ARS', 'AR S'])