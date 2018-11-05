# -*- coding: utf-8 -*-
"""
Created on Thu Jul 26 10:32:17 2018

@author: R550427
"""

import sys
import re
import os
import shutil
import random
import langdetect
import PyPDF2
import textract
from PIL import Image
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
    except:
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
        text = ConvertScanToText(filePath, language)
    
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


def DetectLanguageForScan(filePath, directoryPath, countryLanguage, sample):
    pdfReader = PyPDF2.PdfFileReader(filePath)
    if pdfReader.isEncrypted:
            pdfReader.decrypt('')
    pagesNumber = CountPagesNumber(filePath)
    lan1 = ""
    temporaryFolderPath = directoryPath + "/Temporary_Folder"
    if os.path.exists(temporaryFolderPath):
        shutil.rmtree(temporaryFolderPath)
    os.makedirs(temporaryFolderPath)
    
    while (len(lan1) < sample):
        pdfWriter = PyPDF2.PdfFileWriter()
        randomPage = random.sample(range(0, pagesNumber), 1)[0]
        pdfWriter.addPage(pdfReader.getPage(randomPage))
        temporaryFilePath = temporaryFolderPath +"/out_" + str(randomPage) + ".pdf"
        
        stream = open(temporaryFilePath, "wb")
        pdfWriter.write(stream)
        stream.close()
        
        PyGs({}).make_img_from_pdf(temporaryFilePath)
        imagePath = temporaryFilePath[:-4] + '_1.JPG'
        image = Image.open(imagePath, mode='r')
        
        lan1 = "" + pytesseract.image_to_string(image, lang = countryLanguage)
 
        if (len(lan1) < sample):
            for file in os.listdir(temporaryFolderPath):
                os.remove(temporaryFolderPath + '/' + file)
    
    lan2 = "" + pytesseract.image_to_string(image,lang='eng')
    shutil.rmtree(temporaryFolderPath)
    out_lan1 = langdetect.detect_langs(lan1)[0]
    out_lan2 = langdetect.detect_langs(lan2)[0]
    lang=re.findall(r"[a-zA-Z]+",str(max(out_lan1,out_lan2)))[0]
    
    return lang


def DetectLanguage(filePath, directoryPath, countryLanguage):
    if isScannedFile(filePath) == 1:
        sample = 300
        lang = DetectLanguageForScan(filePath, directoryPath, countryLanguage, sample)
    else:
        lang = langdetect.detect(ConvertPdftoText(filePath))
    lang = LanguageName(str(lang)[:2])
    return lang


def RetrieveFileInfosAfterRenamingWithPdfType(filePath, pdfTypes):
    fileName = filePath[filePath.rfind('/')+1:-4]
    for type in pdfTypes:
        if (fileName.find(type) != -1):
            pdfType = fileName[fileName.find(type):fileName.find(type)+len(type)]
            break
    
    language = fileName[fileName.find(pdfType)-4:fileName.find(pdfType)-1]
    year = fileName[fileName.find(language)-5:fileName.find(language)-1]
    country = fileName.split(' ')[0]
    companyName = fileName[fileName.find(country)+len(country)+1:fileName.find(year)-1]
    return country, companyName, year, language, pdfType


def KeepOnlyPdfsNeeded(directoryPath, pdfTypes):
    count = 0
    ithFile = 0
    for country in os.listdir(directoryPath):
        if ((country[-4] != '.') & (country not in ['Duplicates', 'Temporary_Folder', 'Error'])):
            for year in os.listdir(directoryPath + '/' + country):
                for pdf in os.listdir(directoryPath + '/' + country + '/' + year):
                    count += 1
    sys.stdout.write("[" + count * " " + "] 0%")           
    
    for country in os.listdir(directoryPath):
        if ((country[-4] != '.') & (country not in ['Duplicates', 'Temporary_Folder', 'Error'])):
            for year in os.listdir(directoryPath + '/' + country):
                for pdf in os.listdir(directoryPath + '/' + country + '/' + year):
                    try:
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
                    except:
                        print('\n ERROR KeepOnlyPdfsNeeded: ' + str(filePath))
                        pass
                    
    print('\nKeep only PDFs needed: OK\n')
    
    
def RenamePdfs(directoryPath, pdfTypes):
    count = 0
    ithFile = 0
    for country in os.listdir(directoryPath):
        if ((country[-4] != '.') & (country not in ['Duplicates', 'Temporary_Folder', 'Error'])):
            countryLanguage = FindCountryLanguage(country)
            for year in os.listdir(directoryPath + '/' + country):
                for pdf in os.listdir(directoryPath + '/' + country + '/' + year):
                    count += 1
    sys.stdout.write("[" + count * " " + "] 0%")
    
    duplicates = []
    
    if not (os.path.exists(directoryPath + '/Duplicates')):
        os.makedirs(directoryPath +  '/Duplicates')
        
    if not (os.path.exists(directoryPath + '/Error')):
        os.makedirs(directoryPath +  '/Error')
    
    for country in os.listdir(directoryPath):
        if ((country[-4] != '.') & (country not in ['Duplicates', 'Temporary_Folder', 'Error'])):
            countryLanguage = FindCountryLanguage(country)
            for year in os.listdir(directoryPath + '/' + country):
                for pdf in os.listdir(directoryPath + '/' + country + '/' + year):
                    ithFile += 1
                    isDuplicate = 0
                    possibleOtherPdfNames = []
                    filePath = directoryPath + '/' + country + '/' + year + '/' + pdf
                    pdfType = ""
                    pdfName = ""
                    
                    try:
                        try:
                            pdfLanguage = DetectLanguage(filePath, directoryPath, countryLanguage)
                        except:
                            print('\n DetectLanguage failed: ' + filePath) 
                            os.rename(filePath, directoryPath + '/Error/' + filePath[filePath.rfind('/')+1:])
                            pass
    
                        if (pdfLanguage in [countryLanguage, 'eng']):
                            for type in pdfTypes:
                                if (pdf.find(type) != -1):
                                    pdfType = type
                                    companyName = pdf[:pdf.rfind(((pdf[:pdf.rfind(pdfType)-1]).split())[-1])-1]
                                    pdfName = country + ' ' + companyName + ' ' + str(year) + ' ' + pdfLanguage + ' ' + pdfType + '.pdf'
                                    break
                    
                            # Find all possible other pdf names
                            for type in pdfTypes:
                                if (type == pdfType):
                                    if (countryLanguage != 'eng'):
                                        if (pdfLanguage != countryLanguage) :
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
                        else:
                            os.remove(filePath)
                        
                
                        if (isDuplicate == 0):
                            if (os.path.exists(directoryPath + '/' + pdfName)):
                                if (os.path.exists(directoryPath + '/Error/' + pdfName)):
                                    pdfName = pdfName[:-4] + ' 2.pdf'
                                    number = 3
                                    while (os.path.exists(directoryPath + '/Error/' + pdfName)):
                                        pdfName = pdfName[:-6] + ' ' + str(number) + '.pdf'
                                        number += 1 
                                    os.rename(filePath, directoryPath + '/Error/' + pdfName)
                                else:
                                    os.rename(filePath, directoryPath + '/Error/' + pdfName)
                            else:
                                os.rename(filePath, directoryPath + '/' + pdfName)
                        else:
                            if (os.path.exists(directoryPath + '/Duplicates/' + pdfName)):
                                if (os.path.exists(directoryPath + '/Error/' + pdfName)):
                                    pdfName = pdfName[:-4] + ' 2.pdf'
                                    number = 3
                                    while (os.path.exists(directoryPath + '/Error/' + pdfName)):
                                        pdfName = pdfName[:-6] + ' ' + str(number) + '.pdf'
                                        number += 1 
                                    os.rename(filePath, directoryPath + '/Error/' + pdfName)
                                else:
                                    os.rename(filePath, directoryPath + '/Error/' + pdfName)
                            else:
                                os.rename(filePath, directoryPath + '/Duplicates/' + pdfName)
                                
                                
                                
                    
                        percentage = int(ithFile * 100 / (count+1))
                        spacesNumber = (count+1) - ithFile
                        sys.stdout.write("\r")
                        sys.stdout.write("[" + ithFile * "#" + spacesNumber * " " + "] " + str(percentage) + "%")
                        
                    except:
                        print('\n ERROR RenamePdfs: ' + str(filePath))
                        try:
                            os.rename(filePath, directoryPath + '/Error/' + filePath[filePath.rfind('/')+1:])
                        except:
                            try:
                                os.rename(filePath, directoryPath + '/Error/' + pdfName)
                            except:
                                pass
                            pass
                        pass
                    
    for pdfName in set(duplicates):
        if (os.path.exists(directoryPath + '/' + pdfName)):
            os.rename(directoryPath + '/' + pdfName, directoryPath + '/Duplicates/' + pdfName)
        
    for pdf in os.listdir(directoryPath):
        if (pdf[-4:] == '.pdf'):
            for type in pdfTypes:
                if (pdf.find(type) != -1):
                    newPdfName = pdf[:pdf.find(type)-1] + '.pdf'
                    os.rename(directoryPath + '/' + pdf, directoryPath + '/' + newPdfName)
                    break
    
    for country in os.listdir(directoryPath):
        if ((country[-4] != '.') & (country not in ['Duplicates', 'Temporary_Folder', 'Error'])):
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
    pdfsToDelete = []
    pdfsPairsToMerge = []

    for element1 in os.listdir(duplicatesDirectoryPath):
        element1Infos = RetrieveFileInfosAfterRenamingWithPdfType(duplicatesDirectoryPath + '/' + element1, pdfTypes)
        shortName1 = element1Infos[0] + ' ' + element1Infos[1] + ' ' + element1Infos[2]
        element1CountryLanguage = FindCountryLanguage(element1Infos[0])
        for element2 in os.listdir(duplicatesDirectoryPath):
            if (os.listdir(duplicatesDirectoryPath).index(element2) >  os.listdir(duplicatesDirectoryPath).index(element1)):
                element2Infos = RetrieveFileInfosAfterRenamingWithPdfType(duplicatesDirectoryPath + '/' + element2, pdfTypes)
                shortName2 = element2Infos[0] + ' ' + element2Infos[1] + ' ' + element2Infos[2]
                element2CountryLanguage = FindCountryLanguage(element2Infos[0])
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
            try:
                for type in pdfTypes:
                    if (pdf.find(type) != -1):
                        newPdfName = pdf[:pdf.find(type)-1] + '.pdf'
                        os.rename(duplicatesDirectoryPath + '/' + pdf, directoryPath + '/' + newPdfName)
                        break
            except:
                print('\n ERROR DeleteDuplicates: ' + str(duplicatesDirectoryPath) + '/' + str(pdf))
                pass
    
    # PDFs to delete
    for pdf in pdfsToDelete:
        os.remove(duplicatesDirectoryPath + '/' + pdf)
        
    print("\n[" + pdfsNumber * "#" + "] 100%")
    print('\n Delete duplicates: OK')
    
    
def SplitPdfsIntoFolders(directoryPath, foldersNumber):
    years = []
    ithFolder = 1
    pdfsNumber = 0
    
    for folder in range(foldersNumber):
        if not (os.path.exists(directoryPath + '/Folder' + str(folder+1))):
            os.makedirs(directoryPath + '/Folder' + str(folder+1))
    
    for pdf in os.listdir(directoryPath):
        if (pdf[-4:] == '.pdf'):
            pdfsNumber += 1
            years.append(((pdf[:-4]).split())[-2])
            
    for year in set(years):
        for pdf in os.listdir(directoryPath):
            if (pdf[-4:] == '.pdf'):
                try:
                    if (((pdf[:-4]).split())[-2] == year):
                        if (ithFolder > foldersNumber):
                            ithFolder = 1
                        os.rename(directoryPath + '/' + pdf, directoryPath + '/Folder' + str(ithFolder) + '/' + pdf)
                        ithFolder += 1
                except:
                    print('\n ERROR SplitPdfsIntoFolders: ' + str(directoryPath) + '/' + str(pdf))
                    pass
                    
    print("\n[" + pdfsNumber * "#" + "] 100%")
    print('\n Split PDFs into folders: OK')
    
    
def CleanPdfs(directoryPath):
    pdfTypes = ['ARS-None', 'FullYear', 'ARS', 'AR S']
    foldersNumber = 5
    duplicatesDirectoryPath = directoryPath + '/Duplicates'
    KeepOnlyPdfsNeeded(directoryPath, pdfTypes)
    RenamePdfs(directoryPath, pdfTypes)
    DeleteDuplicates(duplicatesDirectoryPath, pdfTypes)
    SplitPdfsIntoFolders(directoryPath, foldersNumber)
    

CleanPdfs(directoryPath = 'C:/temp/Annual_Reports')
