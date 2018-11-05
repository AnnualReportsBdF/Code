# -*- coding: utf-8 -*-
"""
Created on Thu Jul 26 10:32:17 2018

@author: R550427
"""
import numpy as np
import collections
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
pytesseract.pytesseract.tesseract_cmd = "Lib/site-packages/pytesseract/tesseract/tesseract"

import warnings
warnings.filterwarnings("ignore")

#==============================================================================
def CountPagesNumber(path):
    pdf = PyPDF2.PdfFileReader(path)
    if pdf.isEncrypted:
        pdf.decrypt("")
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
            pdf.decrypt("")
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
def isScannedFile(path):
    text = ConvertPdftoText(path)
    pagesNumber = CountPagesNumber(path)
    if text in ["\x0c" * pagesNumber, ""]:
        scannedFile = 1
    else:
        scannedFile = 0  
    return scannedFile
#==============================================================================
def ConvertFileToText(path, language):
    text = ConvertPdftoText(path)
    pagesNumber = CountPagesNumber(path)
    if text in ["\x0c" * pagesNumber, ""]:
        text = ConvertScanToText(path, language)
    return text
#==============================================================================
def FindCountryLanguage(country):
    countries = ["France", "Spain", "Italy", "Germany", "England", "Netherlands", 
                 "Belgium", "Cyprus", "Estonia", "Finland", "Greece", "Ireland", 
                 "Lithuania", "Luxembourg", "Malta", "Austria", "Portugal", 
                 "Slovenia", "Slovakia"]
    langs = [["fr"], ["es"], ["it"], ["de"], ["en"], ["nl"], ["nl", "fr", "de"], 
             ["el", "tr"], ["et"], ["fi", "sv"], ["el"], ["ga"], ["lt"], 
             ["fr", "de", "lb"], ["mt"], ["de"], ["pt"], ["sl"], ["sk"]]
    #https://github.com/tesseract-ocr/tesseract/wiki/Data-Files
    return langs[countries.index(country)]
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
def DetectLanguageForScan(filePath, directoryPath, countryLanguages, nbCharacters): #countryLanguage: 3 letters, lang: 2 letters
    text = ""
    texts = []
    langs = []
    probs = []
    temporaryFolderPath = directoryPath + "/Temporary_Folder"
    if os.path.exists(temporaryFolderPath):
        shutil.rmtree(temporaryFolderPath)
    os.makedirs(temporaryFolderPath)
    pdf = PyPDF2.PdfFileReader(filePath)
    if pdf.isEncrypted:
            pdf.decrypt('')
    pagesNumber = CountPagesNumber(filePath) 
    
    while (len(text) < nbCharacters):
        pdfWriter = PyPDF2.PdfFileWriter()
        randomPage = random.sample(range(0, pagesNumber), 1)[0]
        pdfWriter.addPage(pdf.getPage(randomPage))
        temporaryFilePath = temporaryFolderPath +"/out_" + str(randomPage) + ".pdf"
        stream = open(temporaryFilePath, "wb")
        pdfWriter.write(stream)
        stream.close()
        PyGs({}).make_img_from_pdf(temporaryFilePath)
        imagePath = temporaryFilePath[:-4] + "_1.JPG"
        image = Image.open(imagePath, mode="r")
        text = "" + pytesseract.image_to_string(image, lang = countryLanguages[0])
        if (len(text) < nbCharacters):
            for file in os.listdir(temporaryFolderPath):
                os.remove(temporaryFolderPath + "/" + file)
                
    texts.append(text)
    for i in range(len(countryLanguages)-1):
        texts.append("" + pytesseract.image_to_string(image, lang = countryLanguages[i+1]))
    texts.append("" + pytesseract.image_to_string(image, lang="eng"))
    shutil.rmtree(temporaryFolderPath)
    for i in range(len(texts)):
        probs.append(float(re.sub(r"[^0-9.]+", "", str(langdetect.detect_langs(texts[i])[0]))))
        langs.append(re.sub(r"[^a-zA-Z]+", "", str(langdetect.detect_langs(texts[i])[0])))
    lang = langs[probs.index(np.max(probs))]
    return lang
#==============================================================================
def DetectLanguage(filePath, directoryPath, countryLanguages, nbCharacters): #countryLanguage: 3 letters, lang: 2 letters
    if (isScannedFile(filePath) == 1):
        lang = DetectLanguageForScan(filePath, directoryPath, countryLanguages, nbCharacters)
    else:
        lang = re.sub(r"[^a-zA-Z]+", "", str(langdetect.detect_langs(ConvertPdftoText(filePath))[0]))
    return lang
#==============================================================================
#==============================================================================
def KeepOnlyPdfsNeeded(directoryPath, pdfTypes):
    pdfsToDelete = []
    total_nb = 0
    ithFile = 0
    for country in os.listdir(directoryPath):
        if ((country[-4] != '.') & (country not in ["Duplicates", "Temporary_Folder", "Errors"])):
            for year in os.listdir(directoryPath + "/" + country):
                total_nb += len(os.listdir(directoryPath + "/" + country + "/" + year))
    sys.stdout.write("\r")
    sys.stdout.write("KeepOnlyPdfsNeeded -- [" + total_nb * " " + "] 0%")           
    
    for country in os.listdir(directoryPath):
        if ((country[-4] != '.') & (country not in ["Duplicates", "Temporary_Folder", "Errors"])):
            for year in os.listdir(directoryPath + "/" + country):
                for pdf in os.listdir(directoryPath + "/" + country + "/" + year):
                    ithFile += 1
                    sys.stdout.write("\r")
                    sys.stdout.write("KeepOnlyPdfsNeeded -- [" + ithFile * "#" + (total_nb-ithFile) * " " + "] " + str(int((float(ithFile)/total_nb)*100)) + "%")
                    toDelete = True
                    for pdfType in pdfTypes:
                        if ((pdf.find(" " + pdfType + " ") != -1) | (pdf.find(" " + pdfType + ".") != -1)):
                            toDelete = False
                    if (toDelete == True):
                       pdfsToDelete.append(directoryPath + "/" + country + "/" + year + "/" + pdf)        
    for p in pdfsToDelete:
        os.remove(p)
    print("\nKeep only PDFs needed: OK\n")
#==============================================================================  
def RenamePdfs(directoryPath, pdfTypes, nbCharacters):
    duplicates_main_folder = []
    total_nb = 0
    ithFile = 0
    #Create Duplicates and Errors folders
    if not (os.path.exists(directoryPath + "/Duplicates")):
        os.makedirs(directoryPath +  "/Duplicates")
    if not (os.path.exists(directoryPath + "/Errors")):
        os.makedirs(directoryPath +  "/Errors")
        
    #Count the total number of pdfs to rename
    for country in os.listdir(directoryPath):
        if ((country[-4] != '.') & (country not in ["Duplicates", "Temporary_Folder", "Errors"])):
            for year in os.listdir(directoryPath + "/" + country):
                total_nb += len(os.listdir(directoryPath + "/" + country + "/" + year))
    sys.stdout.write("\r")
    sys.stdout.write("RenamePdfs -- [" + total_nb * " " + "] 0%")
    
    # For each pdf of each year of each country
    for country in os.listdir(directoryPath):
        if ((country[-4] != '.') & (country not in ["Duplicates", "Temporary_Folder", "Errors"])):
            countryLanguages = [LanguageName(lang) for lang in FindCountryLanguage(country)]
            for year in os.listdir(directoryPath + "/" + country):
                for pdf in os.listdir(directoryPath + "/" + country + "/" + year):
                    ithFile += 1
                    sys.stdout.write("\r")
                    sys.stdout.write("RenamePdfs -- [" + ithFile * "#" + (total_nb - ithFile) * " " + "] " + str(int((float(ithFile)/total_nb)*100)) + "%")
                    filePath = directoryPath + "/" + country + "/" + str(year) + "/" + pdf
                    isDuplicate = 0
                    possibleOtherPdfNames = [] 
                    pdfName = ""
                    pdfLanguage = ""
                    
                    # Detect the language
                    try:
                        pdfLanguage = LanguageName(DetectLanguage(filePath, directoryPath, countryLanguages, nbCharacters))
                    except Exception as e:
                        date_start_pos = [m.start() for m in re.compile("[0-9]{2}[-][A-Z]{3}[-][0-9]{4}").finditer(pdf)][0]
                        pdfName = country + " " + pdf[:date_start_pos-1] + " " + str(year) + " " + pdf[date_start_pos+12:pdf.rfind(pdf.split()[-1])-1] + " " + pdf.split()[-1]
                        os.rename(filePath, directoryPath + "/Errors/" + pdfName)
                        print("\nDetectLanguage failed: " + country + "/" + str(year) + "/" + pdf)
                        print(str(e))
                        pass
                    
                    # If the pdf language is one of the languages ​​spoken in the country or English, put it in the good folder
                    if (pdfLanguage in countryLanguages + ["eng"]):
                        date_start_pos = [m.start() for m in re.compile("[0-9]{2}[-][A-Z]{3}[-][0-9]{4}").finditer(pdf)][0]
                        pdfName = country + " " + pdf[:date_start_pos-1] + " " + str(year) + " " + pdfLanguage + " " + pdf[date_start_pos+12:pdf.rfind(pdf.split()[-1])-1] + ".pdf"

                        # Find all possible other pdf names
                        for pdfType in pdfTypes:
                            for lang in countryLanguages+["eng"]:
                                if ((pdfType != pdf[date_start_pos+12:pdf.rfind(pdf.split()[-1])-1]) | (lang != pdfLanguage)):
                                    possibleOtherPdfNames.append(country + " " + pdf[:date_start_pos-1] + " " + str(year) + " " + lang + ' ' + pdfType + ".pdf")
                    
                        # If the pdf has already duplicates, remane it
                        if os.path.exists(directoryPath + "/" + pdfName):
                            isDuplicate = 1
                            duplicates_main_folder.append(pdfName) # For now, keep the first duplicate in the main folder to be able to compare it with others and stock it in the duplicates_main_folder list to put it later in the Duplicates folder 
                            pdfName = pdfName[:-4] + " 2.pdf"
                            number = 3
                            while (os.path.exists(directoryPath + "/Duplicates/" + pdfName)):
                                pdfName = pdfName[:-6] + " " + str(number) + ".pdf"
                                number += 1       
                        else:
                            for possiblePdfName in possibleOtherPdfNames:
                                if (os.path.exists(directoryPath + "/" + possiblePdfName)):
                                    isDuplicate = 1
                                    duplicates_main_folder.append(possiblePdfName) # Same reason than above
                                    if (os.path.exists(directoryPath + "/Duplicates/" + pdfName)):
                                        pdfName = pdfName[:-4] + " 2.pdf"
                                        number = 3
                                        while (os.path.exists(directoryPath + "/Duplicates/" + pdfName)):
                                            pdfName = pdfName[:-6] + ' ' + str(number) + ".pdf"
                                            number += 1 
                        
                        # If the pdf has no duplicate yet, put it in the main folder
                        if (isDuplicate == 0):
                            os.rename(filePath, directoryPath + "/" + pdfName)
                        # If the pdf has already duplicates, put it in the Duplicates folder
                        else:
                            os.rename(filePath, directoryPath + "/Duplicates/" + pdfName)
                    # If the pdf language is neither one of the languages ​​spoken in the country nor English, delete it
                    else:
                        if (len(pdfLanguage) > 0):
                            os.remove(filePath)             
     
    # Put all first duplicates from the main folder to the Duplicates folder              
    for pdfName in set(duplicates_main_folder):
        if (os.path.exists(directoryPath + "/" + pdfName)):
            os.rename(directoryPath + "/" + pdfName, directoryPath + "/Duplicates/" + pdfName)
    
    # Delete all useless folders
    for country in os.listdir(directoryPath):
        if ((country[-4] != ".") & (country not in ["Duplicates", "Temporary_Folder", "Errors"])):
            for year in os.listdir(directoryPath + "/" + country):
                if (len(os.listdir(directoryPath + "/" + country + "/" + year)) == 0):
                    shutil.rmtree(directoryPath + "/" + country + "/" + year)
            if (len(os.listdir(directoryPath + "/" + country)) == 0):
                shutil.rmtree(directoryPath + "/" + country)
                         
    sys.stdout.write("\r")
    sys.stdout.write("RenamePdfs -- [" + total_nb * "#" + "] " + "100%")
    print("\nRename all PDFs: OK\n")
#==============================================================================   
def FindDuplicatesToDelete(duplicatesDirectoryPath, pdfTypes):
    pdfsToDelete = []
    uniquePdfs = []
    shortNames = []
    # Retrieve all shortNames (names that should be unique)
    for duplicate in os.listdir(duplicatesDirectoryPath):
        date_start_pos = [m.start() for m in re.compile("[ ][0-9]{4}[ ]").finditer(duplicate)][0]
        shortNames.append(duplicate[:date_start_pos+5])
    shortNames = list(set(shortNames))
    
    # Stock all duplicates together
    pdfs = []
    for shortName in shortNames:
        duplicates = []
        for duplicate in os.listdir(duplicatesDirectoryPath):
             if (duplicate.rfind(shortName) != -1):
                 duplicates.append(duplicate)
        pdfs.append(duplicates) 
    sys.stdout.write("\r")
    sys.stdout.write("FindDuplicatesToDelete -- [" + len(pdfs) * " " + "] 0%")
    
    # Keep in priority PDFs in English then in the language of the country
    for i in range(len(pdfs)):
        d = pdfs[i]
        toRemove = []
        langs = []
        countryLanguages = FindCountryLanguage(d[0].split()[0])
        for duplicate in d:
            date_start_pos = [m.start() for m in re.compile("[ ][0-9]{4}[ ]").finditer(duplicate)][0]
            langs.append(LanguageName(duplicate[date_start_pos+6:date_start_pos+9]))
        langs = list(set(langs))
        if (len(langs) > 1):
            if ("en" in langs):
                for duplicate in d:
                    if (duplicate.find(" eng ") == -1):
                        pdfsToDelete.append(duplicate)
                        toRemove.append(duplicate)
            elif (len(list(set(countryLanguages) & set(langs))) > 0):
                for duplicate in d:
                    for lang in list(set(countryLanguages) & set(langs)):
                        if (duplicate.find(lang) == -1):
                            pdfsToDelete.append(duplicate)
                            toRemove.append(duplicate)            
            else:
                for duplicate in d:
                    pdfsToDelete.append(duplicate)
                    toRemove.append(duplicate)
        for j in range(len(set(toRemove))):
            pdfs[i].remove(toRemove[j])                    

    # Keep in priority pdf types: ARS-None then FullYear then ARS then AR S
    for i in range(len(pdfs)):
        d = pdfs[i]
        if (len(d) > 0):
            toRemove = []
            types = []
            for duplicate in d:
                if ((duplicate.rfind(" AR S ") != -1) | (duplicate.rfind(" AR S.") != -1)):
                    types.append("AR S")
                else:
                    if (duplicate.split()[-1][:-4].isdigit()):
                        types.append(duplicate.split()[-2])
                    else:
                        types.append(duplicate.split()[-1][:-4])
            types = list(set(types))
            if (len(types) > 1):
                ind = []
                for t in types:
                    ind.append(pdfTypes.index(t))
                typeToKeep = types[ind.index(np.min(ind))]
                for duplicate in d:
                    if (duplicate.find(typeToKeep) == -1):
                        pdfsToDelete.append(duplicate)
                        toRemove.append(duplicate) 
            for j in range(len(set(toRemove))):
                pdfs[i].remove(toRemove[j])
                
    # For the remaining duplicates: 
    # 1) if their text is identical then keep only one
    # 2) if their text is different:
    #      a - if the entire text of one PDF is included in the text of the other then keep the biggest one
    #      b - else keep both PDFs to merge them later
    for i in range(len(pdfs)):
        d = pdfs[i]
        sys.stdout.write("\r")
        sys.stdout.write("FindDuplicatesToDelete -- [" + (i+1) * "#" + (len(pdfs)-(i+1)) * " " + "] " + str(int((float(i+1)/len(pdfs))*100)) + "%")
        if(len(d) > 1):
            toRemove = []
            text1 = ""
            text2 = ""
            for j in range(len(d)):
                for k in range(len(d)):
                    if ((k > j) & (d[j] not in toRemove) & (d[k] not in toRemove)):  
                        j=0
                        k=1
                        lang1 = [l for l in [LanguageName(x) for x in FindCountryLanguage(d[j].split()[0])] + ['eng'] if (d[j].find(l) != -1)][0]
                        lang2 = [l for l in [LanguageName(x) for x in FindCountryLanguage(d[j].split()[0])] + ['eng'] if (d[k].find(l) != -1)][0]
                        text1 = re.sub(r" {2,}", " ", re.sub(r"[^A-Za-z ]+", " ", ConvertFileToText(duplicatesDirectoryPath + '/' + d[j], lang1)))
                        text2 = re.sub(r" {2,}", " ", re.sub(r"[^A-Za-z ]+", " ", ConvertFileToText(duplicatesDirectoryPath + '/' + d[k], lang2)))
                        if (text1 == text2):
                            pdfsToDelete.append(d[j])
                            toRemove.append(d[j])
                        else:
                            if (len(text1) > len(text2)):
                                if (text1.find(text2) != -1):
                                    pdfsToDelete.append(d[k])
                                    toRemove.append(d[k])
                            else:
                                if (text2.find(text1) != -1):
                                    pdfsToDelete.append(d[j])
                                    toRemove.append(d[j])                 
            for r in range(len(set(toRemove))):
                pdfs[i].remove(toRemove[r])
                
    # Stock all unique PDFs        
    for i in range(len(pdfs)):
        if (len(pdfs[i]) == 1):
            uniquePdfs.append(pdfs[i][0])
            
    return pdfsToDelete, uniquePdfs
#==============================================================================  
def DeleteDuplicates(directoryPath, pdfTypes):
    duplicatesDirectoryPath = directoryPath + "/Duplicates"
    total_nb = len(os.listdir(duplicatesDirectoryPath))
    pdfsWhichNeedAction = FindDuplicatesToDelete(duplicatesDirectoryPath, pdfTypes)
    pdfsToDelete = pdfsWhichNeedAction[0]
    uniquePdfs = pdfsWhichNeedAction[1]
    
    # Put unique PDFs in the main folder
    for pdf in uniquePdfs:
        os.rename(duplicatesDirectoryPath + "/" + pdf, directoryPath + "/" + pdf)
        
    # Delete useless duplicates
    for pdf in pdfsToDelete:
        os.remove(duplicatesDirectoryPath + "/" + pdf)
    
    sys.stdout.write("\r")
    sys.stdout.write("DeleteDuplicates -- [" + total_nb * "#" + "] 100%")
    print("\n Delete duplicates: OK")
#==============================================================================
def VerifyUniquePdfs(directoryPath):
    shortNames = []
    # Rename all unique PDFs such as: country company year language type (whitout digit at the end)
    for pdf in os.listdir(directoryPath):
        if ((pdf[-4:] == ".pdf") & (pdf.split()[-1][:-4].isdigit())):
            os.rename(directoryPath + "/" + pdf, directoryPath + "/" + pdf[:pdf.find(pdf.split()[-1])-1] + ".pdf")
    
    # Find remaining duplicates in the main folder
    for pdf in os.listdir(directoryPath):
        if (pdf[-4:] == ".pdf"):
            date_start_pos = [m.start() for m in re.compile("[ ][0-9]{4}[ ]").finditer(pdf)][0]
            shortNames.append(pdf[:date_start_pos+5])
    duplicates = [item for item, count in collections.Counter(shortNames).items() if count > 1]

    if (len(duplicates) > 0):
        print("\n Verify Unique Pdfs: STILL DUPLICATES!!")
        print(duplicates)
    else:
        print("\n Verify Unique Pdfs: OK")
#==============================================================================
def SplitPdfsIntoFolders(directoryPath, foldersNumber):
    years = []
    ithFolder = 1
    
    for folder in range(foldersNumber):
        if not (os.path.exists(directoryPath + "/Folder" + str(folder+1))):
            os.makedirs(directoryPath + "/Folder" + str(folder+1))
    
    for pdf in os.listdir(directoryPath):
        if (pdf[-4:] == ".pdf"):
            date_start_pos = [m.start() for m in re.compile("[ ][0-9]{4}[ ]").finditer(pdf)][0]
            years.append(pdf[date_start_pos+1:date_start_pos+5])
            
    for year in set(years):
        for pdf in os.listdir(directoryPath):
            if (pdf[-4:] == ".pdf"):
                date_start_pos = [m.start() for m in re.compile("[ ][0-9]{4}[ ]").finditer(pdf)][0]
                if (pdf[date_start_pos+1:date_start_pos+5] == year):
                    if (ithFolder > foldersNumber):
                        ithFolder = 1
                    os.rename(directoryPath + "/" + pdf, directoryPath + "/Folder" + str(ithFolder) + "/" + pdf)
                    ithFolder += 1
                    
    print("\n Split PDFs into folders: OK")
#==============================================================================    
def CleanPdfs(directoryPath, pdfTypes, nbCharacters, foldersNumber):
    KeepOnlyPdfsNeeded(directoryPath, pdfTypes)
    RenamePdfs(directoryPath, pdfTypes, nbCharacters)
    DeleteDuplicates(directoryPath, pdfTypes)
#==============================================================================    
   
# MAIN
directoryPath = "C:/temp/ANNUAL REPORTS - NEW"
pdfTypes = ["ARS-None", "FullYear", "ARS", "AR S"]
nbCharacters = 300
foldersNumber = 5
CleanPdfs(directoryPath, pdfTypes, nbCharacters, foldersNumber)
VerifyUniquePdfs(directoryPath)
SplitPdfsIntoFolders(directoryPath, foldersNumber)
