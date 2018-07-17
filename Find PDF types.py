import os

def FindPdfTypes(directoryPath):
    pdfTypes = []
    for year in os.listdir(directoryPath):
        for file in os.listdir(directoryPath + '/' + year):
            pdfType = ""
            if file.find(' SE ') != -1:
                name = file[file.find(' SE ')+4:]
            elif file.find(' SA ') != -1:
                name = file[file.find(' SA ')+4:]
            elif file.find(' SCA ') != -1:
                name = file[file.find(' SCA ')+5:]
            elif file.find(' SAS ') != -1:
                name = file[file.find(' SAS ')+5:]
            else:
                print(file)
            words = name.split(' ')
            for word in words[1:-1]:
                pdfType += word + " "
            
            pdfTypes.append(pdfType[:-1])
    
    return set(pdfTypes)

FindPdfTypes('France base/en')