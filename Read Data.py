import os
import pandas as pd
import re
import unidecode
import unicodedata


def CreateDataframe(directoryPath):
    csvsCombined = pd.DataFrame([])
    for csv in os.listdir(directoryPath):
        if (csv[-4:] == '.csv'):
            df = pd.read_csv(directoryPath + '/' + csv)
            csvsCombined = csvsCombined.append(df)
    csvsCombined = csvsCombined.drop('Unnamed: 0', axis=1)
    csvsCombined.index = range(len(csvsCombined))    
    del df
    return csvsCombined


def EncodeText(text):
    try:
        text = unicode(text, 'utf-8')
    except TypeError:
        pass
    text = unicodedata.normalize('NFD', text)
    text = text.encode('ascii', 'ignore')
    text = text.decode("utf-8")
    return str(text)


def CleanText(text):
    #text = unicode(text, 'utf-8')
    #text = unidecode.unidecode(text)
    text = EncodeText(text)
    #text = re.sub(r"[^A-Za-z ]+", " ", text)
    text = re.sub(r"[\n\r\x0c]", " ", text)
    text = re.sub(r" {2,}", " ", text)
    return str(text)


df = CreateDataframe('Annual reports')
df.head()
df['Text'][0]
CleanText(df['Text'][0])