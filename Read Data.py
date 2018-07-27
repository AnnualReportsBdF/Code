import os
import pandas as pd
import re
import unidecode
import unicodedata
from googletrans import Translator
#https://pypi.org/project/googletrans/#description


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


df = CreateDataframe('C:/temp/Annual reports')
print(len(df))
df.head()



"""
print(df[df.LanguageEstimated == 'fra'].index.values)
df[df.LanguageEstimated == 'fra']['Text'][12]


translator = Translator()
translation = translator.translate("un deux trois", src = "fr", dest = "en").text
print(translation)


from requests.auth import HTTPProxyAuth
proxyDict = { 'http' : '127.0.0.1:3128', 'https' : '127.0.0.1:3128' } 
auth = HTTPProxyAuth('R550427', 'QwV33993#')
r = requests.get("http://www.google.com", proxies=proxyDict, auth=auth)
print r
"""