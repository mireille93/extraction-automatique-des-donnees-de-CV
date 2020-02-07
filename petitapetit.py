from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox, LTTextLine
import nltk, os, subprocess, code, glob, re, traceback, sys, inspect
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pprint import pprint

#ouvrir plusieurs fichiers

from xlwt import Workbook
 
# création 
path = r"E:\cours\reconnaissancesdeformes\extractionCV\Infoextrait.xls"
book = Workbook()
 
# création de la feuille 1
feuil1 = book.add_sheet('feuil1')
 
# ajout des en-têtes
feuil1.write(0,0,'NOM')
feuil1.write(0,1,'EMAILS')
feuil1.write(0,2,'CONTACTS')
feuil1.write(0,3,'COMPETENCES')
feuil1.write(0,4,'LANGUES')
feuil1.write(0,5,'DILOMES')

# ajout des valeurs dans la ligne suivante
ligne1 = feuil1.row(1)
ligne2 = feuil1.row(2)
ligne3 = feuil1.row(3)
ligne4 = feuil1.row(4)
ligne5 = feuil1.row(5)
ligne6 = feuil1.row(6)
 
# ajustement éventuel de la largeur d'une colonne
feuil1.col(0).width = 10000
feuil1.col(1).width = 10000
feuil1.col(2).width = 10000
feuil1.col(3).width = 10000
feuil1.col(4).width = 10000
feuil1.col(5).width = 10000
feuil1.col(6).width = 10000
# création matérielle du fichier résultant
 
#print u"Fichier créé: {}".format(path)





def pdf_to_string(pdf_file):
    fp = open(pdf_file, 'rb')

    parser = PDFParser(fp)
    doc = PDFDocument(parser)
    parser.set_document(doc)
    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    extracted_text = ''

    for page in PDFPage.create_pages(doc):
        interpreter.process_page(page)
        layout = device.get_result()
        for lt_obj in layout:
            if isinstance(lt_obj, (LTTextBox, LTTextLine)):
                extracted_text += ' '+lt_obj.get_text()

    return extracted_text


with open('cv.txt', 'w', encoding='utf-8') as f:
    f.write(pdf_to_string('cv.pdf'))
    

#extraire les mails
pattern_emails= re.compile(r'\S*@\S*')
emails=pattern_emails.finditer(pdf_to_string('cv.pdf'))

emmail=''
for email in emails:
    print(email.group(0))
    emmail+='  '+str(email.group(0))

ligne1.write(1,emmail)  


#extraire les contacts

phone = ''
inputString=str(pdf_to_string('cv.pdf'))
pattern = re.compile(r'([+84]+[+(]?\d+[)\-]?[ \t\r\f\v]*[(]?\d{2,}[()\-]?[ \t\r\f\v]*\d{2,}[()\-]?[ \t\r\f\v]*\d*[ \t\r\f\v]*\d*[ \t\r\f\v]*)')
match = pattern.finditer(inputString)

for tel in match:
    print(tel.group(0))
    phone+='  '+str(tel.group(0))

ligne1.write(2,phone)

#extraire les noms
document=pdf_to_string('cv.pdf')
lines = [el.strip() for el in document.split("\n") if len(el) > 0]   
lines = [nltk.word_tokenize(el) for el in lines]    
lines = [nltk.pos_tag(el) for el in lines]  
sentences = nltk.sent_tokenize(document)    
sentences = [nltk.word_tokenize(sent) for sent in sentences]    
tokens = sentences
sentences = [nltk.pos_tag(sent) for sent in sentences]    
dummy = []
for el in tokens:
    dummy += el
    tokens = dummy
    
def getName(document):
        names = open("allNames.txt", "r").read().lower()
        names = set(names.split())
        otherNameHits = []
        nameHits = []
        name = None
        try:
            grammar = r'NAME: {<NN.*><NN.*><NN.*>*}'
            chunkParser = nltk.RegexpParser(grammar)
            all_chunked_tokens = []
            for tagged_tokens in lines:
                if len(tagged_tokens) == 0: continue 
                chunked_tokens = chunkParser.parse(tagged_tokens)
                all_chunked_tokens.append(chunked_tokens)
                for subtree in chunked_tokens.subtrees():
                    if subtree.label() == 'NAME':
                        for ind, leaf in enumerate(subtree.leaves()):
                            if leaf[0].lower() in names and 'NN' in leaf[1]:
                                hit = " ".join([el[0] for el in subtree.leaves()[ind:ind+3]])
                                if re.compile(r'[\d,:]').search(hit): continue
                                nameHits.append(hit)
            if len(nameHits) > 0:
                nameHits = [re.sub(r'[^a-zA-Z \-]', '', el).strip() for el in nameHits] 
                name = " ".join([el[0].upper()+el[1:].lower() for el in nameHits[0].split() if len(el)>0])
                otherNameHits = nameHits[1:]

        except Exception as e:
            print (traceback.format_exc())
            print (e)  

        print(name)
        return name
    
nom=''
nom=str(getName(document))
ligne1.write(0,nom)

def getCompetences(document):
        comps = open("competences.txt", "r").read().lower()
        comps = set(comps.split())
        otherCompHits = []
        compHits = []
        comp = None

        try: 
            grammar = r'NAME: {<NN.*><NN.*>*}'
            chunkParser = nltk.RegexpParser(grammar)
            all_chunked_tokens = []
            for tagged_tokens in lines:
                if len(tagged_tokens) == 0: continue 
                chunked_tokens = chunkParser.parse(tagged_tokens)
                all_chunked_tokens.append(chunked_tokens)
                for subtree in chunked_tokens.subtrees():
                    if subtree.label() == 'NAME':
                        for ind, leaf in enumerate(subtree.leaves()):
                            if leaf[0].lower() in comps and 'NN' in leaf[1]:
                                hit = " ".join([el[0] for el in subtree.leaves()[ind:ind+3]])
                                if re.compile(r'[\d,:]').search(hit): continue
                                compHits.append(hit)
            if len(compHits) > 0:
                compHits = [re.sub(r'[^a-zA-Z \-]', '', el).strip() for el in compHits] 
                comp = " ".join([el[0].upper()+el[1:].lower() for el in compHits[0].split() if len(el)>0])
                otherCompHits = compHits[1:]

        except Exception as e:
            print (traceback.format_exc())
            print (e)  

        print(comp, otherCompHits)
        return comp,otherCompHits

comp=''
comp=str(getCompetences(document))
ligne1.write(3,comp)

def getLangues(document):
        langues = open("langues.txt", "r").read().lower()
        langues = set(langues.split())
        otherLgsHits = []
        compHits = []
        langue = None

        try: 
            grammar = r'NAME: {<NN.*>*}'
            chunkParser = nltk.RegexpParser(grammar)
            all_chunked_tokens = []
            print('coucou')
            for tagged_tokens in lines:
                if len(tagged_tokens) == 0: continue 
                chunked_tokens = chunkParser.parse(tagged_tokens)
                all_chunked_tokens.append(chunked_tokens)
                for subtree in chunked_tokens.subtrees():
                    if subtree.label() == 'NAME':
                        for ind, leaf in enumerate(subtree.leaves()):
                            if leaf[0].lower() in langues and 'NN' in leaf[1]:
                                hit = " ".join([el[0] for el in subtree.leaves()[ind:ind+3]])
                                if re.compile(r'[\d,:]').search(hit): continue
                                compHits.append(hit)
            if len(compHits) > 0:
                compHits = [re.sub(r'[^a-zA-Z \-]', '', el).strip() for el in compHits] 
                langue = " ".join([el[0].upper()+el[1:].lower() for el in compHits[0].split() if len(el)>0])
                otherLgsHits = compHits[1:]
                
        except Exception as e:
            print (traceback.format_exc())
            print (e)  

        print(langue, otherLgsHits)
        return langue,otherLgsHits

langs=''
langs=str(getLangues(document))
ligne1.write(4,langs)



def getDiplome(document):
        langues = open("diplome.txt", "r").read().lower()
        langues = set(langues.split())
        otherLgsHits = []
        compHits = []
        langue = None

        try: 
            grammar = r'NAME: {<NN.*>*}'
            chunkParser = nltk.RegexpParser(grammar)
            all_chunked_tokens = []
            print('coucou')
            for tagged_tokens in lines:
                if len(tagged_tokens) == 0: continue 
                chunked_tokens = chunkParser.parse(tagged_tokens)
                all_chunked_tokens.append(chunked_tokens)
                for subtree in chunked_tokens.subtrees():
                    if subtree.label() == 'NAME':
                        for ind, leaf in enumerate(subtree.leaves()):
                            if leaf[0].lower() in langues and 'NN' in leaf[1]:
                                hit = " ".join([el[0] for el in subtree.leaves()[ind:ind+3]])
                                if re.compile(r'[\d,:]').search(hit): continue
                                compHits.append(hit)
            if len(compHits) > 0:
                compHits = [re.sub(r'[^a-zA-Z \-]', '', el).strip() for el in compHits] 
                langue = " ".join([el[0].upper()+el[1:].lower() for el in compHits[0].split() if len(el)>0])
                otherLgsHits = compHits[1:]
                
        except Exception as e:
            print (traceback.format_exc())
            print (e)  

        print(langue, otherLgsHits)
        return langue,otherLgsHits

diplom=''
diplom=str(getDiplome(document))
ligne1.write(5,diplom)

book.save(path)