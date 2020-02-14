

from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox, LTTextLine
import nltk, os, subprocess, code, glob, re, traceback, sys, inspect
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pprint import pprint


from xlwt import Workbook
 
# création  du fichier excel
pathexel = r"E:\cours\reconnaissancesdeformes\extractionCV\Infoextrait.xls"
book = Workbook()
 
# création de la feuille 1
feuil1 = book.add_sheet('feuil1', cell_overwrite_ok=True)
 
# ajout des en-têtes
feuil1.write(0,0,'NOM')
feuil1.write(0,1,'EMAILS')
feuil1.write(0,2,'CONTACTS')
feuil1.write(0,3,'COMPETENCES')
feuil1.write(0,4,'LANGUES')
feuil1.write(0,5,'DIPLOMES')


# ajustement éventuel de la largeur d'une colonne
feuil1.col(0).width = 10000
feuil1.col(1).width = 10000
feuil1.col(2).width = 10000
feuil1.col(3).width = 10000
feuil1.col(4).width = 10000
feuil1.col(5).width = 10000
feuil1.col(6).width = 10000
# création matérielle du fichier résultant


#function pour covertir un pdf en texte
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


path='./resumes'
pathtexte='./resumtext'

def listdirectory(path): 
    fichier=[] 
    l = glob.glob(path+'/*') 
    print(l)
    for i in l: 
        if os.path.isdir(i): fichier.extend(listdirectory(i)) 
        else: fichier.append(i) 
    return fichier

resume=listdirectory(path)

def listdirectory(pathtexte): 
    fichiertexte=[] 
    ll = glob.glob(pathtexte+'/*') 
    print(ll)
    for j in ll: 
        if os.path.isdir(i): fichiertexte.extend(listdirectory(j)) 
        else: fichiertexte.append(j) 
    return fichiertexte
resumetexte=listdirectory(pathtexte)
print(resumetexte)


#fonctin pour extraire les contacts
def getTelephone(document):
    phone = ''
    inputString=str(document)
    pattern = re.compile(r'([+84]+[+(]?\d+[)\-]?[ \t\r\f\v]*[(]?\d{2,}[()\-]?[ \t\r\f\v]*\d{2,}[()\-]?[ \t\r\f\v]*\d*[ \t\r\f\v]*\d*[ \t\r\f\v]*)')
    match = pattern.finditer(inputString)
                    
    for tel in match:
        print(tel.group(0))
        phone+='  '+str(tel.group(0))
    return phone



#fonctin pour extraire les emails
def getEmail(document):
    emmail=''
    for email in emails:
        print(email.group(0))
        emmail+='  '+str(email.group(0))
    return emmail


#fontion qui extrait les noms
def getName(document):
        names = open("allNames.txt", "r").read().lower()
        names = set(names.split())
        otherNameHits = []
        nameHits = []
        name = None
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

#fonction pour extraire les competences
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

        print(otherCompHits)
        return comp,otherCompHits




#function pour extraire le diplome
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
                
    print(otherLgsHits)
    return otherLgsHits                 


#fonction pour extraire les langues
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


ii=1

for i in resume:
    iii=str(ii)
    ligne = feuil1.row(ii)
    with open('./resumtext/'+iii+'.txt', 'w', encoding='utf-8') as f:
        f.write(pdf_to_string(i))
        document=pdf_to_string(i)
        
        for ji in resumetexte:
            print(ji)
            #extraire les mails
            pattern_emails= re.compile(r'\S*@\S*')
            emails=pattern_emails.finditer(document)
                
                #extraire les contacts

            phone=getTelephone(document)

            feuil1.row(ii).write(2,phone)

            feuil1.row(ii).write(1,getEmail(document))

                
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
                    
                    
            nom=''
            nom=str(getName(document))
            feuil1.row(ii).write(0,nom)
                
            comp=''
            comp=str(getCompetences(document))
            feuil1.row(ii).write(3,comp)

            langs=''
            langs=str(getLangues(document))
            feuil1.row(ii).write(4,langs)
                
            diplom=''
            diplom=str(getDiplome(document))
            feuil1.row(ii).write(5,diplom)
                
    ii=ii+1 
         
book.save(pathexel)
    
        

   
