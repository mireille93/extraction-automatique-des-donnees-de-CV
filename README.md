# extraction-automatique-des-donnees-de-CV
Extraction automatique des données de Curriculum Vitae avec Python
Apres avoir telecharger et dezipper ce projet dans un repertoire, acceedez au dossier de ce projet sur votre pc

cd extraction-automatique-des-donnees-de-CV

Remplacer vos donnees cvs texte dans le dossier resumes

Dans le code liredossier.py; changer le lien de creation du fichier excel a la ligne 15.
pathexel = r"E:\cours\reconnaissancesdeformes\extractionCV\Infoextrait.xls"

Remplacer l'adresse dans le code.

installer ensuite pdfminer et nltk

python3 -m venv .venv

pip install -r pdfminer3k==1.3.1

pip install -r splitty==0.0.7

pip install -r nltk

et compiler avec la commande:

python liredossier.py



le fichier Infoextrait.xls sera cree avec es donnees extraites.
dans le dossier resumtext, les fichiers cv pdf convertis apparaitront.
