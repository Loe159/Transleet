import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googletrans import Translator
from time import sleep



Translator = Translator()

def Config():
	print('Bienvenue sur ce programme, il va vous permettre de traduire vos mots dans n\'importe quelle langue sur votre feuille de calcul Google Sheet')
	print('Pour commencer, veuillez mettre les informations de connexion dans un fichier "creds.json" dans le même dossier que celui-ci, merci de contacter le créateur si vous avez besoin d\'aide.')
	#input('Appuyez sur entrer pour continuer...')
	FileName = input('Quel est le nom de votre fichier ?')
	return FileName
	

def Googlesheet(FileName):

	print('Connexion au Google sheet en cours...')

	Scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

	Creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", Scope) #Mettre les informations de connexion dans un fichier creds.json dans le même dossier que celui-ci

	Client = gspread.authorize(Creds)

	Sheet = Client.open(FileName).sheet1

	Data = Sheet.get_all_records()

	NumWords = len(Data)+1

	print('Connexion au Google sheet réussie')
	NbTrad = 0
	#Traduction des mots dans le Google sheet
	print('Traduction du tableau en cours...')
	for x in range(NumWords):
		enCell = Sheet.cell(x+2, 1).value
		frCell = Sheet.cell(x+2, 2).value
		if not frCell :
			Translation = Translator.translate(enCell, dest='fr').text
			Sheet.update_cell(x+2, 2, Translation)
			NbTrad += 1
		elif not enCell :
			Translation = Translator.translate(frCell, dest='en').text
			Sheet.update_cell(x+2, 1, Translation)
			NbTrad += 1

	
	NbTrad -= 1
	if NbTrad > 0 :
		print("Traduction de", NbTrad,"mot(s) terminée")
	else :
		print('Aucun nouveau mot a été ajouté depuis la dernière fois')




FileName = Config()
Googlesheet(FileName)
input('Appuyez sur entrer pour quitter...')
