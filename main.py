
# Version 1.0
# Release 2022-11-19
# Author: Micke Kring @mickekring
# E-mail: jag@mickekring.se

# Import modules

import xlrd
import xlsxwriter
import os
from tinydb import TinyDB, Query
from termcolor import colored
import tabula
import pandas as pd
from sys import platform

arskurs = "9"


# Kontrollerar vilket operativsystem

if platform == "linux" or platform == "linux2":
        operativsystem = "linux"
elif platform == "darwin":
        operativsystem = "macos"
elif platform == "win32":
        operativsystem = "windows"
        

# Global lista över de betyg som finns. Används för att iterera över

betygslista = ["BL", "EN", "HKK", "IDH", "MA", "ML SPR", "ML BET", "M1 SPR", "M1 BET", "M2 SPR", 
"M2 BET", "MU", "NO", "BI", "FY", "KE", "SO", "GE", "HI", "RE", "SH", "SL", "SV", "SVA", "TN", 
"TK", "DA", "JU"]


# Databasinställningar ##########################################################

db = TinyDB('db.json')

table_betyg = db.table('betyg')
table_omdommen = db.table('omdommen')
table_nationella_prov = db.table('nationellaprov')
table_franvaro = db.table('franvaro')
table_larare = db.table('larare')

user = Query()


# Rensat konsolen vid start av script ####################################

if operativsystem == "windows":
        clear = lambda: os.system('cls')
else:
        clear = lambda: os.system('clear')

clear()


# Startvariabel för bakgrundfärg #######################################

betyg_color = ""



def Konvertera_pdf_betygskatalog_till_xls():

	df = ("betygskatalog/betyg.pdf")

	output = "betygskatalog/betyg.csv"
	tabula.convert_into(df, output, output_format="csv", pages=1, stream=True)

	if operativsystem == "windows":
		read_file = pd.read_csv(r"betygskatalog/betyg.csv", encoding = "unicode_escape", engine ="python")
	else:
		read_file = pd.read_csv(r"betygskatalog/betyg.csv")

	read_file.to_excel(r"betygskatalog/betyg.xls", index = None, header=False)



# Läser betygsfilen excel och lagrar i databasen ############################

def Läs_från_betygsfil_till_databas():

	print("\nLäser in betyg från excelfilen till databasen.\nDet här kan ta ett litet tag (30 - 60 sekunder)...")

	# betyg.xls - Namn på betygsfilen i excelformat som du vill läsa från. Måste
	# finnas för att scriptet ska starta 

	workbook = xlrd.open_workbook('betygskatalog/betyg.xls')
	worksheet = workbook.sheet_by_index(0)


	# 1. Letar rätt på klassens namn, ex 9A samt termin och år, ex HT2018

	rad = 2
	klass = ""
	personnummer ="1234567890"
	student_rad = 2 # Vilken rad som första eleven börjar på i excelfilen
	student_namn = "Karl Anka"
	bryt_loop_text = "Betygsgivande" # Text som står på raden direkt efter sista eleven, som gör att loopen bryts


	while "Klass" not in klass:

		klass = worksheet.cell(rad, 0)
		klass = klass.value

		if "Klass" in klass:
			klassbeteckning = klass.replace("Klass ", "")
			arskurs = klassbeteckning
			arskurs = arskurs.replace("A", "")
			arskurs = arskurs.replace("B", "")
			arskurs = arskurs.replace("C", "")
			arskurs = arskurs.replace("D", "")
		else:
			pass

		rad += 1

	termin = worksheet.cell(rad - 1, 1)
	termin = termin.value
	termin = termin.replace("Termin ", "")



	# Loop elev för elev #########################

	while bryt_loop_text not in student_namn:

		kolumn = 2 # Startkolumn för loop att börja hämta betyg, i detta fall BL
		
		# 2. Hämtar elevens namn från excel 

		student_namn = worksheet.cell(student_rad, 0)
		student_namn = student_namn.value
		
		# Bryter loopen när sista namnet är läst

		if bryt_loop_text in student_namn:
			#worksheet_new.write(row ,column ,"", grader_no_border_left)
			break
		else:
			pass
		
		# 3. Hämtar personnummer från excel och räknar ut kön

		personnummer = worksheet.cell(student_rad, 1)
		
		# Tar bort .0 från excel-värdet i slutet på persnummret - löser även problem med TF
		
		personnummer = str(personnummer.value)
		personnummer = personnummer.replace(".0", "")
		
		kon = Personnummer_till_kon(personnummer)
		
		# 4. Kollar om elev finns i databas, om inte så skrivs grunduppgifter

		DB_Skriv_Grundelevdata(personnummer, termin, klassbeteckning, arskurs, student_namn, kon)

		#Uppdatera_nya_elevuppgifter(personnummer, termin, klassbeteckning, arskurs, student_namn, kon)

		# 5. Hämtar betyg och lagrar i databas

		for betyg in betygslista:

			amne = worksheet.cell(student_rad, kolumn)
			amne_varde = Formattera_betyg(str(amne.value))

			# Hoppar över omvandling till betygspoäng gällande språk, då ARA och liknande
			# ger 20 poäng eftersom A finns med

			if betyg == "ML SPR" or betyg == "M1 SPR" or betyg == "M2 SPR":
				amne_poang = ""

			else:
				amne_poang = Omvandla_betyg_till_tal(amne_varde)

			# print(betyg + ": " + amne_varde + " | Poäng: " + str(amne_poang))

			table_betyg.update({betyg: amne_varde}, (user.Personnummer == personnummer) & (user.Termin == termin))
			table_betyg.update({betyg + '-P': amne_poang}, (user.Personnummer == personnummer) & (user.Termin == termin))

			kolumn += 1

		# Slut på hämtning av denna elev
		
		student_rad += 1

	print("Klart. Betygen lästes in utan problem.\n")

	Skapa_excelfil(termin, klassbeteckning, arskurs, "Felsökning")



# Funktion som läser betyg ut databasen och printar ut dessa till konsolen

def Printa_inlasta_betyg_konsol_och_skapa_katalog_excel(termin, klassbeteckning, arskurs, syfte):

	row = 10

	students = table_betyg.search((user.Termin == termin) & (user.Klass == klassbeteckning))
	#students = table_betyg.search((user.Termin == termin) & (klassbeteckning in user.Klass))

	betyg_summa_tot = 0
	betyg_summa_F = 0
	betyg_summa_P = 0

	antal_elever = 0
	antal_flickor = 0
	antal_pojkar = 0

	for student in students:

		col = 0

		if student:
			print(colored("Namn: " + student['Namn'], 'green'))
			worksheet_new.write(row, col, student['Namn'], header_white_border)
			
			col = 1
			antal_elever += 1

			if student['Kön'] == "F":
				antal_flickor += 1
			else:
				antal_pojkar += 1

			worksheet_new.write(row, col, student['Kön'], cell_grey_bg_left_noborder)
			

			print(colored("Klass: " + student['Klass'], 'green'))
			print("Termin: " + student['Termin'] + "\n")
			
			col = 3
			
			if syfte == "Felsökning":

				for betyg in betygslista:

					if "ML SPR" in betyg and "2" in student['ML BET'] and " " not in student['ML SPR']:
						# Läser inte modersmål
						pass

					elif "ML BET" in betyg and "2" in student['ML BET'] and " " not in student['ML SPR']:
						# Läser inte modersmål
						pass

					elif "M1 SPR" in betyg and "2" in student['M1 BET'] and " " not in student['M1 SPR']:
						# Läser inte M1
						pass

					elif "M1 BET" in betyg and "2" in student['M1 BET'] and " " not in student['M1 SPR']:
						# Läser inte M1
						pass

					elif "M2 SPR" in betyg and "2" in student['M2 BET'] and " " not in student['M2 SPR']:
						# Läser inte moderna språk M2
						pass

					elif "M2 BET" in betyg and "2" in student['M2 BET'] and " " not in student['M2 SPR']:
						# Läser inte moderna språk
						pass

					elif "DA" in betyg or "JU" in betyg or "TN" in betyg:
						# Exkluderar DA, JU och TN
						pass

					elif "NO" in betyg and "9" in student['Klass'] or "SO" in betyg and "9" in student['Klass']:
						# Exkluderar NO och SO för åk 6-9
						pass

					elif "NO" in betyg and "8" in student['Klass'] or "SO" in betyg and "8" in student['Klass']:
						# Exkluderar NO och SO för åk 6-9
						pass

					elif "NO" in betyg and "7" in student['Klass'] or "SO" in betyg and "7" in student['Klass']:
						# Exkluderar NO och SO för åk 6-9
						pass

					elif "NO" in betyg and "6" in student['Klass'] or "SO" in betyg and "6" in student['Klass']:
						# Exkluderar NO och SO för åk 6-9
						pass

					elif "SVA" in betyg and "2" in student['SVA']:
						# Tar bort SVA där det inte läses
						pass

					elif betyg == "SV" and "2" in student['SV']:
						# Tar bort SV där det inte läses
						pass

					elif "SV" in betyg and "2" not in student['SV'] and "2" not in student['SVA']:
						# Kollar dubbelregistrering av både SV och SVA betyg
						print(colored(betyg + ": " + " " + student[betyg], 'yellow'))
						worksheet_new.write(row, col, student[betyg], grade_ordinary_center_red_bg)

					elif "2" in student[betyg]:
						# Letar efter alla 2:or
						print(colored(betyg + ": " + " " + student[betyg], 'yellow'))
						worksheet_new.write(row, col, student[betyg], grade_ordinary_center_yellow_bg)

					elif not student[betyg] and "SPR" not in betyg:
						# Letar efter alla tomma betygsfält exkluderat de med SPR, som är språk (ex ARA)
						print(colored(betyg + ": " + " " + student[betyg], 'yellow'))
						worksheet_new.write(row, col, student[betyg], grade_ordinary_center_yellow_bg)

					elif "*" not in student[betyg] and "SPR" not in betyg and "9" in student['Klass'] and "VT" in student['Termin']:
						# Letar efter saknad slutbetygsmarkering om det är åk 9 och VT
						print(colored(betyg + ": " + " " + student[betyg], 'red'))
						worksheet_new.write(row, col, student[betyg], grade_ordinary_center_red_bg)

					elif "*" in student[betyg] and "SPR" not in betyg and "9" in student['Klass'] and "HT" in student['Termin']:
						# Letar efter felaktig slutbetygsmarkering om det är åk 6-9HT
						print(colored(betyg + ": " + " " + student[betyg], 'red'))
						worksheet_new.write(row, col, student[betyg], grade_ordinary_center_red_bg)

					elif "*" in student[betyg] and "SPR" not in betyg and "9" not in student['Klass']:
						# Letar efter felaktig slutbetygsmarkering om det är åk 6-9HT
						print(colored(betyg + ": " + " " + student[betyg], 'red'))
						worksheet_new.write(row, col, student[betyg], grade_ordinary_center_red_bg)

					else:
						print(betyg + ": " + " " + student[betyg])
						worksheet_new.write(row, col, student[betyg], grade_ordinary)

					col += 1

			elif syfte == "Statistik":

				for betyg in betygslista:

					if "SPR" in betyg:
						worksheet_new.write(row, col, student[betyg], grade_ordinary)

					elif "2" in student[betyg]:
						worksheet_new.write(row, col, " ", grade_ordinary)

					elif "A" in student[betyg] or "B" in student[betyg]:

						worksheet_new.write(row, col, student[betyg], grade_ordinary_center_blue_bg)

					elif "C" in student[betyg]:
						worksheet_new.write(row, col, student[betyg], grade_ordinary_center_green_bg)

					elif "F" in student[betyg] or "-" in student[betyg] or "3" in student[betyg]:

						worksheet_new.write(row, col, student[betyg], grade_ordinary_center_red_bg)

					else:	

						worksheet_new.write(row, col, student[betyg], grade_ordinary)

					col += 1

			else:
				pass
						
			# Betygsstatistik för enskild elev
			
			antal_amnen = 0
			betyg_summa = 0

			for betyg in betygslista:

				if student[betyg + '-P'] or student[betyg + '-P'] == 0:
					antal_amnen += 1
					betyg_summa += student[betyg + '-P']
					betyg_summa_tot += student[betyg + '-P']

					if student['Kön'] == "F":
						betyg_summa_F += student[betyg + '-P']
					else:
						betyg_summa_P += student[betyg + '-P']
				else:
					pass

			print("\nAntal ämnen: " + str(antal_amnen))
			print("Summa betyg: " + str(betyg_summa))
			worksheet_new.write(row, col, betyg_summa, grade_ordinary)

			col += 1

			if antal_amnen == 0:
				betyg_medelvarde = 0
			else:
				betyg_medelvarde = round(betyg_summa / antal_amnen, 2)
			
			print("Medel betyg: " + str(betyg_medelvarde))
			worksheet_new.write(row, col, betyg_medelvarde, Color_Points(betyg_medelvarde))

			print("--- --- ---\n")


			row += 1

		else:
			pass

	worksheet_new.write(2, 31, round(betyg_summa_tot / antal_elever, 2), grade_ordinary)
	worksheet_new.write(3, 31, round(betyg_summa_F / antal_flickor, 2), grade_ordinary)
	worksheet_new.write(4, 31, round(betyg_summa_P / antal_pojkar, 2), grade_ordinary)

	genomsnittlig_betygspoang_text_1 = "Den genomsnittliga betygspoängen i respektive ämne visar elevernas genomsnittliga betyg omräknat till poäng. Den genomsnittliga betygspoängen beräknas för elever som fått betyg A-F."
	genomsnittlig_betygspoang_text_2 = "Andel (%) elever med A-E. Andel elever som fått godkänt betyg, A-E av de elever som har fått A-F eller streck (-), dvs underlag saknas."

	worksheet_new.write(row + 1, 0, genomsnittlig_betygspoang_text_1, cell_meta_left_noborder)
	worksheet_new.write(row + 2, 0, genomsnittlig_betygspoang_text_2, cell_meta_left_noborder)

	Skapa_amnesstatikstik_for_betygskatalog(termin, klassbeteckning, arskurs)



def Skapa_amnesstatikstik_for_betygskatalog(termin, klassbeteckning, arskurs):
	
	# Söker fram aktuell klass - kan skickas via funktion
	students = table_betyg.search((user.Termin == termin) & (user.Klass == klassbeteckning))

	col = 3

	for betyg in betygslista:

		row = 2

		# Exkluderar språkbeteckning
		if "SPR" in betyg:
			pass

		else:
			summa = 0
			summa_P = 0
			summa_F = 0
			antal_betyg = 0
			antal_betyg_P = 0
			antal_betyg_F = 0
			antal_godkanda_elever = 0
			antal_godkanda_elever_P = 0
			antal_godkanda_elever_F = 0
			antal_elever_streck_till_A = 0
			antal_elever_streck_till_A_P = 0
			antal_elever_streck_till_A_F = 0

			print("\n------------\n" + betyg + "\n")

			for student in students:

				if "2" in student[betyg]:
					pass

				elif "3" in student[betyg]:
					pass

				elif "-" in student[betyg]:
					antal_elever_streck_till_A += 1

					if student['Kön'] == "F":
						antal_elever_streck_till_A_F += 1
					else:
						antal_elever_streck_till_A_P += 1
					
					pass

				elif not student[betyg]:
					pass

				else:
					antal_elever_streck_till_A += 1
					antal_betyg += 1
					summa += (student[betyg + '-P'])

					if student['Kön'] == "F":
						summa_F += (student[betyg + '-P'])
						antal_elever_streck_till_A_F += 1
						antal_betyg_F += 1
					else:
						summa_P += (student[betyg + '-P'])
						antal_elever_streck_till_A_P += 1
						antal_betyg_P += 1

					if (student[betyg + '-P']) > 9:

						antal_godkanda_elever += 1
						
						if student['Kön'] == "F":
							antal_godkanda_elever_F += 1
						else:
							antal_godkanda_elever_P += 1

			if antal_elever_streck_till_A == 0:
				pass

			else:

				print("\nAntal elever med betyg A-F: " + str(antal_betyg))
				print("Summa betyg: " + str(summa))
				
				if antal_betyg > 0:
					print("(X) Genomsnittlig betygspoäng A-F: " + str(round(summa / antal_betyg, 2)))
					worksheet_new.write(row, col, round(summa / antal_betyg, 2), Color_Points(summa / antal_betyg))
				else:
					worksheet_new.write(row, col, "-", grade_ordinary)

				if antal_betyg_F > 0:
					print("(X) Genomsnittlig betygspoäng Flickor A-F: " + str(round(summa_F / antal_betyg_F, 2)))
					worksheet_new.write(row + 1, col, round(summa_F / antal_betyg_F, 2), Color_Points(summa_F / antal_betyg_F))
				else:
					worksheet_new.write(row + 1, col, "-", grade_ordinary)

				if antal_betyg_P > 0:
					print("(X) Genomsnittlig betygspoäng Pojkar A-F: " + str(round(summa_P / antal_betyg_P, 2)))
					worksheet_new.write(row + 2, col, round(summa_P / antal_betyg_P, 2), Color_Points(summa_P / antal_betyg_P))
				else:
					worksheet_new.write(row + 2, col, "-", grade_ordinary)

				print("Antal elever med A-F samt ---: " + str(antal_elever_streck_till_A))
				print("Antal elever med godkänt betyg (A-E): " + str(antal_godkanda_elever))

				print("(X) Antal elever (%) med godkänt betyg (A-E): " + str(round((antal_godkanda_elever / antal_elever_streck_till_A) * 100, 1)) + "%")
				worksheet_new.write(row + 3, col, round((antal_godkanda_elever / antal_elever_streck_till_A) * 100, 1), cell_grey_bg_center_border)

				if antal_elever_streck_till_A_F > 0:
					print("(X) Antal flickor (%) med godkänt betyg (A-E): " + str(round((antal_godkanda_elever_F / antal_elever_streck_till_A_F) * 100, 1)) + "%")
					worksheet_new.write(row + 4, col, round((antal_godkanda_elever_F / antal_elever_streck_till_A_F) * 100, 1), cell_grey_bg_center_border)
				else:
					worksheet_new.write(row + 4, col, "-", grade_ordinary)

				if antal_elever_streck_till_A_P > 0:
					print("(X) Antal pojkar (%) med godkänt betyg (A-E): " + str(round((antal_godkanda_elever_P / antal_elever_streck_till_A_P) * 100, 1)) + "%")
					worksheet_new.write(row + 5, col, round((antal_godkanda_elever_P / antal_elever_streck_till_A_P) * 100, 1), cell_grey_bg_center_border)
				else:
					worksheet_new.write(row + 5, col, "-", grade_ordinary)


		col += 1

	Stäng_excelfil()



# Funktion som tar bort decimaler och annat och snyggar till betygen

def Formattera_betyg(betyg):

	if ".0" in betyg:

		betyg = betyg.replace(".0", "")

	else:

		pass

	return betyg



# Funktion för att räkna ut om elev är pojke eller flicka

def Personnummer_till_kon(personnummer):

	pnr = int(personnummer[10])

	if (pnr % 2) == 0:
		kon = "Flicka"
		kon_kort = "F"
	
	else:  
		kon = "Pojke"
		kon_kort = "P"

	return kon_kort



# Funktion för att omvandla textbetyg till poäng, ex A = 20, B = 17.5 samt ############
# för att sätta formattering för celler med färger 

def Omvandla_betyg_till_tal(betyg):

	# Variabler för poäng för betyg

	betyg_A = 20
	betyg_B = 17.5
	betyg_C = 15
	betyg_D = 12.5
	betyg_E = 10
	betyg_F = 0
	betyg_streck = 0

	if "A" in betyg:
		betygspoang = betyg_A
		#betyg_color = grade_ordinary
	
	elif "B" in betyg:
		betygspoang = betyg_B
		#betyg_color = grade_ordinary
	
	elif "C" in betyg:
		betygspoang = betyg_C
		#betyg_color = grade_ordinary
	
	elif "D" in betyg:
		betygspoang = betyg_D
		#betyg_color = grade_ordinary
	
	elif "E" in betyg:
		betygspoang = betyg_E
		#betyg_color = grade_ordinary

	elif "F" in betyg:
		betygspoang = betyg_F
		#betyg_color = grade_ordinary

	elif "-" in betyg:
		betygspoang = betyg_streck

	elif "2" in betyg:
		betygspoang = ""

	elif "3" in betyg:
		betygspoang = ""
	
	else:
		betygspoang = ""
		#betyg_color = grade_ordinary

	return betygspoang



def DB_Skriv_Grundelevdata(personnummer, termin, klassbeteckning, arskurs, namn, kon):

	# DB - grunduppgifter
	# Kollar av om eleven finns OCH den inlästa terminen. Om inte, så skapas detta.

	if table_betyg.search((user.Personnummer == personnummer) & (user.Termin == termin)):
		pass
	else:
		table_betyg.insert({'Termin': termin, 'Personnummer': personnummer, 'Klass': klassbeteckning, 'Årskurs': arskurs, 'Namn': namn, 'Kön': kon, 'BL': '', 
			'BL-P': '', 'EN': '', 'EN-P': '', 'HKK': '', 'HKK-P': '', 'IDH': '', 'IDH-P': '', 'MA': '', 'MA-P': '', 'MLSPR': '', 
			'MLBET': '', 'MLBET-P': '', 'M1SPR': '', 'M1BET': '', 'M1BET-P': '', 'M2SPR': '', 'M2BET': '', 'M2BET-P': '', 'MU': '', 
			'MU-P': '', 'NO': '', 'NO-P': '', 'BI': '', 'BI-P': '', 'FY': '', 'FY-P': '', 'KE': '', 'KE-P': '', 'SO': '', 'SO-P': '', 
			'GE': '', 'GE-P': '', 'HI': '', 'HI-P': '', 'RE': '', 'RE-P': '', 'SH': '', 'SH-P': '', 'SL': '', 'SL-P': '', 
			'SV': '', 'SV-P': '', 'SVA': '', 'SVA-P': '', 'TN': '', 'TN-P': '', 'TK': '', 'TK-P': '', 'DA': '', 'DA-P': '', 
			'JU': '', 'JU-P': '', 'SV NP': '', 'SV NP-P': '', 'MA NP': '', 'MA NP-P': '', 'EN NP': '', 'EN NP-P': '', 
			'SO NP': '', 'SO NP-P': '', 'NO NP': '', 'NO NP-P': ''})



def Uppdatera_nya_elevuppgifter(personnummer, termin, klassbeteckning, arskurs, student_namn, kon):

	table_betyg.insert({'Termin': termin, 'Personnummer': personnummer, 'Klass': klassbeteckning, 'Årskurs': arskurs, 'Namn': namn, 'Kön': kon, 'BL': '', 
			'BL-P': '', 'EN': '', 'EN-P': '', 'HKK': '', 'HKK-P': '', 'IDH': '', 'IDH-P': '', 'MA': '', 'MA-P': '', 'MLSPR': '', 
			'MLBET': '', 'MLBET-P': '', 'M1SPR': '', 'M1BET': '', 'M1BET-P': '', 'M2SPR': '', 'M2BET': '', 'M2BET-P': '', 'MU': '', 
			'MU-P': '', 'NO': '', 'NO-P': '', 'BI': '', 'BI-P': '', 'FY': '', 'FY-P': '', 'KE': '', 'KE-P': '', 'SO': '', 'SO-P': '', 
			'GE': '', 'GE-P': '', 'HI': '', 'HI-P': '', 'RE': '', 'RE-P': '', 'SH': '', 'SH-P': '', 'SL': '', 'SL-P': '', 
			'SV': '', 'SV-P': '', 'SVA': '', 'SVA-P': '', 'TN': '', 'TN-P': '', 'TK': '', 'TK-P': '', 'DA': '', 'DA-P': '', 
			'JU': '', 'JU-P': '', 'SV NP': '', 'SV NP-P': '', 'MA NP': '', 'MA NP-P': '', 'EN NP': '', 'EN NP-P': '', 
			'SO NP': '', 'SO NP-P': '', 'NO NP': '', 'NO NP-P': ''})



def Skapa_excelfil(termin, klassbeteckning, arskurs, syfte):

	# Namn på den nya excelfil med betyg som skapas. ##############

	global workbook_new, worksheet_new
	global header_white_border, grade_ordinary, grade_ordinary_center_yellow_bg, header_white__no_border_center, grade_ordinary_center_red_bg, grade_ordinary_center_green_bg, cell_grey_bg_left_noborder, cell_grey_bg_center_border, cell_meta_left_noborder, grade_ordinary_center_blue_bg

	workbook_name = (mapp + klassbeteckning + "_" + termin + "_" + syfte + ".xlsx")
	workbook_new = xlsxwriter.Workbook(workbook_name)
	worksheet_new = workbook_new.add_worksheet()

	# SKAPA LAYOUT TILL NY EXCELFIL ######################################

	# Format i excelfilen ############################

	# Headers

	header_white_border = workbook_new.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 12, 'border': 1, 'border_color': '#ffffff'})
	header_white__no_border_center = workbook_new.add_format({'bold': False, 'font_name': 'Arial', 'font_size': 12, 'align': 'center', 'border': 1, 'border_color': '#ffffff'})
	header_setting_border_center = workbook_new.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 12, 'align': 'center', 'border': 1})
	header_setting_border_left = workbook_new.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 12, 'align': 'left', 'border': 1})
	header_grey_bg = workbook_new.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 12, 'bg_color': '#eeece2', 'border': 1})
	header_grey_bg_center = workbook_new.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 12, 'align': 'center', 'bg_color': '#eeece2', 'border': 1})

	
	# Betyg

	grade_ordinary = workbook_new.add_format({'bold': False, 'font_name': 'Arial', 'font_size': 12, 'align': 'center', 'border': 1})
	grade_ordinary_red_bg = workbook_new.add_format({'bold': False, 'font_name': 'Arial', 'font_size': 12, 'align': 'center', 'bg_color': '#ed6e46', 'font_color': '#ffffff', 'border': 1})
	grade_ordinary_green_bg = workbook_new.add_format({'bold': False, 'font_name': 'Arial', 'font_size': 12, 'align': 'center', 'bg_color': '#81c83d', 'font_color': '#ffffff', 'border': 1})

	grade_ordinary_left = workbook_new.add_format({'bold': False, 'font_name': 'Arial', 'font_size': 12, 'align': 'left', 'border': 1})
	grader_no_border_left = workbook_new.add_format({'bold': False, 'font_name': 'Arial', 'font_size': 12, 'align': 'left', 'border': 1, 'border_color': '#ffffff'})
	grade_ordinary_left_red_bg = workbook_new.add_format({'bold': False, 'font_name': 'Arial', 'font_size': 12, 'align': 'left', 'bg_color': '#ed6e46', 'font_color': '#ffffff', 'border': 1})
	grade_ordinary_left_green_bg = workbook_new.add_format({'bold': False, 'font_name': 'Arial', 'font_size': 12, 'align': 'left', 'bg_color': '#81c83d', 'font_color': '#ffffff', 'border': 1})
	grade_ordinary_center_red_bg = workbook_new.add_format({'bold': False, 'font_name': 'Arial', 'font_size': 12, 'align': 'center', 'bg_color': '#ed6e46', 'font_color': '#ffffff', 'border': 1})
	grade_ordinary_center_green_bg = workbook_new.add_format({'bold': False, 'font_name': 'Arial', 'font_size': 12, 'align': 'center', 'bg_color': '#81c83d', 'font_color': '#ffffff', 'border': 1})
	grade_ordinary_center_yellow_bg = workbook_new.add_format({'bold': False, 'font_name': 'Arial', 'font_size': 12, 'align': 'center', 'bg_color': 'yellow', 'font_color': '#000000', 'border': 1})
	grade_ordinary_center_blue_bg = workbook_new.add_format({'bold': False, 'font_name': 'Arial', 'font_size': 12, 'align': 'center', 'bg_color': '#65aac3', 'font_color': '#ffffff', 'border': 1})
	

	# Celler

	empty_cell_white_border = workbook_new.add_format({'border': 1, 'border_color': '#ffffff'})
	cell_grey_bg_left_noborder = workbook_new.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 12, 'align': 'left', 'bg_color': '#eeece2', 'border': 1, 'border_color': '#eeece2'})
	cell_grey_bg_center_border = workbook_new.add_format({'bold': False, 'font_name': 'Arial', 'font_size': 12, 'align': 'center', 'bg_color': '#eeece2', 'border': 1,})
	cell_meta_left_noborder = workbook_new.add_format({'bold': False, 'font_name': 'Arial', 'font_size': 10, 'align': 'left', 'bg_color': 'ffffff', 'border': 1, 'border_color': '#ffffff'})

	# Titel

	title_setting = workbook_new.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 16, 'border': 1, 'border_color': '#ffffff'})

	##################### Format slut


	# Justera bredd på kolumner i excelfilen.

	worksheet_new.set_column(0, 0, 30, empty_cell_white_border)
	worksheet_new.set_column(1, 1, 4, empty_cell_white_border)
	worksheet_new.set_column(2, 2, 1, empty_cell_white_border)
	worksheet_new.set_column(3, 7, 6, empty_cell_white_border)
	worksheet_new.set_column(8, 13, 8, empty_cell_white_border)
	worksheet_new.set_column(14, 30, 6, empty_cell_white_border)
	worksheet_new.set_column(31, 32, 8, empty_cell_white_border)
	worksheet_new.set_column(33, 33, 30, empty_cell_white_border)


	########## GRUNDSKAPANDE AV BETYGSKATALOG ##################

	# Skriv titel och lite annat

	worksheet_new.write(0,0, "Betygskatalog " + klassbeteckning, title_setting)
	worksheet_new.write(0,3, termin + " | " + syfte, title_setting)
	worksheet_new.write(1,0, "")
	worksheet_new.write(2,0, "Betygspoäng A-F Totalt", header_white_border)
	worksheet_new.write(3,0, "Betygspoäng A-F Flickor", header_white_border)
	worksheet_new.write(4,0, "Betygspoäng A-F Pojkar", header_white_border)
	worksheet_new.write(5,0, "Andel (%) med A-E Totalt", header_white_border)
	worksheet_new.write(6,0, "Andel (%) med A-E Flickor", header_white_border)
	worksheet_new.write(7,0, "Andel (%) med A-E Pojkar", header_white_border)

	worksheet_new.write(9,0, "Namn", header_setting_border_left)
	worksheet_new.write(9,1, "Kön", header_grey_bg)
	worksheet_new.write(9,2, "")
	worksheet_new.write(9,28, "")
	worksheet_new.write(9,31, "Summa", header_setting_border_left)
	worksheet_new.write(9,32, "Medel", header_setting_border_left)


	# Excel - Lista med ämnen i ämnesraden

	# Starta från första cellen

	row = 0
	col = 0

	# Skapa celler för ämnena

	for betyg in betygslista:
	    if "NP" in betyg:
	        worksheet_new.write(row +9, col +3, betyg, header_grey_bg_center)
	    else:
	        worksheet_new.write(row +9, col +3, betyg, header_setting_border_center)
	    col += 1


	Printa_inlasta_betyg_konsol_och_skapa_katalog_excel(termin, klassbeteckning, arskurs, syfte)



# Funktion som stänger den aktuella öppna excelfilen

def Stäng_excelfil():

	workbook_new.close()



# Tar emot betygspoäng och returnerar en färgad excel-cell

def Color_Points(poang):
	if poang < 10:
		color_bg = grade_ordinary_center_red_bg

	elif poang >= 17.5:
		color_bg = grade_ordinary_center_blue_bg

	elif poang >= 15:
		color_bg = grade_ordinary_center_green_bg

	else: 
		color_bg = grade_ordinary

	return color_bg



# Meny ####################################################

def Menu():

	global mapp

	print("\n--- BETYG | STATISTIK OCH FELSÖKNING ---\n")
	print("--- MENY ---\n")
	print("1. Läs in betygskatalog till databas och skapa betygskatalog för felsökning")
	print("2. Skapa betygsktalog för statistik")
	print("")
	val = input("Siffra + Enter >>> ")

	if val == "1":
		mapp = "betygskatalog_felsökning/"
		Konvertera_pdf_betygskatalog_till_xls()
		Läs_från_betygsfil_till_databas()
	
	elif val == "2":
		mapp = "betygskatalog_statistik/"
		print("\nVälj termin (ex VT2021)")
		termin = input(">>> ")
		print("\nVälj klass (ex 9A)")
		klassbeteckning = input(">>> ")
		Skapa_excelfil(termin, klassbeteckning, arskurs, "Statistik")
	
	else:
		pass



# -------------- Main --------------- #

def Main():

	Menu()



### MAIN PROGRAM ###

if __name__ == "__main__":
	Main()


