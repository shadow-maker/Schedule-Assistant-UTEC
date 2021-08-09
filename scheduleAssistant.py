# Import standard libraries
import csv
import itertools
import json
import math
import sys
import time
import os
import platform
from datetime import datetime

# Import pre-req third party libraries
pipInstall = {"tabula": "tabula-py"}
try:
	from selenium import webdriver
	from selenium.webdriver.support.ui import WebDriverWait
	from selenium.webdriver.support import expected_conditions as EC
	from selenium.common.exceptions import NoSuchElementException
	from selenium.webdriver.common.by import By
	from selenium.common.exceptions import TimeoutException
	from tabula import read_pdf
except ModuleNotFoundError as error:
	lib = error.msg.split("'")[1]
	command = lib if lib not in pipInstall else pipInstall[lib]
	sys.exit(f"ERROR: Libreria {lib} no encontrada\nInstalar con el comando 'pip install {command}'")

class ScheduleAssistant:

	#
	# URLS
	#

	sisURL = "https://sistema-academico.utec.edu.pe"
	loginPage = "/access"
	downloadPage = "/students/view/enabled-courses"

	#
	# BOOLS
	#

	logCurrentProcess = True
	saveDataCSV = True
	saveDataJSON = True

	#
	# FILE / DIR NAMES
	#

	chromedriverName = "chromedriver"
	geckodriverName = "geckodriver"
	downDir = ".down"
	pdfName = "horarios.pdf"
	csvName = "horarios.csv"
	jsonName = "horarios.json"

	#
	# TEMP DATA CONTAINERS
	#

	coursesDataTable = []
	coursesDataDict = {}

	#
	# TIME VARS
	#

	timeout = 30

	#
	# INIT
	#

	# Constructor - Initializes with optional email and password
	def __init__(self, email="", passw=""):
		print("-" * 29 + "\n\n  SCHEDULE ASSISTANT - UTEC\n\n" + "-" * 29)
		self.email = email
		self.passw = passw
	
	# Destructor - Cleans the console
	def __del__(self):
		self.log("")

	# Initializes the webdriver with the selected browser
	def initWebdriver(self, selBrowser="C"):
		self.log("Inicializando web driver...")

		downDir = os.path.join(os.getcwd(), self.downDir)
		selBrowser = selBrowser.upper()

		if selBrowser == "C":
			options = webdriver.ChromeOptions()
			options.add_experimental_option("prefs", {
				"download.default_directory": downDir,
				"download.prompt_for_download": False,
				"download.directory_upgrade": True,
				"plugins.always_open_pdf_externally": True
			})
		elif selBrowser == "F":
			options = webdriver.FirefoxOptions()
			options.set_preference("browser.download.folderList", 2)
			options.set_preference("browser.download.dir", downDir)
			options.set_preference("browser.download.useDownloadDir", True)
			options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf,application/x-pdf")
			options.set_preference("browser.download.manager.useWindow", False)
			options.set_preference("browser.download.manager.focusWhenStarting", False)
			options.set_preference("browser.download.manager.alertOnEXEOpen", False)
			options.set_preference("browser.download.manager.showAlertOnComplete", False)
			options.set_preference("browser.download.manager.closeWhenDone", True)
			options.set_preference("pdfjs.disabled", True)
		else:
			self.error("\nbrowser solo puede ser 'C' (Chrome) o 'F' (Firefox)")
			return False

		instBr = lambda b, p, o : webdriver.Chrome(executable_path=p, options=o) if b == "C" else webdriver.Firefox(executable_path=p, options=o)

		driver = {"C": self.chromedriverName, "F": self.geckodriverName}[selBrowser]
		driver += ".exe" if platform.system() == "Windows" and ".exe" not in driver else ""
		try:
			self.br = instBr(selBrowser.upper(), driver, options)
		except:
			path = os.path.join(os.getcwd(), driver)
			if not os.path.exists(path):
				self.error(f"webdriver con nombre {driver} no encontrado en el PATH o en el directorio actual")
				return False
			self.br = instBr(selBrowser.upper(), path, options)
		return True
	
	#
	# LOG FUNCS
	#

	# Cleans last log message and prints a new one in the same line
	def log(self, msg):
		if self.logCurrentProcess:
			sys.stdout.write("\r" + (" " * os.get_terminal_size().columns))
			msg = f"# {msg}" if len(msg) > 0 else msg
			sys.stdout.write("\r" + msg)
			sys.stdout.flush()
	
	# Prints error messages
	def error(self, msg):
		print("\nERROR: " + msg)
	
	#
	# SAVE FILE FUNCS
	#
	
	# Saves a Python matrix as a CSV file
	def saveCSV(self, data=[]):
		if self.saveDataCSV:
			data = self.coursesDataTable if len(data) == 0 else data
			self.log("Guardando tabla (matriz) como CSV...")
			with open(self.csvName, "w") as file:
				writer = csv.writer(file)
				for line in data:
					writer.writerow(line)

	# Saves a Python dictionary as a JSON file
	def saveJSON(self, data={}):
		if self.saveDataJSON:
			data = self.coursesDataDict if len(data) == 0 else data
			self.log("Guardando diccionario como JSON...")
			with open(self.jsonName, "w") as file:
				json.dump(self.coursesDataDict, file, indent=4, ensure_ascii=False)

	#
	# WAIT FUNCS
	#

	# Waits for a certain element to appear in the webpage
	def waitForPageLoad(self, elementToCheck, by=By.ID):
		try:
			elementPresent = EC.presence_of_element_located((by, elementToCheck))
			WebDriverWait(self.br, self.timeout).until(elementPresent)
		except TimeoutException:
			self.error("Timeout!")
			return False
		return True
	
	# Waits until a boolean lambda func becomes True
	def waitUntilTrue(self, func, timeElapsed=0, interval=0.1):
		if func():
			return True
		elif timeElapsed > self.timeout:
			self.error("Timeout!")
			return False
		time.sleep(interval)
		return self.waitUntilTrue(func, timeElapsed + interval)

	#
	# WEB SCRAPER FUNCS
	#
	
	# Logs in the user
	def login(self, email="", passw=""):
		email = self.email if email == "" else email
		passw = self.passw if passw == "" else passw

		self.log("Iniciando sesion...")
		self.br.get(self.sisURL)
		if self.loginPage not in self.br.current_url:
			return True
		self.br.find_element_by_tag_name("button").click()

		if not self.waitForPageLoad("form", By.TAG_NAME):
			return False

		form = self.br.find_element_by_tag_name("form")
		field = [i for i in form.find_elements_by_tag_name("input") if i.get_attribute("type") == "email"][0]
		field.send_keys(email)

		buttons = self.br.find_elements_by_tag_name("button")
		[i for i in buttons if len(i.find_elements_by_tag_name("span")) == 1][0].click()

		if not self.waitForPageLoad("profileIdentifier"):
			return False
		time.sleep(1)

		form = self.br.find_element_by_tag_name("form")
		field = [i for i in form.find_elements_by_tag_name("input") if i.get_attribute("type") == "password"][0]
		field.send_keys(passw)

		buttons = self.br.find_elements_by_tag_name("button")
		btn = [i for i in buttons if len(i.find_elements_by_tag_name("span")) == 1][0]
		btn.click()
		return True

	# Navigates to enabled courses download page and downloads the courses pdf
	def downloadScheduleData(self):
		self.log("Navegando a pagina de descarga de cursos disponibles...")
		self.br.get(self.sisURL + self.downloadPage)
		if not self.waitForPageLoad("report"):
			return False

		btn = self.br.find_element_by_id("report")

		homeWindow = self.br.window_handles[0]

		self.log("Descargando cursos disponibles...")
		btn.click()

		WebDriverWait(self.br, self.timeout).until(EC.number_of_windows_to_be(2))

		dtChecked = datetime.now()

		fname = os.path.join(self.downDir, "pdf")

		if not self.waitUntilTrue(lambda : os.path.exists(fname)):
			return False

		if not self.waitUntilTrue(lambda : os.stat(fname).st_size > 0):
			return False
		
		os.rename(fname, self.pdfName)

		if not self.waitUntilTrue(lambda : len(list(os.walk(self.downDir))[0][2]) == 0):
			return False

		os.rmdir(self.downDir)

		popup = [win for win in self.br.window_handles if win != homeWindow][0]
		self.br.switch_to.window(popup)
		self.br.close()
		self.br.switch_to.window(self.br.window_handles[0])
		return dtChecked

	#
	# DATA PARSER FUNCS
	#
	
	# Scrapes the courses' schedules pdf and parses into a Python matrix
	def pdfToTable(self):
		if not os.path.exists(self.pdfName):
			self.error(f"El pdf de horarios con nombre {self.pdfName} no existe")
			return 0

		self.log("Leyendo tablas del pdf...")
		try:
			tables = read_pdf(self.pdfName, pages="all")
		except:
			self.error(f"No se pudo leer tablas del pdf {self.pdfName}. Posiblemente el formato sea el incorrecto")

		self.log("Parse-ando tabla de pdf a matriz de Python...")
		self.coursesDataTable = [tuple(tables[0].columns)] + list(itertools.chain(*[
			list(zip(*[[
				(
					" ".join(i.replace("\r", " ").replace("\n", " ").split(", ")[::-1]) if type(i) == str else ("" if math.isnan(i) else i)
				) for i in table[col].to_list()
			] for col in table])) for table in tables
		]))

		self.saveCSV()
		return self.coursesDataTable

	# Parses the schedule data table into a Python dictionary
	def tableToDict(self):
		self.log("Parse-ando matriz a diccionario de cursos...")
		keys = ("cod", "nom", "prof", "malla", "tipo", "mod", "sec", "ses", "hora", "tipo", "ubic", "vac", "mat")
		mat = [dict(zip(keys, i)) for i in self.coursesDataTable[1:]]

		self.coursesDataDict = {}

		dias = ["lun", "mar", "mie", "jue", "vie", "sab", "dom"]
		blacklist = []
		for i in mat:
			if i["tipo"] != "Semana General" or i["cod"] in blacklist:
				blacklist.append(i["cod"])
				continue
			if i["cod"] not in self.coursesDataDict:
				self.coursesDataDict[i["cod"]] = {
					"nombre": i["nom"],
					#"malla": i["malla"],
					"secciones": {}
				}
			if i["sec"] not in self.coursesDataDict[i["cod"]]["secciones"]:
				self.coursesDataDict[i["cod"]]["secciones"][i["sec"]] = {
					"vacantes": i["vac"],
					"matriculados": i["mat"],
					"sesiones": []
				}
			dia, hora = i["hora"].replace(" ", "").split(".")
			horaIn, horaFin = hora.split("-")
			self.coursesDataDict[i["cod"]]["secciones"][i["sec"]]["sesiones"].append({
				"sesion" : i["ses"],
				"dia" : dias.index(dia.lower()),
				"hora" : int(horaIn.split(":")[0]),
				"duracion" : int(horaFin.split(":")[0]) - int(horaIn.split(":")[0]),
				"docente" : i["prof"]
			})

		self.saveJSON()
		return self.coursesDataDict

	#
	# VALIDATION FUNCS
	#

	# Validates that the sessions in each section of a course don't conflict
	def validateCourse(self, cod):
		valid = True
		for secNum, sec in self.coursesDataDict[cod]["secciones"].items():
			if self.mergeClassesIntoWeekIfPossible([[i] for i in sec["sesiones"]]) == []:
				self.error(f"Conflicto de horarios encontrado entre las sesiones de la seccion {secNum} del curso {cod}")
				valid = False
		return valid

	# Checks that every extracted course is valid
	def validateCoursesData(self):
		self.log("Validando que no hayan conflictos entre las sesiones de cada seccion de cada curso...")
		return sum([int(not self.validateCourse(c)) for c in self.coursesDataDict]) == 0
	
	#
	# SCHEDULE GENERATOR FUNCS
	#

	# Adds the course' info into every session of a section
	def addCourseInfoToSessions(self, cod, sec):
		return [{
			"codigo" : cod,
			"nombre" : self.coursesDataDict[cod]["nombre"],
			"seccion" : sec,
			"vacantes" : self.coursesDataDict[cod]["secciones"][sec]["vacantes"],
			"matriculados" : self.coursesDataDict[cod]["secciones"][sec]["matriculados"],
			**ses
		} for ses in self.coursesDataDict[cod]["secciones"][sec]["sesiones"]]

	# Merges a list of classes (sections) into a week matrix if no conflict is found
	def mergeClassesIntoWeekIfPossible(self, classes=[]):
		week = [[{} for j in range(24)] for i in range(7)]
		for sessions in classes:
			for ses in sessions:
				for i in range(ses["duracion"]):
					if week[ses["dia"]][ses["hora"]+ i] != {}:
						return []
					week[ses["dia"]][ses["hora"]+ i] = ses
		return week
	
	# Gets all possible combinations of classes from a selected list of courses
	def getClassCombinations(self, courses):
		self.log("Generando todas las posibles combinaciones de secciones con los cursos {courses}...")
		return list(itertools.product(*[
			[self.addCourseInfoToSessions(cod, sec) for sec in self.coursesDataDict[cod]["secciones"]] for cod in courses
		]))


	# Gets all possible schedules from a selected list of courses
	def getPossibleSchedules(self, courses):
		self.log("Generando todos los posibles horarios con los cursos {courses}...")
		possibleSchedules = []
		for comb in self.getClassCombinations(courses):
			schedule = self.mergeClassesIntoWeekIfPossible(comb)
			if schedule != []:
				possibleSchedules.append(schedule)
		return possibleSchedules

	#
	# FILTERING FUNCS
	#

	def filterBy(self, func, allSes=True):
		result = {}
		for code, courseData in self.coursesDataDict.items():
			secsFound = {}
			for secNum, secData in courseData["secciones"].items():
				sesFound = len([ses for ses in secData["sesiones"] if func(ses)])
				if sesFound > 0 and (sesFound == len(secData["sesiones"]) or not allSes):
					secsFound[secNum] = secData
			if len(secsFound) > 0:
				result[code] = {"nombre" : courseData["nombre"], "secciones": secsFound}
		return result

	def filterByProf(self, query=""):
		return self.filterBy(lambda ses : query.lower() in ses["docente"].lower(), False)

	def filterByMinBegTime(self, time=0):
		return self.filterBy(lambda ses : time <= ses["hora"])

	def filterByMaxEndTime(self, time=24):
		return self.filterBy(lambda ses : time >= ses["hora"] + ses["duracion"])

	def filterByDurTime(self, time=2):
		return self.filterBy(lambda ses : time == ses["duracion"])	
	
	#
	# UI FUNCS
	#

	def optionSelector(self, options):
		print("Selecciona una opcion:")
		print("\n".join([f"[{i + 1}] {options[i]}" for i in range(len(options))]))
		selection = ""
		while selection not in [str(i + 1) for i in range(len(options))]:
			selection = input(">")
		return int(selection) - 1
	

	def boolSelector(self):
		print("Selecciona [si] o [no]:")
		yOps = ["si", "s", "yes", "y", "1", "true", "t"]
		nOps = ["no", "n", "0", "false", "f"]
		selection = ""
		while selection not in yOps + nOps:
			selection = input(">").lower()
		return selection in yOps
	
	def downloadPDF(self):
		print(f"Puede descargar el pdf de horarios disonibles de {self.sisURL + self.downloadPage} manualmente, o descargarlo automaticamente con este programa")
		print("Para descargar el archivo de data automaticamente deberá descargar el webdriver para el browser de su eleccion")
		print("Desea descargar el archivo de data automaticamente?")

		if self.boolSelector():
			print("Seleccione el browser que desea usar [C] Chrome o [F] Firefox:")
			selBrowser = ""
			while selBrowser not in ["C", "F"]:
				selBrowser = input(">").upper()
			if not self.initWebdriver(selBrowser):
				return False
			self.log("")
			if not self.login(input("Ingrese su email de la UTEC: "), input("Ingrese su password de la UTEC: ")):
				return False
			if not self.downloadScheduleData():
				return False
			if not self.pdfToTable():
				return False
			if not self.tableToDict():
				return False
			self.log("")
			return True
		print("Puede volver a inicar este programa cuando haya descargado el archivo de data de horarios de los cursos disponibles")
		return False
	

	def printAvailableCourses(self):
		for course, data in self.coursesDataDict.items():
			print(f"{course} - {data['nombre']} ({len(data['secciones'])} secciones)")

	
	def begin(self):
		exists = dict(zip(["PDF", "CSV", "JSON"], [os.path.exists(i) for i in [self.pdfName, self.csvName, self.jsonName]]))
		existent = [i for i in exists if exists[i]]

		if len(existent) > 0:
			print(f"Se encontraron archivos {', '.join(existent)} con la data de horarios de los cursos disponibles")
			print("Desea...")
			op = self.optionSelector([f"Cargar data desde archivo {i}" for i in existent] + ["Volver a descargar pdf de horarios de los cursos disponibles"])
			if op == len(existent):
				if not self.downloadPDF():
					return False
			elif existent[op] == "PDF":
				if not self.pdfToTable():
					return False
				if not self.tableToDict():
					return False
			elif existent[op] == "CSV":
				with open(self.csvName, "r") as file:
					self.coursesDataTable = [line.replace("\n", "").split(",") for line in file.readlines()]
				if not self.tableToDict():
					return False
			elif existent[op] == "JSON":
				with open(self.jsonName, "r") as file:
					try:
						self.coursesDataDict = json.load(file)
					except:
						self.error(f"No se pudo leer archivo JSON {self.jsonName}")
						return False
			self.log("")
		else:
			print(f"No se encontro el archivo de data de horarios de los cursos disponibles en PDF ({self.pdfName}), CSV ({self.csvName}), o JSON ({self.jsonName})")
			if not self.downloadPDF():
				return False
		
		if not self.validateCoursesData():
			return False

		self.log("")
		print(f"Se encontraron {len(self.coursesDataDict)} cursos disponibles para la matricula")
		print("Desea mostrarlos?")
		if self.boolSelector():
			self.printAvailableCourses()