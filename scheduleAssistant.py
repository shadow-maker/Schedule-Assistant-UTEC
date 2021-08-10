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
	# CONSTANTS
	#

	dias = ["lun", "mar", "mie", "jue", "vie", "sab", "dom"]

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
			options.add_experimental_option("excludeSwitches", ["enable-automation"])
			options.add_experimental_option('useAutomationExtension', False)
			options.add_argument('--disable-blink-features=AutomationControlled')
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
				self.error(f"Webdriver con nombre {driver} no encontrado en el PATH o en el directorio actual")
				return False
			try:
				self.br = instBr(selBrowser.upper(), path, options)
			except Exception as e:
				print(e)
				self.error("No se pudo inicializar webdriver. Es posible que el webdriver descargado no es el correcto para su sistema operativo y/o version de su browser")
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
		print("\n---\nERROR: " + msg)
	
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
					if len(line) > 0:
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

		if not self.waitUntilTrue(lambda : os.path.exists(".down")):
			return False

		if not self.waitUntilTrue(lambda : len(list(os.walk(".down"))[0][2]) > 0):
			return False
		
		time.sleep(1)

		fname = os.path.join(self.downDir, list(os.walk(".down"))[0][2][0])

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
		except Exception as e:
			print(e)
			self.error(f"No se pudo leer tablas del pdf {self.pdfName}. Posiblemente el formato sea el incorrecto")

		self.log("Parse-ando tabla de pdf a matriz de Python...")
		formatCell = lambda cell : " ".join(cell.replace("\r", " ").replace("\n", " ").split(", ")[::-1]) if type(cell) == str else ("" if math.isnan(cell) else int(cell))
		self.coursesDataTable = list(itertools.chain(*[
			[tuple(formatCell(i) for i in table.columns)]
			+
			list(zip(*[[formatCell(i) for i in table[col].to_list()] for col in table]))
			for table in tables
		]))

		self.saveCSV()
		return self.coursesDataTable

	# Parses the schedule data table into a Python dictionary
	def tableToDict(self):
		self.log("Parse-ando matriz a diccionario de cursos...")
		keys = ("cod", "nom", "prof", "malla", "tipo", "mod", "sec", "ses", "hora", "sem", "ubic", "vac", "mat")
		mat = [dict(zip(keys, i)) for i in self.coursesDataTable[1:]]

		self.coursesDataDict = {}

		blacklist = []
		for i in mat:
			if i["sem"] != "Semana General" or i["cod"] in blacklist:
				blacklist.append(i["cod"])
				continue
			if i["cod"] not in self.coursesDataDict:
				self.coursesDataDict[i["cod"]] = {
					"nombre": i["nom"],
					"malla": i["malla"],
					"secciones": {}
				}
			if str(i["sec"]) not in self.coursesDataDict[i["cod"]]["secciones"]:
				self.coursesDataDict[i["cod"]]["secciones"][str(i["sec"])] = {
					"vacantes": int("0" + str(i["vac"])),
					"matriculados": int("0" + str(i["mat"])),
					"sesiones": []
				}
			dia, hora = i["hora"].replace(" ", "").split(".")
			horaIn, horaFin = hora.split("-")
			self.coursesDataDict[i["cod"]]["secciones"][str(i["sec"])]["sesiones"].append({
				"sesion" : i["ses"],
				"dia" : self.dias.index(dia.lower()),
				"hora" : int(horaIn.split(":")[0]),
				"duracion" : int(horaFin.split(":")[0]) - int(horaIn.split(":")[0]),
				"docente" : i["prof"]
			})
		
		self.log("Ordenando diccionario de cursos...")
		for courseData in self.coursesDataDict.values():
			for secData in courseData["secciones"].values():
				secData["sesiones"] = sorted(secData["sesiones"], key=lambda i : (i["dia"], i["hora"], i["duracion"]))
			courseData["secciones"] = {str(sec) : courseData["secciones"][sec] for sec in sorted(courseData["secciones"])}

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
		self.log("Validando la data de horarios de cursos disponibles...")
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

	def filterBy(self, func, allSes=True, baseData={}):
		baseData = self.coursesDataDict if baseData == {} else baseData
		result = {}
		for code, courseData in baseData.items():
			secsFound = {}
			for secNum, secData in courseData["secciones"].items():
				sesFound = len([ses for ses in secData["sesiones"] if func(ses)])
				if sesFound > 0 and (sesFound == len(secData["sesiones"]) or not allSes):
					secsFound[secNum] = secData
			if len(secsFound) > 0:
				result[code] = {"nombre" : courseData["nombre"], "secciones": secsFound}
		return result

	def filterByProf(self, query="", baseData={}):
		return self.filterBy(lambda ses : query.lower() in ses["docente"].lower(), False, baseData)

	def filterByMinBegTime(self, time=0, baseData={}):
		return self.filterBy(lambda ses : time <= ses["hora"], True, baseData)

	def filterByMaxEndTime(self, time=24, baseData={}):
		return self.filterBy(lambda ses : time >= ses["hora"] + ses["duracion"], True, baseData)

	def filterByDurTime(self, time=2, baseData={}):
		return self.filterBy(lambda ses : time == ses["duracion"], True, baseData)	
	
	#
	# UI FUNCS
	#

	def optionValueSelector(self, options):
		print("Selecciona una opcion:")
		print("\n".join([f"[{i}]" for i in options]))
		selection = ""
		while selection not in options:
			selection = input(">")
		return selection

	def optionIndexSelector(self, options):
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
		print()
		for course, data in self.coursesDataDict.items():
			print(f"{course} - {data['nombre']} ({len(data['secciones'])} secciones)")
	
	def printCourseInfo(self, course, secs={}):
		if course not in self.coursesDataDict:
			return False
		print(f"\n{course} - {self.coursesDataDict[course]['nombre']}")
		secs = self.coursesDataDict[course]["secciones"] if secs == {} else secs
		print(f"Secciones ({len(secs)}):")
		for secNum, secData in secs.items():
			print(f"  {secNum}) Matriculados: {secData['matriculados']} / {secData['vacantes']}")
			for ses in secData["sesiones"]:
				print(f"     + {ses['sesion']}")
				print(f"       {ses['docente']}")
				print(f"       {self.dias[ses['dia']].upper()} {str(ses['hora']).zfill(2)}:00–{str(ses['hora'] + ses['duracion']).zfill(2)}:00 ({ses['duracion']}hrs)")
		return True
	
	def printCoursesInfo(self, data={}):
		data = self.coursesDataDict if data == {} else data
		for course in data:
			self.printCourseInfo(course, data[course]["secciones"])
			print()

	def filterMenu(self, data={}):
		print("\nFILTRAR CURSOS")
		op = self.optionIndexSelector([
			"Filtrar por profesor",
			"Filtrar por hora minima de inicio de las sesiones",
			"Filtrar por hora maxima de fin de las sesiones",
			"Filtrar por hora de duracion de la sesion"
		])

		if op == 0:
			data = self.filterByProf(input("Ingrese profesor a buscar:\n>"), data)
		elif op == 1:
			print("Ingrese la hora minima de inicio de las sesiones:")
			time = -1
			while time < 0 or time > 23:
				try:
					time = int(input(">"))
				except:
					continue
			data = self.filterByMinBegTime(time, data)
		elif op == 2:
			print("Ingrese la hora maxima de fin de las sesiones:")
			time = -1
			while time < 0 or time > 23:
				try:
					time = int(input(">"))
				except:
					continue
			data = self.filterByMaxEndTime(time, data)
		elif op == 3:
			print("Ingrese la hora de duracion de la sesion:")
			time = -1
			while time < 1 or time > 5:
				try:
					time = int(input(">"))
				except:
					continue
			data = self.filterByDurTime(time, data)
		
		secsFound = sum([len(data[course]["secciones"]) for course in data])
		print(f"\nSe encontraron {secsFound} secciones en {len(data)} cursos")
		print("Que desea hacer con la informacion filtrada:")

		op = 0
		while op == 0:
			op = self.optionIndexSelector([
				"Mostrar informacion de los cursos filtrados",
				"Aplicar un nuevo filtro",
				"Salir"
			])

			if op == 0:
				self.printCoursesInfo(data)
			elif op == 1:
				return self.filterMenu(data)

		return data
	
	def mainMenu(self):
		op = 0
		while op != 3:
			print()
			print("-" * 17 + "\n Menu Principal\n")
			print(">>")
			print(f"  Cursos disponibles para la matricula: {len(self.coursesDataDict)}")
			print(">>\n")

			op = self.optionIndexSelector([
				"Mostrar cursos disponibles",
				"Mostrar informacion de un curso",
				"Filtrar cursos",
				"Salir"
			])

			if op == 0:
				self.printAvailableCourses()
			elif op == 1:
				self.printCourseInfo(self.optionValueSelector(list(self.coursesDataDict.keys())))
			elif op == 2:
				self.filterMenu()
	
	def begin(self):
		exists = dict(zip(["JSON", "CSV", "PDF"], [os.path.exists(i) for i in [self.jsonName, self.csvName, self.pdfName]]))
		existent = [i for i in exists if exists[i]]

		if len(existent) > 0:
			print(f"Se encontraron archivos {', '.join(existent)} con la data de horarios de los cursos disponibles")
			print("Desea...")
			op = self.optionIndexSelector([f"Cargar data desde archivo {i}" for i in existent] + ["Volver a descargar pdf de horarios de los cursos disponibles"])
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
					except Exception as e:
						print(e)
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
		
		self.mainMenu()

# BEGIN COMMAND
if __name__ == "__main__":
	assistant = ScheduleAssistant()
	assistant.begin()