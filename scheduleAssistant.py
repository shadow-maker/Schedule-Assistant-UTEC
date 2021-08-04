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

	scheduleDataTable = []
	scheduleDataDict = {}

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
			return None

		instBr = lambda b, p, o : webdriver.Chrome(executable_path=p, options=o) if b == "C" else webdriver.Firefox(executable_path=p, options=o)

		driver = {"C": self.chromedriverName, "F": self.geckodriverName}[selBrowser]
		driver += ".exe" if platform.system() == "Windows" and ".exe" not in driver else ""
		try:
			self.br = instBr(selBrowser.upper(), driver, options)
		except:
			path = os.path.join(os.getcwd(), driver)
			if not os.path.exists(path):
				self.error(f"webdriver con nombre {driver} no encontrado en el PATH o en el directorio actual")
				return None
			self.br = instBr(selBrowser.upper(), path, options)
	
	#
	# LOG FUNCS
	#

	# Cleans last log message and prints a new one in the same line
	def log(self, msg):
		if self.logCurrentProcess:
			sys.stdout.write("\r" + (" " * os.get_terminal_size().columns))
			msg = ">" + msg if len(msg) > 0 else msg
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
			data = self.scheduleDataTable if len(data) == 0 else data
			self.log("Guardando tabla (matriz) como CSV...")
			with open(self.csvName, "w") as file:
				writer = csv.writer(file)
				for line in data:
					writer.writerow(line)

	# Saves a Python dictionary as a JSON file
	def saveJSON(self, data={}):
		if self.saveDataJSON:
			data = self.scheduleDataDict if len(data) == 0 else data
			self.log("Guardando diccionario como JSON...")
			with open(self.jsonName, "w") as file:
				json.dump(self.scheduleDataDict, file, indent=4, ensure_ascii=False)

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
			return None

		form = self.br.find_element_by_tag_name("form")
		field = [i for i in form.find_elements_by_tag_name("input") if i.get_attribute("type") == "email"][0]
		field.send_keys(email)

		buttons = self.br.find_elements_by_tag_name("button")
		[i for i in buttons if len(i.find_elements_by_tag_name("span")) == 1][0].click()

		if not self.waitForPageLoad("profileIdentifier"):
			return None
		time.sleep(1)

		form = self.br.find_element_by_tag_name("form")
		field = [i for i in form.find_elements_by_tag_name("input") if i.get_attribute("type") == "password"][0]
		field.send_keys(passw)

		buttons = self.br.find_elements_by_tag_name("button")
		btn = [i for i in buttons if len(i.find_elements_by_tag_name("span")) == 1][0]
		btn.click()

	# Navigates to enabled courses download page and downloads the courses pdf
	def downloadScheduleData(self):
		self.log("Navegando a pagina de descarga de cursos disponibles...")
		self.br.get(self.sisURL + self.downloadPage)
		if not self.waitForPageLoad("report"):
			return None

		btn = self.br.find_element_by_id("report")

		homeWindow = self.br.window_handles[0]

		self.log("Descargando cursos disponibles...")
		btn.click()

		WebDriverWait(self.br, self.timeout).until(EC.number_of_windows_to_be(2))

		dtChecked = datetime.now()

		fname = os.path.join(self.downDir, "pdf")

		if not self.waitUntilTrue(lambda : os.path.exists(fname)):
			return None

		if not self.waitUntilTrue(lambda : os.stat(fname).st_size > 0):
			return None
		
		os.rename(fname, self.pdfName)

		if not self.waitUntilTrue(lambda : len(list(os.walk(self.downDir))[0][2]) == 0):
			return None

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
		self.log("Parse-ando tabla de pdf a matriz de Python...")
		tables = read_pdf(self.pdfName, pages="all")

		self.scheduleDataTable = [tuple(tables[0].columns)] + list(itertools.chain(*[
			list(zip(*[[
				(
					" ".join(i.replace("\r", " ").replace("\n", " ").split(", ")[::-1]) if type(i) == str else ("" if math.isnan(i) else i)
				) for i in table[col].to_list()
			] for col in table])) for table in tables
		]))

		self.saveCSV()
		return self.scheduleDataTable

	# Parses the schedule data table into a Python dictionary
	def tableToDict(self):
		self.log("Parse-ando matriz a diccionario de cursos...")
		keys = ("cod", "nom", "prof", "malla", "tipo", "mod", "sec", "ses", "hora", "tipo", "ubic", "vac", "mat")
		mat = [dict(zip(keys, i)) for i in self.scheduleDataTable[1:]]

		self.scheduleDataDict = {}

		dias = ["lun", "mar", "mie", "jue", "vie", "sab", "dom"]
		blacklist = []
		for i in mat:
			if i["tipo"] != "Semana General" or i["cod"] in blacklist:
				blacklist.append(i["cod"])
				continue
			if i["cod"] not in self.scheduleDataDict:
				self.scheduleDataDict[i["cod"]] = {
					"nombre": i["nom"],
					#"malla": i["malla"],
					"secciones": {}
				}
			if i["sec"] not in self.scheduleDataDict[i["cod"]]["secciones"]:
				self.scheduleDataDict[i["cod"]]["secciones"][i["sec"]] = {
					"vacantes": i["vac"],
					"matriculados": i["mat"],
					"sesiones": []
				}
			dia, hora = i["hora"].replace(" ", "").split(".")
			horaIn, horaFin = hora.split("-")
			self.scheduleDataDict[i["cod"]]["secciones"][i["sec"]]["sesiones"].append({
				"sesion" : i["ses"],
				"dia" : dias.index(dia.lower()),
				"hora" : int(horaIn.split(":")[0]),
				"duracion" : int(horaFin.split(":")[0]) - int(horaIn.split(":")[0]),
				"docente" : i["prof"]
			})

		self.saveJSON()
		return self.scheduleDataDict