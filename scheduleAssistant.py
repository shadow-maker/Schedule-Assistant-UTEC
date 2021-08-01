# Import standard libraries
import csv
import itertools
import json
import math
import sys
import time
import os
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

	homePage = "https://sistema-academico.utec.edu.pe"
	downloadPage = "https://sistema-academico.utec.edu.pe/students/view/enabled-courses"


	#
	# TEMP DATA CONTAINERS
	#

	scheduleDataTable = []
	scheduleDataDict = {}

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
	# BOOLS
	#

	saveDataCSV = True
	saveDataJSON = True

	#
	# TIME VARS
	#

	timeout = 30

	#
	# INIT
	#

	def __init__(self, email="", passw=""):
		print("-" * 29 + "\n\n  SCHEDULE ASSISTANT - UTEC\n\n" + "-" * 29)
		self.email = email
		self.passw = passw
	
	def __del__(self):
		self.log("")

	def initWebdriver(self):
		self.log(">Initializing web driver...")
		profile = webdriver.FirefoxProfile()

		profile.set_preference("browser.download.dir", os.path.join(os.getcwd(), self.downDir));
		profile.set_preference("browser.download.folderList", 2);
		profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf,application/x-pdf");
		profile.set_preference("browser.download.manager.showWhenStarting", False);
		profile.set_preference("browser.helperApps.neverAsk.openFile", "application/pdf,application/x-pdf");
		profile.set_preference("browser.helperApps.alwaysAsk.force", False);
		profile.set_preference("browser.download.manager.useWindow", False);
		profile.set_preference("browser.download.manager.focusWhenStarting", False);
		profile.set_preference("browser.download.manager.alertOnEXEOpen", False);
		profile.set_preference("browser.download.manager.showAlertOnComplete", False);
		profile.set_preference("browser.download.manager.closeWhenDone", True);
		profile.set_preference("pdfjs.disabled", True);

		path = os.path.join(os.getcwd(), self.geckodriverName)
		self.br = webdriver.Firefox(executable_path=path, firefox_profile=profile)
	
	#
	# LOG FUNCS
	#

	def log(self, msg):
		sys.stdout.write("\r" + (" " * os.get_terminal_size().columns))
		sys.stdout.write("\r" + msg)
		sys.stdout.flush()
	
	def saveCSV(self, data=[]):
		if self.saveDataCSV:
			data = self.scheduleDataTable if len(data) == 0 else data
			self.log(">Saving table as CSV...")
			with open(self.csvName, "w") as file:
				writer = csv.writer(file)
				for line in data:
					writer.writerow(line)

	def saveJSON(self, data={}):
		if self.saveDataJSON:
			data = self.scheduleDataDict if len(data) == 0 else data
			self.log(">Saving dictionary as JSON...")
			with open(self.jsonName, "w") as file:
				json.dump(self.scheduleDataDict, file, indent=4, ensure_ascii=False)
	
	#
	# WEB SCRAPER FUNCS
	#

	def waitForPageLoad(self, elementToCheck, by=By.CLASS_NAME):
		#self.log(">Waiting for page to finish loading...")
		try: # Wait for page to finish loading
			elementPresent = EC.presence_of_element_located((by, elementToCheck))
			WebDriverWait(self.br, self.timeout).until(elementPresent)
		except TimeoutException:
			print(">Loading took too much time!")
			return False
		return True
	
	def login(self, email="", passw=""):
		self.log(">Logging in...")
		email = self.email if email == "" else email
		passw = self.passw if passw == "" else passw
		self.br.get(self.homePage)
		btn = self.br.find_element_by_tag_name("button")
		btn.click()

		self.waitForPageLoad("form", By.TAG_NAME)

		form = self.br.find_element_by_tag_name("form")
		field = [i for i in form.find_elements_by_tag_name("input") if i.get_attribute("type") == "email"][0]
		field.send_keys(email)

		buttons = self.br.find_elements_by_tag_name("button")
		btn = [i for i in buttons if len(i.find_elements_by_tag_name("span")) == 1][0]
		btn.click()

		time.sleep(5)

		form = self.br.find_element_by_tag_name("form")
		field = [i for i in form.find_elements_by_tag_name("input") if i.get_attribute("type") == "password"][0]
		field.send_keys(passw)

		buttons = self.br.find_elements_by_tag_name("button")
		btn = [i for i in buttons if len(i.find_elements_by_tag_name("span")) == 1][0]
		btn.click()

	def downloadScheduleData(self):
		self.log(">Navigating to data download page...")
		self.br.get(self.downloadPage)
		self.waitForPageLoad("report", By.ID)

		btn = self.br.find_element_by_id("report")

		parentWindow = self.br.window_handles[0]

		btn.click()

		self.log(">Downloading schedule data...")

		WebDriverWait(self.br, self.timeout).until(EC.number_of_windows_to_be(2))

		dtChecked = datetime.now()

		while not os.path.exists(self.downDir):
			time.sleep(0.1)
		
		fname = os.path.join(self.downDir, "pdf")
		
		while os.stat(fname).st_size == 0:
			time.sleep(0.1)

		os.rename(fname, self.pdfName)

		while len(list(os.walk(self.downDir))[0][2]) > 0:
			time.sleep(0.1)

		os.rmdir(self.downDir)

		popup = [x for x in self.br.window_handles if x != parentWindow][0]
		self.br.switch_to.window(popup)
		self.br.close()
		self.br.switch_to.window(self.br.window_handles[0])
		return dtChecked

	#
	# DATA PARSER FUNCS
	#
	
	def pdfToTable(self):
		self.log(">Parsing table from pdf into python matrix...")
		tables = read_pdf("horarios.pdf", pages="all")

		self.scheduleDataTable = list(itertools.chain(*[[tuple(table.columns)] + list(zip(*[[(i.replace("\r", " ") if type(i) == str else ("" if math.isnan(i) else i)) for i in table[col].to_list()] for col in table])) for table in tables]))

		self.saveCSV()
		return self.scheduleDataTable

	def tableToDict(self):
		self.log(">Parsing matrix into data dictionary...")
		keys = ("cod", "nom", "prof", "malla", "tipo", "mod", "sec", "ses", "hora", "tipo", "ubic", "vac", "mat")
		mat = [dict(zip(keys, i)) for i in self.scheduleDataTable[1:]]

		dias = ["lun", "mar", "mie", "jue", "vie", "sab", "dom"]

		self.scheduleDataDict = {}

		for i in mat:
			if i["tipo"] == "Semana General":
				if i["cod"] not in self.scheduleDataDict:
					self.scheduleDataDict[i["cod"]] = {
						"nombre": i["nom"],
						"malla": i["malla"],
						"secciones": {}
					}
				if i["sec"] not in self.scheduleDataDict[i["cod"]]["secciones"]:
					self.scheduleDataDict[i["cod"]]["secciones"][i["sec"]] = {
						"vacantes": i["vac"],
						"matriculados": i["mat"],
						"sesiones": []
					}
				i["hora"] = i["hora"].replace(" ", "")
				dia, hora = i["hora"].split(".")
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