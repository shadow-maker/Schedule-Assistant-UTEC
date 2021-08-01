# Import standard libraries
from time import sleep
from datetime import datetime
import sys
import os
import json
import csv
import itertools
import math

# Import pre-req libraries
pipInstall = {"selenium": "selenium", "tabula": "tabula-py"}
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
	sys.exit(f"ERROR: {lib} no encontrado\nInstalar con el comando 'pip install {pipInstall[lib]}'")

class ScheduleAssistant:
	timeout = 30

	scheduleDataTable = []
	scheduleDataDict = {}

	downDir = ".down"
	pdfName = "horarios.pdf"

	saveDataCSV = True
	dataCSV = "horarios.csv"

	saveDataJSON = True
	dataJSON = "horarios.json"

	def __init__(self):
		pass

	def waitForPageLoad(self, elementToCheck, by=By.CLASS_NAME):
		try: # Wait for page to finish loading
			elementPresent = EC.presence_of_element_located((by, elementToCheck))
			WebDriverWait(self.br, self.timeout).until(elementPresent)
		except TimeoutException:
			print(">Loading took too much time!")
			return False
		return True

	def initWebdriver(self):
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

		self.br = webdriver.Firefox(executable_path="/Users/carlosandres/Documents/Python/UCalgaryRoomChecker/geckodriver", firefox_profile=profile)
	
	def login(self, email, passw):
		self.br.get("https://sistema-academico.utec.edu.pe")
		btn = self.br.find_element_by_tag_name("button")
		btn.click()

		self.waitForPageLoad("form", By.TAG_NAME)

		form = self.br.find_element_by_tag_name("form")
		field = [i for i in form.find_elements_by_tag_name("input") if i.get_attribute("type") == "email"][0]
		field.send_keys(email)

		buttons = self.br.find_elements_by_tag_name("button")
		btn = [i for i in buttons if len(i.find_elements_by_tag_name("span")) == 1][0]
		btn.click()

		sleep(5)

		form = self.br.find_element_by_tag_name("form")
		field = [i for i in form.find_elements_by_tag_name("input") if i.get_attribute("type") == "password"][0]
		field.send_keys(passw)

		buttons = self.br.find_elements_by_tag_name("button")
		btn = [i for i in buttons if len(i.find_elements_by_tag_name("span")) == 1][0]
		btn.click()

	def downloadSchedule(self):
		self.br.get("https://sistema-academico.utec.edu.pe/students/view/enabled-courses")

		self.waitForPageLoad("report", By.ID)

		btn = self.br.find_element_by_id("report")

		parentWindow = self.br.window_handles[0]

		btn.click()

		WebDriverWait(self.br, self.timeout).until(EC.number_of_windows_to_be(2))

		dtChecked = datetime.now()

		while not os.path.exists(self.downDir):
			sleep(0.1)
		
		fname = os.path.join(self.downDir, list(os.walk(self.downDir))[0][2][0])
		
		while os.stat(fname).st_size == 0:
			sleep(0.1)

		os.rename(fname, self.pdfName)

		while len(list(os.walk(self.downDir))[0][2]) > 0:
			sleep(0.1)

		os.rmdir(self.downDir)

		popup = [x for x in self.br.window_handles if x != parentWindow][0]
		self.br.switch_to.window(popup)
		self.br.close()
		self.br.switch_to.window(self.br.window_handles[0])
		return dtChecked
	
	def pdfToTable(self):
		tables = read_pdf("horarios.pdf", pages="all")

		self.scheduleDataTable = list(itertools.chain(*[[tuple(table.columns)] + list(zip(*[[(i.replace("\r", " ") if type(i) == str else ("" if math.isnan(i) else i)) for i in table[col].to_list()] for col in table])) for table in tables]))

		if self.saveDataCSV:
			with open(self.dataCSV, "w") as file:
				writer = csv.writer(file)
				for line in self.scheduleDataTable:
					writer.writerow(line)
		return self.scheduleDataTable

	def tableToDict(self):
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

		if self.saveDataJSON:
			with open(self.dataJSON, "w") as file:
				json.dump(self.scheduleDataDict, file, indent=4, ensure_ascii=False)

		return self.scheduleDataDict