from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import pyautogui
import PyPDF2
import pandas as pd
import os
import requests

def login(driver):
	driver.find_element(By.XPATH, '//*[@id="username_input"]').send_keys('kh intern')
	driver.find_element(By.XPATH, '//*[@id="password_input"]').send_keys('kmc123')
	driver.find_element(By.XPATH, '/html/body/form/div/div[1]/table/tbody/tr[4]/td[2]/div/table/tbody/tr[13]/td/input').click()
	pyautogui.hotkey('tab')
	pyautogui.hotkey('enter')
	driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td[1]/div/div[2]/ul[3]/li/a').click()

def settings(driver):
	driver.switch_to.frame(driver.find_element(By.XPATH, '//*[@id="ifm"]'))
	driver.find_element(By.XPATH, '//*[@id="ALL_checkbox"]').click()
	driver.find_element(By.XPATH, '//*[@id="bodyPart"]').send_keys('ABD')
	#driver.find_element(By.XPATH, '//*[@id="start_date_input"]').send_keys('05-10-2023 00:00:00')
	#driver.find_element(By.XPATH, '//*[@id="end_date_input"]').send_keys('04-10-2023 23:59:59')
	driver.find_element(By.XPATH, '//*[@id="CT_checkbox"]').click()
	driver.find_element(By.XPATH, '//*[@id="MR_checkbox"]').click()
	driver.find_element(By.XPATH, '//*[@id="showDiv"]').click()
	nil = input('Select the required dates and press enter on this screen to continue...')
	return driver

def getAccessionNumbers():
	driver = webdriver.Chrome()
	driver.get("https://mahepacs.manipal.edu/")
	login(driver)
	driver = settings(driver)
	acc_lst = list()
	try:
		for k in range(0, 100):
			for i in range(1, 21):
				name = driver.find_element(By.XPATH, '/html/body/table/tbody/tr[1]/td/form/table[2]/tbody/tr/td/table[2]/tbody/tr[2]/td/table/tbody[2]/tr['+str(i)+']/td[16]/font')
				acc_lst.append(name.text)
			driver.find_element(By.XPATH, '/html/body/table/tbody/tr[1]/td/form/table[2]/tbody/tr/td/table[2]/tbody/tr[2]/td/table/tbody[3]/tr[1]/td/table/tbody/tr/td[2]/nobr[3]/label/img').click()
			time.sleep(8)
	except:
		print(acc_lst)
		df = pd.DataFrame({ 'Accession Numbers' : acc_lst })
		df.to_excel('Accession Numbers.xlsx', index = False)
	driver.quit()

def downloadReport(accession):
	driver = webdriver.Chrome()
	driver.get("https://mahepacs.manipal.edu/")
	login(driver)
	time.sleep(10)
	for acc in accession:
		if len(str(acc)) < 5:
			print('Invalid Accession Number')
			valid = False
		else:
			valid = True
			if str(acc)+'.pdf' in os.listdir():
				print(str(acc)+'.pdf exists and hence skipped.')
				valid = False
			else:
				driver.get("https://mahepacs.manipal.edu/get_pdf_report?accessionNo=NUMBER&amp;loadType=REPORT&amp;centerId=1".replace('NUMBER', acc))
				time.sleep(2)
		try:
			if valid:
				driver.find_element(By.XPATH, '/html/body/table/tbody/tr[1]/td/label/h3')
				print('Report Not Found')
		except:
			print('Report Found')
			pyautogui.click(100, 100)
			pyautogui.hotkey('ctrl', 's')
			time.sleep(1)
			pyautogui.write(acc)
			pyautogui.press('enter')
			time.sleep(1)
	driver.quit()

def readPDF(pdf_file_path):
	with open(pdf_file_path, 'rb') as pdf_file:
	    pdf = PyPDF2.PdfReader(pdf_file)
	    text = ""
	    for page in pdf.pages:
	        text += page.extract_text()
	return text

def generateExcel():
	name = list()
	sex = list()
	age = list()
	scan = list()
	impression = list()
	hospno = list()
	date = list()
	biopsies = list()
	reportLink = list()
	n = 0
	for file in os.listdir():
		if file[-3:] == 'pdf':
			n+=1
			print('Doing ' + str(n))
			try:
				text = readPDF(file)
				valid = True
			except:
				print(file + ' not found')
				text = ''
				valid = False

			if valid:
				try:
					name.append(text.split('Patient Name:')[1].split('Sex:')[0].strip())
				except:
					name.append('ERROR')

				try:
					sex.append(text.split('Sex:')[1].split('Age:')[0].strip())
				except:
					sex.append('ERROR')

				try:
					age.append(text.split('Age:')[1].split('Patient ID:')[0].strip().replace(' ', ''))
				except:
					age.append('ERROR')

				try:
					if len(text.split('RADIOLOGY REPORT :')[1].split('\n')[0].strip())>3:
						scan.append(text.split('RADIOLOGY REPORT :')[0].split('\n')[1].strip())
					elif len(text.split('RADIOLOGY REPORT :')[1].split('\n')[1].strip())>3:
						scan.append(text.split('RADIOLOGY REPORT :')[1].split('\n')[1].strip())
					elif len(text.split('RADIOLOGY REPORT :')[1].split('\n')[2].strip())>3:
						scan.append(text.split('RADIOLOGY REPORT :')[1].split('\n')[2].strip())
					elif len(text.split('RADIOLOGY REPORT :')[1].split('\n')[3].strip())>3:
						scan.append(text.split('RADIOLOGY REPORT :')[1].split('\n')[3].strip())
					else:
						scan.append(text.split('RADIOLOGY REPORT :')[1].split('\n')[4].strip())
				except:
					scan.append('ERROR')

				try:
					hospno.append(text.split('Patient ID:')[1].split('Order')[0].replace(' ', ''))
				except:
					hospno.append('ERROR')

				try:
					datestring = text.split('CLINICAL DETAILS:')[0].split('Performed on')[1].strip()
					slashloc = datestring.find('/')
					date.append(datestring[slashloc-11:slashloc].strip())
				except:
					date.append('ERROR')

				try:
					impression.append(text.split('IMPRESSION :')[1].split('REPORTED BY:')[0].strip())

				except:
					impression.append('ERROR')
				try:
					reportLink.append("https://mahepacs.manipal.edu/get_pdf_report?accessionNo=NUMBER&amp;loadType=REPORT&amp;centerId=1".replace('NUMBER', str(file)[:-4]))
				except:
					reportLink.append('ERROR')

				try:
					document = requests.get("http://172.16.7.74/DISCHSUM/Dissumbio1.aspx?hp=" + str(text.split('Patient ID:')[1].split('Order')[0].replace(' ', '')) + "&adate=01/01/2023/")
					biopsyText = ""
					biopsyFound = False
					investText = document.text
					for investigation in investText.split('\n'):
						if 'Biopsy' in investigation:
							biopsy = investigation.split('<')[0]
							biopsyText += biopsy
							biopsyFound = True
					if biopsyFound:
						biopsies.append(biopsyText)
					else:
						biopsies.append('No Biopsy Report Found')
				except:
					biopsies.append('ERROR')

	findict = {'DATE' : date, 'HOSPITAL NUMBER' : hospno, 'NAME' : name, 'SEX' : sex, 'AGE': age, 'SCAN' : scan, 'IMPRESSION' : impression, 'HISTOPATHOLOGY' : biopsies, 'REPORT LINK' : reportLink}
	df = pd.DataFrame(findict)
	df.to_excel('Final.xlsx', index = False)

	n=0
	date = list()
	hospno = list()
	name = list()
	sex = list()
	age = list()
	scan = list()
	impressions = list()
	biopsies = list()
	reportLink = list()
	for i in range(len(df['IMPRESSION'])):
		imp = df['IMPRESSION'][i]
		buzzFound = False
		buzzwords = ['k/c/o ca', 'cancer', 'carcinoma', ' ca ', 'tumor', 'tumour', 'mets', 'metastasis', 'sarcoma', 'malignant', 'neoplastic', 'neoplasm', 'metastatic', 'benign', 'adenoma', 'angioma', 'myoma', 'CA-125', 'CA125', 'CA19-9', 'CA 19-9', ' CEA', 'AFP', 'biopsy', 'PET', 'stage', 'staging', 'HPE']
		for buzz in buzzwords:
			if buzzFound == False:
				if '\uf0b7' in str(imp):
					impLines = str(imp).split('\uf0b7')
				elif '\uf0d8' in str(imp):
					impLines = str(imp).split('\uf0d8')
				else:
					impLines = str(imp).split('. ')
				for impLine in impLines:
					if buzz.lower() in impLine.lower():
						buzzFound = True
						impLineDetect = impLine.replace('\n', '').replace('Page 2/2', '').replace('Page 3/3', '').strip()

		if buzzFound:
			name.append(df['DATE'][i])
			sex.append(df['SEX'][i])
			age.append(df['AGE'][i])
			scan.append(df['SCAN'][i])
			impressions.append(impLineDetect)
			date.append(df['DATE'][i])
			hospno.append(df['HOSPITAL NUMBER'][i])
			biopsies.append(df['HISTOPATHOLOGY'][i])
			reportLink.append(df['REPORT LINK'][i])

	findict = {'DATE' : date, 'HOSPITAL NUMBER' : hospno, 'NAME' : name, 'SEX' : sex, 'AGE': age, 'SCAN' : scan, 'IMPRESSION' : impressions, 'HISTOPATHOLOGY' : biopsies, 'REPORT LINK' : reportLink}
	df = pd.DataFrame(findict)
	df.to_excel('CompletedProject.xlsx', index = False)

def start():
	print('Welcome! Please select what you would like to do by selecting the option.')
	print('1. Get Accession Numbers.')
	print('2. Download Reports From Accession Numbers.')
	print('3. Generate Final Excel')

	dec = input('Input an option number to continue... ')
	if dec == '1':
		getAccessionNumbers()
	elif dec == '2':
		print('Please do not use the system while this part of the program is running.')
		df = pd.read_excel('Accession Numbers.xlsx')
		numbers = df['Accession Numbers']
		downloadReport(numbers)
	elif dec == '3':
		nil = input('Please ensure pdf files are present in the Project folder. They might need to be shifted from the Downloads folder. Once done, press enter on this screen to continue.')
		generateExcel()
	else:
		print('Incorrect input. Make sure the input is a number only.')
		start()

start()
