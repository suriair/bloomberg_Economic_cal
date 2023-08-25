import requests
from bs4 import BeautifulSoup
import openpyxl
import os
import pprint

res = requests.get('https://econcal.forexprostools.com/?features=datepicker,timezone,country_importance,filters&calType=week')

soup = BeautifulSoup(res.text, 'html.parser')

all_section = []
for item in soup.section.find_all('section'):
	all_section.append(item)

eco_cal = []

for section in all_section:
	date = section.h2.getText()
	
	for li in section.ul.find_all('li'):
		try:
			time = li.button.time.getText()
		except AttributeError:
			time = 'None'
		
		try:
			event = li.button.find('div', attrs={'class':'left event'}).getText()
		except AttributeError:
			event = 'None'

		try:
			flagCur = li.button.find('div', attrs={'class':'left flagCur'}).getText()
			flagCur = flagCur.splitlines()[1].lstrip()
		except AttributeError:
			flagCur = 'None'

		try:
			imp = li.button.get_attribute_list('aria-label')
			imp = imp[0].split('|')
			imp = imp[2]
		except AttributeError:
			imp = 'No importance found'

		try:
			act = li.button.find('div', attrs={'class':'act'}).getText()
			act = act.splitlines()[2]
			act = act.replace("\xa0"," ")
		
			fore = li.button.find('div', attrs={'class':'fore'}).getText()
			fore = fore.splitlines()[2]
			fore = fore.replace("\xa0"," ")	
		
			prev = li.button.find('div', attrs={'class':'prev'}).getText()
			prev = prev.splitlines()[2]
			prev = prev.replace("\xa0"," ")
		except:
			act = 'None'
			prev = 'None'
			fore = 'None'

		base_path = os.getcwd()
		file_path = os.path.join(base_path, 'eco_cal.xlsx')

		if not os.path.exists(file_path):
			file = openpyxl.Workbook()
			sheet = file.active
			columns = ['Time','Event','Currency','Importance','Actual','Forecast','Previous']
			sheet.append(columns)
			file.save(file_path)
		file = openpyxl.load_workbook(file_path)
		sheet = file.active
		sheet.append([time, event, flagCur, imp, act, fore, prev])
		file.save(file_path)


		eco_cal.append({'Time': time,
						'Event': event,
						'Currency': flagCur,
						'Importance': imp,
						'Actual':act,
						'Forecast': fore,
						'Previous': prev
						})


pprint.pprint(eco_cal)

