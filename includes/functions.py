from openpyxl import load_workbook, Workbook
from datetime import datetime
from pprint import pprint
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

month_names = [
	u'января',
	u'февраля',
	u'марта',
	u'апреля',
	u'мая',
	u'июня',
	u'июля',
	u'августа',
	u'сентября',
	u'октября',
	u'ноября',
	u'декабря'
	]

def get_data_xls(filename):

	wb = load_workbook(filename=filename)
	ws = wb.active
	entries = []

	for row in ws.rows:
		entry = []
		for cell in row:
			entry.append(cell.value)
		# if not all(entry):
		# 	pprint(u'Запись %r заполнена не полностью!' % (''.join([str(cell.value)+' ' for cell in row])))
		entries.append(entry)

	# check if last rows is None
	for entry in reversed(entries):
		if any(entry):
			break
		else:
			del entries[-1]

	# check if last columns is None
	for _ in entries[0]:
		last_column = [_ for entry in entries if entry[-1]]
		if last_column:
			break
		else:
			for index,_ in enumerate(entries):
				del entries[index][-1]

	return entries

def save_xls_data(filename, entries):

	wb = Workbook()
	ws = wb.active
	ws.title = u'База'

	for row_index,entry in enumerate(entries):
		for col_index,value in enumerate(entry):
			ws.cell(row=row_index+1,column=col_index+1).value = value

	wb.save(filename)

def get_doc_data(filename):

	document = Document(filename)
	lines = []
	for paragraph in document.paragraphs:
		line = paragraph.text
		if line != '':
			lines.append(line)

	for table in document.tables:
		for row in table.rows:
			for cell in row.cells:
				for paragraph in cell.paragraphs:
					line = paragraph.text
					if line != '':
						lines.append(line)

	return lines

def create_new_replaced_doc(fin,fout,replacements):

	document = Document(fin)

	for paragraph in document.paragraphs:
		for old_data,new_data in replacements.items():
			if old_data in paragraph.text:
				# Loop added to work with runs (strings with same style)
				for run in paragraph.runs:
					if old_data in run.text:
						run.text = run.text.replace(old_data, new_data)

	for table in document.tables:
		for row in table.rows:
			for cell in row.cells:
				for paragraph in cell.paragraphs:
					for old_data,new_data in replacements.items():
						if old_data in paragraph.text:
							# Loop added to work with runs (strings with same style)
							for run in paragraph.runs:
								if old_data in run.text:
									run.text = run.text.replace(old_data, new_data)

	document.save(fout)
