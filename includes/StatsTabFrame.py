# -*- coding: utf-8 -*-
from tkinter.filedialog import *
from tkinter.messagebox import *
from tkinter.font import Font
from tkinter import *
import tkinter.ttk as ttk
from .functions import *
from collections import Counter,OrderedDict
from pprint import pprint
import configparser

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
# from matplotlib import style
# style.use("ggplot")

class StatsTabFrame(Frame):

	def __init__(self, mainWindow,title_text):

		Frame.__init__(self)

		self.mainWindow = mainWindow
		self.small_font = self.mainWindow.small_font
		self.default_font = self.mainWindow.default_font
		self.big_font = self.mainWindow.big_font

		self.statsList = []
		# num of entries which will be displayed on figure
		self.topStats = IntVar()
		# num of columns in statistci table
		self.num_of_columns = IntVar()

		# it displays when files didn't selected
		self.stats_frame_plug = Label(
			self,
			text=u'Место для статистики.',
			font=self.big_font
			)

		self.title = Label(
			self,
			text=title_text,
			font=self.big_font,
			pady=15
			)

		self.optionsFrame = Frame(self,padx=50,pady=10)
		self.columnChoiceLabel = Label(
			self.optionsFrame,
			text=u'Столбец: ',
			font=self.default_font
			)
		self.colIndex = IntVar()
		self.colIndex.set(1)
		self.columnChoiceMenu = OptionMenu(
			self.optionsFrame,
			self.colIndex,
			range(1,2)
			)

		self.statsFrame = Frame(self,padx=50,pady=10)
		self.statsTableFrame = Frame(self.statsFrame)
		
		ttk.Style().configure('Treeview',rowheight=25)						# set row height

	def pack_all(self):

		'''Pack all elements in a tab frame'''
		self.stats_frame_plug.pack_forget()
		self.title.pack_forget()
		self.optionsFrame.pack_forget()
		self.columnChoiceLabel.pack_forget()
		self.columnChoiceMenu.pack_forget()
		self.statsFrame.pack_forget()
		self.statsTableFrame.pack_forget()
		self.statsTable.pack_forget()
		self.statsTableScroll.pack_forget()
		self.figureCanvas.get_tk_widget().pack_forget()
		self.figureCanvas._tkcanvas.pack_forget()

		if self.mainWindow.base_file.get():
			self.title.pack(side=TOP,fill=X)
			self.optionsFrame.pack(side=TOP,fill=X)
			self.columnChoiceLabel.pack(side=LEFT,anchor="nw")
			self.columnChoiceMenu.pack(side=LEFT,anchor="nw")
			self.statsFrame.pack(side=TOP,fill=X)
			self.statsTableFrame.pack(side=LEFT)
			self.statsTable.pack(side=LEFT,anchor="nw")
			self.statsTableScroll.pack(side=LEFT,fill=Y)
			# self.figureCanvas.get_tk_widget().pack(side=RIGHT)
			self.figureCanvas._tkcanvas.pack(side=RIGHT)
		else:
			self.stats_frame_plug.pack(side=TOP,fill=BOTH,expand=True)

	def bind_all(self):

		'''Bind all events in a tab frame'''
		self.colIndex.trace('w',self.changeColIndex)

	def readOptions(self):

		'''Read options from .ini file'''
		configs = configparser.ConfigParser()
		configs.read(self.mainWindow.configsFileName)
		self.colIndex.set(int(configs['Statistic']['col_index']))
		self.topStats.set(int(configs['Statistic']['top_stats']))

	def saveOptions(self):

		configs = configparser.ConfigParser()
		configs.read(self.mainWindow.configsFileName)

		configs['Statistic']['col_index'] = str(self.colIndex.get())
		with open(self.mainWindow.configsFileName, 'w') as configfile:
			configs.write(configfile)
			configfile.close()

	def changeColIndex(self,*args):

		self.saveOptions()
		col_index = self.colIndex.get()-1
		data_list = self.mainWindow.base_frame.getColumnAsList(col_index=col_index)
		heading_text = self.mainWindow.base_frame.getHeadingText(col_index=col_index)
		self.getStatsFromList(data_list=data_list)

		headings = (
			'№',
			heading_text,
			'Всего',
			'Процент'
			)

		self.fillTable(headings)

		self.fillFigure()

	def changeColumnChoices(self,amountOfColumns):

		self.columnChoiceMenu['menu'].delete(0, 'end')
		for index in range(1,amountOfColumns+1):
			self.columnChoiceMenu['menu'].add_command(
				label=index,
				command=lambda index=index: self.colIndex.set(index)
				)

	def getStatsFromList(self,data_list):

		'''Get statistic from base frame'''
		data_counter = Counter(data_list)
		stats_list = []
		for name,total in data_counter.most_common():
			percentage = round(total/len(data_list) * 100.0, 2)
			stats_list.append(OrderedDict(name=name,total=total,percentage=percentage))

		self.statsList = stats_list
		self.num_of_columns.set(len(self.statsList[0])+1)

	def destroyTable(self):

		try:
			self.statsTable.destroy()
			self.statsTableScroll.destroy()
		except AttributeError:
			pass

	def initTable(self):

		'''Initialize statistic table'''
		self.statsTable = ttk.Treeview(
			self.statsTableFrame,
			show = 'headings',
			selectmode='none'
			)
		# create scrollbar for table
		self.statsTableScroll = Scrollbar(
			self.statsTableFrame,
			orient=VERTICAL,
			command=self.statsTable.yview
			)
		self.statsTable.configure(yscrollcommand=self.statsTableScroll.set)	# attach scrollbar

		self.statsTable['columns'] = ['']*self.num_of_columns.get()
		self.statsTable.column(0,width=50)
		self.statsTable.column(1,width=250)
		self.statsTable.column(2,width=70)
		self.statsTable.column(3,width=70)

	def fillTable(self,headings):

		'''Fill table with statistic'''
		self.statsTable.delete(*self.statsTable.get_children())

		for col_index, heading in enumerate(headings):
			self.statsTable.heading(col_index,text=heading)
		
		for index, stats_entry in enumerate(self.statsList):
			values=['']*self.num_of_columns.get()
			values[0] = index+1
			values[1] = stats_entry['name']
			values[2] = stats_entry['total']
			values[3] = stats_entry['percentage']
			self.statsTable.insert('', 'end', values=values, tags=[])

	def destroyFigure(self):

		try:
			self.figureCanvas.get_tk_widget().destroy()
		except AttributeError:
			pass

	def initFigure(self):

		if not self.topStats.get():
			return
		if not self.statsList:
			return

		f = Figure(figsize=(9,5), dpi=60)
		self.pie = f.add_subplot(111)

		self.figureCanvas = FigureCanvasTkAgg(f, self.statsFrame)

	def fillFigure(self):

		self.pie.clear()

		data = []
		labels = []

		top = self.topStats.get()
		# if more then 5 entries - put last entries to the 'others'
		if len(self.statsList)<=top:
			for item in self.statsList:
				data.append(item['total'])
				labels.append(item['name'])

			explode = [0]*len(self.statsList)
		else:
			for item in self.statsList[:top]:
				data.append(item['total'])
				labels.append(item['name'])
			data.append(0)
			for item in self.statsList[top:]:
				data[top] += item['total']
			labels.append(u'Другие')
			explode = [0]*(top+1)
		
		explode[0] = 0.1

		self.pie.pie(
			data, 
			explode=explode,
			labels=labels,
			autopct='%1.1f%%', 
			shadow=True)

		self.figureCanvas.draw()
