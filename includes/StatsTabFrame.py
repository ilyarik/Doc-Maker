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

class StatsTabFrame(Frame):

	def __init__(self, mainWindow):

		Frame.__init__(self)

		self.mainWindow = mainWindow
		self.small_font = self.mainWindow.small_font
		self.default_font = self.mainWindow.default_font
		self.big_font = self.mainWindow.big_font

		self.statsList = []
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
			text='Здесь отображается статистика по выбранному столбцу',
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

		self.statsTableFrame = Frame(self,padx=50,pady=10)
		self.statsTable = ttk.Treeview(
			self.statsTableFrame,
			show = 'headings'
			)
		# create scrollbar for table
		self.statsTableScroll = Scrollbar(
			self.statsTableFrame,
			orient=VERTICAL,
			command=self.statsTable.yview
			)
		self.statsTable.configure(yscrollcommand=self.statsTableScroll.set)	# attach scrollbar
		ttk.Style().configure('Treeview',rowheight=25)						# set row height

	def pack_all(self):

		'''Pack all elements in a tab frame'''
		self.stats_frame_plug.pack_forget()
		self.title.pack_forget()
		self.optionsFrame.pack_forget()
		self.columnChoiceLabel.pack_forget()
		self.columnChoiceMenu.pack_forget()
		self.statsTableFrame.pack_forget()
		self.statsTable.pack_forget()
		self.statsTableScroll.pack_forget()

		if self.mainWindow.base_file.get():
			self.title.pack(side=TOP,fill=X)
			self.optionsFrame.pack(side=TOP,fill=X)
			self.columnChoiceLabel.pack(side=LEFT,anchor="nw")
			self.columnChoiceMenu.pack(side=LEFT,anchor="nw")
			self.statsTableFrame.pack(side=TOP,fill=X)
			self.statsTable.pack(side=LEFT,anchor="nw")
			self.statsTableScroll.pack(side=LEFT,fill=Y)
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

	def saveOptions(self):

		configs = configparser.ConfigParser()
		configs.read(self.mainWindow.configsFileName)

		configs['Statistic']['col_index'] = str(self.colIndex.get())
		with open(self.mainWindow.configsFileName, 'w') as configfile:
			configs.write(configfile)
			configfile.close()

	def changeColIndex(self,*args):

		self.saveOptions()
		data_list = self.mainWindow.base_frame.getColumnAsList(col_index=self.colIndex.get()-1)
		self.getStatsFromList(data_list=data_list)
		self.fillTable()
		self.changeColumnChoices(self.mainWindow.base_frame.num_of_fields)

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

	def initTable(self):

		'''Initialize statistic table'''
		self.statsTable['columns'] = ['']*self.num_of_columns.get()
		self.statsTable.column(0,width=50)
		self.statsTable.column(1,width=250)
		self.statsTable.column(2,width=70)
		self.statsTable.column(3,width=70)

	def fillTable(self):

		'''Fill table with statistic'''
		self.statsTable.delete(*self.statsTable.get_children())

		self.statsTable.heading(0,text='№')
		self.statsTable.heading(1,text='ФИО сотрудника')
		self.statsTable.heading(2,text='Всего')
		self.statsTable.heading(3,text='Процент')
		
		for index, stats_entry in enumerate(self.statsList):
			values=['']*self.num_of_columns.get()
			values[0] = index+1
			values[1] = stats_entry['name']
			values[2] = stats_entry['total']
			values[3] = stats_entry['percentage']
			self.statsTable.insert('', 'end', values=values, tags=[])