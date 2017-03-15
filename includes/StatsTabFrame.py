# -*- coding: utf-8 -*-
from tkinter.filedialog import *
from tkinter.messagebox import *
from tkinter.font import Font
from tkinter import *
import tkinter.ttk as ttk
from .functions import *
from collections import Counter,OrderedDict
from pprint import pprint

class StatsTabFrame(Frame):

	def __init__(self, mainWindow):

		Frame.__init__(self)

		self.mainWindow = mainWindow
		self.small_font = self.mainWindow.small_font
		self.default_font = self.mainWindow.default_font
		self.big_font = self.mainWindow.big_font

		self.statsList = []

		# it displays when files didn't selected
		self.base_frame_plug = Label(
			self,
			text=u'Место для статистики.',
			font=self.big_font
			)

		self.statsTableFrame = Frame(self,padx=10,pady=10)
		self.statsTable = ttk.Treeview(
			self.statsTableFrame,
			show = 'headings'
			)
		# create scrollbar for table
		self.statsTableScroll = Scrollbar(
			self.statsTableFrame,
			orient=VERTICAL,
			command=self.statsTable.yview)
		self.statsTable.configure(yscrollcommand=self.statsTableScroll.set)	# attach scrollbar
		ttk.Style().configure('Treeview',rowheight=25)						# set row height

	def pack_all(self):

		'''Pack all elements in a tab frame'''
		self.statsTableFrame.pack(side=TOP)
		self.statsTable.pack(side=LEFT)
		self.statsTableScroll.pack(side=RIGHT,fill=Y)

	def bind_all(self):

		'''Bind all events in a tab frame'''
		pass

	def getStatsFromList(self,data_list,amount=3):

		'''Get statistic from base frame'''
		data_counter = Counter(data_list)
		stats_list = []
		for name,total in data_counter.most_common(amount):
			percentage = round(total/len(data_list) * 100.0, 2)
			stats_list.append(OrderedDict(name=name,total=total,percentage=percentage))

		self.statsList = stats_list
		pprint(self.statsList)

	def fillTable(self):

		'''Fill table with statistic'''
		self.statsTable.delete(*self.statsTable.get_children())
		num_of_columns = (len(self.statsList[0])+1)
		self.statsTable['columns'] = ['']*num_of_columns
		self.statsTable.column(0,width=100)
		self.statsTable.column(1,width=300)
		self.statsTable.column(2,width=100)
		self.statsTable.column(3,width=100)
		self.statsTable.heading(0,text='№')
		self.statsTable.heading(1,text='ФИО сотрудника')
		self.statsTable.heading(2,text='Всего')
		self.statsTable.heading(3,text='Процент')
		self.statsTable.update()
		for index, stats_entry in enumerate(self.statsList):
			print(index)
			print(stats_entry)
			values=[u'']*num_of_columns
			values[0] = index
			values[1] = stats_entry['name']
			values[2] = stats_entry['total']
			values[3] = stats_entry['percentage']
			self.statsTable.insert('', 'end', values=values, tags=[])