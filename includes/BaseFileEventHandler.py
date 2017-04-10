# -*- coding: utf-8 -*-
from watchdog.events import PatternMatchingEventHandler
from .functions import *

class BaseFileEventHandler(PatternMatchingEventHandler):

	'''Looking for FileSystem events for base file and do something'''
	def __init__(self, patterns, mainWindow):

		super().__init__(patterns)
		self.mainWindow = mainWindow

	def on_modified(self,event):

		'''Load base file, compare with local base and load it to table if they different'''
		try:
			from_base = get_data_xls(self.mainWindow.base_file.get())
			from_table = self.mainWindow.base_frame.getAllEntriesAsList()
			if from_base[1:] != from_table:
				self.mainWindow.base_frame.loadBase()
				self.mainWindow.status_bar['text'] = u'База обновлена'
		except:
			pass