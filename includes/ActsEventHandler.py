# -*- coding: utf-8 -*-
from watchdog.events import PatternMatchingEventHandler
from .functions import *

class ActsEventHandler(PatternMatchingEventHandler):

	'''Looking for FileSystem events for base file and do something'''
	def __init__(self, patterns, sync_func):

		super().__init__(patterns)
		self.sync_func = sync_func

	def on_created(self,event):

		self.sync_func()

	def on_modified(self,event):

		self.sync_func()

	def on_deleted(self,event):

		self.sync_func()

