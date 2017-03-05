# -*- coding: utf-8 -*-
import os, sys
import re
import configparser
from includes.functions import *
from includes.DocMaker import DocMaker
from pprint import pprint
from tkinter import *

if __name__=="__main__":

	if getattr(sys, 'frozen', False):
		# frozen
		dir_ = os.path.dirname(sys.executable)
	else:
		# unfrozen
		dir_ = os.path.dirname(os.path.realpath(__file__))

	# read configs
	configs = configparser.ConfigParser()
	configs.read(u'%s\\USER\\configs.ini' % (dir_))

	root = Tk()
	root.geometry('1200x720')
	root.title(u'Составитель актов 2000')

	maker = DocMaker(root)

	root.mainloop()