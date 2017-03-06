# -*- coding: utf-8 -*-
import os, sys
import re
from includes.functions import *
from includes.DocMaker import DocMaker
from pprint import pprint
from tkinter import *

if __name__=="__main__":

	if getattr(sys, 'frozen', False):
		# frozen
		root_dir = os.path.dirname(sys.executable)
	else:
		# unfrozen
		root_dir = os.path.dirname(os.path.realpath(__file__))

	root = Tk()
	root.geometry('1200x720+50+10')
	root.title(u'Составитель актов 2000')
	root.update()

	maker = DocMaker(root,root_dir)

	root.mainloop()