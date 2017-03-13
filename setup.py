from distutils.core import setup
import py2exe, sys, os
from glob import glob

sys.argv.append('py2exe')

datafiles = [('dlls', glob(r'bin_64\*.*')), ('includes',glob(r'includes\*'))]

setup(name="Doc maker",
	version="1.0",
	windows=['doc_maker.py'],
	data_files=datafiles,
	options={"py2exe": {"includes": ["openpyxl","docx",'lxml.etree','lxml._elementpath','gzip','tkinter','tkinter.ttk']}})