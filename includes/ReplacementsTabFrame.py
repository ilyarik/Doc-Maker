# -*- coding: utf-8 -*-
from tkinter.filedialog import *
from tkinter.messagebox import *
from tkinter.font import Font
from tkinter import *
import tkinter.ttk as ttk
from .functions import *
import configparser

class ReplacementsTabFrame(Frame):

	def __init__(self, mainWindow, act_name, act_var, plug_text):

		Frame.__init__(self)

		self.mainWindow = mainWindow
		self.act_name = act_name
		self.act_var = act_var
		self.replacements = {}

		self.frame_plug = Label(
			self,
			text=plug_text,
			font=self.mainWindow.big_font
			)

		# create text field for display primary text from doc file
		self.plain_text_frame = Frame(self, padx=10,pady=10)
		self.plain_text_label = Label(
			self.plain_text_frame,
			text=u'Начальный текст',
			font=self.mainWindow.default_font
			)
		self.plain_text = Text(
			self.plain_text_frame,
			width=45,
			font=self.mainWindow.default_font,
			height=30
			)
		self.plain_text.tag_config('replaced',background='yellow')

		# create text field for display text with replaced values
		self.result_text_frame = Frame(self,padx=10,pady=10)
		self.result_text_label = Label(
			self.result_text_frame,
			text=u'Результирующий текст',
			font=self.mainWindow.default_font)
		self.result_text = Text(
			self.result_text_frame,
			width=45,
			font=self.mainWindow.default_font,
			height=30
			)
		self.result_text.tag_config('replaced',background='yellow')

		# create frame for replacements
		# each replacement contain one entry for primary value,
		# one label and one combobox for new value
		# separators split replacements in frame
		self.primary_values_for_replacement = []
		self.labels_for_replacement = []
		self.new_values_for_replacement = []
		self.replacements_separators = []
		self.num_of_replacements = 0
		# create canvas for scrollable frame
		self.replacement_canvas = Canvas(
			self,
			borderwidth=0,
			width=290,
			height=500)
		self.replacement_frame = Frame(self.replacement_canvas,padx=10,pady=10)
		self.replacement_frameScroll = Scrollbar(
			self,
			orient=VERTICAL,
			command=self.replacement_canvas.yview)
		self.add_replacement_button = Button(
			self.replacement_frame,
			text=u'Добавить замену',
			font=self.mainWindow.small_font,
			width=22,
			height=1
			)
		# some magic for create scrollable frame with replacement fields
		self.replacement_canvas.configure(yscrollcommand=self.replacement_frameScroll.set)
		self.replacement_canvas.create_window((4,4), window=self.replacement_frame, anchor=CENTER, 
								  tags="self.replacement_frame")

	def pack_all(self):

		'''Pack all elements in a tab frame'''
		self.frame_plug.pack_forget()
		self.plain_text_frame.pack_forget()
		self.plain_text.pack_forget()
		self.result_text_frame.pack_forget()
		self.result_text.pack_forget()
		self.replacement_canvas.pack_forget()
		self.replacement_frameScroll.pack_forget()
		for index in range(self.num_of_replacements):
			self.primary_values_for_replacement[index].grid_forget()
			self.labels_for_replacement[index].grid_forget()
			self.new_values_for_replacement[index].grid_forget()
		for replacement_separator in self.replacements_separators:
			replacement_separator.grid_forget()
		self.add_replacement_button.grid_forget()

		if self.act_var.get():
			self.plain_text_frame.pack(side=LEFT,fill=Y)
			self.plain_text_label.pack(side=TOP)
			self.plain_text.pack(side=LEFT,fill=Y)
			self.result_text_frame.pack(side=RIGHT,fill=Y)
			self.result_text_label.pack(side=TOP)
			self.result_text.pack(side=RIGHT,fill=Y)
			self.replacement_frameScroll.pack(side=RIGHT, fill=Y)
			self.replacement_canvas.pack(side=LEFT,fill=BOTH,expand=True)
			for index in range(self.num_of_replacements):
				self.primary_values_for_replacement[index].grid(row=index*4,column=0,pady=2,sticky=N)
				self.labels_for_replacement[index].grid(row=index*4+1,column=0,pady=2,sticky=N)
				self.new_values_for_replacement[index].grid(row=index*4+2,column=0,pady=2,sticky=N)
			for index, replacement_separator in enumerate(self.replacements_separators):
				replacement_separator.grid(row=index*4+3,column=0,padx=10,pady=5,sticky=W+E)
			self.add_replacement_button.grid(row=self.num_of_replacements*5,column=0,pady=10,sticky=N)
		else:
			self.frame_plug.pack(side=TOP,fill=BOTH,expand=True)

	def bind_all(self):

		'''Bind all events in a tab frame'''
		if self.act_var.get():
			self.add_replacement_button.bind('<Button-1>',self.add_replacement)
			self.replacement_frame.bind("<Configure>", self.replaceFrameConfigure)		# bind mouse scroll
			[primary_val.bind("<Return>",self.replace) for primary_val in self.primary_values_for_replacement]
			[new_val.bind("<Return>",self.replace) for new_val in self.new_values_for_replacement]

	def load_act(self,event=None):

		if not self.act_var.get():
			return

		try:
			doc_text = get_doc_data(self.act_var.get())
		except Exception as e:
			showerror(u'Ошибка!',e)
			self.act_var.set(u'')
			return

		if not doc_text:
			showerror(u'Ошибка!',u'Пустой .doc файл.')
			self.act_var.set(u'')
			return

		# clear text fields and fill it
		self.plain_text.delete(1.0,END)
		self.result_text.delete(1.0,END)
		for line in doc_text:
			self.plain_text.insert(END,line+'\n')
		self.replace()			# fill result text field and add tags

		# delete replacements if exists and init new replacements
		self.read_replacements()
		self.destroy_replacements()
		for _ in range(len(self.replacements)):
			self.add_replacement()
		self.set_replacements_entries()
		self.replace()

		self.mainWindow.status_bar['text'] = u'Документ загружен'

		self.pack_all()
		self.bind_all()

	def get_replace_variants(self):

		'''Add fields from base for replace in choicefields'''
		for combobox in self.new_values_for_replacement:
			replace_values = []
			if self.mainWindow.base_file.get():
				replace_values.extend([u'*Столбец %r*' % (index+1) for index in range(self.mainWindow.base_frame.num_of_fields)])
			replace_values.append(u'…')
			combobox['values'] = replace_values

	def read_replacements(self):

		'''Read replacements dict from .ini file'''

		configs = configparser.ConfigParser()
		configs.read(u'%s\\USER\\configs.ini' % (self.mainWindow.root_dir))
		
		if self.act_name == 'aot':
			self.replacements = eval(configs['Act_of_transfer']['replacements'])
		elif self.act_name == 'ra':
			self.replacements = eval(configs['Return_act']['replacements'])
		elif self.act_name == 'aoe':
			self.replacements = eval(configs['Act_of_elimination']['replacements'])

		'''Set some replacements customly'''
		if 'ЗАМЕНАДАТА' in self.replacements.keys():
			self.replacements['ЗАМЕНАДАТА'] = get_current_date_russian()
		if 'ЗАМЕНАДАТА1' in self.replacements.keys():
			self.replacements['ЗАМЕНАДАТА1'] = get_current_date_russian()
		if 'ЗАМЕНАДАТА2' in self.replacements.keys():
			self.replacements['ЗАМЕНАДАТА2'] = get_current_date_russian()
	def destroy_replacements(self):

		for primary_val in self.primary_values_for_replacement:
			primary_val.destroy()
		for label in self.labels_for_replacement:
			label.destroy()
		for new_val in self.new_values_for_replacement:
			new_val.destroy()
		for separator in self.replacements_separators:
			separator.destroy()

		self.primary_values_for_replacement = []
		self.labels_for_replacement = []
		self.new_values_for_replacement = []
		self.replacements_separators = []
		self.num_of_replacements = 0

	def add_replacement(self,event=None):

		self.primary_values_for_replacement.append(
			Entry(
				self.replacement_frame,
				width=30,
				font=self.mainWindow.default_font
				)
			)
		self.labels_for_replacement.append(
			Label(
				self.replacement_frame,
				text=u'Заменить на:',
				font=self.mainWindow.small_font
				)
			)
		self.new_values_for_replacement.append(
			ttk.Combobox(
				self.replacement_frame,
				width=35,
				font=self.mainWindow.small_font,
				values=[]
				)
			)
		if self.num_of_replacements>0:
			self.replacements_separators.append(
				ttk.Separator(
					self.replacement_frame,
					orient=HORIZONTAL
					)
				)

		self.num_of_replacements += 1
		self.get_replace_variants()

		self.pack_all()
		self.bind_all()

	def set_replacements_entries(self):

		'''Set entries'''
		for index,(old_val,new_val) in enumerate(self.replacements.items()):
			self.primary_values_for_replacement[index].delete(0,END)
			self.primary_values_for_replacement[index].insert(0,old_val)
			self.new_values_for_replacement[index].set(new_val)

	def replaceFrameConfigure(self,event=None):

		'''Reset the scroll region to encompass the inner frame'''

		self.replacement_canvas.configure(scrollregion=self.replacement_canvas.bbox("all"))

	def get_result_value(self,index,index_row=0):

		num_of_column = self.new_values_for_replacement[index].current()
		if num_of_column == -1 or num_of_column == self.mainWindow.base_frame.num_of_fields:
			result_value = self.new_values_for_replacement[index].get()
		elif self.mainWindow.base_file and self.mainWindow.base_frame.base_table.selection():
			result_value = self.mainWindow.base_frame.base_table.item(self.mainWindow.base_frame.base_table.selection()[index_row])['values'][num_of_column]
		elif self.mainWindow.base_file:
			result_value = self.mainWindow.base_frame.base_table.item(self.mainWindow.base_frame.base_table.get_children()[0])['values'][num_of_column]
		else:
			showerror(u'Ошибка!', u'Не удается получить данные для замены.')
			return
		return result_value

	def replace(self,event=None):

		# remove tags
		for tag in self.plain_text.tag_names():
			self.plain_text.tag_remove(tag,1.0,END)
		for tag in self.result_text.tag_names():
			self.result_text.tag_remove(tag,1.0,END)
		# copy text from plain to result
		text = self.plain_text.get(1.0,END)
		self.result_text.delete(1.0,END)
		self.result_text.insert(1.0,text)
		plain_lines = text.split('\n')
		for index_line,line in enumerate(plain_lines):

			if not line:
				continue

			start_of_line = float(index_line+1)
			end_of_line = '%r.end' % (index_line+1)

			for index in range(self.num_of_replacements):
				plain = str(self.primary_values_for_replacement[index].get())
				if not plain:
					continue
				if plain not in line:
					continue
				result_value = str(self.get_result_value(index))
				line = line.replace(plain,result_value)

				# rewrite line on new replaced line
				self.result_text.delete(start_of_line,end_of_line)
				self.result_text.insert(start_of_line,line)

				# add tags to plain text
				search_start = 1.0
				while True:
					start_of_tag = self.plain_text.search(plain,search_start,stopindex=end_of_line)
					if not start_of_tag:
						break
					end_of_tag = start_of_tag.split('.')[0]+'.'+str(int(start_of_tag.split('.')[1])+len(plain))
					self.plain_text.tag_add('replaced',start_of_tag,end_of_tag)
					search_start = end_of_tag

				# add tags to result text
				if not result_value:
					continue
				search_start = 1.0
				while True:
					start_of_tag = self.result_text.search(result_value,search_start,stopindex=end_of_line)
					if not start_of_tag:
						break
					end_of_tag = start_of_tag.split('.')[0]+'.'+str(int(start_of_tag.split('.')[1])+len(result_value))
					self.result_text.tag_add('replaced',start_of_tag,end_of_tag)
					search_start = end_of_tag
