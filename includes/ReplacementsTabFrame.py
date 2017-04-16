# -*- coding: utf-8 -*-
from tkinter.filedialog import *
from tkinter.messagebox import *
from tkinter.font import Font
from tkinter import *
import tkinter.ttk as ttk
from .functions import *
import configparser

class ReplacementsTabFrame(Frame):

	def __init__(self, mainWindow, act_name, act_var, plug_text, title_text):

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

		self.title = Label(
			self,
			text=title_text,
			font=self.mainWindow.big_font,
			pady=15
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
			state=DISABLED,
			width=37,
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
			state=DISABLED,
			width=37,
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
			width=440,
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
		self.title.pack_forget()
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
			self.title.pack(side=TOP,fill=X)
			self.plain_text_frame.pack(side=LEFT,fill=Y)
			self.plain_text_label.pack(side=TOP)
			self.plain_text.pack(side=LEFT,fill=Y)
			self.result_text_frame.pack(side=RIGHT,fill=Y)
			self.result_text_label.pack(side=TOP)
			self.result_text.pack(side=RIGHT,fill=Y)
			
			self.replacement_canvas.pack(side=LEFT,fill=BOTH)
			for index in range(self.num_of_replacements):
				self.primary_values_for_replacement[index].grid(row=index*2,column=0,pady=2,sticky=W)
				self.labels_for_replacement[index].grid(row=index*2,column=1,pady=2,sticky=W)
				self.new_values_for_replacement[index].grid(row=index*2,column=2,pady=2,sticky=E)
				# add separator for all replacements except last
				if index != self.num_of_replacements-1:
					self.replacements_separators[index].grid(row=index*2+1,column=0,columnspan=3,padx=10,pady=5,sticky=W+E)
			self.add_replacement_button.grid(row=self.num_of_replacements*5,column=0,columnspan=3,pady=10,sticky=N)
			self.replacement_frameScroll.pack(side=LEFT, fill=Y)
		else:
			self.frame_plug.pack(side=TOP,fill=BOTH,expand=True)

	def bind_all(self):

		'''Bind all events in a tab frame'''
		if self.act_var.get():
			self.plain_text.bind("<Key>",lambda event=None,text=self.plain_text:self.text_clipboard(event,text))
			self.result_text.bind("<Key>",lambda event=None,text=self.plain_text:self.text_clipboard(event,text))
			self.add_replacement_button.bind('<Button-1>',self.add_replacement)
			self.replacement_frame.bind("<Configure>", self.replaceFrameConfigure)		# bind mouse scroll
			[primary_val.bind("<Return>",self.replace) for primary_val in self.primary_values_for_replacement]
			[primary_val.bind("<Key>",lambda event=None,entry=primary_val:self.entry_clipboard(event,entry)) for primary_val in self.primary_values_for_replacement]
			[new_val.bind("<Return>",self.replace) for new_val in self.new_values_for_replacement]
			[new_val.bind("<Key>",lambda event=None,entry=new_val:self.entry_clipboard(event,entry)) for new_val in self.new_values_for_replacement]

	def text_clipboard(self,event=None,text=0):

		'''Make clipboard great again for russian symbols'''
		if not text:
			return
		try:
			char = event.char.encode('cp1251')
			sym = event.keysym

			# ctrl+c
			if char==b'\x03' and sym=='ntilde':
				try:
					selected = text.selection_get()
				except:
					return
				if not selected:
					return
				self.mainWindow.clipboard_clear()
				self.mainWindow.clipboard_append(selected)
		
		except:
			return

	def entry_clipboard(self,event=None,entry=0):

		'''Make clipboard great again for russian symbols'''
		if not entry:
			return
		try:
			char = event.char.encode('cp1251')
			sym = event.keysym

			# ctrl+c
			if char==b'\x03' and sym=='ntilde':
				try:
					selected = entry.selection_get()
				except:
					return
				if not selected:
					return
				self.mainWindow.clipboard_clear()
				self.mainWindow.clipboard_append(selected)
			# ctrl+x
			elif char==b'\x18' and sym=='division':
				try:
					selected = entry.selection_get()
				except:
					return
				if not selected:
					return
				self.mainWindow.clipboard_clear()
				self.mainWindow.clipboard_append(selected)
				try:
					first = entry.index("sel.first")
					last = entry.index("sel.last")
					entry.delete(first,last)
				except Exception as e:
					first = entry.index(INSERT)
					last = first+len(clip_val)
				entry.delete(first,last)
			# ctrl+v
			elif char==b'\x16' and sym=='igrave':
				# cursor position in entry
				try:
					clip_val = self.mainWindow.selection_get(selection = "CLIPBOARD")
				except:
					return
				if not clip_val:
					return

				try:
					first = entry.index("sel.first")
					last = entry.index("sel.last")
					entry.delete(first,last)
				except Exception as e:
					first = entry.index(INSERT)
					last = first+len(clip_val)
				entry.insert(first,clip_val)
			# ctrl+a
			elif char ==b'\x01' and sym=='ocircumflex':
				entry.select_range(0,END)
		except:
			return

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

		# make text fields editable
		self.plain_text.config(state=NORMAL)
		self.result_text.config(state=NORMAL)

		self.plain_text.delete(1.0,END)
		self.result_text.delete(1.0,END)
		for line in doc_text:
			self.plain_text.insert(END,line+'\n')

		# make text fields not editable
		self.plain_text.config(state=DISABLED)
		self.result_text.config(state=DISABLED)

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
		configs.read(self.mainWindow.configsFileName)
		
		if self.act_name == 'aot':
			self.replacements = eval(configs['Act_of_transfer']['replacements'])
		elif self.act_name == 'ra':
			self.replacements = eval(configs['Return_act']['replacements'])
		elif self.act_name == 'aoe':
			self.replacements = eval(configs['Act_of_elimination']['replacements'])

		'''Set some replacements customly'''
		for old_val, new_val in self.replacements.items():
			if 'ЗАМЕНАДАТА' in old_val:
				self.replacements[old_val] = get_current_date_russian()

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
				width=20,
				font=self.mainWindow.default_font
				)
			)
		self.labels_for_replacement.append(
			Label(
				self.replacement_frame,
				text=u'заменить:',
				font=self.mainWindow.small_font
				)
			)
		self.new_values_for_replacement.append(
			ttk.Combobox(
				self.replacement_frame,
				width=20,
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

		if not self.mainWindow.base_file.get():
			return

		# remove tags
		for tag in self.plain_text.tag_names():
			self.plain_text.tag_remove(tag,1.0,END)
		for tag in self.result_text.tag_names():
			self.result_text.tag_remove(tag,1.0,END)
		# copy text from plain to result
		text = self.plain_text.get(1.0,END)

		# make text field editable
		self.result_text.config(state=NORMAL)

		self.result_text.delete(1.0,END)
		self.result_text.insert(1.0,text)

		# make text field not editable
		self.result_text.config(state=DISABLED)

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
				# process specific replacement for initials
				if 'ФИО' in plain:
					try:
						surname, name, patronymic = result_value.split()
						result_value = '%s %s. %s.' % (surname,name[0],patronymic[0])
					except Exception as e:
						showerror(u'Ошибка.',u'Не могу разделить ФИО на фамилию и инициалы для %s.\n%s' % (plain,e))
						continue

				line = line.replace(plain,result_value)

				# rewrite line on new replaced line
				# make text field editable
				self.result_text.config(state=NORMAL)

				self.result_text.delete(start_of_line,end_of_line)
				self.result_text.insert(start_of_line,line)

				# make text field not editable
				self.result_text.config(state=DISABLED)

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

				# process specific replacement for initials
				if 'ФИО' in plain:
					try:
						surname, name, patronymic = result_value.split()
						result_value = '%s %s. %s.' % (surname,name[0],patronymic[0])
					except Exception as e:
						showerror(u'Ошибка.',u'Не могу разделить ФИО на фамилию и инициалы для %s.\n%s' % (plain,e))
						continue

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
