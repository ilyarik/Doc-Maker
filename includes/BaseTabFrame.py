# -*- coding: utf-8 -*-
from tkinter.filedialog import *
from tkinter.messagebox import *
from tkinter.font import Font
from tkinter import *
import tkinter.ttk as ttk
from .functions import *
from .AutocompleteCombobox import AutocompleteCombobox
from .Calendar import Calendar
import configparser
import glob
import datetime
import calendar

class BaseTabFrame(Frame):

	def __init__(self, mainWindow,title_text):

		Frame.__init__(self)

		self.mainWindow = mainWindow
		self.small_font = self.mainWindow.small_font
		self.default_font = self.mainWindow.default_font
		self.big_font = self.mainWindow.big_font

		# num of columns with date of act of transfer and act of elimination
		# set values in self.readDateCols
		self.aot_date_col = IntVar()
		self.aoe_date_col = IntVar()
		self.date_format = StringVar()

		# it displays when files didn't selected
		self.base_frame_plug = Label(
			self,
			text=u'Место для базы.',
			font=self.big_font
			)

		# create table with entries
		self.num_of_entries = 0
		self.num_of_fields = 0

		self.title = Label(
			self,
			text=title_text,
			font=self.big_font,
			pady=10
			)

		self.table_frame = Frame(self,pady=10,padx=10)
		ttk.Style().configure('Treeview',rowheight=25)						# set row height

		# labels for entry inputs and addition modes
		self.entry_frame = Frame(self,pady=10,padx=10)
		self.entry_label = Label(
			self.entry_frame,
			text=u"Текущая запись:",
			font=self.default_font
			)
		self.modes_label = Label(
			self.entry_frame,
			text=u'При добавлении:',
			font=self.small_font
			)
		self.aotDateButton = Button(
			self.entry_frame,
			text=u'Дата',
			font=self.small_font
			)

		# fields for current entry and addition mode
		self.ADDITION_MODES = (u'Не заполнять', u'Инкремент', u'Константа', u'Выбрать из списка')
		self.addition_modes_default = []
		self.entry_inputs = []
		self.entry_options = []
		self.entry_option_vars = []

		# frame for add, delete and save buttons
		self.action_frame = Frame(self,pady=10,padx=10)
		self.add_entry_button = Button(
			self.action_frame, 
			text=u'Добавить (Ctrl+Enter)',
			font=self.default_font
			)
		self.del_entry_button = Button(
			self.action_frame, 
			text=u'Удалить',
			font=self.default_font
			)
		self.save_base_button = Button(
			self.action_frame, 
			text=u'Сохранить базу (Ctrl+S)',
			font=self.default_font
			)

		self.readBaseData()

	def pack_all(self):

		'''Pack all elements in a tab frame'''
		self.title.pack_forget()
		self.base_frame_plug.pack_forget()
		self.table_frame.pack_forget()
		self.base_table.grid_forget()
		self.base_tableScroll.grid_forget()
		self.entry_frame.pack_forget()
		self.entry_label.grid_forget()
		self.modes_label.grid_forget()
		for index in range(self.num_of_fields):
			self.entry_inputs[index].grid_forget()
			self.entry_options[index].grid_forget()
		self.action_frame.pack_forget()
		self.add_entry_button.grid_forget()
		self.del_entry_button.grid_forget()
		self.save_base_button.grid_forget()

		if self.mainWindow.base_file.get():
			self.title.pack(side=TOP,fill=X)
			self.table_frame.pack(side=TOP,fill=X)
			self.base_table.pack(side=LEFT,fill=X,expand=True)
			self.base_tableScroll.pack(side=RIGHT,fill=Y)
			self.entry_frame.pack(side=LEFT,fill=Y)
			self.entry_label.grid(row=0,column=1,sticky=W+N)
			self.modes_label.grid(row=0,column=2,sticky=W+N,padx=10)
			for index, entry_input in enumerate(self.entry_inputs):
				# don't display auto date column
				if index == self.aoe_date_col.get():
					continue
				Label(
					self.entry_frame,
					text=get_truncated_line(self.base_table.heading(index)['text'],20),
					font=self.small_font
					).grid(row=index+1,column=0,sticky=W+N)
				entry_input.grid(row=index+1,column=1,sticky=W)
				entry_input.lift()												# set tab order
			for index, entry_option in enumerate(self.entry_options):
				# check entry is autocompletable
				if index == self.aot_date_col.get():
					self.aotDateButton.grid(row=index+1,column=2,sticky=W+E)
				elif index == self.aoe_date_col.get():
					continue
				else:
					entry_option.grid(row=index+1,column=2,sticky=W+N+S)
			self.action_frame.pack(side=LEFT,fill=Y)
			self.add_entry_button.grid(row=4,column=2,sticky=W+N+E+S)
			self.del_entry_button.grid(row=5,column=2,sticky=W+N+E+S)
		else:
			self.base_frame_plug.pack(side=TOP,fill=BOTH,expand=True)

	def bind_all(self):

		'''Bind all events in a tab frame'''
		if self.mainWindow.base_file.get():
			self.base_table.bind("<<TreeviewSelect>>", self.setCurrentEntry)
			for index,entry_input in enumerate(self.entry_inputs):
				entry_input.bind("<Return>",self.saveEntry)
				entry_input.bind("<Key>",lambda event=None,entry=entry_input:self.clipboard(event,entry))
			self.aotDateButton.bind('<Button-1>',self.setAotDate)
			self.add_entry_button.bind('<Button-1>',self.addEntry)
			self.del_entry_button.bind('<Button-1>',self.delEntry)

	def clipboard(self,event=None,entry=0):

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

	def readBaseData(self):

		'''Read nums of 2 cols from configs file:
			one with date of transfer act, one with date of elimination act'''
		configs = configparser.RawConfigParser()
		configs.read(self.mainWindow.configsFileName)

		self.mainWindow.base_file.set(configs['Base']['filename'])
		self.aot_date_col.set(int(configs['Base']['aot_date_col'])-1)
		self.aoe_date_col.set(int(configs['Base']['aoe_date_col'])-1)
		self.date_format.set(configs['Base']['date_format'])
		self.addition_modes_default = eval(configs['Base']['addition_modes'])

	def saveAdditionModes(self):

		'''Save default values for optionmenus for addition mode in .ini file'''
		configs = configparser.ConfigParser()
		configs.read(self.mainWindow.configsFileName)

		configs['Base']['addition_modes'] = str([entry_option.get() for entry_option in self.entry_option_vars])
		with open(self.mainWindow.configsFileName, 'w') as configfile:
			configs.write(configfile)
			configfile.close()

	def sync_exist_acts(self):

		'''Looks for docx files in destination folder,
			add info about existing files into last column'''
		# clear last column
		for item in self.base_table.get_children():
			values = self.base_table.item(item)['values']
			values[-1] = ''
			self.base_table.item(item,values=values)

		if not self.mainWindow.base_file.get():
			return

		if not self.mainWindow.destination_folder.get():
			return

		act_filepathes = glob.glob(self.mainWindow.destination_folder.get()+'/*.docx')

		acts = {}
		for act_filepath in act_filepathes:
			try:
				act_filename = act_filepath.split('\\')[-1]
				entry_num = int(re.findall(r'\w+ \w+ №\d{4}-(\d+)-\w{1}.docx',act_filename)[0])
				act_type = re.findall(r'\w+ \w+ №\d{4}-\d+-(\w{1}).docx',act_filename)[0]
				if entry_num in acts.keys():
					acts[entry_num].add(act_type[0])
				else:
					acts[entry_num] = set(act_type)
			except:
				continue

		for row in self.base_table.get_children():
			values = self.base_table.item(row)['values']
			index = values[0]
			new_val = []
			identifiers = [u'П',u'В',u'У']
			for identifier in identifiers:
				if index in acts.keys():
					if identifier in acts[index]:
						new_val.append(identifier)
					else:
						new_val.append(u'_')
				else:
					new_val.append(u'_')
			values[self.num_of_fields] = ', '.join(new_val)
			self.base_table.item(row,values=values,tags=[])

		self.refreshTags()

	def set_base_table_cols(self, columns):

		self.base_table['columns'] = ['']*columns
		col_width = int(self.winfo_width()*0.95//columns)
		for index in range(columns):
			self.base_table.column(
				index, 
				width=col_width
			)

	def setAotDate(self,event=None):

		toplevel = Toplevel(self.mainWindow)
		entry = self.entry_inputs[self.aot_date_col.get()]
		ttkcal = Calendar(entry,self.date_format.get(),master=toplevel)
		ttkcal.pack(expand=1, fill='both')
		toplevel.update()
		toplevel.minsize(toplevel.winfo_reqwidth(), toplevel.winfo_reqheight())

	def setCurrentEntry(self,event=None):

		selitems = self.base_table.selection()
		if selitems:
			selitem = selitems[-1]
			values = self.base_table.item(selitem)['values']
			for index in range(self.num_of_fields):
				value = values[index]
				if isinstance(self.entry_inputs[index], AutocompleteCombobox):
					self.entry_inputs[index].set(value)
				elif isinstance(self.entry_inputs[index], Entry):
					self.entry_inputs[index].delete(0,END)
					self.entry_inputs[index].insert(0,value)

	def saveEntry(self,event=None):

		selitems = self.base_table.selection()
		if not selitems:
			return
		selitem = selitems[-1]
		tags = []
		values_tmp = self.base_table.item(selitem)['values']
		values = [entry_input.get() for entry_input in self.entry_inputs]
		values.append(self.base_table.item(selitem)['values'][self.num_of_fields])
		self.base_table.item(selitem,values=values,tags=tags)
		try:
			self.refreshAutoDate()
		except Exception as e:
			self.base_table.item(selitem,values=values_tmp,tags=tags)
			showerror(u'Ошибка.',u'Неверный формат даты (столбец %r).\n%s' % (self.aot_date_col.get()+1,e))
			return
		self.save(self.mainWindow.base_file.get())
		self.change_combobox_values()

	def addEntry(self,event=None):

		tags = []
		rows = self.base_table.get_children()
		if not rows:
			return

		# get values of last entry in treeview
		values = self.base_table.item(rows[-1])['values']
		for index in range(self.num_of_fields):
			mode = self.entry_option_vars[index].get()	# get addition mode from comboboxes
			if mode == self.ADDITION_MODES[0] or mode == self.ADDITION_MODES[3]:
				values[index] = u''
			elif mode == self.ADDITION_MODES[1]:
				value_type = type(values[index])
				if value_type == int:
					values[index] += 1
				elif value_type == str:
					try:
						number = re.findall(r'\d+$',values[index])
						if number:
							values[index] = re.sub(
								r'\d+$',
								str(int(number[-1])+1),
								values[index])
					except Exception as e:
						showerror(u'Ошибка!', e)
						return
		if u'П' in values[-1]:
			tags = ["green_row"]
		if u'В' in values[-1]:
			tags = ["yellow_row"]
		if u'У' in values[-1]:
			tags = ["red_row"]
		
		self.base_table.insert('', 'end', values=values, tags=tags)
		self.num_of_entries+=1
		self.base_table.selection_set(self.base_table.get_children()[-1])
		self.base_table.see(self.base_table.get_children()[-1])
		self.refreshAutoDate()
		self.save(self.mainWindow.base_file.get())
		self.change_combobox_values()
		self.mainWindow.status_bar['text'] = u'Запись добавлена'

	def delEntry(self,event=None):

		selitems = self.base_table.selection()
		if not selitems:
			return
		if not askyesno(u'Удаление',u'Удалить?'):
			return
		for selitem in selitems:
			self.base_table.delete(selitem)
			self.num_of_entries-=1
		self.mainWindow.status_bar['text'] = u'Запись удалена'
		self.save(self.mainWindow.base_file.get())
		self.change_combobox_values()

	def destroyBaseTable(self):

		try:
			self.base_table.destroy()
			self.base_tableScroll.destroy()
		except AttributeError:
			pass
		
	def initBaseTable(self,event=None):

		'''Initialize base tab, crate entry inputs, repack and rebind all'''
		if not self.mainWindow.base_file.get():
			return

		entries = get_data_xls(self.mainWindow.base_file.get())

		if not entries:
			return

		self.destroyBaseTable()

		self.base_table = ttk.Treeview(
			self.table_frame,
			show = 'headings'
			)
		# create scrollbar for table
		self.base_tableScroll = Scrollbar(
			self.table_frame,
			orient=VERTICAL,
			command=self.base_table.yview)
		self.base_table.configure(yscrollcommand=self.base_tableScroll.set)	# attach scrollbar
		
		self.base_table.tag_configure('green_row',background='#CCFFCC')
		self.base_table.tag_configure('yellow_row',background='#FFFFCC')		# set tags
		self.base_table.tag_configure('red_row',background='#FFCCCC')
		self.base_table.tag_configure('table_text',font=self.small_font)
		
		self.num_of_entries = len(entries)
		self.num_of_fields = len(entries[0])
		self.set_base_table_cols(self.num_of_fields+1)

		self.base_table.delete(*self.base_table.get_children())
		self.fill_table(entries)

		self.sync_exist_acts()

		# destroy entry inputs and comboboxes if exists and create new
		for entry_input in self.entry_inputs:
			entry_input.destroy()
		self.entry_inputs = []
		for entry_option in self.entry_options:
			entry_option.destroy()
		self.entry_options = []
		self.entry_option_vars = []
		
		# repack all and rebind all in this method
		self.create_entry_inputs(self.num_of_fields)

		self.mainWindow.status_bar['text'] = u'База Загружена'
		self.mainWindow.base_file_label_text['text'] = get_truncated_line(self.mainWindow.base_file.get(),40)

	def loadBase(self):

		'''Reload entries into Treeview object'''
		if not self.mainWindow.base_file.get():
			return

		entries = get_data_xls(self.mainWindow.base_file.get())

		if not entries:
			return

		if self.num_of_fields != len(entries[0]):
			showerror('Ошибка!','Изменился файл базы. Перезагрузите программу.')
		self.num_of_entries = len(entries)

		self.base_table.delete(*self.base_table.get_children())
		self.fill_table(entries)

		self.sync_exist_acts()

	def save(self,filename):

		'''Save base to file'''
		if not filename:
			return
		rows = self.base_table.get_children()
		if not rows:
			return
		entries = []
		for row in rows:
			entry = self.base_table.item(row)['values'][:-1]
			# format date columns before put into xlsx
			try:
				aot_value = self.aot_date_col.get()
				if aot_value:
					entry[aot_value] = datetime.datetime.strptime(entry[aot_value], self.date_format.get()).date()
				aoe_value = self.aoe_date_col.get()
				if aoe_value:
					entry[aoe_value] = datetime.datetime.strptime(entry[aoe_value], self.date_format.get()).date()
			except:
				showerror(u'Ошибка.',u'Ошибка во время форматирования столбца с датой перед сохранением.\n%s' % e)
				return
			entries.append(entry)
		# fill rows after base with None values
		[entries.append([None]*self.num_of_fields) for _ in range(100)]
		try:
			save_xls_data(filename,entries)
		except Exception as e:
			showerror(u'Ошибка.',u'Ошибка во время записи в файл %s.\n%s' % (filename,e))
			return

	def fill_table(self, entries):

		# set columns title
		for col_index,cell in enumerate(entries[0]):
			col_name = str(col_index+1)+'. '+str(cell)
			self.base_table.heading(col_index,text=col_name)

		# add column with info about existing acts
		col_name = str(self.num_of_fields+1)+u'. В наличии'
		self.base_table.heading(self.num_of_fields,text=col_name)

		for entry in entries[1:]:
			values=[u'']*self.num_of_fields
			for index_col,cell in enumerate(entry):
				if not cell:
					continue
				if self.aot_date_col.get() and index_col==self.aot_date_col.get():
					try:
						values[index_col] = cell.strftime(self.date_format.get())
					except Exception as e:
						showerror(u'Ошибка.',u'Дата в столбце для даты должна иметь тип Дата и формат %s.\n%s' % (self.date_format.get(),e))
						return
				else:
					values[index_col] = cell
			values.append('')
			self.base_table.insert('', 'end', values=values, tags=[])

		self.refreshAutoDate()
		self.refreshTags()

	def refreshTags(self):

		for item in self.base_table.get_children():
			tags=[]
			values = self.base_table.item(item)['values']
			if u'П' in values[-1]:
				tags = ["green_row"]
			if u'В' in values[-1]:
				tags = ["yellow_row"]
			if u'У' in values[-1]:
				tags = ["red_row"]
			self.base_table.item(item,tags=tags)

	def refreshAutoDate(self):

		'''Check all cells where need to input date of aot and insert aoe date'''
		self.readBaseData()

		if self.aot_date_col.get()<1 or self.aot_date_col.get()>self.num_of_fields:
			return

		if self.aoe_date_col.get()<1 or self.aoe_date_col.get()>self.num_of_fields:
			return

		items = self.base_table.get_children()
		for row_index,item in enumerate(items):
			values = self.base_table.item(item)['values']
			date = values[self.aot_date_col.get()]
			if not date:
				continue
			if type(date) != datetime.datetime:
				date = datetime.datetime.strptime(date, self.date_format.get())
			values[self.aoe_date_col.get()] = add_date(date,years=1,months=3).strftime(self.date_format.get())
			self.base_table.item(item,values=values)

	def create_entry_inputs(self,amount):

		'''Create fields to change the current entry'''
		for index in range(amount):
			self.entry_inputs.append(
				Entry(
				self.entry_frame,
				width=25,
				font=self.default_font
				)
			)
			var = StringVar()
			# set to default values from .ini file if exist
			try:
				mode = self.addition_modes_default[index]
				if mode in self.ADDITION_MODES:
					var.set(mode)
			except:
				var.set(self.ADDITION_MODES[0])
			self.entry_option_vars.append(var)
			optionmenu = OptionMenu(
				self.entry_frame, 
				var, 
				*self.ADDITION_MODES,
				command = self.change_entry_inputs
				)
			optionmenu.configure(width=15)
			self.entry_options.append(optionmenu)
		self.change_entry_inputs()
			
	def change_entry_inputs(self,event=None):

		'''Change input when entry option has chaged'''
		self.saveAdditionModes()

		for index in range(self.num_of_fields):
			cur_val = self.entry_inputs[index].get()
			if self.entry_option_vars[index].get() == self.ADDITION_MODES[3]:
				# get set of values for combobox choices, sorted alphabetically
				values = set()
				for row in self.base_table.get_children():
					value = str(self.base_table.item(row)['values'][index])
					if value:
						values.add(value)	
				values = sorted(values, key=str.lower)
				values = list(filter(lambda value: value.lower().startswith(self.entry_inputs[index].get().lower()),values))
				item = AutocompleteCombobox(
						self.entry_frame,
						width=23,
						font=self.default_font
					)
				item.set_completion_list(values)
				item.set(cur_val)
			else:
				item = Entry(
						self.entry_frame,
						width=25,
						font=self.default_font
					)
				item.insert(0,cur_val)

			self.entry_inputs[index].destroy()
			self.entry_inputs[index] = item
		self.pack_all()
		self.bind_all()
		self.change_combobox_values()

	def change_combobox_values(self, event=None):

		'''Variants for comboboxes'''
		for index in range(self.num_of_fields):
			# filter combobox input values
			if isinstance(self.entry_inputs[index], AutocompleteCombobox):
				values = set()
				for row in self.base_table.get_children():
					value = str(self.base_table.item(row)['values'][index])
					if value:
						values.add(value)
				values = sorted(values, key=str.lower)
				self.entry_inputs[index].set_completion_list(values)

	def getHeadingText(self,col_index):

		return self.base_table.heading(col_index)['text']

	def getColumnAsList(self, col_index):

		'''Get list of values in col_index column in base'''
		rows = self.base_table.get_children()
		if not rows:
			return

		column_list = []

		for row in rows:

			values = self.base_table.item(row)['values']
			if not values:
				return
			if col_index<0 or col_index>self.num_of_fields:
				return

			column_list.append(values[col_index])

		return column_list

	def getAllEntriesAsList(self):

		'''Read all rows in table and put values into list, put these lists into other list and return it'''
		rows = self.base_table.get_children()
		if not rows:
			return []
		entries = []
		
		for row in rows:
			entry = self.base_table.item(row)['values'][:-1]
			# format date columns before put into xlsx
			try:
				aot_value = self.aot_date_col.get()
				if aot_value:
					entry[aot_value] = datetime.datetime.strptime(entry[aot_value], self.date_format.get())
				aoe_value = self.aoe_date_col.get()
				if aoe_value:
					entry[aoe_value] = datetime.datetime.strptime(entry[aoe_value], self.date_format.get())
			except:
				return
			entries.append(entry)
		return entries