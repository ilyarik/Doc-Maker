from .functions import *
from tkinter.filedialog import *
from tkinter.messagebox import *
from tkinter.font import Font
import tkinter.ttk as ttk
import configparser
from .OptionsWindow import OptionsWindow
from collections import OrderedDict

class DocMaker:

	def __init__(self, root,root_dir):

		self.root = root
		self.root_dir = root_dir
		self.big_font = Font(family="Helvetica",size=14)
		self.default_font = Font(family="Helvetica",size=12)
		self.small_font = Font(family="Helvetica",size=10)

		self.act_of_transfer = StringVar()
		self.return_act = StringVar()
		self.act_of_elimination = StringVar()
		self.destination_folder = StringVar()
		self.base_file = StringVar()

		self.menu = Menu(self.root)
		self.filemenu = Menu(self.menu,tearoff=0)
		self.filemenu.add_command(label=u"Выход", command=self.exit)
		self.optionsmenu = Menu(self.menu,tearoff=0)
		self.optionsmenu.add_command(label=u"Настройки", command=self.call_options_window)
		self.menu.add_cascade(label=u"Файл",menu=self.filemenu)
		self.menu.add_cascade(label=u"Настройки", menu=self.optionsmenu)
		self.root.config(menu=self.menu)

		# create frame for info about files
		self.info_frame = Frame(self.root,pady=10,padx=10)
		# input file label and filename and change-button
		self.base_file_label = Label(
			self.info_frame,
			text=u'Файл базы: ',
			font=self.default_font
			)
		self.base_file_label_text = Label(
			self.info_frame,
			textvariable=self.base_file,
			font=self.default_font,
			padx=10
			)
		self.base_file_change_button = Button(
			self.info_frame,
			text=u"…",
			font=self.default_font
			)
		# generate button
		self.generate_button = Button(
			self.info_frame,
			text=u'Сгенерировать',
			font=self.default_font
			)

		# create notebook with two tabs
		self.notebook = ttk.Notebook(
			self.root
			)
		self.base_frame = Frame(
			self.notebook
			)
		self.generator_frame = Frame(
			self.notebook
			)
		self.notebook.add(self.base_frame, text=u'База')
		self.notebook.add(self.generator_frame, text=u'Генератор')

		# create plugs for tabs
		# it displays when files didn't selected
		self.base_frame_plug = Label(
			self.base_frame,
			text=u'Место для базы.',
			font=self.big_font
			)
		self.generator_frame_plug = Label(
			self.generator_frame,
			text=u'Место для генератора документов.',
			font=self.big_font
			)

		# create table with entries
		self.num_of_entries = 0
		self.num_of_fields = 0
		self.table_frame = Frame(self.base_frame,pady=10,padx=10)
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
		ttk.Style().configure('Treeview',rowheight=25)						# set row height
		self.base_table.tag_configure('yellow_row',background='yellow')		# set tags
		self.base_table.tag_configure('table_text',font=self.small_font)

		# labels for entry inputs and addition modes
		self.entry_frame = Frame(self.base_frame,pady=10,padx=10)
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
		# fields for current entry and addition mode
		self.ADDITION_MODES = (u'Не заполнять', u'Инкремент', u'Константа', u'Выбрать из списка')
		self.entry_inputs = []
		self.entry_options = []
		self.entry_option_vars = []


		# frame for add, delete and save buttons
		self.action_frame = Frame(self.base_frame,pady=10,padx=10)
		self.add_entry_button = Button(
			self.action_frame, 
			text=u'Добавить (ctrl+Enter)',
			font=self.default_font
			)
		self.del_entry_button = Button(
			self.action_frame, 
			text=u'Удалить (Del)',
			font=self.default_font
			)
		self.save_base_button = Button(
			self.action_frame, 
			text=u'Сохранить (ctrl+s)',
			font=self.default_font
			)

		# create text field for display primary text from doc file
		self.plain_text_frame = Frame(self.generator_frame, padx=10,pady=10)
		self.plain_text_label = Label(
			self.plain_text_frame,
			text=u'Начальный текст',
			font=self.default_font
			)
		self.plain_text = Text(
			self.plain_text_frame,
			width=45,
			font=self.default_font,
			height=30
			)
		self.plain_text.tag_config('replaced',background='yellow')

		# create text field for display text with replaced values
		self.result_text_frame = Frame(self.generator_frame,padx=10,pady=10)
		self.result_text_label = Label(
			self.result_text_frame,
			text=u'Результирующий текст',
			font=self.default_font)
		self.result_text = Text(
			self.result_text_frame,
			width=45,
			font=self.default_font,
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
			self.generator_frame,
			borderwidth=0,
			width=290,
			height=500)
		self.replacement_frame = Frame(self.replacement_canvas,padx=10,pady=10)
		self.replacement_frameScroll = Scrollbar(
			self.generator_frame,
			orient=VERTICAL,
			command=self.replacement_canvas.yview)
		self.add_replacement_button = Button(
			self.replacement_frame,
			text=u'Добавить',
			font=self.small_font,
			width=20,
			height=1
			)
		# some magic for create scrollable frame with replacement fields
		self.replacement_canvas.configure(yscrollcommand=self.replacement_frameScroll.set)
		self.replacement_canvas.create_window((4,4), window=self.replacement_frame, anchor=CENTER, 
								  tags="self.replacement_frame")

		# create status bar in bottom of main window
		self.status_bar = Label(
			self.root,
			border=1,
			relief=SUNKEN,
			anchor=W,
			text = u'',
			font=self.default_font,
			padx=10
			)

		self.pack_all()
		self.bind_all()

		self.read_options()

	def pack_all(self):

		'''Pack all elements in a window'''
		self.info_frame.pack(side=TOP,fill=X)
		self.base_file_label.grid(row=1,column=0,sticky=W+N)
		self.base_file_label_text.grid(row=1,column=1,sticky=W+N)
		self.base_file_change_button.grid(row=1,column=2,sticky=W+N)
		if self.base_file:
			self.generate_button.place(relx=0.6,rely=0,width=130,height=40)
		self.notebook.pack(side=TOP,fill=BOTH,expand=True,padx=10,pady=10)

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
		self.generator_frame_plug.pack_forget()
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

		if self.base_file:
			self.table_frame.pack(side=TOP,fill=X)
			self.base_table.grid(row=2,column=0,sticky=W+N)
			self.base_tableScroll.grid(row=2,column=1,sticky=N+W+S)
			self.entry_frame.pack(side=LEFT,fill=Y)
			self.entry_label.grid(row=0,column=0,sticky=W+N)
			self.modes_label.grid(row=0,column=1,sticky=W+N,padx=10)
			for index, entry_input in enumerate(self.entry_inputs):
				entry_input.grid(row=index+1,column=0,sticky=W+N)
				entry_input.lift()												# set tab order
			for index, entry_combobox in enumerate(self.entry_options):
				entry_combobox.grid(row=index+1,column=1,sticky=W+N,padx=10)
			self.action_frame.pack(side=LEFT,fill=Y)
			self.add_entry_button.grid(row=4,column=2,sticky=W+N+E+S)
			self.del_entry_button.grid(row=5,column=2,sticky=W+N+E+S)
			self.save_base_button.grid(row=6,column=2,sticky=W+N+E+S)
		else:
			self.base_frame_plug.pack(side=TOP,fill=BOTH,expand=True)

		# if self.example_file:

		# 	self.plain_text_frame.pack(side=LEFT,fill=Y)
		# 	self.plain_text_label.pack(side=TOP)
		# 	self.plain_text.pack(side=LEFT,fill=Y)
		# 	self.result_text_frame.pack(side=RIGHT,fill=Y)
		# 	self.result_text_label.pack(side=TOP)
		# 	self.result_text.pack(side=RIGHT,fill=Y)
		# 	self.replacement_frameScroll.pack(side=RIGHT, fill=Y)
		# 	self.replacement_canvas.pack(side=LEFT,fill=BOTH,expand=True)
		# 	for index in range(self.num_of_replacements):
		# 		self.primary_values_for_replacement[index].grid(row=index*4,column=0,pady=2,sticky=N)
		# 		self.labels_for_replacement[index].grid(row=index*4+1,column=0,pady=2,sticky=N)
		# 		self.new_values_for_replacement[index].grid(row=index*4+2,column=0,pady=2,sticky=N)
		# 	for index, replacement_separator in enumerate(self.replacements_separators):
		# 		replacement_separator.grid(row=index*4+3,column=0,padx=10,pady=5,sticky=W+E)
		# 	self.add_replacement_button.grid(row=self.num_of_replacements*5,column=0,pady=10,sticky=N)
		# else:
		self.generator_frame_plug.pack(side=TOP,fill=BOTH,expand=True)

		self.status_bar.pack(side=BOTTOM,fill=X)

	def bind_all(self):

		'''Bind all events in a window'''
		self.base_file_change_button.bind('<Button-1>',self.load_base)
		self.generate_button.bind('<Button-1>',self.generate_acts)
		if self.base_file:
			self.base_table.bind("<<TreeviewSelect>>", self.setCurrentEntry)
			[entry_input.bind("<Return>",self.saveEntry) for entry_input in self.entry_inputs]
			self.add_entry_button.bind('<Button-1>',self.add_entry)
			self.del_entry_button.bind('<Button-1>',self.del_entry)
			self.save_base_button.bind('<Button-1>',self.save_base)
			self.root.bind('<Control-Return>',self.add_entry)
			self.root.bind('<Delete>',self.del_entry)
			self.root.bind('<Control-s>',self.save_base)

		# if self.example_file:
		# 	self.add_replacement_button.bind('<Button-1>',self.add_replacement)
		# 	self.replacement_frame.bind("<Configure>", self.replaceFrameConfigure)		# bind mouse scroll
		# 	[primary_val.bind("<Return>",self.replace) for primary_val in self.primary_values_for_replacement]
		# 	[new_val.bind("<Return>",self.replace) for new_val in self.new_values_for_replacement]

	def setCurrentEntry(self,event=None):

		selitems = self.base_table.selection()
		if selitems:
			selitem = selitems[-1]
			values = self.base_table.item(selitem)['values']
			for index,value in enumerate(values):
				if isinstance(self.entry_inputs[index], ttk.Combobox):
					self.entry_inputs[index].set(value)
				elif isinstance(self.entry_inputs[index], Entry):
					self.entry_inputs[index].delete(0,END)
					self.entry_inputs[index].insert(0,value)

	def read_options(self):

		'''Read options from .ini file'''

		configs = configparser.ConfigParser()
		configs.read(u'%s\\USER\\configs.ini' % (self.root_dir))

		self.base_file.set(configs['DEFAULT']['base_file'])
		self.act_of_transfer.set(configs['DEFAULT']['act_of_transfer'])
		self.return_act.set(configs['DEFAULT']['return_act'])
		self.act_of_elimination.set(configs['DEFAULT']['act_of_elimination'])
		self.destination_folder.set(configs['DEFAULT']['destination_folder'])

	def write_options(self):

		configs = configparser.ConfigParser()
		configs['DEFAULT']['base_file'] = self.base_file.get()
		configs['DEFAULT']['act_of_transfer'] = self.act_of_transfer.get()
		configs['DEFAULT']['return_act'] = self.return_act.get()
		configs['DEFAULT']['act_of_elimination'] = self.act_of_elimination.get()
		configs['DEFAULT']['destination_folder'] = self.destination_folder.get()
		with open(u'%s\\USER\\configs.ini' % (self.root_dir), 'w') as configfile:
			configs.write(configfile)
			configfile.close()

	def exit(self):

		self.root.quit()

	def call_options_window(self):

		optionsWindow = Toplevel(self.root)
		optionsWindow.geometry('780x500')
		optionsWindow.title(u'Настройки')

		window = OptionsWindow(self,optionsWindow)

	def saveEntry(self,event=None):

		selitems = self.base_table.selection()
		if not selitems:
			return
		selitem = selitems[-1]
		tags = []
		values = [entry_input.get() for entry_input in self.entry_inputs]
		if u'' in values:
			tags.append("yellow_row")
		self.base_table.item(selitem,values=values,tags=tags)

	def add_entry(self,event=None):

		tags = []
		rows = self.base_table.get_children()
		if rows:
			values = self.base_table.item(rows[-1])['values']
			for index in range(self.num_of_fields):
				mode = self.entry_option_vars[index].get()	# get addition mode from comboboxes
				if mode == u'Не заполнять':
					values[index] = u''
				if mode == u'Инкремент':
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
							showerror(u'Ошибка!', 'Ошибка во время добавления.')
		else:
			values = ['']*self.num_of_fields
		if u'' in values:
			tags.append("yellow_row")
		self.base_table.insert('', 'end', values=values, tags=tags)
		self.num_of_entries+=1
		self.base_table.selection_set(self.base_table.get_children()[-1])
		self.base_table.see(self.base_table.get_children()[-1])
		self.status_bar['text'] = u'Запись добавлена'

	def del_entry(self,event=None):

		selitems = self.base_table.selection()
		if not selitems:
			return
		if askyesno(u'Удаление',u'Удалить?'):
			tags = []
			for selitem in selitems:
				self.base_table.delete(selitem)
				self.num_of_entries-=1
		self.status_bar['text'] = u'Запись удалена'

	def load_base(self,event=None):

		filename = askopenfilename(filetypes=(("XLS files", "*.xls;*.xlsx"),('All files','*.*')))
		if not filename:
			return
		
		try:
			entries = get_data_xls(filename)
		except Exception as e:
			showerror(u'Ошибка!',e)
			return

		if not entries:
			showerror(u'Ошибка!',u'Пустой файл базы.')
			return
		
		self.base_file_label_text['text'] = filename.split('/')[-1]
		self.base_file = filename
		self.num_of_entries = len(entries)
		self.num_of_fields = len(entries[0])
		self.base_table['columns'] = ['']*self.num_of_fields

		self.base_table.delete(*self.base_table.get_children())
		self.fill_table(entries)
		# if self.example_file:
		# 	self.get_replace_variants()

		# set columns width
		col_width = int(self.root.winfo_width()*0.95//self.num_of_fields)
		for index in range(self.num_of_fields):
			self.base_table.column(
				index, 
				width=col_width
			)

		# destroy entry inputs and comboboxes if exists and create new
		for entry_input in self.entry_inputs:
			entry_input.destroy()
		self.entry_inputs = []
		for entry_option in self.entry_options:
			entry_option.destroy()
		self.entry_options = []
		self.entry_option_vars = []
		
		self.create_entry_inputs(self.num_of_fields)

		self.pack_all()
		self.bind_all()

		self.status_bar['text'] = u'База Загружена'

	def load_act_of_transfer(self,event=None):

		filename = askopenfilename(filetypes=(("Doc files", "*.doc;*.docx"),('All files','*.*')))
		if not filename:
			return

		try:
			doc_text = get_doc_data(filename)
		except Exception as e:
			showerror(u'Ошибка!',e)
			return

		if not doc_text:
			showerror(u'Ошибка!',u'Пустой .doc файл.')
			return

		self.act_of_transfer = filename
		# clear text fields and fill it
		self.act_of_transfer_plain_text.delete(1.0,END)
		self.act_of_transfer_result_text.delete(1.0,END)
		for line in doc_text:
			self.act_of_transfer_plain_text.insert(END,line+'\n')
		self.replace()			# fill result text field and add tags

		# delete replacements if exists and init new replacements
		self.destroy_replacements()
		for _ in range(3):
			self.add_replacement()

		self.status_bar['text'] = u'Документ загружен'

		self.pack_all()
		self.bind_all()

	def save_base(self,event=None):

		filename = asksaveasfilename(filetypes=(("XLS files", "*.xls;*.xlsx"),('All files','*.*')))
		if not filename:
			return
		rows = self.base_table.get_children()
		if not rows:
			return
		entries = []
		for row in rows:
			entry = self.base_table.item(row)['values']
			entries.append(entry)
		save_xls_data(filename,entries)
		self.status_bar['text'] = u'База сохранена'

	def get_replace_variants(self,event=None):

		'''Add fields from base to replace in choicefields'''

		for combobox in self.new_values_for_replacement:
			replace_values = []
			if self.base_file:
				replace_values.extend([u'Столбец %r' % (index+1) for index in range(self.num_of_fields)])
			replace_values.append(u'…')
			combobox['values'] = replace_values

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
				font=self.default_font
				)
			)
		self.labels_for_replacement.append(
			Label(
				self.replacement_frame,
				text=u'Заменить на:',
				font=self.small_font
				)
			)
		self.new_values_for_replacement.append(
			ttk.Combobox(
				self.replacement_frame,
				width=35,
				font=self.small_font,
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

	def replaceFrameConfigure(self,event=None):

		'''Reset the scroll region to encompass the inner frame'''

		self.replacement_canvas.configure(scrollregion=self.replacement_canvas.bbox("all"))

	def get_result_value(self,index):

		num_of_column = self.new_values_for_replacement[index].current()
		if num_of_column == -1 or num_of_column == self.num_of_fields:
			result_value = self.new_values_for_replacement[index].get()
		elif self.base_file:
			result_value = self.base_table.item(self.base_table.get_children()[-1])['values'][num_of_column]
		else:
			showerror(u'Ошибка!', u'Не удается получить данные для замены.')
			return
		return result_value

	# def replace(self,event=None):

	# 	text = self.plain_text.get(1.0,END)
	# 	for index in range(self.num_of_replacements):
	# 		plain = self.primary_values_for_replacement[index].get()
	# 		if not plain:
	# 			continue

	# 		result_value = self.get_result_value(index)
			
	# 		text = text.replace(
	# 			plain,
	# 			result_value
	# 			)

	# 	self.result_text.delete(1.0,END)
	# 	self.result_text.insert(1.0,text)

	# 	self.add_tags()

	# def add_tags(self):

	# 	'''Remove tags from text field and add anew'''

	# 	for tag in self.plain_text.tag_names():
	# 		self.plain_text.tag_remove(tag,1.0,END)
	# 	for tag in self.result_text.tag_names():
	# 		self.result_text.tag_remove(tag,1.0,END)

	# 	# add yellow tag for replaced words
	# 	for index in range(self.num_of_replacements):
	# 		plain = self.primary_values_for_replacement[index].get()
	# 		if not plain:
	# 			continue

	# 		num_of_column = self.new_values_for_replacement[index].current()
	# 		if num_of_column == -1 or num_of_column == self.num_of_fields:
	# 			result_value = self.new_values_for_replacement[index].get()
	# 		elif self.base_file:
	# 			result_value = self.base_table.item(self.base_table.get_children()[-1])['values'][num_of_column]
	# 		else:
	# 			showerror(u'Ошибка!', u'Ошибка во время замены.')
	# 		if not result_value:
	# 			continue

	# 		search_start = 1.0
	# 		while True:
	# 			start = self.plain_text.search(plain,search_start,stopindex=END)
	# 			if not start:
	# 				break
	# 			end = start.split('.')[0]+'.'+str(int(start.split('.')[1])+len(plain))
	# 			search_start = end
	# 			self.plain_text.tag_add('replaced',start,end)
	
	# 		search_start = 1.0
	# 		while True:
	# 			start = self.result_text.search(result_value,search_start,stopindex=END)
	# 			if not start:
	# 				break
	# 			end = start.split('.')[0]+'.'+str(int(start.split('.')[1])+len(result_value))
	# 			search_start = end
	# 			self.result_text.tag_add('replaced',start,end)

	def generate_acts(self,event=None):

		'''Generate new word documents with replaced values from base'''

		if not self.base_file:
			return

		direct_folder = askdirectory(mustexist=True)
		if not direct_folder:
			return

		# get data from table
		entries = [self.base_table.item(children)['values'] for children in self.base_table.get_children()]
		# get replacements for every row in base_table and generate new files
		for entry_index in range(self.num_of_entries):
			replacements = {}
			for index in range(self.num_of_replacements):
				primary_val = self.primary_values_for_replacement[index].get()
				if not primary_val:
					continue
				result_val = self.get_result_value(index)
				replacements.update({primary_val:result_val})
			doc_filename = direct_folder + '/' + u'Акт %r.docx' % (entry_index+1)
			create_new_replaced_doc(self.example_file, doc_filename, replacements)

	def fill_table(self, entries):

		for index,entry in enumerate(entries):
			values=[]
			tags = []
			for cell in entry:
				if not cell:
					values.append(u'')
				else:
					values.append(str(cell))
			if u'' in values:
				tags.append("yellow_row")
			self.base_table.insert('', 'end', values=values, tags=tags)

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
			var.set(self.ADDITION_MODES[0])
			self.entry_option_vars.append(var)
			self.entry_options.append(
				OptionMenu(
					self.entry_frame, 
					var, 
					*self.ADDITION_MODES,
					command = self.change_entry_inputs
					)
			)
			
	def change_entry_inputs(self,event=None):

		'''Change input when entry option has chaged'''
		for index in range(self.num_of_fields):
			
			cur_val = self.entry_inputs[index].get()
			if self.entry_option_vars[index].get() == u'Выбрать из списка':
				values = list(set([self.base_table.item(item)['values'][index] for item in self.base_table.get_children()]))
				item = ttk.Combobox(
						self.entry_frame,
						width=23,
						font=self.default_font,
						values=values
					)
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