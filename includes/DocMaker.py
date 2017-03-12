from .functions import *
from tkinter import *
from tkinter.filedialog import *
from tkinter.messagebox import *
from tkinter.font import Font
import tkinter.ttk as ttk
import configparser
import datetime
from .OptionsWindow import OptionsWindow
from .GenerateInfoWindow import GenerateInfoWindow
from .ReplacementsTabFrame import ReplacementsTabFrame

class DocMaker(Tk):

	def __init__(self,root_dir):

		Tk.__init__(self)

		self.geometry('1200x720+50+10')
		self.title(u'Составитель актов 2000')
		self.update()

		self.root_dir = root_dir
		self.big_font = Font(family="Helvetica",size=14)
		self.default_font = Font(family="Helvetica",size=12)
		self.small_font = Font(family="Helvetica",size=10)

		self.act_of_transfer = StringVar()
		self.return_act = StringVar()
		self.act_of_elimination = StringVar()
		self.destination_folder = StringVar()
		self.base_file = StringVar()
		self.create_aot = BooleanVar()
		self.create_ra = BooleanVar()
		self.create_aoe = BooleanVar()

		self.menu = Menu(self)
		self.filemenu = Menu(self.menu,tearoff=0)
		self.filemenu.add_command(label=u"Выход", command=self.exit)
		self.optionsmenu = Menu(self.menu,tearoff=0)
		self.optionsmenu.add_command(label=u"Настройки", command=self.call_options_window)
		self.menu.add_cascade(label=u"Файл",menu=self.filemenu)
		self.menu.add_cascade(label=u"Настройки", menu=self.optionsmenu)
		self.config(menu=self.menu)

		# create frame for info about files
		self.info_frame = Frame(self,pady=10,padx=10)
		# input file label and filename and change-button
		self.base_file_label = Label(
			self.info_frame,
			text=u'Файл базы: ',
			font=self.default_font
			)
		self.base_file_label_text = Label(
			self.info_frame,
			text='',
			font=self.default_font,
			padx=10
			)
		self.base_file_change_button = Button(
			self.info_frame,
			text=u"…",
			font=self.default_font
			)
		self.create_aot_check = Checkbutton(
			self.info_frame,
			text=u'Акт передачи',
			variable=self.create_aot,
			onvalue=True,
			offvalue=False
			)
		self.create_ra_check = Checkbutton(
			self.info_frame,
			text=u'Акт возврата',
			variable=self.create_ra,
			onvalue=True,
			offvalue=False
			)
		self.create_aoe_check = Checkbutton(
			self.info_frame,
			text=u'Акт уничтожения',
			variable=self.create_aoe,
			onvalue=True,
			offvalue=False
			)
		# generate button
		self.generate_button = Button(
			self.info_frame,
			text=u'Сгенерировать',
			font=self.default_font
			)

		# create notebook with tabs
		# aot - act of transfer
		# ra - return act
		# aoe - act of elimination
		self.notebook = ttk.Notebook(
			self
			)
		self.base_frame = Frame(
			self.notebook
			)
		self.aot_frame = ReplacementsTabFrame(
			self,
			act_name='aot',
			act_var=self.act_of_transfer,
			plug_text=u'Место для акта передачи'
			)
		self.ra_frame = ReplacementsTabFrame(
			self,
			act_name='ra',
			act_var=self.return_act,
			plug_text=u'Место для акта возврата'
			)
		self.aoe_frame = ReplacementsTabFrame(
			self,
			act_name='aoe',
			act_var=self.act_of_elimination,
			plug_text=u'Место для акта уничтожения'
			)

		self.notebook.add(self.base_frame, text=u'База')
		self.notebook.add(self.aot_frame, text=u'Акт передачи')
		self.notebook.add(self.ra_frame, text=u'Акт возврата')
		self.notebook.add(self.aoe_frame, text=u'Акт уничтожения')

		# create plugs for tabs
		# it displays when files didn't selected
		self.base_frame_plug = Label(
			self.base_frame,
			text=u'Место для базы.',
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

		# create status bar in bottom of main window
		self.status_bar = Label(
			self,
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
		self.load_base()
		self.aot_frame.load_act()
		self.ra_frame.load_act()
		self.aoe_frame.load_act()

	def pack_all(self):

		'''Pack all elements in a window'''
		self.info_frame.pack(side=TOP,fill=X)
		self.base_file_label.grid(row=1,column=0,sticky=W+N)
		self.base_file_label_text.grid(row=1,column=1,sticky=W+N)
		self.base_file_change_button.grid(row=1,column=2,sticky=W+N)
		if self.base_file:
			self.create_aot_check.place(relx=0.5,rely=0.2)
			self.create_ra_check.place(relx=0.6,rely=0.2)
			self.create_aoe_check.place(relx=0.7,rely=0.2)
			self.generate_button.place(relx=0.85,rely=0,width=130,height=40)
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

		if self.base_file.get():
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

		self.aot_frame.pack_all()
		self.ra_frame.pack_all()
		self.aoe_frame.pack_all()

		self.status_bar.pack(side=BOTTOM,fill=X)

	def bind_all(self):

		'''Bind all events in a window'''
		self.base_file_change_button.bind('<Button-1>',self.set_base_file)
		self.generate_button.bind('<Button-1>',self.generate_info)

		if self.base_file.get():
			self.base_table.bind("<<TreeviewSelect>>", self.setCurrentEntry)
			for entry_input in self.entry_inputs:
				entry_input.bind("<Return>",self.saveEntry)
				if isinstance(entry_input,ttk.Combobox):
					entry_input.bind('<KeyRelease>',self.change_combobox_values)
			self.add_entry_button.bind('<Button-1>',self.add_entry)
			self.del_entry_button.bind('<Button-1>',self.del_entry)
			self.save_base_button.bind('<Button-1>',self.save_base)
			self.bind('<Control-Return>',self.add_entry)
			self.bind('<Delete>',self.del_entry)
			self.bind('<Control-s>',self.save_base)

		self.aot_frame.bind_all()
		self.ra_frame.bind_all()
		self.aoe_frame.bind_all()

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
		self.act_of_transfer.set(configs['Act_of_transfer']['filename'])
		self.return_act.set(configs['Return_act']['filename'])
		self.act_of_elimination.set(configs['Act_of_elimination']['filename'])
		self.destination_folder.set(configs['DEFAULT']['destination_folder'])

	def write_options(self):

		configs = configparser.ConfigParser()
		configs.read(u'%s\\USER\\configs.ini' % (self.root_dir))
		configs['DEFAULT']['base_file'] = self.base_file.get()
		configs['DEFAULT']['destination_folder'] = self.destination_folder.get()

		configs['Act_of_transfer']['filename'] = self.act_of_transfer.get()
		configs['Return_act']['filename'] = self.return_act.get()
		configs['Act_of_elimination']['filename'] = self.act_of_elimination.get()
		
		with open(u'%s\\USER\\configs.ini' % (self.root_dir), 'w') as configfile:
			configs.write(configfile)
			configfile.close()

	def exit(self):

		self.quit()

	def call_options_window(self):

		window = OptionsWindow(self)

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
							showerror(u'Ошибка!', 'Ошибка во время добавления.\n'+e)
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

	def set_base_file(self,event=None):

		filename = askopenfilename(filetypes=(("XLS files", "*.xls;*.xlsx"),('All files','*.*')))
		if not filename:
			return

		self.base_file.set(filename)
		self.write_options()
		self.load_base()

	def load_base(self,event=None):

		if not self.base_file.get():
			return

		try:
			entries = get_data_xls(self.base_file.get())
		except Exception as e:
			showerror(u'Ошибка!',e)
			self.act_of_transfer.set(u'')
			return

		if not entries:
			showerror(u'Ошибка!',u'Пустой файл базы.')
			self.act_of_transfer.set(u'')
			return
		
		self.base_file_label_text['text'] = get_truncated_line(self.base_file.get(),40)
		self.num_of_entries = len(entries)
		self.num_of_fields = len(entries[0])
		self.base_table['columns'] = ['']*self.num_of_fields

		self.base_table.delete(*self.base_table.get_children())
		self.fill_table(entries)

		# change variants in choice fields into tabs
		if self.act_of_transfer.get():
			self.aot_frame.get_replace_variants()
		if self.return_act.get():
			self.ra_frame.get_replace_variants()
		if self.act_of_elimination.get():
			self.aoe_frame.get_replace_variants()

		# set columns width
		col_width = int(self.winfo_width()*0.95//self.num_of_fields)
		self.base_frame.update()
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

	def generate_info(self,event=None):

		'''Show this window before generate'''

		if not self.base_file.get():
			return

		if not self.base_table.selection():
			return

		if not self.create_aot.get() and not self.create_ra.get() and not self.create_aoe.get():
			return

		window = GenerateInfoWindow(self)

	def generate_acts(self,event=None):

		'''Generate new word documents with replaced values from base'''

		if not self.base_file.get():
			return

		if not self.base_table.selection():
			return

		# get data from table
		entries = [self.base_table.item(selitem)['values'] for selitem in self.base_table.selection()]
		# get replacements for selected rows in base_table and generate new files
		for index_row,entry in enumerate(entries):
			if self.create_aot.get():
				replacements = {}
				for index in range(self.aot_frame.num_of_replacements):
					primary_val = self.aot_frame.primary_values_for_replacement[index].get()
					if not primary_val:
						continue
					result_val = self.aot_frame.get_result_value(index,index_row)
					replacements.update({primary_val:str(result_val)})
				doc_filename = self.destination_folder.get() + '/' + u'Акт %s №%s-%s-%s.docx' % (
								u'Передачи',
								str(datetime.datetime.now().year),
								str(entry[0]),
								u'П'
								)
				create_new_replaced_doc(self.act_of_transfer.get(), doc_filename, replacements)

			if self.create_ra.get():
				replacements = {}
				for index in range(self.ra_frame.num_of_replacements):
					primary_val = self.ra_frame.primary_values_for_replacement[index].get()
					if not primary_val:
						continue
					result_val = self.ra_frame.get_result_value(index,index_row)
					replacements.update({primary_val:str(result_val)})
				doc_filename = self.destination_folder.get() + '/' + u'Акт %s №%s-%s-%s.docx' % (
								u'Возврата',
								str(datetime.datetime.now().year),
								str(entry[0]),
								u'В'
								)
				create_new_replaced_doc(self.return_act.get(), doc_filename, replacements)

			if self.create_aoe.get():
				replacements = {}
				for index in range(self.aoe_frame.num_of_replacements):
					primary_val = self.aoe_frame.primary_values_for_replacement[index].get()
					if not primary_val:
						continue
					result_val = self.aoe_frame.get_result_value(index,index_row)
					replacements.update({primary_val:str(result_val)})
				doc_filename = self.destination_folder.get() + '/' + u'Акт %s №%s-%s-%s.docx' % (
								u'Уничтожения',
								str(datetime.datetime.now().year),
								str(entry[0]),
								u'У'
								)
				create_new_replaced_doc(self.act_of_elimination.get(), doc_filename, replacements)

		self.status_bar['text'] = u'Готово'

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
				values = list(set(
					[self.base_table.item(item)['values'][index] \
					for item in self.base_table.get_children() \
					if self.base_table.item(item)['values'][index]
					]))
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

	def change_combobox_values(self, event=None):

		'''Autocompletion for comboboxes'''

		for index in range(self.num_of_fields):
			# filter combobox input values
			if isinstance(self.entry_inputs[index], ttk.Combobox):
				values = list(set(
					[self.base_table.item(item)['values'][index] \
					for item in self.base_table.get_children() \
					if self.base_table.item(item)['values'][index]
					]))
				values = list(filter(lambda value: value.startswith(self.entry_inputs[index].get()),values))
				self.entry_inputs[index]['values'] = values