from .functions import *
from tkinter.filedialog import *
from tkinter.messagebox import *
from tkinter.font import Font
import tkinter.ttk as ttk
import configparser
from .OptionsWindow import OptionsWindow
from .GenerateInfoWindow import GenerateInfoWindow
from collections import OrderedDict
import datetime

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
		self.create_aot = BooleanVar()
		self.create_ra = BooleanVar()
		self.create_aoe = BooleanVar()

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
			self.root
			)
		self.base_frame = Frame(
			self.notebook
			)
		self.aot_frame = Frame(
			self.notebook
			)
		self.ra_frame = Frame(
			self.notebook
			)
		self.aoe_frame = Frame(
			self.notebook
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
		self.aot_frame_plug = Label(
			self.aot_frame,
			text=u'Место для акта передачи.',
			font=self.big_font
			)
		self.ra_frame_plug = Label(
			self.ra_frame,
			text=u'Место для акта возврата.',
			font=self.big_font
			)
		self.aoe_frame_plug = Label(
			self.aoe_frame,
			text=u'Место для акта уничтожения.',
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
		self.aot_plain_text_frame = Frame(self.aot_frame, padx=10,pady=10)
		self.aot_plain_text_label = Label(
			self.aot_plain_text_frame,
			text=u'Начальный текст',
			font=self.default_font
			)
		self.aot_plain_text = Text(
			self.aot_plain_text_frame,
			width=45,
			font=self.default_font,
			height=30
			)
		self.aot_plain_text.tag_config('replaced',background='yellow')

		# create text field for display text with replaced values
		self.aot_result_text_frame = Frame(self.aot_frame,padx=10,pady=10)
		self.aot_result_text_label = Label(
			self.aot_result_text_frame,
			text=u'Результирующий текст',
			font=self.default_font)
		self.aot_result_text = Text(
			self.aot_result_text_frame,
			width=45,
			font=self.default_font,
			height=30
			)
		self.aot_result_text.tag_config('replaced',background='yellow')

		# create frame for replacements
		# each replacement contain one entry for primary value,
		# one label and one combobox for new value
		# separators split replacements in frame
		self.aot_primary_values_for_replacement = []
		self.aot_labels_for_replacement = []
		self.aot_new_values_for_replacement = []
		self.aot_replacements_separators = []
		self.aot_num_of_replacements = 0
		# create canvas for scrollable frame
		self.aot_replacement_canvas = Canvas(
			self.aot_frame,
			borderwidth=0,
			width=290,
			height=500)
		self.aot_replacement_frame = Frame(self.aot_replacement_canvas,padx=10,pady=10)
		self.aot_replacement_frameScroll = Scrollbar(
			self.aot_frame,
			orient=VERTICAL,
			command=self.aot_replacement_canvas.yview)
		self.aot_add_replacement_button = Button(
			self.aot_replacement_frame,
			text=u'Добавить замену',
			font=self.small_font,
			width=22,
			height=1
			)
		# some magic for create scrollable frame with replacement fields
		self.aot_replacement_canvas.configure(yscrollcommand=self.aot_replacement_frameScroll.set)
		self.aot_replacement_canvas.create_window((4,4), window=self.aot_replacement_frame, anchor=CENTER, 
								  tags="self.aot_replacement_frame")

		# create text field for display primary text from doc file
		self.ra_plain_text_frame = Frame(self.ra_frame, padx=10,pady=10)
		self.ra_plain_text_label = Label(
			self.ra_plain_text_frame,
			text=u'Начальный текст',
			font=self.default_font
			)
		self.ra_plain_text = Text(
			self.ra_plain_text_frame,
			width=45,
			font=self.default_font,
			height=30
			)
		self.ra_plain_text.tag_config('replaced',background='yellow')

		# create text field for display text with replaced values
		self.ra_result_text_frame = Frame(self.ra_frame,padx=10,pady=10)
		self.ra_result_text_label = Label(
			self.ra_result_text_frame,
			text=u'Результирующий текст',
			font=self.default_font)
		self.ra_result_text = Text(
			self.ra_result_text_frame,
			width=45,
			font=self.default_font,
			height=30
			)
		self.ra_result_text.tag_config('replaced',background='yellow')

		# create frame for replacements
		# each replacement contain one entry for primary value,
		# one label and one combobox for new value
		# separators split replacements in frame
		self.ra_primary_values_for_replacement = []
		self.ra_labels_for_replacement = []
		self.ra_new_values_for_replacement = []
		self.ra_replacements_separators = []
		self.ra_num_of_replacements = 0
		# create canvas for scrollable frame
		self.ra_replacement_canvas = Canvas(
			self.ra_frame,
			borderwidth=0,
			width=290,
			height=500)
		self.ra_replacement_frame = Frame(self.ra_replacement_canvas,padx=10,pady=10)
		self.ra_replacement_frameScroll = Scrollbar(
			self.ra_frame,
			orient=VERTICAL,
			command=self.ra_replacement_canvas.yview)
		self.ra_add_replacement_button = Button(
			self.ra_replacement_frame,
			text=u'Добавить замену',
			font=self.small_font,
			width=22,
			height=1
			)
		# some magic for create scrollable frame with replacement fields
		self.ra_replacement_canvas.configure(yscrollcommand=self.ra_replacement_frameScroll.set)
		self.ra_replacement_canvas.create_window((4,4), window=self.ra_replacement_frame, anchor=CENTER, 
								  tags="self.ra_replacement_frame")

		# create text field for display primary text from doc file
		self.aoe_plain_text_frame = Frame(self.aoe_frame, padx=10,pady=10)
		self.aoe_plain_text_label = Label(
			self.aoe_plain_text_frame,
			text=u'Начальный текст',
			font=self.default_font
			)
		self.aoe_plain_text = Text(
			self.aoe_plain_text_frame,
			width=45,
			font=self.default_font,
			height=30
			)
		self.aoe_plain_text.tag_config('replaced',background='yellow')

		# create text field for display text with replaced values
		self.aoe_result_text_frame = Frame(self.aoe_frame,padx=10,pady=10)
		self.aoe_result_text_label = Label(
			self.aoe_result_text_frame,
			text=u'Результирующий текст',
			font=self.default_font)
		self.aoe_result_text = Text(
			self.aoe_result_text_frame,
			width=45,
			font=self.default_font,
			height=30
			)
		self.aoe_result_text.tag_config('replaced',background='yellow')

		# create frame for replacements
		# each replacement contain one entry for primary value,
		# one label and one combobox for new value
		# separators split replacements in frame
		self.aoe_primary_values_for_replacement = []
		self.aoe_labels_for_replacement = []
		self.aoe_new_values_for_replacement = []
		self.aoe_replacements_separators = []
		self.aoe_num_of_replacements = 0
		# create canvas for scrollable frame
		self.aoe_replacement_canvas = Canvas(
			self.aoe_frame,
			borderwidth=0,
			width=290,
			height=500)
		self.aoe_replacement_frame = Frame(self.aoe_replacement_canvas,padx=10,pady=10)
		self.aoe_replacement_frameScroll = Scrollbar(
			self.aoe_frame,
			orient=VERTICAL,
			command=self.aoe_replacement_canvas.yview)
		self.aoe_add_replacement_button = Button(
			self.aoe_replacement_frame,
			text=u'Добавить замену',
			font=self.small_font,
			width=22,
			height=1
			)
		# some magic for create scrollable frame with replacement fields
		self.aoe_replacement_canvas.configure(yscrollcommand=self.aoe_replacement_frameScroll.set)
		self.aoe_replacement_canvas.create_window((4,4), window=self.aoe_replacement_frame, anchor=CENTER, 
								  tags="self.aoe_replacement_frame")

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
		self.load_base()
		self.load_act_of_transfer()
		self.load_return_act()
		self.load_act_of_elimination()

	def pack_all(self):

		'''Pack all elements in a window'''
		self.info_frame.pack(side=TOP,fill=X)
		self.base_file_label.grid(row=1,column=0,sticky=W+N)
		self.base_file_label_text.grid(row=1,column=1,sticky=W+N)
		self.base_file_change_button.grid(row=1,column=2,sticky=W+N)
		if self.base_file:
			self.create_aot_check.place(relx=0.45,rely=0)
			self.create_ra_check.place(relx=0.55,rely=0)
			self.create_aoe_check.place(relx=0.65,rely=0)
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

		self.aot_frame_plug.pack_forget()
		self.aot_plain_text_frame.pack_forget()
		self.aot_plain_text.pack_forget()
		self.aot_result_text_frame.pack_forget()
		self.aot_result_text.pack_forget()
		self.aot_replacement_canvas.pack_forget()
		self.aot_replacement_frameScroll.pack_forget()
		for index in range(self.aot_num_of_replacements):
			self.aot_primary_values_for_replacement[index].grid_forget()
			self.aot_labels_for_replacement[index].grid_forget()
			self.aot_new_values_for_replacement[index].grid_forget()
		for replacement_separator in self.aot_replacements_separators:
			replacement_separator.grid_forget()
		self.aot_add_replacement_button.grid_forget()

		self.ra_frame_plug.pack_forget()
		self.ra_plain_text_frame.pack_forget()
		self.ra_plain_text.pack_forget()
		self.ra_result_text_frame.pack_forget()
		self.ra_result_text.pack_forget()
		self.ra_replacement_canvas.pack_forget()
		self.ra_replacement_frameScroll.pack_forget()
		for index in range(self.ra_num_of_replacements):
			self.ra_primary_values_for_replacement[index].grid_forget()
			self.ra_labels_for_replacement[index].grid_forget()
			self.ra_new_values_for_replacement[index].grid_forget()
		for replacement_separator in self.ra_replacements_separators:
			replacement_separator.grid_forget()
		self.ra_add_replacement_button.grid_forget()

		self.aoe_frame_plug.pack_forget()
		self.aoe_plain_text_frame.pack_forget()
		self.aoe_plain_text.pack_forget()
		self.aoe_result_text_frame.pack_forget()
		self.aoe_result_text.pack_forget()
		self.aoe_replacement_canvas.pack_forget()
		self.aoe_replacement_frameScroll.pack_forget()
		for index in range(self.aoe_num_of_replacements):
			self.aoe_primary_values_for_replacement[index].grid_forget()
			self.aoe_labels_for_replacement[index].grid_forget()
			self.aoe_new_values_for_replacement[index].grid_forget()
		for replacement_separator in self.aoe_replacements_separators:
			replacement_separator.grid_forget()
		self.aoe_add_replacement_button.grid_forget()

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

		if self.act_of_transfer.get():
			self.aot_plain_text_frame.pack(side=LEFT,fill=Y)
			self.aot_plain_text_label.pack(side=TOP)
			self.aot_plain_text.pack(side=LEFT,fill=Y)
			self.aot_result_text_frame.pack(side=RIGHT,fill=Y)
			self.aot_result_text_label.pack(side=TOP)
			self.aot_result_text.pack(side=RIGHT,fill=Y)
			self.aot_replacement_frameScroll.pack(side=RIGHT, fill=Y)
			self.aot_replacement_canvas.pack(side=LEFT,fill=BOTH,expand=True)
			for index in range(self.aot_num_of_replacements):
				self.aot_primary_values_for_replacement[index].grid(row=index*4,column=0,pady=2,sticky=N)
				self.aot_labels_for_replacement[index].grid(row=index*4+1,column=0,pady=2,sticky=N)
				self.aot_new_values_for_replacement[index].grid(row=index*4+2,column=0,pady=2,sticky=N)
			for index, replacement_separator in enumerate(self.aot_replacements_separators):
				replacement_separator.grid(row=index*4+3,column=0,padx=10,pady=5,sticky=W+E)
			self.aot_add_replacement_button.grid(row=self.aot_num_of_replacements*5,column=0,pady=10,sticky=N)
		else:
			self.aot_frame_plug.pack(side=TOP,fill=BOTH,expand=True)

		if self.return_act.get():
			self.ra_plain_text_frame.pack(side=LEFT,fill=Y)
			self.ra_plain_text_label.pack(side=TOP)
			self.ra_plain_text.pack(side=LEFT,fill=Y)
			self.ra_result_text_frame.pack(side=RIGHT,fill=Y)
			self.ra_result_text_label.pack(side=TOP)
			self.ra_result_text.pack(side=RIGHT,fill=Y)
			self.ra_replacement_frameScroll.pack(side=RIGHT, fill=Y)
			self.ra_replacement_canvas.pack(side=LEFT,fill=BOTH,expand=True)
			for index in range(self.ra_num_of_replacements):
				self.ra_primary_values_for_replacement[index].grid(row=index*4,column=0,pady=2,sticky=N)
				self.ra_labels_for_replacement[index].grid(row=index*4+1,column=0,pady=2,sticky=N)
				self.ra_new_values_for_replacement[index].grid(row=index*4+2,column=0,pady=2,sticky=N)
			for index, replacement_separator in enumerate(self.ra_replacements_separators):
				replacement_separator.grid(row=index*4+3,column=0,padx=10,pady=5,sticky=W+E)
			self.ra_add_replacement_button.grid(row=self.ra_num_of_replacements*5,column=0,pady=10,sticky=N)
		else:
			self.ra_frame_plug.pack(side=TOP,fill=BOTH,expand=True)

		if self.act_of_elimination.get():
			self.aoe_plain_text_frame.pack(side=LEFT,fill=Y)
			self.aoe_plain_text_label.pack(side=TOP)
			self.aoe_plain_text.pack(side=LEFT,fill=Y)
			self.aoe_result_text_frame.pack(side=RIGHT,fill=Y)
			self.aoe_result_text_label.pack(side=TOP)
			self.aoe_result_text.pack(side=RIGHT,fill=Y)
			self.aoe_replacement_frameScroll.pack(side=RIGHT, fill=Y)
			self.aoe_replacement_canvas.pack(side=LEFT,fill=BOTH,expand=True)
			for index in range(self.aoe_num_of_replacements):
				self.aoe_primary_values_for_replacement[index].grid(row=index*4,column=0,pady=2,sticky=N)
				self.aoe_labels_for_replacement[index].grid(row=index*4+1,column=0,pady=2,sticky=N)
				self.aoe_new_values_for_replacement[index].grid(row=index*4+2,column=0,pady=2,sticky=N)
			for index, replacement_separator in enumerate(self.aoe_replacements_separators):
				replacement_separator.grid(row=index*4+3,column=0,padx=10,pady=5,sticky=W+E)
			self.aoe_add_replacement_button.grid(row=self.aoe_num_of_replacements*5,column=0,pady=10,sticky=N)
		else:
			self.aoe_frame_plug.pack(side=TOP,fill=BOTH,expand=True)

		self.status_bar.pack(side=BOTTOM,fill=X)

	def bind_all(self):

		'''Bind all events in a window'''
		self.base_file_change_button.bind('<Button-1>',self.set_base_file)
		self.generate_button.bind('<Button-1>',self.generate_info)
		if self.base_file:
			self.base_table.bind("<<TreeviewSelect>>", self.setCurrentEntry)
			[entry_input.bind("<Return>",self.saveEntry) for entry_input in self.entry_inputs]
			self.add_entry_button.bind('<Button-1>',self.add_entry)
			self.del_entry_button.bind('<Button-1>',self.del_entry)
			self.save_base_button.bind('<Button-1>',self.save_base)
			self.root.bind('<Control-Return>',self.add_entry)
			self.root.bind('<Delete>',self.del_entry)
			self.root.bind('<Control-s>',self.save_base)

		if self.act_of_transfer:
			self.aot_add_replacement_button.bind('<Button-1>',self.aot_add_replacement)
			self.aot_replacement_frame.bind("<Configure>", self.aot_replaceFrameConfigure)		# bind mouse scroll
			[primary_val.bind("<Return>",self.aot_replace) for primary_val in self.aot_primary_values_for_replacement]
			[new_val.bind("<Return>",self.aot_replace) for new_val in self.aot_new_values_for_replacement]

		if self.return_act:
			self.ra_add_replacement_button.bind('<Button-1>',self.ra_add_replacement)
			self.ra_replacement_frame.bind("<Configure>", self.ra_replaceFrameConfigure)		# bind mouse scroll
			[primary_val.bind("<Return>",self.ra_replace) for primary_val in self.ra_primary_values_for_replacement]
			[new_val.bind("<Return>",self.ra_replace) for new_val in self.ra_new_values_for_replacement]

		if self.act_of_elimination:
			self.aoe_add_replacement_button.bind('<Button-1>',self.aoe_add_replacement)
			self.aoe_replacement_frame.bind("<Configure>", self.aoe_replaceFrameConfigure)		# bind mouse scroll
			[primary_val.bind("<Return>",self.aoe_replace) for primary_val in self.aoe_primary_values_for_replacement]
			[new_val.bind("<Return>",self.aoe_replace) for new_val in self.aoe_new_values_for_replacement]

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
		optionsWindow.update()

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
		
		self.base_file_label_text['text'] = self.base_file.get()
		self.num_of_entries = len(entries)
		self.num_of_fields = len(entries[0])
		self.base_table['columns'] = ['']*self.num_of_fields

		self.base_table.delete(*self.base_table.get_children())
		self.fill_table(entries)
		if self.act_of_transfer:
			self.get_replace_variants()

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

	def set_act_of_transfer(self,event=None):

		filename = askopenfilename(filetypes=(("Doc files", "*.doc;*.docx"),('All files','*.*')))
		if not filename:
			return

		self.act_of_transfer.set(filename)
		self.load_act_of_transfer()

	def load_act_of_transfer(self,event=None):

		if not self.act_of_transfer.get():
			return

		try:
			doc_text = get_doc_data(self.act_of_transfer.get())
		except Exception as e:
			showerror(u'Ошибка!',e)
			self.act_of_transfer.set(u'')
			return

		if not doc_text:
			showerror(u'Ошибка!',u'Пустой .doc файл.')
			self.act_of_transfer.set(u'')
			return

		# clear text fields and fill it
		self.aot_plain_text.delete(1.0,END)
		self.aot_result_text.delete(1.0,END)
		for line in doc_text:
			self.aot_plain_text.insert(END,line+'\n')
		self.aot_replace()			# fill result text field and add tags

		# delete replacements if exists and init new replacements
		self.aot_destroy_replacements()
		for _ in range(3):
			self.aot_add_replacement()

		self.status_bar['text'] = u'Документ загружен'

		self.pack_all()
		self.bind_all()

	def set_return_act(self,event=None):

		filename = askopenfilename(filetypes=(("Doc files", "*.doc;*.docx"),('All files','*.*')))
		if not filename:
			return

		self.return_act.set(filename)
		self.load_return_act()

	def load_return_act(self,event=None):

		if not self.return_act.get():
			return

		try:
			doc_text = get_doc_data(self.return_act.get())
		except Exception as e:
			showerror(u'Ошибка!',e)
			self.return_act.set(u'')
			return

		if not doc_text:
			showerror(u'Ошибка!',u'Пустой .doc файл.')
			self.return_act.set(u'')
			return

		# clear text fields and fill it
		self.ra_plain_text.delete(1.0,END)
		self.ra_result_text.delete(1.0,END)
		for line in doc_text:
			self.ra_plain_text.insert(END,line+'\n')
		self.ra_replace()			# fill result text field and add tags

		# delete replacements if exists and init new replacements
		self.ra_destroy_replacements()
		for _ in range(3):
			self.ra_add_replacement()

		self.status_bar['text'] = u'Документ загружен'

		self.pack_all()
		self.bind_all()

	def set_act_of_elimination(self,event=None):

		filename = askopenfilename(filetypes=(("Doc files", "*.doc;*.docx"),('All files','*.*')))
		if not filename:
			return

		self.act_of_elimination.set(filename)
		self.load_act_of_elimination()

	def load_act_of_elimination(self,event=None):

		if not self.act_of_elimination.get():
			return

		try:
			doc_text = get_doc_data(self.act_of_elimination.get())
		except Exception as e:
			showerror(u'Ошибка!',e)
			self.act_of_elimination.set(u'')
			return

		if not doc_text:
			showerror(u'Ошибка!',u'Пустой .doc файл.')
			self.act_of_elimination.set(u'')
			return

		# clear text fields and fill it
		self.aoe_plain_text.delete(1.0,END)
		self.aoe_result_text.delete(1.0,END)
		for line in doc_text:
			self.aoe_plain_text.insert(END,line+'\n')
		self.aoe_replace()			# fill result text field and add tags

		# delete replacements if exists and init new replacements
		self.aoe_destroy_replacements()
		for _ in range(3):
			self.aoe_add_replacement()

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

		'''Add fields from base for replace in choicefields'''

		for combobox in self.aot_new_values_for_replacement:
			replace_values = []
			if self.base_file:
				replace_values.extend([u'Столбец %r' % (index+1) for index in range(self.num_of_fields)])
			replace_values.append(u'…')
			combobox['values'] = replace_values

		for combobox in self.ra_new_values_for_replacement:
			replace_values = []
			if self.base_file:
				replace_values.extend([u'Столбец %r' % (index+1) for index in range(self.num_of_fields)])
			replace_values.append(u'…')
			combobox['values'] = replace_values

		for combobox in self.aoe_new_values_for_replacement:
			replace_values = []
			if self.base_file:
				replace_values.extend([u'Столбец %r' % (index+1) for index in range(self.num_of_fields)])
			replace_values.append(u'…')
			combobox['values'] = replace_values

	def aot_destroy_replacements(self):

		for primary_val in self.aot_primary_values_for_replacement:
			primary_val.destroy()
		for label in self.aot_labels_for_replacement:
			label.destroy()
		for new_val in self.aot_new_values_for_replacement:
			new_val.destroy()
		for separator in self.aot_replacements_separators:
			separator.destroy()

		self.aot_primary_values_for_replacement = []
		self.aot_labels_for_replacement = []
		self.aot_new_values_for_replacement = []
		self.aot_replacements_separators = []
		self.aot_num_of_replacements = 0

	def aot_add_replacement(self,event=None):

		self.aot_primary_values_for_replacement.append(
			Entry(
				self.aot_replacement_frame,
				width=30,
				font=self.default_font
				)
			)
		self.aot_labels_for_replacement.append(
			Label(
				self.aot_replacement_frame,
				text=u'Заменить на:',
				font=self.small_font
				)
			)
		self.aot_new_values_for_replacement.append(
			ttk.Combobox(
				self.aot_replacement_frame,
				width=35,
				font=self.small_font,
				values=[]
				)
			)
		if self.aot_num_of_replacements>0:
			self.aot_replacements_separators.append(
				ttk.Separator(
					self.aot_replacement_frame,
					orient=HORIZONTAL
					)
				)

		self.aot_num_of_replacements += 1
		self.get_replace_variants()

		self.pack_all()
		self.bind_all()

	def aot_replaceFrameConfigure(self,event=None):

		'''Reset the scroll region to encompass the inner frame'''

		self.aot_replacement_canvas.configure(scrollregion=self.aot_replacement_canvas.bbox("all"))

	def aot_get_result_value(self,index,index_row=0):

		num_of_column = self.aot_new_values_for_replacement[index].current()
		if num_of_column == -1 or num_of_column == self.num_of_fields:
			result_value = self.aot_new_values_for_replacement[index].get()
		elif self.base_file and self.base_table.selection():
			result_value = self.base_table.item(self.base_table.selection()[index_row])['values'][num_of_column]
		elif self.base_file:
			result_value = self.base_table.item(self.base_table.get_children()[-1])['values'][num_of_column]
		else:
			showerror(u'Ошибка!', u'Не удается получить данные для замены.')
			return
		return result_value

	def aot_replace(self,event=None):

		# remove tags
		for tag in self.aot_plain_text.tag_names():
			self.aot_plain_text.tag_remove(tag,1.0,END)
		for tag in self.aot_result_text.tag_names():
			self.aot_result_text.tag_remove(tag,1.0,END)
		# copy text from plain to result
		text = self.aot_plain_text.get(1.0,END)
		self.aot_result_text.delete(1.0,END)
		self.aot_result_text.insert(1.0,text)
		plain_lines = text.split('\n')
		for index_line,line in enumerate(plain_lines):

			if not line:
				continue

			start_of_line = float(index_line+1)
			end_of_line = '%r.end' % (index_line+1)

			for index in range(self.aot_num_of_replacements):
				plain = str(self.aot_primary_values_for_replacement[index].get())
				if not plain:
					continue
				if plain not in line:
					continue
				result_value = str(self.aot_get_result_value(index))
				line = line.replace(plain,result_value)

				# rewrite line on new replaced line
				self.aot_result_text.delete(start_of_line,end_of_line)
				self.aot_result_text.insert(start_of_line,line)

				# add tags to plain text
				search_start = 1.0
				while True:
					start_of_tag = self.aot_plain_text.search(plain,search_start,stopindex=end_of_line)
					if not start_of_tag:
						break
					end_of_tag = start_of_tag.split('.')[0]+'.'+str(int(start_of_tag.split('.')[1])+len(plain))
					self.aot_plain_text.tag_add('replaced',start_of_tag,end_of_tag)
					search_start = end_of_tag

				# add tags to result text
				if not result_value:
					continue
				search_start = 1.0
				while True:
					start_of_tag = self.aot_result_text.search(result_value,search_start,stopindex=end_of_line)
					if not start_of_tag:
						break
					end_of_tag = start_of_tag.split('.')[0]+'.'+str(int(start_of_tag.split('.')[1])+len(result_value))
					self.aot_result_text.tag_add('replaced',start_of_tag,end_of_tag)
					search_start = end_of_tag

	def ra_destroy_replacements(self):

		for primary_val in self.ra_primary_values_for_replacement:
			primary_val.destroy()
		for label in self.ra_labels_for_replacement:
			label.destroy()
		for new_val in self.ra_new_values_for_replacement:
			new_val.destroy()
		for separator in self.ra_replacements_separators:
			separator.destroy()

		self.ra_primary_values_for_replacement = []
		self.ra_labels_for_replacement = []
		self.ra_new_values_for_replacement = []
		self.ra_replacements_separators = []
		self.ra_num_of_replacements = 0

	def ra_add_replacement(self,event=None):

		self.ra_primary_values_for_replacement.append(
			Entry(
				self.ra_replacement_frame,
				width=30,
				font=self.default_font
				)
			)
		self.ra_labels_for_replacement.append(
			Label(
				self.ra_replacement_frame,
				text=u'Заменить на:',
				font=self.small_font
				)
			)
		self.ra_new_values_for_replacement.append(
			ttk.Combobox(
				self.ra_replacement_frame,
				width=35,
				font=self.small_font,
				values=[]
				)
			)
		if self.ra_num_of_replacements>0:
			self.ra_replacements_separators.append(
				ttk.Separator(
					self.ra_replacement_frame,
					orient=HORIZONTAL
					)
				)

		self.ra_num_of_replacements += 1
		self.get_replace_variants()

		self.pack_all()
		self.bind_all()

	def ra_replaceFrameConfigure(self,event=None):

		'''Reset the scroll region to encompass the inner frame'''

		self.ra_replacement_canvas.configure(scrollregion=self.ra_replacement_canvas.bbox("all"))

	def ra_get_result_value(self,index,index_row=0):

		num_of_column = self.ra_new_values_for_replacement[index].current()
		if num_of_column == -1 or num_of_column == self.num_of_fields:
			result_value = self.ra_new_values_for_replacement[index].get()
		elif self.base_file and self.base_table.selection():
			result_value = self.base_table.item(self.base_table.selection()[index_row])['values'][num_of_column]
		elif self.base_file:
			result_value = self.base_table.item(self.base_table.get_children()[-1])['values'][num_of_column]
		else:
			showerror(u'Ошибка!', u'Не удается получить данные для замены.')
			return
		return result_value

	def ra_replace(self,event=None):

		# remove tags
		for tag in self.ra_plain_text.tag_names():
			self.ra_plain_text.tag_remove(tag,1.0,END)
		for tag in self.ra_result_text.tag_names():
			self.ra_result_text.tag_remove(tag,1.0,END)
		# copy text from plain to result
		text = self.ra_plain_text.get(1.0,END)
		self.ra_result_text.delete(1.0,END)
		self.ra_result_text.insert(1.0,text)
		plain_lines = text.split('\n')
		for index_line,line in enumerate(plain_lines):

			if not line:
				continue

			start_of_line = float(index_line+1)
			end_of_line = '%r.end' % (index_line+1)

			for index in range(self.ra_num_of_replacements):
				plain = str(self.ra_primary_values_for_replacement[index].get())
				if not plain:
					continue
				if plain not in line:
					continue
				result_value = str(self.ra_get_result_value(index))
				line = line.replace(plain,result_value)

				# rewrite line on new replaced line
				self.ra_result_text.delete(start_of_line,end_of_line)
				self.ra_result_text.insert(start_of_line,line)

				# add tags to plain text
				search_start = 1.0
				while True:
					start_of_tag = self.ra_plain_text.search(plain,search_start,stopindex=end_of_line)
					if not start_of_tag:
						break
					end_of_tag = start_of_tag.split('.')[0]+'.'+str(int(start_of_tag.split('.')[1])+len(plain))
					self.ra_plain_text.tag_add('replaced',start_of_tag,end_of_tag)
					search_start = end_of_tag

				# add tags to result text
				if not result_value:
					continue
				search_start = 1.0
				while True:
					start_of_tag = self.ra_result_text.search(result_value,search_start,stopindex=end_of_line)
					if not start_of_tag:
						break
					end_of_tag = start_of_tag.split('.')[0]+'.'+str(int(start_of_tag.split('.')[1])+len(result_value))
					self.ra_result_text.tag_add('replaced',start_of_tag,end_of_tag)
					search_start = end_of_tag

	def aoe_destroy_replacements(self):

		for primary_val in self.aoe_primary_values_for_replacement:
			primary_val.destroy()
		for label in self.aoe_labels_for_replacement:
			label.destroy()
		for new_val in self.aoe_new_values_for_replacement:
			new_val.destroy()
		for separator in self.aoe_replacements_separators:
			separator.destroy()

		self.aoe_primary_values_for_replacement = []
		self.aoe_labels_for_replacement = []
		self.aoe_new_values_for_replacement = []
		self.aoe_replacements_separators = []
		self.aoe_num_of_replacements = 0

	def aoe_add_replacement(self,event=None):

		self.aoe_primary_values_for_replacement.append(
			Entry(
				self.aoe_replacement_frame,
				width=30,
				font=self.default_font
				)
			)
		self.aoe_labels_for_replacement.append(
			Label(
				self.aoe_replacement_frame,
				text=u'Заменить на:',
				font=self.small_font
				)
			)
		self.aoe_new_values_for_replacement.append(
			ttk.Combobox(
				self.aoe_replacement_frame,
				width=35,
				font=self.small_font,
				values=[]
				)
			)
		if self.aoe_num_of_replacements>0:
			self.aoe_replacements_separators.append(
				ttk.Separator(
					self.aoe_replacement_frame,
					orient=HORIZONTAL
					)
				)

		self.aoe_num_of_replacements += 1
		self.get_replace_variants()

		self.pack_all()
		self.bind_all()

	def aoe_replaceFrameConfigure(self,event=None):

		'''Reset the scroll region to encompass the inner frame'''

		self.aoe_replacement_canvas.configure(scrollregion=self.aoe_replacement_canvas.bbox("all"))

	def aoe_get_result_value(self,index,index_row=0):

		num_of_column = self.aoe_new_values_for_replacement[index].current()
		if num_of_column == -1 or num_of_column == self.num_of_fields:
			result_value = self.aoe_new_values_for_replacement[index].get()
		elif self.base_file and self.base_table.selection():
			result_value = self.base_table.item(self.base_table.selection()[index_row])['values'][num_of_column]
		elif self.base_file:
			result_value = self.base_table.item(self.base_table.get_children()[-1])['values'][num_of_column]
		else:
			showerror(u'Ошибка!', u'Не удается получить данные для замены.')
			return
		return result_value

	def aoe_replace(self,event=None):

		# remove tags
		for tag in self.aoe_plain_text.tag_names():
			self.aoe_plain_text.tag_remove(tag,1.0,END)
		for tag in self.aoe_result_text.tag_names():
			self.aoe_result_text.tag_remove(tag,1.0,END)
		# copy text from plain to result
		text = self.aoe_plain_text.get(1.0,END)
		self.aoe_result_text.delete(1.0,END)
		self.aoe_result_text.insert(1.0,text)
		plain_lines = text.split('\n')
		for index_line,line in enumerate(plain_lines):

			if not line:
				continue

			start_of_line = float(index_line+1)
			end_of_line = '%r.end' % (index_line+1)

			for index in range(self.aoe_num_of_replacements):
				plain = str(self.aoe_primary_values_for_replacement[index].get())
				if not plain:
					continue
				if plain not in line:
					continue
				result_value = str(self.aoe_get_result_value(index))
				line = line.replace(plain,result_value)

				# rewrite line on new replaced line
				self.aoe_result_text.delete(start_of_line,end_of_line)
				self.aoe_result_text.insert(start_of_line,line)

				# add tags to plain text
				search_start = 1.0
				while True:
					start_of_tag = self.aoe_plain_text.search(plain,search_start,stopindex=end_of_line)
					if not start_of_tag:
						break
					end_of_tag = start_of_tag.split('.')[0]+'.'+str(int(start_of_tag.split('.')[1])+len(plain))
					self.aoe_plain_text.tag_add('replaced',start_of_tag,end_of_tag)
					search_start = end_of_tag

				# add tags to result text
				if not result_value:
					continue
				search_start = 1.0
				while True:
					start_of_tag = self.aoe_result_text.search(result_value,search_start,stopindex=end_of_line)
					if not start_of_tag:
						break
					end_of_tag = start_of_tag.split('.')[0]+'.'+str(int(start_of_tag.split('.')[1])+len(result_value))
					self.aoe_result_text.tag_add('replaced',start_of_tag,end_of_tag)
					search_start = end_of_tag

	def generate_info(self,event=None):

		'''Show this window before generate'''

		if not self.base_file:
			return

		if not self.base_table.selection():
			return

		if not self.create_aot.get() and not self.create_ra.get() and not self.create_aoe.get():
			return

		generateInfoWindow = Toplevel(self.root)
		generateInfoWindow.geometry('600x400')
		generateInfoWindow.title(u'Настройки')
		generateInfoWindow.update()

		window = GenerateInfoWindow(self,generateInfoWindow)

	def generate_acts(self,event=None):

		'''Generate new word documents with replaced values from base'''

		if not self.base_file:
			return

		if not self.base_table.selection():
			return

		# get data from table
		entries = [self.base_table.item(selitem)['values'] for selitem in self.base_table.selection()]
		# get replacements for selected rows in base_table and generate new files
		for index_row,entry in enumerate(entries):
			if self.create_aot.get():
				replacements = {}
				for index in range(self.aot_num_of_replacements):
					primary_val = self.aot_primary_values_for_replacement[index].get()
					if not primary_val:
						continue
					result_val = self.aot_get_result_value(index,index_row)
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
				for index in range(self.ra_num_of_replacements):
					primary_val = self.ra_primary_values_for_replacement[index].get()
					if not primary_val:
						continue
					result_val = self.ra_get_result_value(index,index_row)
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
				for index in range(self.aoe_num_of_replacements):
					primary_val = self.aoe_primary_values_for_replacement[index].get()
					if not primary_val:
						continue
					result_val = self.aoe_get_result_value(index,index_row)
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