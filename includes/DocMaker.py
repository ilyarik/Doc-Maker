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
from .BaseTabFrame import BaseTabFrame
from .StatsTabFrame import StatsTabFrame
from .BaseFileEventHandler import BaseFileEventHandler
from watchdog.observers import Observer

class DocMaker(Tk):

	def __init__(self,root_dir):

		Tk.__init__(self)

		self.geometry('1200x800+50+10')
		self.minsize(1200,800)
		self.title(u'Составитель актов 2000')
		self.update()

		self.root_dir = root_dir
		self.configsFileName = u'%s\\USER\\configs.ini' % (self.root_dir)
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
		# num of column from base for statistic
		self.statsColIndex = IntVar()

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
		self.notebook = ttk.Notebook(self)
		self.base_frame = BaseTabFrame(self,title_text=u'Здесь отображается база записей СКЗИ')
		self.stats_frame = StatsTabFrame(self,title_text=u'Здесь отображается статистика по выбранному столбцу')
		self.aot_frame = ReplacementsTabFrame(
			self,
			act_name='aot',
			act_var=self.act_of_transfer,
			plug_text=u'Место для акта передачи',
			title_text=u'Настройки для генерации акта передачи (Enter в поле для подстановки)'
			)
		self.ra_frame = ReplacementsTabFrame(
			self,
			act_name='ra',
			act_var=self.return_act,
			plug_text=u'Место для акта возврата',
			title_text=u'Настройки для генерации акта возврата (Enter в поле для подстановки)'
			)
		self.aoe_frame = ReplacementsTabFrame(
			self,
			act_name='aoe',
			act_var=self.act_of_elimination,
			plug_text=u'Место для акта уничтожения',
			title_text=u'Настройки для генерации акта уничтожения (Enter в поле для подстановки)'
			)

		self.notebook.add(self.base_frame, text=u'База')
		self.notebook.add(self.stats_frame, text=u'Статистика')
		self.notebook.add(self.aot_frame, text=u'Акт передачи')
		self.notebook.add(self.ra_frame, text=u'Акт возврата')
		self.notebook.add(self.aoe_frame, text=u'Акт уничтожения')

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

		# init syncronization manager
		self.baseSyncObserver = Observer()

		# read options from .ini file
		self.read_options()
		self.base_frame.initBase()
		
		# start syncronization manager
		self.startSyncManager()

		self.loadStats()
		self.aot_frame.load_act()
		self.ra_frame.load_act()
		self.aoe_frame.load_act()
		self.pack_all()
		self.bind_all()

	def pack_all(self):

		'''Pack all elements in a window'''
		self.info_frame.pack_forget()
		self.base_file_label.grid_forget()
		self.base_file_label_text.grid_forget()
		self.base_file_change_button.grid_forget()
		self.create_aot_check.place_forget()
		self.create_ra_check.place_forget()
		self.create_aoe_check.place_forget()
		self.generate_button.place_forget()

		self.info_frame.pack(side=TOP,fill=X)
		self.base_file_label.grid(row=1,column=0,sticky=W+N)
		self.base_file_label_text.grid(row=1,column=1,sticky=W+N)
		self.base_file_change_button.grid(row=1,column=2,sticky=W+N)
		
		if self.base_file.get():
			self.create_aot_check.place(relx=0.5,rely=0.2)
			self.create_ra_check.place(relx=0.6,rely=0.2)
			self.create_aoe_check.place(relx=0.7,rely=0.2)
			self.generate_button.place(relx=0.85,rely=0,width=130,height=40)
		self.notebook.pack(side=TOP,fill=BOTH,expand=True,padx=10,pady=10)

		self.base_frame.pack_all()
		self.stats_frame.pack_all()
		self.aot_frame.pack_all()
		self.ra_frame.pack_all()
		self.aoe_frame.pack_all()

		self.status_bar.pack(side=BOTTOM,fill=X)

	def bind_all(self):

		'''Bind all events in a window'''
		self.base_file_change_button.bind('<Button-1>',self.set_base_file)
		self.generate_button.bind('<Button-1>',self.generate_info)

		self.base_frame.bind_all()
		self.bind('<Control-Return>',self.base_frame.addEntry)

		self.stats_frame.bind_all()
		self.aot_frame.bind_all()
		self.ra_frame.bind_all()
		self.aoe_frame.bind_all()

	def read_options(self):

		'''Read options from .ini file'''
		configs = configparser.ConfigParser()
		configs.read(self.configsFileName)

		self.base_file.set(configs['Base']['filename'])
		self.act_of_transfer.set(configs['Act_of_transfer']['filename'])
		self.return_act.set(configs['Return_act']['filename'])
		self.act_of_elimination.set(configs['Act_of_elimination']['filename'])
		self.destination_folder.set(configs['DEFAULT']['destination_folder'])


	def write_options(self):

		configs = configparser.ConfigParser()
		configs.read(self.configsFileName)
		configs['Base']['filename'] = self.base_file.get()
		configs['DEFAULT']['destination_folder'] = self.destination_folder.get()
		configs['Act_of_transfer']['filename'] = self.act_of_transfer.get()
		configs['Return_act']['filename'] = self.return_act.get()
		configs['Act_of_elimination']['filename'] = self.act_of_elimination.get()
		
		with open(self.configsFileName, 'w') as configfile:
			configs.write(configfile)
			configfile.close()

	def exit(self):

		self.quit()

	def call_options_window(self):

		window = OptionsWindow(self)

	def startSyncManager(self):

		'''Create observer which will see modifications in base file and starts his work'''
		patterns = [self.base_file.get()]
		_dir = os.path.split(self.base_file.get())[0]
		handler = BaseFileEventHandler(patterns,self)
		
		self.baseSyncObserver.schedule(handler,_dir)
		self.baseSyncObserver.start()

	def loadStats(self):

		'''Take data from base frame and put to statistic frame'''
		if not self.base_file.get():
			return
		self.stats_frame.readOptions()
		data_list = self.base_frame.getColumnAsList(col_index=self.stats_frame.colIndex.get()-1)
		self.stats_frame.getStatsFromList(data_list=data_list)
		self.stats_frame.initTable()
		self.stats_frame.fillTable()
		self.stats_frame.initFigure()
		self.stats_frame.changeColumnChoices(self.base_frame.num_of_fields)
		self.stats_frame.pack_all()

	def set_base_file(self,event=None):

		filename = askopenfilename(filetypes=(("XLS files", "*.xls;*.xlsx"),('All files','*.*')))
		if not filename:
			return

		self.base_file.set(filename)
		self.write_options()
		self.base_frame.initBase()
		self.loadStats()

	def generate_info(self,event=None):

		'''Show this window before generate'''
		if not self.base_file.get():
			return

		if not self.base_frame.base_table.selection():
			return

		if not self.create_aot.get() and not self.create_ra.get() and not self.create_aoe.get():
			return

		window = GenerateInfoWindow(self)

	def generate_acts(self,event=None):

		'''Generate new word documents with replaced values from base'''

		if not self.base_file.get():
			return

		if not self.base_frame.base_table.selection():
			return

		# get data from table
		entries = [self.base_frame.base_table.item(selitem)['values'] for selitem in self.base_frame.base_table.selection()]
		if [entry for entry in entries if not all(entry)]:
			showerror(u'Ошибка.', u'Есть незаполненные значения в строке.')
			return
		# get replacements for selected rows in base_table and generate new files
		for index_row,entry in enumerate(entries):
			if self.create_aot.get():
				replacements = {}
				for index in range(self.aot_frame.num_of_replacements):
					primary_val = self.aot_frame.primary_values_for_replacement[index].get()
					if not primary_val:
						continue
					result_val = self.aot_frame.get_result_value(index,index_row)
					if 'ФИО' in primary_val:
						try:
							surname, name, patronymic = result_val.split()
							result_val = '%s %s. %s.' % (surname,name[0],patronymic[0])
						except Exception as e:
							showerror(u'Ошибка.',u'Не могу разделить ФИО на фамилию и инициалы для %s.\n%s' % (primary_val,e))
							return
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
					if 'ФИО' in primary_val:
						try:
							surname, name, patronymic = result_val.split()
							result_val = '%s %s. %s.' % (surname,name[0],patronymic[0])
						except Exception as e:
							showerror(u'Ошибка.',u'Не могу разделить ФИО на фамилию и инициалы для %s.\n%s' % (primary_val,e))
							return
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
					if 'ФИО' in primary_val:
						try:
							surname, name, patronymic = result_val.split()
							result_val = '%s %s. %s.' % (surname,name[0],patronymic[0])
						except Exception as e:
							showerror(u'Ошибка.',u'Не могу разделить ФИО на фамилию и инициалы для %s.\n%s' % (primary_val,e))
							return
					replacements.update({primary_val:str(result_val)})
				doc_filename = self.destination_folder.get() + '/' + u'Акт %s №%s-%s-%s.docx' % (
								u'Уничтожения',
								str(datetime.datetime.now().year),
								str(entry[0]),
								u'У'
								)
				create_new_replaced_doc(self.act_of_elimination.get(), doc_filename, replacements)

		showinfo(u'Успех',u'Выполнено')
		self.base_frame.sync_exist_acts()
		self.status_bar['text'] = u'Готово'
