# -*- coding: utf-8 -*-
from tkinter.filedialog import *
from tkinter.messagebox import *
from tkinter.font import Font
from tkinter import *
import tkinter.ttk as ttk
from pprint import pprint
import datetime
from .functions import *

class GenerateInfoWindow(Toplevel):

	def __init__(self, mainWindow):

		Toplevel.__init__(self)
		self.geometry('600x400')
		self.title(u'Генерация')
		self.update()

		self.mainWindow = mainWindow
		self.small_font = Font(family="Helvetica",size=10)

		self.entries_selected = IntVar()
		self.act_types = StringVar()
		self.total_create = IntVar()
		self.example_filename = StringVar()
		self.destination_folder = StringVar()

		self.entries_selected.set(len(self.mainWindow.base_table.selection()))
		self.act_types.set(self.get_act_types_strings())
		self.total_create.set(self.get_total_create())
		self.example_filename.set(self.get_example_filename())
		self.destination_folder.set(self.mainWindow.destination_folder.get())

		self.options_frame = Frame(self,padx=10,pady=10)
		self.act_types_label = Label(
				self.options_frame,
				text = u"Типы актов: ",
				padx=20,
				pady=5,
				anchor = "nw",
				font=self.small_font
				)
		self.act_types_label_text = Label(
				self.options_frame,
				textvariable=self.act_types,
				padx=20,
				pady=5,
				anchor = "ne",
				font=self.small_font
				)
		self.total_create_label = Label(
				self.options_frame,
				text = u"Всего будет создано документов: ",
				padx=20,
				pady=5,
				anchor = "nw",
				font=self.small_font
				)
		self.total_create_label_text = Label(
				self.options_frame,
				textvariable=self.total_create,
				padx=20,
				pady=5,
				anchor = "ne",
				font=self.small_font
				)
		self.example_filename_label = Label(
				self.options_frame,
				text = u"Имя файла (пример): ",
				padx=20,
				pady=5,
				anchor = "nw",
				font=self.small_font
				)
		self.example_filename_label_text = Label(
				self.options_frame,
				textvariable=self.example_filename,
				padx=20,
				pady=5,
				anchor = "ne",
				font=self.small_font
				)
		self.destination_folder_label = Label(
				self.options_frame,
				text = u"Папка с результатами: ",
				padx=20,
				pady=5,
				anchor = "nw",
				font=self.small_font
				)
		self.destination_folder_label_text = Label(
				self.options_frame,
				text=get_truncated_line(self.destination_folder.get(),60),
				padx=20,
				pady=5,
				anchor = "ne",
				font=self.small_font
				)

		self.apply_changes_button = Button(
			self,
			text = u'ОК',
			font = self.small_font
			)
		self.cancel_changes_button = Button(
			self,
			text = u'Отмена',
			font = self.small_font
			)

		self.pack_all()
		self.bind_all()

	def pack_all(self):

		self.options_frame.pack(side=TOP,fill=X)
		self.act_types_label.grid(row=0,column=0,sticky=W+N)
		self.act_types_label_text.grid(row=0,column=1,sticky=N+E)
		self.total_create_label.grid(row=1,column=0,sticky=W+N)
		self.total_create_label_text.grid(row=1,column=1,sticky=N+E)
		self.example_filename_label.grid(row=2,column=0,sticky=W+N)
		self.example_filename_label_text.grid(row=2,column=1,sticky=N+E)
		self.destination_folder_label.grid(row=3,column=0,sticky=W+N)
		self.destination_folder_label_text.grid(row=3,column=1,sticky=N+E)

		self.apply_changes_button.place(relx=0.3,rely=0.9,width=60,height=30)
		self.cancel_changes_button.place(relx=0.6,rely=0.9,width=80,height=30)

	def bind_all(self):

		self.apply_changes_button.bind('<Button-1>',self.ok)
		self.cancel_changes_button.bind('<Button-1>',self.exit)

	def ok(self,event=None):

		self.mainWindow.generate_acts()
		self.exit()

	def exit(self,event=None):

		self.destroy()

	def get_act_types_strings(self):

		create_acts = (
			(self.mainWindow.create_aot,u'Передачи'),
			(self.mainWindow.create_ra,u'Возврата'),
			(self.mainWindow.create_aoe,u'Уничтожения')
			)
		return ', '.join([create_act[1] for create_act in create_acts if create_act[0].get()])

	def get_total_create(self):

		return len(
			self.mainWindow.base_table.selection()
			)*sum(
				[self.mainWindow.create_aot.get(),
				self.mainWindow.create_ra.get(),
				self.mainWindow.create_aoe.get()
				]
			)

	def get_example_filename(self):

		act_type = None
		if self.mainWindow.create_aot.get():
			act_type = u'Передачи'
		elif self.mainWindow.create_ra.get():
			act_type = u'Возврата'
		elif self.mainWindow.create_aoe.get():
			act_type = u'Уничтожения'

		return u'Акт %s №%s-%s-%s.docx' % (
			act_type,
			str(datetime.datetime.now().year),
			str(self.mainWindow.base_table.item(self.mainWindow.base_table.selection()[0])['values'][0]),
			str(act_type[0])
			)