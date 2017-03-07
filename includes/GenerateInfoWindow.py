# -*- coding: utf-8 -*-
from tkinter.filedialog import *
from tkinter.messagebox import *
from tkinter.font import Font
from tkinter import *
import tkinter.ttk as ttk
from pprint import pprint

class OptionsWindow:

	def __init__(self, mainWindow, root):

		self.mainWindow = mainWindow
		self.root = root
		self.small_font = Font(family="Helvetica",size=10)

		self.entries_selected = IntVar()
		self.act_types = StringVar()
		self.total_create = IntVar()
		self.aot_filename = StringVar()
		self.ra_filename = StringVar()
		self.aoe_filename = StringVar()
		self.destination_folder = StringVar()

		self.entries_selected.set(len(self.mainWindow.base_table.selection()))
		self.act_types.set()
		self.act_of_elimination.set(self.mainWindow.act_of_elimination.get())
		self.destination_folder.set(self.mainWindow.destination_folder.get())

		self.options_frame = Frame(self.root,padx=10,pady=10)
		self.act_of_transfer_label = Label(
				self.options_frame,
				text = u"Акт передачи: ",
				padx=20,
				pady=5,
				anchor = "nw",
				font=self.small_font
				)
		self.act_of_transfer_label_text = Label(
				self.options_frame,
				textvariable=self.act_of_transfer,
				padx=20,
				pady=5,
				anchor = "ne",
				font=self.small_font
				)

		self.action_frame = Frame(self.root)
		self.apply_changes_button = Button(
			self.root,
			text = u'ОК',
			font = self.small_font
			)
		self.cancel_changes_button = Button(
			self.root,
			text = u'Отмена',
			font = self.small_font
			)

		self.pack_all()
		self.bind_all()

	def pack_all(self):

		self.options_frame.pack(side=TOP,fill=X)
		self.act_of_transfer_label.grid(row=0,column=0,sticky=W+N)
		self.act_of_transfer_label_text.grid(row=0,column=1,sticky=N+E)
		self.act_of_transfer_button.grid(row=0,column=2,sticky=N+E)


		self.apply_changes_button.place(relx=0.3,rely=0.9,width=60,height=30)
		self.cancel_changes_button.place(relx=0.6,rely=0.9,width=80,height=30)

	def bind_all(self):

		self.apply_changes_button.bind('<Button-1>',self.exit_with_save)
		self.cancel_changes_button.bind('<Button-1>',self.exit)

	def exit_with_save(self,event=None):

		self.save_changes()
		self.mainWindow.status_bar['text'] = u'Изменения сохранены'
		self.exit()

	def exit(self,event=None):

		self.root.destroy()

	def save_changes(self):

		pass
