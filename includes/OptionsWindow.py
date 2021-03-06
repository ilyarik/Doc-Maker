# -*- coding: utf-8 -*-
from tkinter.filedialog import *
from tkinter.messagebox import *
from tkinter.font import Font
from tkinter import *
import tkinter.ttk as ttk
from .functions import *

class OptionsWindow(Toplevel):

	def __init__(self, mainWindow):

		Toplevel.__init__(self)
		self.geometry('700x500')
		self.title(u'Настройки')
		self.update()

		self.mainWindow = mainWindow
		self.pathesLength = 60

		self.act_of_transfer = StringVar()
		self.return_act = StringVar()
		self.act_of_elimination = StringVar()
		self.destination_folder = StringVar()

		self.act_of_transfer.set(self.mainWindow.act_of_transfer.get())
		self.return_act.set(self.mainWindow.return_act.get())
		self.act_of_elimination.set(self.mainWindow.act_of_elimination.get())
		self.destination_folder.set(self.mainWindow.destination_folder.get())

		self.options_frame = Frame(self,padx=10,pady=10)
		self.act_of_transfer_label = Label(
				self.options_frame,
				text = u"Акт передачи: ",
				padx=20,
				pady=5,
				anchor = "nw",
				font=self.mainWindow.small_font
				)
		self.act_of_transfer_label_text = Label(
				self.options_frame,
				text=get_truncated_line(self.act_of_transfer.get(),self.pathesLength),
				padx=20,
				pady=5,
				anchor = "ne",
				font=self.mainWindow.small_font
				)
		self.act_of_transfer_button = Button(
				self.options_frame,
				text = u'…',
				pady=5,
				font=self.mainWindow.small_font
				)

		self.return_act_label = Label(
				self.options_frame,
				text = u"Акт возврата: ",
				padx=20,
				pady=5,
				anchor = "nw",
				font=self.mainWindow.small_font
				)
		self.return_act_label_text = Label(
				self.options_frame,
				text=get_truncated_line(self.return_act.get(),self.pathesLength),
				padx=20,
				pady=5,
				anchor = "ne",
				font=self.mainWindow.small_font
				)
		self.return_act_button = Button(
				self.options_frame,
				text = u'…',
				pady=5,
				font=self.mainWindow.small_font
				)

		self.act_of_elimination_label = Label(
				self.options_frame,
				text = u"Акт уничтожения: ",
				padx=20,
				pady=5,
				anchor = "nw",
				font=self.mainWindow.small_font
				)
		self.act_of_elimination_label_text = Label(
				self.options_frame,
				text=get_truncated_line(self.act_of_elimination.get(),self.pathesLength),
				padx=20,
				pady=5,
				anchor = "ne",
				font=self.mainWindow.small_font
				)
		self.act_of_elimination_button = Button(
				self.options_frame,
				text = u'…',
				pady=5,
				font=self.mainWindow.small_font
				)

		self.destination_folder_label = Label(
				self.options_frame,
				text = u"Папка с актами: ",
				padx=20,
				pady=5,
				anchor = "nw",
				font=self.mainWindow.small_font
				)
		self.destination_folder_label_text = Label(
				self.options_frame,
				text=get_truncated_line(self.destination_folder.get(),self.pathesLength),
				padx=20,
				pady=5,
				anchor = "ne",
				font=self.mainWindow.small_font
				)
		self.destination_folder_button = Button(
				self.options_frame,
				text = u'…',
				pady=5,
				font=self.mainWindow.small_font
				)

		self.apply_changes_button = Button(
				self,
				text = u'ОК',
				font = self.mainWindow.small_font
				)
		self.cancel_changes_button = Button(
				self,
				text = u'Отмена',
				font = self.mainWindow.small_font
				)

		self.pack_all()
		self.bind_all()

	def pack_all(self):

		self.options_frame.pack(side=TOP,fill=X)
		self.act_of_transfer_label.grid(row=0,column=0,sticky=W+N)
		self.act_of_transfer_label_text.grid(row=0,column=1,sticky=N+E)
		self.act_of_transfer_button.grid(row=0,column=2,sticky=N+E)

		self.return_act_label.grid(row=1,column=0,sticky=W+N)
		self.return_act_label_text.grid(row=1,column=1,sticky=N+E)
		self.return_act_button.grid(row=1,column=2,sticky=N+E)

		self.act_of_elimination_label.grid(row=2,column=0,sticky=W+N)
		self.act_of_elimination_label_text.grid(row=2,column=1,sticky=N+E)
		self.act_of_elimination_button.grid(row=2,column=2,sticky=N+E)

		self.destination_folder_label.grid(row=3,column=0,sticky=W+N)
		self.destination_folder_label_text.grid(row=3,column=1,sticky=N+E)
		self.destination_folder_button.grid(row=3,column=2,sticky=N+E)

		self.apply_changes_button.place(relx=0.3,rely=0.9,width=60,height=30)
		self.cancel_changes_button.place(relx=0.6,rely=0.9,width=80,height=30)

	def bind_all(self):

		self.act_of_transfer_button.bind('<Button-1>',self.change_act_of_transfer)
		self.return_act_button.bind('<Button-1>',self.change_return_act)
		self.act_of_elimination_button.bind('<Button-1>',self.change_act_of_elimination)
		self.destination_folder_button.bind('<Button-1>',self.change_destination_folder)

		self.apply_changes_button.bind('<Button-1>',self.exit_with_save)
		self.cancel_changes_button.bind('<Button-1>',self.exit)

	def exit_with_save(self,event=None):

		self.save_changes()
		self.mainWindow.aot_frame.load_act()
		self.mainWindow.ra_frame.load_act()
		self.mainWindow.aoe_frame.load_act()
		self.mainWindow.status_bar['text'] = u'Изменения сохранены'
		self.exit()

	def exit(self,event=None):

		self.destroy()

	def save_changes(self):

		self.mainWindow.act_of_transfer.set(self.act_of_transfer.get())
		self.mainWindow.return_act.set(self.return_act.get())
		self.mainWindow.act_of_elimination.set(self.act_of_elimination.get())
		self.mainWindow.destination_folder.set(self.destination_folder.get())
		self.mainWindow.base_frame.sync_exist_acts()
		self.mainWindow.write_options()

	def change_act_of_transfer(self,event=None):

		filename = askopenfilename(filetypes=(("Doc files", "*.doc;*.docx"),('All files','*.*')))
		self.focus_force()
		if not filename:
			return
		self.act_of_transfer.set(filename)
		self.act_of_transfer_label_text['text'] = get_truncated_line(self.act_of_transfer.get(), self.pathesLength)

	def change_return_act(self,event=None):

		filename = askopenfilename(filetypes=(("Doc files", "*.doc;*.docx"),('All files','*.*')))
		self.focus_force()
		if not filename:
			return
		self.return_act.set(filename)
		self.return_act_label_text['text'] = get_truncated_line(self.return_act.get(), self.pathesLength)

	def change_act_of_elimination(self,event=None):

		filename = askopenfilename(filetypes=(("Doc files", "*.doc;*.docx"),('All files','*.*')))
		self.focus_force()
		if not filename:
			return
		self.act_of_elimination.set(filename)
		self.act_of_elimination_label_text['text'] = get_truncated_line(self.act_of_elimination.get(), self.pathesLength)

	def change_destination_folder(self,event=None):

		dirname = askdirectory(mustexist=True)
		self.focus_force()
		if not dirname:
			return
		self.destination_folder.set(dirname)
		self.destination_folder_label_text['text'] = get_truncated_line(self.destination_folder.get(), self.pathesLength)