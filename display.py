import tkinter as tk
from tkinter import ttk
import scrape
import datetime
import threading
import os
import openpyxl

class FIIConsole:
    def __init__(self, root, manual=True):

        # Tkinter Bug Work Around
    
        self.root = root
        self.manual = manual
        self.root.title("Script Frequency")
        self.root.iconbitmap('icon.ico')
        self.style = ttk.Style()
        self.style.configure("Treeview", font=('Britannic', 11, 'bold'), rowheight=25)
        self.style.configure("Treeview.Heading", font=('Britannic' ,13, 'bold'))

        if self.root.getvar('tk_patchLevel')=='8.6.9': #and OS_Name=='nt':
            def fixed_map(option):
                # Fix for setting text colour for Tkinter 8.6.9
                # From: https://core.tcl.tk/tk/info/509cafafae
                #
                # Returns the style map for 'option' with any styles starting with
                # ('!disabled', '!selected', ...) filtered out.
                #
                # style.map() returns an empty list for missing options, so this
                # should be future-safe.
                return [elm for elm in self.style.map('Treeview', query_opt=option) if elm[:2] != ('!disabled', '!selected')]
        self.style.map('Treeview', foreground=fixed_map('foreground'), background=fixed_map('background'))


        if not manual:
            self.lst_updt_var = tk.StringVar(value="Last Updated: ")
            self.last_updated = tk.Label(self.root, textvariable=self.lst_updt_var, font=('Britannic', 11))
            self.last_updated.pack()
        self.__columns = ('#1', '#2')
        self.tree = ttk.Treeview(self.root, columns=self.__columns, show='headings', selectmode='none')
        self.tree.tag_configure('top5', background='orange')
        self.tree.column('#1', anchor=tk.CENTER)
        self.tree.column('#2', anchor=tk.CENTER)
        self.tree.heading('#1', text='Company')
        self.tree.heading('#2', text='Count')
        self.tree.pack()
        if manual:
            self.open_button = ttk.Button(self.root, text='Open File', command=lambda : self.open_file())
            self.open_button.pack()
            self.load_data_button = ttk.Button(self.root, text='Load Data', command=self.load_data)
            self.load_data_button.pack()
    
    def open_file(self):
        os.system('start excel.exe data.xlsx')
    
    def load_data(self):
        self.__data_count = {}
        self.__wb = openpyxl.load_workbook('data.xlsx')
        self.__ws = self.__wb['main']
        self.__row_count = self.__ws.max_row
        for row in range(4, self.__row_count + 1):
            endpoint = self.__ws[f"A{row}"].value
            if not endpoint in self.__data_count:
                self.__data_count[endpoint] = 1
                continue
            self.__data_count[endpoint] += 1
        self.__sorted_data = sorted(self.__data_count.items(), key=lambda x: x[1], reverse=True)
        self.tree.delete(*self.tree.get_children())
        for i, line in enumerate(self.__sorted_data):
            self.tree.insert("", tk.END, iid=i, value=line)
        for i in range(5):
            self.tree.item(i, tags="top5")

        
            
        
        


    def extract_data(self):
        print("extraction Started")
        data_map = scrape.get_data()
        
        sorted_data = sorted(data_map.items(), key=lambda x: x[1], reverse=True)
        return sorted_data
    
    def map_data(self):
        self.__extracted_data = self.extract_data()
        self.tree.delete(*self.tree.get_children())
        for i, line in enumerate(self.__extracted_data):
            self.tree.insert("", tk.END, iid=i, value=line)
        for i in range(5):
            self.tree.item(i, tags="top5")
        if not self.manual:
            now = datetime.datetime.now().strftime("%H:%M:%S")
            self.lst_updt_var.set("Last Updated: " + now)
    
    def refresh(self):
        threading.Thread(target=self.map_data, daemon=True).start()
        self.root.after(120000, self.refresh)



