import tkinter as tk
from tkinter import ttk
import scrape
import datetime
import threading
import os
import openpyxl

class FIIConsole:
    def __init__(self, root, manual=True):
        self.root = root
        self.manual = manual
        self.root.title("Script Frequency")
        if not manual:
            self.lst_updt_var = tk.StringVar(value="Last Updated: ")
            self.last_updated = tk.Label(self.root, textvariable=self.lst_updt_var)
            self.last_updated.pack()
        self.__columns = ('#1', '#2')
        self.tree = ttk.Treeview(self.root, columns=self.__columns, show='headings')
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
        for line in self.__sorted_data:
            self.tree.insert("", tk.END, value=line)
        


    def extract_data(self):
        print("extraction Started")
        data_map = scrape.get_data()
        
        sorted_data = sorted(data_map.items(), key=lambda x: x[1], reverse=True)
        return sorted_data
    
    def map_data(self):
        self.__extracted_data = self.extract_data()
        self.tree.delete(*self.tree.get_children())
        for line in self.__extracted_data:
            self.tree.insert("", tk.END, value=line)
        if not self.manual:
            now = datetime.datetime.now().strftime("%H:%M:%S")
            self.lst_updt_var.set("Last Updated: " + now)
    
    def refresh(self):
        threading.Thread(target=self.map_data, daemon=True).start()
        self.root.after(120000, self.refresh)



