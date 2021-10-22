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
        self.delay = 120000
        self.override = False
        self.active_process = False
        self.root.title("Script Frequency")
        self.root.iconbitmap('icon.ico')
        self.style = ttk.Style()
        self.style.configure("Treeview", font=('Britannic', 11, 'bold'), rowheight=25)
        self.style.configure("Treeview.Heading", font=('Britannic' ,13, 'bold'))
        
        # Tkinter Bug Work Around
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
            self.lst_updt_var = tk.StringVar(value="Last Updated: Loading...")
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
        else:
            self.control_frame = tk.Frame(root)
            self.control_frame.pack(padx=5, pady=10)
            self.refresh_button = ttk.Button(self.control_frame, text='Refresh', command=self.manual_refresh)
            self.refresh_button.grid(column=0, row=0, padx=2)
            self.delay_var = tk.StringVar()
            self.delay_input = ttk.Entry(self.control_frame, textvariable=self.delay_var, width=8)
            self.delay_input.grid(column=1, row=0, padx=2)
            self.set_delay_btn = ttk.Button(self.control_frame, text="Set delay", command=self.set_delay)
            self.set_delay_btn.grid(column=2, row=0, padx=2)
    
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
        try:
            for i in range(5):
                self.tree.item(i, tags="top5")
        except tk.TclError:
            print("No Data")
        if not self.manual:
            now = datetime.datetime.now().strftime("%H:%M:%S")
            self.lst_updt_var.set("Last Updated: " + now)
    
    def refresh(self):
        if not self.override:
            self.active_process = True
            print("auto process started")
            self.t1 = threading.Thread(target=self.map_data, daemon=True)
            self.t1.start()
            threading.Thread(target=self.setStateActiveProcess, daemon=True).start()
        else:
            print("Process paused.")
        self.root.after(self.delay, self.refresh)
    
    def manual_refresh(self):
        if not self.active_process:
            self.override = True
            self.t2 = threading.Thread(target=self.map_data, daemon=True)
            self.t2.start()
            threading.Thread(target=self.setStateOverride, daemon=True).start()
        else:
            print("Override denied")
    
    def setStateActiveProcess(self):
        while self.t1.is_alive():
                pass
        self.active_process = False

    def setStateOverride(self):
        while self.t2.is_alive():
                pass
        self.override = False

    def set_delay(self):
        self.delay = int(self.delay_var.get()) * 60 * 1000
        print(f"set delay to {self.delay}")
